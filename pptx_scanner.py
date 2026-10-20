#!/usr/bin/env python3
"""
PowerPoint Accessibility Auditor
================================

Checks PPTX files for WCAG 2.2 AA issues and writes per-file HTML reports, a
rollup report with D3 charts, and a CSV.

How it works:
1. Two-pass scan: the first pass checks every slide without screenshots; only
   slides with issues are then screenshotted.
2. LibreOffice converts a deck to PDF once, and Poppler extracts the problem slides.
3. Results are cached by file hash, so unchanged decks are skipped on re-scan.
4. Screenshots not referenced by any report are removed.

Dependencies:
- python-pptx, Pillow, pdf2image (pip install)
- LibreOffice (soffice) and Poppler for screenshots
  (Mac: brew install poppler; Linux: apt-get install poppler-utils)
- Tesseract, optional, for --ocr

Usage:
python resource_scanner.py /path/to/presentations --mode strict
"""

import os
import sys
import argparse
import subprocess
import tempfile
import csv
import json
import hashlib
from pathlib import Path
from collections import defaultdict
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import pytesseract
from datetime import datetime
from tqdm import tqdm

# Global cache for contrast calculations
_contrast_cache = {}

# Issue severity levels (for per-slide impact)
ISSUE_SEVERITY = {
    # Severe issues (WCAG Level A violations) - high impact per slide
    'missing_alt_text': 'severe',
    'missing_language': 'severe',
    'insufficient_headings': 'severe',
    'scanned_image': 'severe',
    
    # Major issues (WCAG Level AA violations) - moderate impact per slide
    'insufficient_contrast': 'major',
    'table_no_header': 'major',
    'poor_alt_text': 'major',
    
    # Minor issues (best practices) - low impact per slide
    'missing_title': 'minor',
    'incorrect_language': 'minor',
    'poor_link_text': 'minor',
    'reading_order': 'minor',
    'image_of_text': 'minor',
    'auto_advance_transition': 'minor',
    'excessive_animations': 'minor',
    'auto_play_animation': 'minor',
}

# Per-slide deduction by severity (how much one issue hurts a single slide)
# Updated to closely match Ally/Panorama industry standards
SLIDE_DEDUCTIONS = {
    'severe': 70,   # One severe issue reduces slide score by 70% (critical accessibility barrier)
    'major': 45,    # One major issue reduces slide score by 45% (significant usability problem)
    'minor': 20,    # One minor issue reduces slide score by 20% (best practice violation)
}

# Issue categories for grouping
ISSUE_CATEGORIES = {
    'missing_alt_text': 'Images & Media',
    'missing_language': 'Language & Structure',
    'missing_title': 'Language & Structure',
    'insufficient_headings': 'Language & Structure',
    'scanned_image': 'Images & Media',
    'insufficient_contrast': 'Visual Design',
    'table_no_header': 'Tables & Data',
    'poor_alt_text': 'Images & Media',
    'incorrect_language': 'Language & Structure',
    'poor_link_text': 'Navigation',
    'reading_order': 'Language & Structure',
    'image_of_text': 'Images & Media',
    'auto_advance_transition': 'Animations & Timing',
    'excessive_animations': 'Animations & Timing',
    'auto_play_animation': 'Animations & Timing',
}

# How to fix instructions for each issue type
FIX_INSTRUCTIONS = {
    'missing_alt_text': {
        'title': 'How to Fix: Missing Alt Text',
        'steps': [
            '1. Right-click the image, chart, table, or diagram',
            '2. Select "Edit Alt Text" or "Format Picture"',
            '3. In the Alt Text pane, add a brief description of the image',
            '4. If the image is purely decorative, leave the alt text field empty (don\'t write "decorative")',
            '5. Good alt text is concise (1-2 sentences) and describes the content and function of the image'
        ],
        'wcag': 'WCAG 2.2 - 1.1.1 Non-text Content'
    },
    'missing_language': {
        'title': 'How to Fix: Missing Language',
        'steps': [
            '1. Go to File > Info',
            '2. Click "Properties" > "Advanced Properties"',
            '3. Go to the "Summary" tab',
            '4. Set the "Language" field to the appropriate language (e.g., "English (United States)")',
            '5. Click OK to save'
        ],
        'wcag': 'WCAG 2.2 - 3.1.1 Language of Page'
    },
    'incorrect_language': {
        'title': 'How to Fix: Incorrect Language',
        'steps': [
            '1. Go to File > Info',
            '2. Click "Properties" > "Advanced Properties"',
            '3. Go to the "Summary" tab',
            '4. Update the "Language" field to the correct language',
            '5. Click OK to save'
        ],
        'wcag': 'WCAG 2.2 - 3.1.1 Language of Page'
    },
    'insufficient_headings': {
        'title': 'How to Fix: Insufficient Headings',
        'steps': [
            '1. Ensure at least 1 out of every 7 slides has a title',
            '2. Use the built-in Title placeholder (not a text box)',
            '3. Go to Home > Layout and choose a layout with a title placeholder',
            '4. Enter a descriptive title that summarizes the slide content',
            '5. Titles help screen reader users navigate the presentation'
        ],
        'wcag': 'WCAG 2.2 - 1.3.1 Info and Relationships'
    },
    'missing_title': {
        'title': 'How to Fix: Missing Title',
        'steps': [
            '1. Click on the title placeholder at the top of the slide',
            '2. If no title placeholder exists, go to Home > Layout and choose a layout with a title',
            '3. Enter a descriptive title that summarizes the slide content',
            '4. Keep titles concise (5-10 words) but meaningful',
            '5. Screen readers use titles to help users navigate between slides',
            '6. Note: Some slides like dividers may not need titles, but most content slides should have them'
        ],
        'wcag': 'WCAG 2.2 - 1.3.1 Info and Relationships'
    },
    'insufficient_contrast': {
        'title': 'How to Fix: Insufficient Contrast',
        'steps': [
            '1. Select the text with low contrast',
            '2. Go to Home > Font Color or Shape Fill',
            '3. Choose colors with a contrast ratio of at least 4.5:1',
            '4. Use online contrast checkers (e.g., WebAIM Contrast Checker) to verify',
            '5. Dark text on light backgrounds or light text on dark backgrounds usually works best'
        ],
        'wcag': 'WCAG 2.2 - 1.4.3 Contrast (Minimum)'
    },
    'poor_alt_text': {
        'title': 'How to Fix: Poor Quality Alt Text',
        'steps': [
            '1. Right-click the image and select "Edit Alt Text"',
            '2. Replace generic text (e.g., "image", "picture") with a meaningful description',
            '3. Describe what\'s important about the image in context',
            '4. Keep it concise (1-2 sentences)',
            '5. Don\'t start with "Image of..." or "Picture of..." - just describe the content'
        ],
        'wcag': 'WCAG 2.2 - 1.1.1 Non-text Content'
    },
    'table_no_header': {
        'title': 'How to Fix: Table Missing Headers',
        'steps': [
            '1. Select the first row of the table',
            '2. Go to Table Design > Header Row (check the box)',
            '3. Or manually bold the text in the first row',
            '4. Make sure header text clearly describes each column',
            '5. Screen readers use header rows to help users navigate tables'
        ],
        'wcag': 'WCAG 2.2 - 1.3.1 Info and Relationships'
    },
    'poor_link_text': {
        'title': 'How to Fix: Poor Link Text',
        'steps': [
            '1. Select the hyperlink text',
            '2. Replace generic text like "click here" with descriptive text',
            '3. The link text should describe where the link goes or what it does',
            '4. Example: Instead of "Click here for the report", use "Download the Q3 Financial Report"',
            '5. Screen reader users often navigate by links, so descriptive text is essential'
        ],
        'wcag': 'WCAG 2.2 - 2.4.4 Link Purpose (In Context)'
    },
    'reading_order': {
        'title': 'How to Fix: Reading Order Issues',
        'steps': [
            '1. Go to Home > Arrange > Selection Pane',
            '2. The reading order is bottom-to-top in the Selection Pane',
            '3. Drag items to reorder them so they match the visual order',
            '4. Test with a screen reader (e.g., NVDA, JAWS) to verify',
            '5. Alternatively, use Home > Arrange > Send to Back/Bring to Front'
        ],
        'wcag': 'WCAG 2.2 - 1.3.2 Meaningful Sequence'
    },
    'image_of_text': {
        'title': 'How to Fix: Image of Text',
        'steps': [
            '1. Replace the image with actual text using PowerPoint text boxes',
            '2. If you must use an image, ensure it has descriptive alt text that includes the text content',
            '3. Use high-resolution images if text must be in an image',
            '4. Consider providing a text alternative in the slide notes',
            '5. Text in images cannot be resized, recolored, or read by screen readers'
        ],
        'wcag': 'WCAG 2.2 - 1.4.5 Images of Text'
    },
    'scanned_image': {
        'title': 'How to Fix: Scanned Image',
        'steps': [
            '1. If this is a scanned document, consider recreating it as actual PowerPoint content',
            '2. Extract text from the scanned image using OCR tools',
            '3. Add the extracted text to the slide as actual text (not an image)',
            '4. If scanning is necessary, add comprehensive alt text describing all content',
            '5. Scanned images are inaccessible to screen readers and cannot be searched or resized'
        ],
        'wcag': 'WCAG 2.2 - 1.1.1 Non-text Content'
    },
    'auto_advance_transition': {
        'title': 'How to Fix: Auto-Advance Transition',
        'steps': [
            '1. Click on the slide thumbnail',
            '2. Go to Transitions tab',
            '3. Under "Timing", uncheck "After" or set timing to 0',
            '4. Check "On Mouse Click" to require user interaction',
            '5. Auto-advancing slides prevent users from reading at their own pace'
        ],
        'wcag': 'WCAG 2.2 - 2.2.2 Pause, Stop, Hide'
    },
    'excessive_animations': {
        'title': 'How to Fix: Excessive Animations',
        'steps': [
            '1. Go to Animations tab',
            '2. Click Animation Pane to see all animations',
            '3. Select unnecessary animations and click "Remove"',
            '4. Keep only essential animations (3-5 maximum per slide)',
            '5. Too many animations can be distracting and cause motion sickness'
        ],
        'wcag': 'WCAG 2.2 - 2.3.3 Animation from Interactions'
    },
    'auto_play_animation': {
        'title': 'How to Fix: Auto-Play Animation',
        'steps': [
            '1. Go to Animations tab and open Animation Pane',
            '2. Select the auto-playing animation',
            '3. Click the dropdown arrow and choose "Effect Options"',
            '4. Under "Timing", change "Start" from "With Previous" or "After Previous" to "On Click"',
            '5. Users should control when animations play, not have them play automatically'
        ],
        'wcag': 'WCAG 2.2 - 2.2.2 Pause, Stop, Hide'
    },
}

def parse_args():
    parser = argparse.ArgumentParser(
        description='Audit PowerPoint files for accessibility issues. Only slides with issues are screenshotted.',
        epilog='Two-pass scanning, batch PDF conversion, hash-based caching and screenshot cleanup.')
    parser.add_argument('input_dir', help='Directory containing PPTX files (searches recursively)')
    parser.add_argument('--output-dir', default='public/reports', help='Output directory for reports')
    parser.add_argument('--mode', choices=['lenient', 'strict'], default='lenient',
                        help='Lenient: check alt text presence only. Strict: also check alt text quality.')
    parser.add_argument('--ocr', action='store_true', help='Enable OCR to detect images of text (requires tesseract)')
    parser.add_argument('--no-screenshots', action='store_true',
                        help='Skip ALL screenshots (fastest - good for quick scans during development)')
    parser.add_argument('--verbose', action='store_true',
                        help='Show detailed error messages for screenshot failures')
    parser.add_argument('--expected-language', default='en-US',
                        help='Expected language code for the presentation (default: en-US)')
    return parser.parse_args()

def get_luminance(rgb):
    def adjust(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * adjust(r) + 0.7152 * adjust(g) + 0.0722 * adjust(b)

def get_contrast_ratio(rgb1, rgb2):
    L1 = get_luminance(rgb1)
    L2 = get_luminance(rgb2)
    lighter = max(L1, L2)
    darker = min(L1, L2)
    return (lighter + 0.05) / (darker + 0.05)

def has_poor_quality_alt_text(alt_text):
    if not alt_text:
        return False
    alt_lower = alt_text.lower().strip()
    poor_indicators = [
        'image', 'picture', 'photo', 'img', 'screenshot', 'graphic',
        'untitled', 'placeholder', 'default'
    ]
    if len(alt_text) < 5:
        return True
    if any(alt_lower == indicator for indicator in poor_indicators):
        return True
    return False

def has_poor_link_text(link_text):
    if not link_text:
        return False
    link_lower = link_text.lower().strip()
    poor_phrases = [
        'click here', 'here', 'link', 'read more', 'more', 'this',
        'click', 'visit', 'see more', 'learn more'
    ]
    return link_lower in poor_phrases

def detect_text_in_image(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img).strip()
        return len(text) > 10
    except Exception as e:
        print(f"    OCR error on {image_path}: {e}")
        return False

def check_screenshot_dependencies():
    """Check if required tools for screenshot generation are available."""
    has_soffice = False
    has_pdf2image = False
    
    try:
        subprocess.run(['soffice', '--version'], 
                      stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=5)
        has_soffice = True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    
    try:
        from pdf2image import convert_from_path
        has_pdf2image = True
    except ImportError:
        pass
    
    return (has_soffice, has_pdf2image)

def export_slide_as_image(prs, slide_idx, output_path, verbose=False):
    """Export slide as image using LibreOffice + pdf2image. Returns (success, error_msg)."""
    try:
        from pdf2image import convert_from_path
    except ImportError:
        return (False, "pdf2image not installed")
    
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_pptx = os.path.join(tmpdir, 'temp.pptx')
            single_slide_prs = Presentation()
            single_slide_prs.slide_width = prs.slide_width
            single_slide_prs.slide_height = prs.slide_height
            slide_layout = single_slide_prs.slide_layouts[6]
            new_slide = single_slide_prs.slides.add_slide(slide_layout)
            
            # Copy shapes from original slide
            try:
                for shape in prs.slides[slide_idx].shapes:
                    el = shape.element
                    new_slide.shapes._spTree.insert_element_before(el, 'p:extLst')
            except Exception as e:
                if verbose:
                    print(f"\n      ⚠️  Slide {slide_idx+1}: Shape copy failed - {str(e)[:50]}")
                return (False, f"shape_copy: {str(e)[:50]}")
            
            single_slide_prs.save(tmp_pptx)
            tmp_pdf = os.path.join(tmpdir, 'temp.pdf')
            
            # Convert single-slide PPTX to PDF with LibreOffice
            result = subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, tmp_pptx],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30
            )
            
            if result.returncode != 0 or not os.path.exists(tmp_pdf):
                stderr = result.stderr.decode('utf-8', errors='ignore')[:100] if result.stderr else 'No error'
                if verbose:
                    print(f"\n      ⚠️  Slide {slide_idx+1}: LibreOffice failed - {stderr}")
                return (False, f"libreoffice: {stderr}")
            
            # Convert PDF to PNG with Poppler (pdf2image)
            images = convert_from_path(
                tmp_pdf,
                dpi=96,  # Optimized DPI
                fmt='png'
            )
            
            if images:
                images[0].save(output_path, 'PNG', optimize=True, quality=75)
                if os.path.exists(output_path):
                    return (True, None)
                else:
                    return (False, "save_failed")
            
            return (False, "no_images")
    except subprocess.TimeoutExpired:
        return (False, "timeout")
    except Exception as e:
        if verbose:
            print(f"\n      ⚠️  Slide {slide_idx+1}: Exception - {str(e)[:50]}")
        return (False, f"exception: {str(e)[:50]}")

def update_rollup_reports(output_dir, books, book_scores, book_slide_counts, book_issue_counts):
    """
    Generate/update rollup reports (CSV and HTML).
    Can be called multiple times as books are processed.
    
    Args:
        output_dir: Output directory for reports
        books: Dictionary of book titles to file lists
        book_scores: Dictionary of book scores
        book_slide_counts: Dictionary of total slides per book
        book_issue_counts: Dictionary of issue counts per book
    """
    # Generate CSV rollup
    csv_path = os.path.join(output_dir, 'accessibility_rollup.csv')
    all_issue_types = set()
    for counts in book_issue_counts.values():
        all_issue_types.update(counts.keys())
    
    # IMPORTANT: Always include ALL possible issue types in columns
    # Even if no instances are found, we want to show 0 to indicate the check was performed
    # This ensures consistent table structure across all scans
    for issue_type in ISSUE_SEVERITY.keys():
        all_issue_types.add(issue_type)
    
    # Sort issues by severity (severe, major, minor), then alphabetically
    # This groups related issues together (e.g., insufficient_headings and missing_title)
    def sort_by_severity(issue_type):
        severity = ISSUE_SEVERITY.get(issue_type, 'minor')
        severity_order = {'severe': 0, 'major': 1, 'minor': 2}
        return (severity_order.get(severity, 3), issue_type)
    
    all_issue_types = sorted(all_issue_types, key=sort_by_severity)
    
    # For HTML display, we'll combine insufficient_headings into missing_title column
    # Keep both in CSV for data completeness
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Book Title', 'Files', 'Total Slides', 'Score'] + all_issue_types
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for book_title in sorted(books.keys()):
            # Only write rows for books that have been processed (have scores)
            if book_title in book_scores:
                row = {
                    'Book Title': book_title, 
                    'Files': len(books[book_title]),
                    'Total Slides': book_slide_counts.get(book_title, 0),
                    'Score': book_scores.get(book_title, 0)
                }
                for issue_type in all_issue_types:
                    row[issue_type] = book_issue_counts[book_title].get(issue_type, 0)
                writer.writerow(row)
    
    # Generate rollup HTML
    rollup_html_path = os.path.join(output_dir, 'accessibility_rollup.html')
    generate_rollup_html(csv_path, rollup_html_path, output_dir, book_issue_counts)
    
    return csv_path, rollup_html_path

def cleanup_unreferenced_screenshots(screenshot_dir, referenced_slides):
    """
    Delete screenshots that are not referenced in reports.
    Keeps only slide images for slides with accessibility issues.
    
    This prevents orphaned screenshots from accumulating (e.g., from previous scans
    where different slides had issues).
    
    Args:
        screenshot_dir: Directory containing screenshots
        referenced_slides: Set of slide numbers that have issues (should be kept)
    
    Returns:
        Number of files deleted
    """
    if not os.path.exists(screenshot_dir):
        return 0
    
    deleted_count = 0
    saved_space = 0
    
    for root, dirs, files in os.walk(screenshot_dir):
        for file in files:
            if file.startswith('slide_') and file.endswith('.png'):
                try:
                    # Extract slide number from filename (e.g., slide_005.png -> 5)
                    slide_num = int(file.split('_')[1].split('.')[0])
                    
                    # Delete if not referenced in report
                    if slide_num not in referenced_slides:
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        os.remove(file_path)
                        deleted_count += 1
                        saved_space += file_size
                except (ValueError, IndexError):
                    # Invalid filename format, skip
                    pass
    
    if deleted_count > 0:
        mb_saved = saved_space / (1024 * 1024)
        print(f"    Cleaned up {deleted_count} unreferenced screenshots ({mb_saved:.1f} MB)")
    
    return deleted_count

def get_file_hash(filepath):
    """
    Get MD5 hash of file for caching purposes.
    Used to detect if a presentation has changed since last scan.
    """
    import hashlib
    hash_md5 = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            # Read in chunks to handle large files efficiently
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        print(f"Warning: Could not hash {filepath}: {e}")
        return None

def check_screenshot_cache(pptx_path, screenshot_dir):
    """
    Check if screenshots are already up-to-date for this presentation.
    Returns True if cached screenshots can be used, False if regeneration needed.
    
    Caching logic:
    - Calculates MD5 hash of presentation file
    - Compares with stored hash from previous scan
    - If hashes match and thumbnails exist: use cache
    - If different or missing: regenerate screenshots
    """
    cache_file = os.path.join(screenshot_dir, '.cache_info')
    current_hash = get_file_hash(pptx_path)
    
    if current_hash is None:
        return False  # Can't verify, regenerate
    
    # Check if cache file exists and hash matches
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r') as f:
                cached_data = json.loads(f.read())
                cached_hash = cached_data.get('hash', '')
                
            if cached_hash == current_hash:
                # Verify screenshots actually exist
                slide_files = list(Path(screenshot_dir).glob('slide_*.png'))
                if len(slide_files) > 0:
                    return True  # Cache is valid, use existing screenshots
        except:
            pass  # Cache file corrupted, regenerate
    
    # Save new hash for future scans
    try:
        os.makedirs(screenshot_dir, exist_ok=True)
        with open(cache_file, 'w') as f:
            json.dump({
                'hash': current_hash,
                'timestamp': datetime.now().isoformat(),
                'filename': os.path.basename(pptx_path)
            }, f)
    except Exception as e:
        print(f"Warning: Could not save cache info: {e}")
    
    return False  # Generate new screenshots

def generate_screenshots_batch(pptx_path, screenshot_dir, slide_numbers, verbose=False):
    """
    Generate screenshots using batch conversion (original strategy).
    Converts entire PPTX to PDF, then extracts all slides, keeping only problem slides.
    
    Settings:
    - 72 DPI (small file size, still readable)
    - Thumbnail resize to max 800px wide
    - Batch conversion (one PDF for entire presentation)
    - Keeps only problem slide images
    - Optimized compression
    
    Args:
        pptx_path: Path to presentation
        screenshot_dir: Directory to save screenshots
        slide_numbers: List of slide numbers to screenshot (1-indexed)
        verbose: If True, print detailed error messages
    
    Returns:
        dict: {slide_num: screenshot_path}
    """
    if not slide_numbers:
        return {}
    
    screenshots = {}
    os.makedirs(screenshot_dir, exist_ok=True)
    
    try:
        from pdf2image import convert_from_path
    except ImportError:
        print(f"    ⚠️  pdf2image not installed. Install with: pip install pdf2image")
        return screenshots
    
    print(f"    Converting {len(slide_numbers)} slides (batch method)...", end='', flush=True)
    
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            # Convert entire PPTX to PDF (LibreOffice)
            subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf',
                 '--outdir', tmpdir, pptx_path],
                check=True,
                capture_output=True,
                timeout=120
            )
            
            # Find the generated PDF
            pdf_files = list(Path(tmpdir).glob('*.pdf'))
            if not pdf_files:
                print(f" failed (no PDF generated)")
                return screenshots
            
            pdf_path = str(pdf_files[0])
            
            # Convert ALL pages from PDF to images at once (72 DPI for smaller files)
            images = convert_from_path(pdf_path, dpi=72)
            
            successful = 0
            failed = 0
            
            # Save only the problem slides
            for i, img in enumerate(images, start=1):
                if i in slide_numbers:
                    try:
                        # Resize to max 800px wide while maintaining aspect ratio
                        max_width = 800
                        if img.width > max_width:
                            ratio = max_width / img.width
                            new_height = int(img.height * ratio)
                            img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
                        
                        slide_png = os.path.join(screenshot_dir, f"slide_{i:03d}.png")
                        img.save(slide_png, 'PNG', optimize=True, quality=65)
                        screenshots[i] = slide_png
                        successful += 1
                    except Exception as e:
                        failed += 1
                        if verbose:
                            print(f"\n    ⚠️  Could not save slide {i}: {e}")
            
            print(f" done ({successful} created", end='')
            if failed > 0:
                print(f", {failed} failed)", end='')
            print(")")
            
        except subprocess.TimeoutExpired:
            print(f" failed (timeout)")
            if verbose:
                print(f"    ⚠️  LibreOffice conversion timed out after 120 seconds")
        except subprocess.CalledProcessError as e:
            print(f" failed (LibreOffice error)")
            if verbose:
                stderr = e.stderr.decode('utf-8', errors='ignore')[:200] if e.stderr else 'No error output'
                print(f"    ⚠️  LibreOffice error: {stderr}")
        except Exception as e:
            print(f" failed ({str(e)[:50]})")
            if verbose:
                print(f"    ⚠️  Exception: {e}")
    
    return screenshots

def audit_pptx_fast(pptx_path, mode='lenient', expected_language='en-US'):
    """
    Fast accessibility audit WITHOUT screenshot generation.
    Used for first pass to identify slides with issues.
    
    Returns:
        tuple: (issues_by_slide, presentation_issues, total_slides)
    """
    issues_by_slide = defaultdict(list)
    presentation_issues = []
    
    try:
        prs = Presentation(pptx_path)
    except Exception as e:
        presentation_issues.append(('error', f"Could not open presentation: {e}"))
        return issues_by_slide, presentation_issues, 0
    
    # Check presentation-level properties
    try:
        lang = prs.core_properties.language
        if not lang:
            presentation_issues.append(('missing_language', 'Presentation is missing language metadata'))
        elif lang != expected_language:
            presentation_issues.append(('incorrect_language', f'Language is set to "{lang}" but expected "{expected_language}"'))
    except:
        presentation_issues.append(('missing_language', 'Could not retrieve language metadata'))
    
    # Check heading ratio
    slides_with_titles = sum(1 for slide in prs.slides if slide.shapes.title and slide.shapes.title.text.strip())
    total_slides = len(prs.slides)
    if total_slides > 0 and slides_with_titles / total_slides < 1/7:
        presentation_issues.append(('insufficient_headings', f'Only {slides_with_titles} of {total_slides} slides have titles'))
    
    # Fast per-slide checks (no screenshots)
    for slide_idx, slide in enumerate(prs.slides, start=1):
        # Check for missing title on this slide
        if not slide.shapes.title or not slide.shapes.title.text.strip():
            issues_by_slide[slide_idx].append(('missing_title', 'Slide is missing a title'))
        
        for shape in slide.shapes:
            # Skip placeholder shapes (they belong to the master/layout, not this slide)
            try:
                if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                    continue
            except:
                pass
            
            # Check images for alt text and scanned images
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                alt_text = get_alt_text_from_shape(shape)
                if not alt_text or alt_text.strip() == '':
                    issues_by_slide[slide_idx].append(('missing_alt_text', f'Image "{shape.name}" has no alt text'))
                elif mode == 'strict' and has_poor_quality_alt_text(alt_text):
                    issues_by_slide[slide_idx].append(('poor_alt_text', f'Image "{shape.name}" has poor quality alt text'))
                
                # Check for scanned images (large images covering most of slide)
                try:
                    slide_width = prs.slide_width
                    slide_height = prs.slide_height
                    img_area = shape.width * shape.height
                    slide_area = slide_width * slide_height
                    
                    # If image covers >70% of slide and has generic/no alt text
                    if img_area > (slide_area * 0.7):
                        if not alt_text or alt_text.strip() == '' or len(alt_text.strip()) < 10:
                            issues_by_slide[slide_idx].append(('scanned_image', f'Large image "{shape.name}" may be scanned document'))
                except:
                    pass
                
                # Check for image-of-text (only if OCR is enabled or obvious from filename)
                # Note: Charts are NOT images-of-text, so we avoid heuristics like aspect ratio
                try:
                    # Only flag if filename strongly suggests it's a screenshot/text image
                    # Removed aspect ratio check - it was flagging vertical bar charts
                    if any(keyword in shape.name.lower() for keyword in ['screenshot', 'screencap', 'screen shot', 'screen_shot']):
                        issues_by_slide[slide_idx].append(('image_of_text', f'Image "{shape.name}" may contain text (filename suggests screenshot)'))
                except:
                    pass
            
            # Check tables
            if shape.has_table:
                table = shape.table
                if len(table.rows) > 0:
                    first_row = table.rows[0]
                    has_header = any(cell.text_frame and cell.text_frame.text.strip() for cell in first_row.cells)
                    if not has_header:
                        issues_by_slide[slide_idx].append(('table_no_header', 'Table is missing header row'))
            
            # Check text for links and contrast
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.hyperlink.address:
                            link_text = run.text.strip()
                            if has_poor_link_text(link_text):
                                issues_by_slide[slide_idx].append(('poor_link_text', f'Poor link text: "{link_text}"'))
                        
                        # Improved contrast check (handles theme colors and font size)
                        try:
                            # Try to get actual RGB color (handles both direct RGB and theme colors)
                            text_color = None
                            if run.font.color:
                                # Try to get RGB directly
                                try:
                                    if run.font.color.rgb:
                                        text_color = tuple(run.font.color.rgb)
                                except:
                                    # Theme color - try to get theme RGB
                                    try:
                                        if run.font.color.theme_color:
                                            # For theme colors, we'll flag for manual review
                                            # since we can't reliably get the actual rendered color
                                            pass
                                    except:
                                        pass
                            
                            # Get background color
                            bg_color = None
                            try:
                                # Try slide background first
                                if slide.background.fill.type == 1:  # SOLID
                                    try:
                                        bg_color = tuple(slide.background.fill.fore_color.rgb)
                                    except:
                                        pass
                            except:
                                pass
                            
                            # If no slide background, try shape fill
                            if not bg_color:
                                try:
                                    if shape.fill and shape.fill.type == 1:  # SOLID
                                        try:
                                            bg_color = tuple(shape.fill.fore_color.rgb)
                                        except:
                                            pass
                                except:
                                    pass
                            
                            # Default to white if can't determine background
                            if not bg_color:
                                bg_color = (255, 255, 255)
                            
                            # Check contrast if we have text color
                            if text_color:
                                ratio = get_contrast_ratio(text_color, bg_color)
                                
                                # Determine required contrast ratio based on font size (WCAG 2.2)
                                # Large text: ≥18pt or ≥14pt bold = 3:1 minimum
                                # Small text: <18pt or <14pt bold = 4.5:1 minimum
                                is_large_text = False
                                try:
                                    font_size = run.font.size
                                    is_bold = run.font.bold
                                    
                                    if font_size:
                                        # Convert EMU to points (914400 EMU = 1 inch = 72 points)
                                        size_pt = font_size / 12700  # EMU to points
                                        
                                        # Check if large text per WCAG
                                        if size_pt >= 18:
                                            is_large_text = True
                                        elif size_pt >= 14 and is_bold:
                                            is_large_text = True
                                except:
                                    pass
                                
                                # Apply appropriate threshold
                                required_ratio = 3.0 if is_large_text else 4.5
                                
                                if ratio < required_ratio:
                                    text_type = "large text" if is_large_text else "text"
                                    issues_by_slide[slide_idx].append(('insufficient_contrast', 
                                        f'{text_type.capitalize()} contrast {ratio:.2f}:1 (needs {required_ratio}:1)'))
                        except:
                            pass
        
        # Check for reading order issues (unusual z-order or positioning)
        # Count text boxes and their positions
        text_shapes = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                try:
                    # Only include shapes with valid positions (not None)
                    if shape.top is not None and shape.left is not None:
                        text_shapes.append({
                            'top': shape.top,
                            'left': shape.left,
                            'text': shape.text_frame.text.strip()
                        })
                except:
                    pass
        
        # If multiple text shapes, check if reading order seems unusual
        if len(text_shapes) >= 3:
            try:
                # Sort by top position (expected reading order: top to bottom)
                sorted_by_top = sorted(text_shapes, key=lambda x: x['top'])
                # Sort by left position (for RTL languages, this would be different)
                sorted_by_left = sorted(text_shapes, key=lambda x: x['left'])
                
                # Check if shapes are in neither top-to-bottom nor left-to-right order
                # This is a heuristic - if there's significant position variation, flag it
                top_positions = [s['top'] for s in text_shapes]
                left_positions = [s['left'] for s in text_shapes]
                
                # Calculate variance in positions
                if len(set(top_positions)) > 2 and len(set(left_positions)) > 2:
                    # Multiple rows and columns - potential reading order issue
                    # Flag if in strict mode
                    if mode == 'strict':
                        issues_by_slide[slide_idx].append(('reading_order', 'Slide has complex layout that may affect reading order'))
            except:
                pass
        
        # Check for animations and transitions
        try:
            # Check for auto-advance transition
            if slide.slide_layout.name:  # Has transition settings
                try:
                    # Check if slide has transition timing (auto-advance)
                    if hasattr(slide, '_element'):
                        timing = slide._element.xpath('.//p:transition')
                        if timing:
                            # Check for automatic advance timing
                            advance_time = slide._element.xpath('.//p:transition/@advTm')
                            if advance_time and int(advance_time[0]) > 0:
                                issues_by_slide[slide_idx].append(('auto_advance_transition', f'Slide auto-advances after {int(advance_time[0])/1000:.1f} seconds'))
                except:
                    pass
            
            # Check for animations on shapes
            animation_count = 0
            auto_play_found = False
            
            for shape in slide.shapes:
                try:
                    # Check if shape has animation effects
                    if hasattr(shape, '_element'):
                        # Look for animation elements
                        animations = shape._element.xpath('.//p:animEffect')
                        if animations:
                            animation_count += len(animations)
                            
                            # Check for auto-play animations
                            for anim in animations:
                                start_type = anim.xpath('.//p:cTn/@nodeType')
                                if start_type and 'afterEffect' in str(start_type):
                                    auto_play_found = True
                except:
                    pass
            
            # Flag excessive animations (>5 per slide)
            if animation_count > 5:
                issues_by_slide[slide_idx].append(('excessive_animations', f'Slide has {animation_count} animations (can be distracting)'))
            
            # Flag auto-play animations
            if auto_play_found:
                issues_by_slide[slide_idx].append(('auto_play_animation', 'Slide has auto-playing animations'))
        except:
            pass
    
    return issues_by_slide, presentation_issues, total_slides

def get_alt_text_from_shape(shape):
    """Helper to extract alt text from shape."""
    if not shape.has_text_frame:
        alt_text = shape.name
        try:
            if hasattr(shape, '_element') and hasattr(shape._element, 'xpath'):
                alt_text_elem = shape._element.xpath('.//p:cNvPr/@descr')
                if alt_text_elem:
                    alt_text = alt_text_elem[0]
        except:
            pass
        return alt_text
    return ""

def audit_pptx(pptx_path, screenshot_dir, mode='lenient', ocr_enabled=False, expected_language='en-US', no_screenshots=False, verbose=False):
    """
    Two-pass accessibility audit:
    Pass 1: Fast scan (no screenshots) - identify problem slides
    Pass 2: Batch screenshot ONLY problem slides
    """
    screenshots = {}
    
    # Check if we can use cached results
    use_cached_screenshots = False
    if not no_screenshots:
        use_cached_screenshots = check_screenshot_cache(pptx_path, screenshot_dir)
        if use_cached_screenshots:
            # Load existing screenshots
            for slide_file in Path(screenshot_dir).glob('slide_*.png'):
                try:
                    slide_num = int(slide_file.stem.split('_')[1])
                    screenshots[slide_num] = str(slide_file)
                except:
                    pass
            
            # Still need to run fast scan to get issues
            issues_by_slide, presentation_issues, total_slides = audit_pptx_fast(
                pptx_path, mode, expected_language
            )
            return issues_by_slide, screenshots, presentation_issues, total_slides
    
    # PASS 1: Scan without screenshots
    print(f"  Pass 1/2: Fast scan (no screenshots)...", end='', flush=True)
    issues_by_slide, presentation_issues, total_slides = audit_pptx_fast(
        pptx_path, mode, expected_language
    )
    
    # Identify slides with issues
    problem_slides = [slide_num for slide_num, issues in issues_by_slide.items() if issues]
    print(f" found {len(problem_slides)} slides with issues")
    
    # PASS 2: Screenshot ONLY problem slides (per-slide method)
    if not no_screenshots and problem_slides:
        print(f"  Pass 2/2: Generating screenshots for {len(problem_slides)} problem slides...")
        
        # Check if screenshot tools are available
        has_soffice, has_pdf2image = check_screenshot_dependencies()
        if has_soffice and has_pdf2image:
            # Use batch method (convert entire PPTX to PDF once, extract problem slides)
            screenshots = generate_screenshots_batch(
                pptx_path, screenshot_dir, problem_slides, verbose=verbose
            )
            
            # OCR check on generated screenshots (if enabled)
            if ocr_enabled and TESSERACT_AVAILABLE:
                for slide_num in screenshots:
                    screenshot_path = screenshots[slide_num]
                    try:
                        if detect_text_in_image(screenshot_path):
                            issues_by_slide[slide_num].append(
                                ('image_of_text', 'Slide may contain image-of-text')
                            )
                    except:
                        pass
        else:
            print(f" skipped (LibreOffice or pdf2image not available)")
    elif no_screenshots:
        print(f"  Pass 2/2: Skipped (--no-screenshots)")
    else:
        print(f"  Pass 2/2: No problem slides to screenshot!")
    
    return issues_by_slide, screenshots, presentation_issues, total_slides

def calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides):
    """
    Calculate accessibility score (0-100%) using Ally-style per-slide scoring.
    
    This approach calculates a score for each slide based on its issues,
    then averages across all slides. This makes the score proportional to
    the size of the presentation - a few issues in a 50-slide deck is much
    better than the same issues in a 5-slide deck.
    """
    if total_slides == 0:
        return 0.0, 0, 0, defaultdict(int)
    
    # Initialize slide scores (all start at 100%)
    slide_scores = {i: 100.0 for i in range(1, total_slides + 1)}
    
    # Count issues for reporting
    issue_counts = defaultdict(int)
    error_count = 0
    warning_count = 0
    
    # Apply slide-level issues (reduce individual slide scores)
    for slide_num, issues in issues_by_slide.items():
        for issue_type, description in issues:
            issue_counts[issue_type] += 1
            
            # Get severity and deduction for this issue type
            severity = ISSUE_SEVERITY.get(issue_type, 'minor')
            deduction = SLIDE_DEDUCTIONS.get(severity, 3)
            
            # Apply diminishing returns for multiple issues on same slide
            # First issue: full deduction
            # Second issue: 60% of deduction
            # Third+ issues: 40% of deduction each
            current_issues_on_slide = len([i for i in issues if i[0] == issue_type])
            if current_issues_on_slide == 1:
                multiplier = 1.0
            elif current_issues_on_slide == 2:
                multiplier = 0.6
            else:
                multiplier = 0.4
            
            slide_scores[slide_num] -= (deduction * multiplier)
            
            # Count errors vs warnings
            if severity in ['severe', 'major']:
                error_count += 1
            else:
                warning_count += 1
    
    # Apply presentation-level issues (percentage penalty to all slides)
    for issue_type, description in presentation_issues:
        issue_counts[issue_type] += 1
        
        severity = ISSUE_SEVERITY.get(issue_type, 'minor')
        
        # Presentation-level issues affect all slides but with smaller per-slide impact
        # Updated to match stricter slide deductions
        if severity == 'severe':
            penalty_per_slide = 6.0  # 6% per slide (proportional to 70% slide deduction)
        elif severity == 'major':
            penalty_per_slide = 3.0  # 3% per slide (proportional to 45% slide deduction)
        else:
            penalty_per_slide = 1.5  # 1.5% per slide (proportional to 20% slide deduction)
        
        for slide_num in slide_scores:
            slide_scores[slide_num] -= penalty_per_slide
        
        # Count errors vs warnings
        if severity in ['severe', 'major']:
            error_count += 1
        else:
            warning_count += 1
    
    # Clamp all slide scores to 0-100 range
    for slide_num in slide_scores:
        slide_scores[slide_num] = max(0, min(100, slide_scores[slide_num]))
    
    # Calculate average score across all slides
    average_score = sum(slide_scores.values()) / len(slide_scores)
    
    return round(average_score, 1), error_count, warning_count, issue_counts

def get_score_class(score):
    if score >= 90:
        return 'score-excellent'
    elif score >= 75:
        return 'score-good'
    elif score >= 60:
        return 'score-fair'
    else:
        return 'score-poor'

def generate_html_report(book_title, files_data, output_path):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Accessibility Report: {book_title}</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body {{ 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      padding: 2rem 0;
    }}
    .container {{ 
      max-width: 1400px;
      background: white;
      padding: 2rem;
      border-radius: 15px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.2);
    }}
    h1 {{
      color: #333;
      border-bottom: 3px solid #667eea;
      padding-bottom: 1rem;
      margin-bottom: 2rem;
    }}
    h2 {{
      color: #555;
      margin-top: 2rem;
      margin-bottom: 1rem;
      font-size: 1.5rem;
    }}
    .score-excellent {{ color: #28a745; font-weight: bold; }}
    .score-good {{ color: #5cb85c; font-weight: bold; }}
    .score-fair {{ color: #f0ad4e; font-weight: bold; }}
    .score-poor {{ color: #d9534f; font-weight: bold; }}
    .issue-severe {{ background-color: #f8d7da; border-left: 4px solid #d9534f; }}
    .issue-major {{ background-color: #fff3cd; border-left: 4px solid #f0ad4e; }}
    .issue-minor {{ background-color: #d1ecf1; border-left: 4px solid #17a2b8; }}
    .card {{
      margin-bottom: 1.5rem;
      border: none;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      transition: transform 0.2s, box-shadow 0.2s;
    }}
    .card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    }}
    .card-header {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      font-weight: 600;
    }}
    .screenshot {{
      max-width: 100%;
      height: auto;
      border: 2px solid #ddd;
      border-radius: 8px;
      margin: 1rem 0;
    }}
    .fix-instructions {{
      background-color: #e7f3ff;
      border-left: 4px solid #0066cc;
      padding: 1rem;
      margin: 1rem 0;
      border-radius: 4px;
    }}
    .wcag-reference {{
      font-size: 0.9rem;
      color: #666;
      font-style: italic;
    }}
    .back-link {{
      display: inline-block;
      margin-bottom: 1.5rem;
      padding: 0.5rem 1.5rem;
      background: #667eea;
      color: white;
      text-decoration: none;
      border-radius: 25px;
      transition: background 0.3s;
    }}
    .back-link:hover {{
      background: #764ba2;
      color: white;
    }}
    .summary-stats {{
      display: flex;
      gap: 1rem;
      margin: 1.5rem 0;
      flex-wrap: wrap;
    }}
    .stat-card {{
      flex: 1;
      min-width: 200px;
      padding: 1.5rem;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border-radius: 10px;
      text-align: center;
    }}
    .stat-card h3 {{
      font-size: 2.5rem;
      margin: 0;
      font-weight: bold;
    }}
    .stat-card p {{
      margin: 0.5rem 0 0 0;
      opacity: 0.9;
    }}
  </style>
</head>
<body>
  <div class="container">
    <a href="accessibility_rollup.html" class="back-link">← Back to Roll-up Report</a>
    <h1>📊 Accessibility Report: {book_title}</h1>
"""
    
    # Calculate overall statistics (weighted by slide count)
    weighted_score = 0
    total_slides_in_book = 0
    total_errors = 0
    total_warnings = 0
    file_count = len(files_data)
    
    for pptx_path, issues_by_slide, screenshots, presentation_issues, total_slides in files_data:
        score, errors, warnings, _ = calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides)
        weighted_score += (score * total_slides)
        total_slides_in_book += total_slides
        total_errors += errors
        total_warnings += warnings
    
    avg_score = weighted_score / total_slides_in_book if total_slides_in_book > 0 else 0
    
    # Summary statistics
    html += f"""
    <div class="summary-stats">
      <div class="stat-card">
        <h3>{file_count}</h3>
        <p>Total Files</p>
      </div>
      <div class="stat-card">
        <h3 class="{get_score_class(avg_score)}">{avg_score:.1f}%</h3>
        <p>Average Score</p>
      </div>
      <div class="stat-card">
        <h3>{total_errors}</h3>
        <p>Total Errors</p>
      </div>
      <div class="stat-card">
        <h3>{total_warnings}</h3>
        <p>Total Warnings</p>
      </div>
    </div>
"""
    
    # Separate files by type
    pptx_files = []
    docx_files = []
    
    for file_data in files_data:
        pptx_path = file_data[0]
        file_ext = os.path.splitext(pptx_path)[1].lower()
        if file_ext == '.docx':
            docx_files.append(file_data)
        else:  # .pptx or other
            pptx_files.append(file_data)
    
    # Create accordion structure for file types
    html += """
    <h2>📁 Documents by Type</h2>
    <div class="accordion" id="documentTypesAccordion">
"""
    
    # PowerPoint Presentations Accordion
    if pptx_files:
        # Calculate stats for PowerPoint files
        pptx_score = 0
        pptx_slides = 0
        pptx_errors = 0
        pptx_warnings = 0
        pptx_issue_counts = {}
        
        for pptx_path, issues_by_slide, screenshots, presentation_issues, total_slides in pptx_files:
            score, errors, warnings, _ = calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides)
            pptx_score += (score * total_slides)
            pptx_slides += total_slides
            pptx_errors += errors
            pptx_warnings += warnings
            
            # Count issues
            for issue_type, _ in presentation_issues:
                pptx_issue_counts[issue_type] = pptx_issue_counts.get(issue_type, 0) + 1
            for slide_num, issues in issues_by_slide.items():
                for issue_type, _ in issues:
                    pptx_issue_counts[issue_type] = pptx_issue_counts.get(issue_type, 0) + 1
        
        pptx_avg_score = pptx_score / pptx_slides if pptx_slides > 0 else 0
        pptx_score_class = get_score_class(pptx_avg_score)
        
        html += f"""
      <div class="accordion-item">
        <h2 class="accordion-header" id="headingPowerPoint">
          <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapsePowerPoint" aria-expanded="false" aria-controls="collapsePowerPoint">
            📊 PowerPoint Presentations ({len(pptx_files)} files) - Score: <span class="{pptx_score_class}">{pptx_avg_score:.1f}%</span>
          </button>
        </h2>
        <div id="collapsePowerPoint" class="accordion-collapse collapse" aria-labelledby="headingPowerPoint" data-bs-parent="#documentTypesAccordion">
          <div class="accordion-body">
"""
        
        # PowerPoint Issue Summary
        if pptx_issue_counts:
            sorted_pptx_issues = sorted(pptx_issue_counts.items(), key=lambda x: (-x[1], x[0]))
            
            html += """
            <h3>📋 PowerPoint Issue Summary</h3>
            <div class="table-responsive">
              <table class="table table-striped table-hover table-sm">
                <thead class="table-dark">
                  <tr>
                    <th>Issue Type</th>
                    <th>Severity</th>
                    <th>Count</th>
                    <th>WCAG Guideline</th>
                  </tr>
                </thead>
                <tbody>
"""
            
            for issue_type, count in sorted_pptx_issues:
                severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                severity_badge_class = {
                    'severe': 'danger',
                    'major': 'warning',
                    'minor': 'info'
                }.get(severity, 'secondary')
                
                wcag_ref = ""
                if issue_type in FIX_INSTRUCTIONS:
                    wcag_ref = FIX_INSTRUCTIONS[issue_type]['wcag']
                
                html += f"""
                  <tr>
                    <td><strong>{issue_type.replace('_', ' ').title()}</strong></td>
                    <td><span class="badge bg-{severity_badge_class}">{severity.upper()}</span></td>
                    <td><strong>{count}</strong></td>
                    <td><small class="text-muted">{wcag_ref}</small></td>
                  </tr>
"""
            
            html += """
                </tbody>
              </table>
            </div>
            <hr class="my-3">
"""
        
        # Individual PowerPoint files
        for pptx_path, issues_by_slide, screenshots, presentation_issues, total_slides in pptx_files:
            file_name = os.path.splitext(os.path.basename(pptx_path))[0]
            score, error_count, warning_count, issue_counts = calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides)
            score_class = get_score_class(score)
            
            html += f"""
            <div class="card">
              <div class="card-header">
                <h3 class="mb-0">📄 {file_name}</h3>
              </div>
              <div class="card-body">
                <p><strong>Accessibility Score:</strong> <span class="{score_class}">{score}%</span></p>
                <p><strong>Errors:</strong> {error_count} | <strong>Warnings:</strong> {warning_count}</p>
"""
            
            # Presentation-level issues
            if presentation_issues:
                html += "<h4>Presentation-Level Issues</h4><ul>"
                for issue_type, description in presentation_issues:
                    severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                    severity_class = f'issue-{severity}'
                    html += f'<li class="{severity_class}" style="padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;"><strong>{issue_type.replace("_", " ").title()}:</strong> {description}</li>'
                    
                    # Add fix instructions
                    if issue_type in FIX_INSTRUCTIONS:
                        fix_info = FIX_INSTRUCTIONS[issue_type]
                        html += f"""
                        <div class="fix-instructions">
                            <h5>{fix_info['title']}</h5>
                            <ol>
"""
                        for step in fix_info['steps']:
                            html += f"                            <li>{step}</li>\n"
                        html += f"""
                            </ol>
                            <p class="wcag-reference">{fix_info['wcag']}</p>
                        </div>
"""
                html += "</ul>"
            
            # Slide-level issues
            if issues_by_slide:
                for slide_num in sorted(issues_by_slide.keys()):
                    issues = issues_by_slide[slide_num]
                    html += f"<h4>Slide {slide_num}</h4>"
                    
                    # Screenshot
                    if slide_num in screenshots:
                        rel_screenshot = os.path.relpath(screenshots[slide_num], os.path.dirname(output_path))
                        html += f'<img src="{rel_screenshot}" alt="Screenshot of slide {slide_num}" class="screenshot">'
                    
                    html += "<ul>"
                    seen_issue_types = set()
                    for issue_type, description in issues:
                        severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                        severity_class = f'issue-{severity}'
                        html += f'<li class="{severity_class}" style="padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;"><strong>{issue_type.replace("_", " ").title()}:</strong> {description}</li>'
                        
                        # Add fix instructions (only once per issue type)
                        if issue_type in FIX_INSTRUCTIONS and issue_type not in seen_issue_types:
                            seen_issue_types.add(issue_type)
                            fix_info = FIX_INSTRUCTIONS[issue_type]
                            html += f"""
                            <div class="fix-instructions">
                                <h5>{fix_info['title']}</h5>
                                <ol>
"""
                            for step in fix_info['steps']:
                                html += f"                                <li>{step}</li>\n"
                            html += f"""
                                </ol>
                                <p class="wcag-reference">{fix_info['wcag']}</p>
                            </div>
"""
                    html += "</ul>"
            else:
                html += "<p>✅ No slide-level issues found!</p>"
            
            html += """
              </div>
            </div>
"""
        
        html += """
          </div>
        </div>
      </div>
"""
    
    # Word Documents Accordion
    if docx_files:
        # Calculate stats for Word files
        docx_score = 0
        docx_pages = 0
        docx_errors = 0
        docx_warnings = 0
        docx_issue_counts = {}
        
        for docx_path, issues_by_slide, screenshots, presentation_issues, total_slides in docx_files:
            score, errors, warnings, _ = calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides)
            docx_score += (score * total_slides)
            docx_pages += total_slides
            docx_errors += errors
            docx_warnings += warnings
            
            # Count issues
            for issue_type, _ in presentation_issues:
                docx_issue_counts[issue_type] = docx_issue_counts.get(issue_type, 0) + 1
            for page_num, issues in issues_by_slide.items():
                for issue_type, _ in issues:
                    docx_issue_counts[issue_type] = docx_issue_counts.get(issue_type, 0) + 1
        
        docx_avg_score = docx_score / docx_pages if docx_pages > 0 else 0
        docx_score_class = get_score_class(docx_avg_score)
        
        html += f"""
      <div class="accordion-item">
        <h2 class="accordion-header" id="headingWord">
          <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseWord" aria-expanded="false" aria-controls="collapseWord">
            📝 Word Documents ({len(docx_files)} files) - Score: <span class="{docx_score_class}">{docx_avg_score:.1f}%</span>
          </button>
        </h2>
        <div id="collapseWord" class="accordion-collapse collapse" aria-labelledby="headingWord" data-bs-parent="#documentTypesAccordion">
          <div class="accordion-body">
"""
        
        # Word Issue Summary
        if docx_issue_counts:
            sorted_docx_issues = sorted(docx_issue_counts.items(), key=lambda x: (-x[1], x[0]))
            
            html += """
            <h3>📋 Word Document Issue Summary</h3>
            <div class="table-responsive">
              <table class="table table-striped table-hover table-sm">
                <thead class="table-dark">
                  <tr>
                    <th>Issue Type</th>
                    <th>Severity</th>
                    <th>Count</th>
                    <th>WCAG Guideline</th>
                  </tr>
                </thead>
                <tbody>
"""
            
            for issue_type, count in sorted_docx_issues:
                severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                severity_badge_class = {
                    'severe': 'danger',
                    'major': 'warning',
                    'minor': 'info'
                }.get(severity, 'secondary')
                
                wcag_ref = ""
                if issue_type in FIX_INSTRUCTIONS:
                    wcag_ref = FIX_INSTRUCTIONS[issue_type]['wcag']
                
                html += f"""
                  <tr>
                    <td><strong>{issue_type.replace('_', ' ').title()}</strong></td>
                    <td><span class="badge bg-{severity_badge_class}">{severity.upper()}</span></td>
                    <td><strong>{count}</strong></td>
                    <td><small class="text-muted">{wcag_ref}</small></td>
                  </tr>
"""
            
            html += """
                </tbody>
              </table>
            </div>
            <hr class="my-3">
"""
        
        # Individual Word files
        for docx_path, issues_by_slide, screenshots, presentation_issues, total_slides in docx_files:
            file_name = os.path.splitext(os.path.basename(docx_path))[0]
            score, error_count, warning_count, issue_counts = calculate_accessibility_score(issues_by_slide, presentation_issues, total_slides)
            score_class = get_score_class(score)
            
            html += f"""
            <div class="card">
              <div class="card-header">
                <h3 class="mb-0">📄 {file_name}</h3>
              </div>
              <div class="card-body">
                <p><strong>Accessibility Score:</strong> <span class="{score_class}">{score}%</span></p>
                <p><strong>Errors:</strong> {error_count} | <strong>Warnings:</strong> {warning_count}</p>
"""
            
            # Document-level issues
            if presentation_issues:
                html += "<h4>Document-Level Issues</h4><ul>"
                for issue_type, description in presentation_issues:
                    severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                    severity_class = f'issue-{severity}'
                    html += f'<li class="{severity_class}" style="padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;"><strong>{issue_type.replace("_", " ").title()}:</strong> {description}</li>'
                    
                    # Add fix instructions
                    if issue_type in FIX_INSTRUCTIONS:
                        fix_info = FIX_INSTRUCTIONS[issue_type]
                        html += f"""
                        <div class="fix-instructions">
                            <h5>{fix_info['title']}</h5>
                            <ol>
"""
                        for step in fix_info['steps']:
                            html += f"                            <li>{step}</li>\n"
                        html += f"""
                            </ol>
                            <p class="wcag-reference">{fix_info['wcag']}</p>
                        </div>
"""
                html += "</ul>"
            
            # Page/section-level issues
            if issues_by_slide:
                for page_num in sorted(issues_by_slide.keys()):
                    issues = issues_by_slide[page_num]
                    html += f"<h4>Page/Section {page_num}</h4>"
                    
                    html += "<ul>"
                    seen_issue_types = set()
                    for issue_type, description in issues:
                        severity = ISSUE_SEVERITY.get(issue_type, 'minor')
                        severity_class = f'issue-{severity}'
                        html += f'<li class="{severity_class}" style="padding: 0.5rem; margin: 0.5rem 0; border-radius: 4px;"><strong>{issue_type.replace("_", " ").title()}:</strong> {description}</li>'
                        
                        # Add fix instructions (only once per issue type)
                        if issue_type in FIX_INSTRUCTIONS and issue_type not in seen_issue_types:
                            seen_issue_types.add(issue_type)
                            fix_info = FIX_INSTRUCTIONS[issue_type]
                            html += f"""
                            <div class="fix-instructions">
                                <h5>{fix_info['title']}</h5>
                                <ol>
"""
                            for step in fix_info['steps']:
                                html += f"                                <li>{step}</li>\n"
                            html += f"""
                                </ol>
                                <p class="wcag-reference">{fix_info['wcag']}</p>
                            </div>
"""
                    html += "</ul>"
            else:
                html += "<p>✅ No page-level issues found!</p>"
            
            html += """
              </div>
            </div>
"""
        
        html += """
          </div>
        </div>
      </div>
"""
    
    html += """
    </div>
    
  </div>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

def generate_rollup_html(csv_path, output_path, output_dir, book_issue_counts):
    import pandas as pd
    df = pd.read_csv(csv_path)
    
    # Calculate aggregate statistics
    total_files = df['Files'].sum()
    total_slides = df['Total Slides'].sum()
    
    # Calculate weighted average score (weighted by number of slides in each book)
    df['weighted_score'] = df['Score'] * df['Total Slides']
    avg_score = df['weighted_score'].sum() / total_slides if total_slides > 0 else 0
    
    # Exclude insufficient_headings from totals (it's shown as flag in missing_title column)
    total_issues = df.drop(['Book Title', 'Files', 'Total Slides', 'Score', 'weighted_score', 'insufficient_headings'], axis=1, errors='ignore').sum().sum()
    
    # Aggregate issues by category
    category_totals = defaultdict(int)
    issue_type_totals = defaultdict(int)
    
    for col in df.columns:
        if col not in ['Book Title', 'Files', 'Total Slides', 'Score', 'weighted_score', 'insufficient_headings']:
            issue_total = df[col].sum()
            issue_type_totals[col] = issue_total
            category = ISSUE_CATEGORIES.get(col, 'Other')
            category_totals[category] += issue_total
    
    # Prepare data for D3 visualizations
    category_data = [{'category': k, 'count': int(v)} for k, v in category_totals.items()]
    issue_type_data = [{'type': k.replace('_', ' ').title(), 'count': int(v)} 
                       for k, v in sorted(issue_type_totals.items(), key=lambda x: x[1], reverse=True)[:10]]
    
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Accessibility Roll-up Report</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://d3js.org/d3.v7.min.js"></script>
  <style>
    body {{ 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      padding: 2rem 0;
    }}
    .container {{ 
      max-width: 1600px;
      background: white;
      padding: 3rem;
      border-radius: 15px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.2);
    }}
    h1 {{
      color: #333;
      border-bottom: 4px solid #667eea;
      padding-bottom: 1rem;
      margin-bottom: 2rem;
      font-size: 2.5rem;
    }}
    h2 {{
      color: #555;
      margin-top: 3rem;
      margin-bottom: 1.5rem;
      font-size: 1.8rem;
      border-left: 4px solid #667eea;
      padding-left: 1rem;
    }}
    .summary-cards {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 1.5rem;
      margin: 2rem 0;
    }}
    .summary-card {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 2rem;
      border-radius: 15px;
      text-align: center;
      box-shadow: 0 4px 15px rgba(0,0,0,0.1);
      transition: transform 0.3s, box-shadow 0.3s;
    }}
    .summary-card:hover {{
      transform: translateY(-5px);
      box-shadow: 0 8px 25px rgba(0,0,0,0.2);
    }}
    .summary-card h3 {{
      font-size: 3rem;
      margin: 0;
      font-weight: bold;
    }}
    .summary-card p {{
      margin: 1rem 0 0 0;
      font-size: 1.1rem;
      opacity: 0.95;
    }}
    .score-excellent {{ color: #28a745; }}
    .score-good {{ color: #5cb85c; }}
    .score-fair {{ color: #f0ad4e; }}
    .score-poor {{ color: #d9534f; }}
    .score-badge {{
      padding: 0.5rem 1rem;
      border-radius: 20px;
      font-weight: bold;
      display: inline-block;
    }}
    .score-badge.score-excellent {{ background-color: #d4edda; color: #155724; }}
    .score-badge.score-good {{ background-color: #d1e7dd; color: #0a3622; }}
    .score-badge.score-fair {{ background-color: #fff3cd; color: #856404; }}
    .score-badge.score-poor {{ background-color: #f8d7da; color: #721c24; }}
    .text-green {{ color: #28a745; font-weight: bold; }}
    .text-yellow {{ color: #ffc107; font-weight: bold; }}
    .text-red {{ color: #dc3545; font-weight: bold; }}
    .chart-container {{
      background: white;
      padding: 2rem;
      border-radius: 10px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      margin: 2rem 0;
    }}
    
    .chart-legend {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }}
    .chart-row {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 2rem;
      margin: 2rem 0;
    }}
    @media (max-width: 992px) {{
      .chart-row {{
        grid-template-columns: 1fr;
      }}
    }}
    table {{
      background: white;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      border-radius: 10px;
      overflow: hidden;
    }}
    thead {{
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
    }}
    thead th {{
      padding: 1rem !important;
      font-weight: 600;
    }}
    tbody tr {{
      transition: background 0.2s;
    }}
    tbody tr:hover {{
      background: #f8f9fa;
    }}
    tbody td {{
      padding: 1rem !important;
    }}
    a {{
      color: #667eea;
      text-decoration: none;
      font-weight: 600;
      transition: color 0.2s;
    }}
    a:hover {{
      color: #764ba2;
      text-decoration: underline;
    }}
    .axis text {{
      font-size: 12px;
    }}
    .bar {{
      transition: opacity 0.2s;
    }}
    .bar:hover {{
      opacity: 0.7;
    }}
    .tooltip {{
      position: absolute;
      text-align: center;
      padding: 8px;
      font-size: 12px;
      background: rgba(0, 0, 0, 0.8);
      color: white;
      border-radius: 4px;
      pointer-events: none;
      opacity: 0;
      transition: opacity 0.2s;
    }}
  </style>
</head>
<body>
  <div class="container">
    <h1>📊 Accessibility Roll-up Report</h1>
    <p class="lead text-muted">Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
    
    <div class="summary-cards">
      <div class="summary-card">
        <h3>{total_files}</h3>
        <p>Total Files Audited</p>
      </div>
      <div class="summary-card">
        <h3>{avg_score:.1f}%</h3>
        <p>Average Score</p>
      </div>
      <div class="summary-card">
        <h3>{total_issues:,}</h3>
        <p>Total Issues Found</p>
      </div>
      <div class="summary-card">
        <h3>{len(df)}</h3>
        <p>Books/Collections</p>
      </div>
    </div>
    
    <h2>📈 Issue Distribution</h2>
    <div class="chart-row">
      <div class="chart-container">
        <h4 class="text-center mb-4">Issues by Category</h4>
        <div id="categoryChart"></div>
      </div>
      <div class="chart-container">
        <h4 class="text-center mb-4">Top 10 Issue Types</h4>
        <div id="issueTypeChart"></div>
      </div>
    </div>
    
    <h2>📚 Issues by Book/Collection</h2>
    <div class="table-responsive">
      <table class="table table-bordered table-hover">
        <thead>
          <tr>
"""
    
    # Filter out weighted_score and insufficient_headings from display
    # insufficient_headings will only be used internally to determine missing_title severity
    display_columns = [col for col in df.columns if col not in ['weighted_score', 'insufficient_headings']]
    
    # Table headers (no severity indicators)
    for col in display_columns:
        col_display = col.replace('_', ' ').title()
        html += f"            <th>{col_display}</th>\n"
    
    html += """          </tr>
        </thead>
        <tbody>
"""
    
    # Table rows
    for _, row in df.iterrows():
        html += "          <tr>\n"
        book_title = row['Book Title']
        
        for col in display_columns:
            val = row[col]
            if col == 'Book Title':
                # Make book title a link to the detailed report
                book_slug = str(val).replace(' ', '_')
                report_link = f"{book_slug}_report.html"
                html += f'            <td><a href="{report_link}">{val}</a></td>\n'
            elif col == 'Score':
                try:
                    score_val = float(val)
                    score_class = get_score_class(score_val)
                    html += f'            <td><span class="score-badge {score_class}">{score_val:.1f}%</span></td>\n'
                except:
                    html += f"            <td>{val}</td>\n"
            elif col in ['Files', 'Total Slides']:
                html += f"            <td>{val}</td>\n"
            else:
                # Issue columns - color code by count only
                try:
                    ival = int(val)
                    
                    # Special handling for missing_title: show severe warning if book fails insufficient_headings threshold
                    if col == 'missing_title':
                        has_insufficient_headings = book_issue_counts.get(book_title, {}).get('insufficient_headings', 0) > 0
                        
                        if ival == 0:
                            color = "text-green"
                            severity_badge = ""
                        else:
                            # Color based on count
                            if ival < 10:
                                color = "text-yellow"
                            else:
                                color = "text-red"
                            
                            # Add severe warning only if also fails insufficient_headings threshold
                            if has_insufficient_headings:
                                severity_badge = ' <span class="badge badge-danger" style="font-size: 0.7em;">🚨 SEVERE: < 14% slides have titles</span>'
                            else:
                                severity_badge = ""
                    else:
                        # All other issue columns: no badges, just color coding
                        if ival == 0:
                            color = "text-green"
                        elif ival < 10:
                            color = "text-yellow"
                        else:
                            color = "text-red"
                        severity_badge = ""
                    
                    html += f'            <td class="{color}">{ival}{severity_badge}</td>\n'
                except:
                    html += f"            <td>{val}</td>\n"
        html += "          </tr>\n"
    
    html += """        </tbody>
      </table>
    </div>
  </div>
  
  <script>
    // Data for visualizations
    const categoryData = """ + json.dumps(category_data) + """;
    const issueTypeData = """ + json.dumps(issue_type_data) + """;
    
    // Create tooltip
    const tooltip = d3.select("body")
      .append("div")
      .attr("class", "tooltip");
    
    // Category Pie Chart
    function createPieChart(data, containerId) {
      const width = 400;
      const height = 400;
      const radius = Math.min(width, height) / 2;
      
      const color = d3.scaleOrdinal()
        .domain(data.map(d => d.category))
        .range(['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b', '#fa709a']);
      
      const svg = d3.select(`#${containerId}`)
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${width / 2},${height / 2})`);
      
      const pie = d3.pie()
        .value(d => d.count)
        .sort(null);
      
      const arc = d3.arc()
        .innerRadius(radius * 0.5)
        .outerRadius(radius * 0.8);
      
      const arcs = svg.selectAll("arc")
        .data(pie(data))
        .enter()
        .append("g");
      
      arcs.append("path")
        .attr("d", arc)
        .attr("fill", d => color(d.data.category))
        .attr("stroke", "white")
        .attr("stroke-width", 2)
        .on("mouseover", function(event, d) {
          d3.select(this).style("opacity", 0.7);
          tooltip
            .style("opacity", 1)
            .html(`<strong>${d.data.category}</strong><br/>Count: ${d.data.count}`)
            .style("left", (event.pageX + 10) + "px")
            .style("top", (event.pageY - 10) + "px");
        })
        .on("mouseout", function() {
          d3.select(this).style("opacity", 1);
          tooltip.style("opacity", 0);
        });
      
      // Add legend instead of labels on slices
      const legend = d3.select(`#${containerId}`)
        .append("div")
        .attr("class", "chart-legend")
        .style("margin-top", "20px")
        .style("display", "flex")
        .style("flex-wrap", "wrap")
        .style("justify-content", "center")
        .style("gap", "15px");
      
      data.forEach(d => {
        const legendItem = legend.append("div")
          .style("display", "flex")
          .style("align-items", "center")
          .style("gap", "8px");
        
        legendItem.append("div")
          .style("width", "20px")
          .style("height", "20px")
          .style("background-color", color(d.category))
          .style("border-radius", "3px")
          .style("flex-shrink", "0");
        
        legendItem.append("span")
          .style("font-size", "14px")
          .text(`${d.category} (${d.count})`);
      });
    }
    
    // Bar Chart for Issue Types
    function createBarChart(data, containerId) {
      const margin = {top: 20, right: 30, bottom: 100, left: 60};
      const width = 500 - margin.left - margin.right;
      const height = 400 - margin.top - margin.bottom;
      
      const svg = d3.select(`#${containerId}`)
        .append("svg")
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);
      
      const x = d3.scaleBand()
        .domain(data.map(d => d.type))
        .range([0, width])
        .padding(0.2);
      
      const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.count) * 1.1])
        .range([height, 0]);
      
      const color = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.count)])
        .range(['#667eea', '#d9534f']);
      
      // Add bars
      svg.selectAll("rect")
        .data(data)
        .enter()
        .append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d.type))
        .attr("y", d => y(d.count))
        .attr("width", x.bandwidth())
        .attr("height", d => height - y(d.count))
        .attr("fill", d => color(d.count))
        .on("mouseover", function(event, d) {
          tooltip
            .style("opacity", 1)
            .html(`<strong>${d.type}</strong><br/>Count: ${d.count}`)
            .style("left", (event.pageX + 10) + "px")
            .style("top", (event.pageY - 10) + "px");
        })
        .on("mouseout", function() {
          tooltip.style("opacity", 0);
        });
      
      // Add X axis
      svg.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x))
        .selectAll("text")
        .attr("transform", "rotate(-45)")
        .style("text-anchor", "end")
        .style("font-size", "10px");
      
      // Add Y axis
      svg.append("g")
        .call(d3.axisLeft(y));
      
      // Add Y axis label
      svg.append("text")
        .attr("transform", "rotate(-90)")
        .attr("y", 0 - margin.left)
        .attr("x", 0 - (height / 2))
        .attr("dy", "1em")
        .style("text-anchor", "middle")
        .style("font-size", "12px")
        .text("Issue Count");
    }
    
    // Create visualizations
    if (categoryData.length > 0) {
      createPieChart(categoryData, 'categoryChart');
    }
    if (issueTypeData.length > 0) {
      createBarChart(issueTypeData, 'issueTypeChart');
    }
  </script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

def process_file_sequential(pptx_path, screenshot_dir, mode, ocr_enabled, expected_language, no_screenshots, verbose):
    """
    Process a single PowerPoint file sequentially.
    
    NOTE: Files are processed ONE AT A TIME to avoid LibreOffice conflicts.
    LibreOffice headless mode cannot handle multiple simultaneous instances.
    """
    # Create screenshot directory
    os.makedirs(screenshot_dir, exist_ok=True)
    
    # Audit the file (optimized: only screenshots problem slides)
    issues_by_slide, screenshots, presentation_issues, total_slides = audit_pptx(
        pptx_path, screenshot_dir, mode=mode, ocr_enabled=ocr_enabled,
        expected_language=expected_language, no_screenshots=no_screenshots,
        verbose=verbose
    )
    
    # Calculate score
    score, error_count, warning_count, issue_counts = calculate_accessibility_score(
        issues_by_slide, presentation_issues, total_slides
    )
    
    return {
        'pptx_path': pptx_path,
        'issues_by_slide': issues_by_slide,
        'screenshots': screenshots,
        'presentation_issues': presentation_issues,
        'total_slides': total_slides,
        'score': score,
        'issue_counts': issue_counts
    }

def main():
    args = parse_args()
    input_dir = args.input_dir

    # Timestamped output directory
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_dir = f"{args.output_dir}_{timestamp}"

    mode = args.mode
    ocr_enabled = args.ocr
    no_screenshots = args.no_screenshots
    expected_language = args.expected_language
    
    # Warn if OCR is enabled but screenshots are disabled
    if ocr_enabled and no_screenshots:
        print("Warning: OCR requires screenshots. OCR will be disabled.")
        ocr_enabled = False

    if not os.path.isdir(input_dir):
        print(f"Error: {input_dir} is not a valid directory")
        sys.exit(1)
    os.makedirs(output_dir, exist_ok=True)
    print(f"Scanning {input_dir} for PPTX files...")
    print(f"Mode: {mode}")
    print(f"OCR: {'enabled' if ocr_enabled else 'disabled'}")
    print(f"Screenshots: {'disabled' if no_screenshots else 'enabled'}")
    print(f"Expected language: {expected_language}")
    
    print(f"\nOptions in effect:")
    print(f"  - Two-pass scanning: scan first, then screenshot only problem slides")
    print(f"  - Hash-based caching: unchanged files are skipped")
    print(f"  - Sequential processing: one file at a time (avoids LibreOffice conflicts)")
    print(f"  - Auto-cleanup: unreferenced screenshots are removed after processing")
    
    books = defaultdict(list)
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.startswith('._'):
                continue
            if file.lower().endswith('.pptx') and not file.startswith('~$'):
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(root, input_dir)
                book_title = rel_path.split(os.sep)[0] if rel_path != '.' else 'Root'
                books[book_title].append(file_path)
    if not books:
        print("No PPTX files found.")
        sys.exit(0)
    
    total_files = sum(len(files) for files in books.values())
    print(f"Found {total_files} PPTX files across {len(books)} book(s)\n")
    
    book_scores = {}
    book_slide_counts = {}  # Track total slides per book for weighted averaging
    book_issue_counts = defaultdict(lambda: defaultdict(int))
    
    # Create initial rollup reports (empty)
    print("\n" + "="*60)
    print("Creating initial rollup reports...")
    csv_path, rollup_html_path = update_rollup_reports(
        output_dir, books, book_scores, book_slide_counts, book_issue_counts
    )
    print(f"Initial rollup created: {rollup_html_path}")
    print("(Will be updated as each title completes)")
    print("="*60)
    
    # Progress bar for all files
    with tqdm(total=total_files, desc="Processing files", unit="file") as pbar:
        for book_title, pptx_files in books.items():
            print(f"\nProcessing book: {book_title} ({len(pptx_files)} files)")
            book_output_dir = os.path.join(output_dir, book_title.replace(' ', '_'))
            os.makedirs(book_output_dir, exist_ok=True)
            files_data = []
            book_weighted_score = 0  # Score weighted by slide count
            book_total_slides = 0     # Total slides across all files in book
            
            # Process files SEQUENTIALLY (LibreOffice can't handle parallel execution)
            for pptx_path in pptx_files:
                file_name = os.path.splitext(os.path.basename(pptx_path))[0]
                screenshot_dir = os.path.join(book_output_dir, f"{file_name}_screenshots")
                
                # Process file
                result = process_file_sequential(
                    pptx_path, screenshot_dir, mode, ocr_enabled,
                    expected_language, no_screenshots, args.verbose
                )
                
                issues_by_slide = result['issues_by_slide']
                screenshots = result['screenshots']
                presentation_issues = result['presentation_issues']
                total_slides = result['total_slides']
                score = result['score']
                issue_counts = result['issue_counts']
                
                # Clean up unreferenced screenshots (slides without issues)
                if not no_screenshots and issues_by_slide:
                    referenced_slides = set(issues_by_slide.keys())
                    cleanup_unreferenced_screenshots(screenshot_dir, referenced_slides)
                
                # Weight score by number of slides in this file
                book_weighted_score += (score * total_slides)
                book_total_slides += total_slides
                
                # Update book issue counts
                for issue_type, count in issue_counts.items():
                    book_issue_counts[book_title][issue_type] += count
                
                files_data.append((pptx_path, issues_by_slide, screenshots, presentation_issues, total_slides))
                
                # Update progress bar
                pbar.set_postfix_str(f"{os.path.basename(pptx_path)} - Score: {score}%")
                pbar.update(1)
            
            # Calculate weighted average score for the book
            # (Each file's score is weighted by its number of slides)
            avg_book_score = book_weighted_score / book_total_slides if book_total_slides > 0 else 0
            book_scores[book_title] = round(avg_book_score, 1)
            book_slide_counts[book_title] = book_total_slides
            
            report_path = os.path.join(output_dir, f"{book_title.replace(' ', '_')}_report.html")
            generate_html_report(book_title, files_data, report_path)
            print(f"  Report saved: {report_path}")
            print(f"  Book Score: {book_scores[book_title]}%")
            
            # Update rollup reports after each book completes
            print(f"  Updating rollup reports...")
            csv_path, rollup_html_path = update_rollup_reports(
                output_dir, books, book_scores, book_slide_counts, book_issue_counts
            )
            print(f"  Rollup updated: {rollup_html_path}")
    
    print("\n" + "="*60)
    print("✅ Audit complete!")
    print(f"📊 Final reports:")
    print(f"  • Rollup: {rollup_html_path}")
    print(f"  • CSV: {csv_path}")
    print(f"  • Processed {len(book_scores)} collection(s) with {total_files} files")

if __name__ == '__main__':
    main()