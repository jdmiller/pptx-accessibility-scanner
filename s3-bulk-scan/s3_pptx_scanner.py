#!/usr/bin/env python3
"""
S3 PowerPoint Accessibility Scanner
===================================

Scans an S3 bucket for ZIP files containing "ppt" in their name.
Downloads ZIPs to memory, extracts PPTX files, and performs accessibility checks
based on Anthology Ally requirements.

Outputs a single CSV file with all issues.

NO local screenshots.
NO PDF conversion.
Fully in-memory processing.
PARALLEL ZIP PROCESSING: Processing multiple ZIP files concurrency.
INCREMENTAL SCANNING: Skips files if S3 LastModified matches report.
"""

import argparse
import boto3
import csv
import io
import os
import sys
import zipfile
import traceback
import tempfile
import concurrent.futures
import multiprocessing
from datetime import datetime
from collections import defaultdict
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from tqdm import tqdm
from docx import Document as DocxDocument

# Anthology Ally-aligned Issue Severity and Categorization
ISSUE_SEVERITY = {
    'malformed_presentation': 'Severe',
    'encrypted_presentation': 'Severe',
    'missing_language': 'Minor',
    'incorrect_language': 'Minor',
    'missing_alt_text': 'Major',
    'insufficient_contrast': 'Major',
    'no_headings': 'Major',
    'table_no_header': 'Major',
    'missing_slide_title': 'Major',
    'reading_order': 'Minor',
    'poor_link_text': 'Minor'
}

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

def get_alt_text_from_shape(shape):
    """Helper to extract alt text from shape."""
    if not shape.has_text_frame:
        # Default to empty string. Do NOT use shape.name (e.g. "Picture 1") as valid alt text.
        alt_text = ""
        try:
            if hasattr(shape, '_element') and hasattr(shape._element, 'xpath'):
                # Try to get 'descr' (Description) - the primary Alt Text field
                alt_text_elem = shape._element.xpath('.//p:cNvPr/@descr')
                if alt_text_elem:
                    alt_text = alt_text_elem[0]
                
                # Some versions might put it in 'title'
                if not alt_text:
                    title_elem = shape._element.xpath('.//p:cNvPr/@title')
                    if title_elem:
                         alt_text = title_elem[0]

                # Check if alt text looks like a filename (e.g. ends with common image extensions)
                if alt_text:
                    common_exts = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tif', '.tiff', '.svg', '.emf', '.wmf')
                    if alt_text.lower().strip().endswith(common_exts):
                        # Treat filename as missing alt text
                        alt_text = ""
        except:
            pass
        return alt_text
    return ""

def check_contrast(slide, shape, slide_idx, issues):
    """Check text contrast in a shape."""
    if not shape.has_text_frame:
        return

    try:
        # Simplified contrast check
        # Hard to get exact background color without rendering, so we try best effort
        # assuming white background if not found
        bg_color = (255, 255, 255) 
        
        # Try to find slide background
        try:
             if slide.background.fill.type == 1: # Solid
                 bg_color = tuple(slide.background.fill.fore_color.rgb)
        except:
             pass

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_color = None
                try:
                    if run.font.color and run.font.color.rgb:
                        text_color = tuple(run.font.color.rgb)
                except:
                    continue
                
                if text_color:
                    ratio = get_contrast_ratio(text_color, bg_color)
                    if ratio < 4.5:
                        issues.append({
                            'issue_type': 'insufficient_contrast',
                            'severity': ISSUE_SEVERITY['insufficient_contrast'],
                            'description': f'Contrast ratio {ratio:.2f}:1 < 4.5:1 (Text: "{run.text[:20]}...")',
                            'slide': slide_idx
                        })
    except Exception:
        pass

def audit_pptx_bytes(file_content, filename, strict_mode=False):
    """
    Audit a PPTX file from bytes.
    Returns a list of issues.
    """
    issues = []
    
    try:
        prs = Presentation(io.BytesIO(file_content))
    except Exception as e:
        msg = str(e).lower()
        if "encrypted" in msg or "password" in msg:
             return [{
                'issue_type': 'encrypted_presentation',
                'severity': ISSUE_SEVERITY['encrypted_presentation'],
                'description': 'Presentation is encrypted or password protected',
                'slide': 'N/A'
            }]
        
        return [{
            'issue_type': 'malformed_presentation',
            'severity': ISSUE_SEVERITY['malformed_presentation'],
            'description': f'Could not open presentation: {str(e)}',
            'slide': 'N/A'
        }]

    # 1. Language Check
    try:
        lang = prs.core_properties.language
        if not lang:
            issues.append({
                'issue_type': 'missing_language',
                'severity': ISSUE_SEVERITY['missing_language'],
                'description': 'Presentation does not have a language set',
                'slide': 'All'
            })
        elif strict_mode and not lang.startswith('en'):
             issues.append({
                'issue_type': 'incorrect_language',
                'severity': ISSUE_SEVERITY['incorrect_language'],
                'description': f'Language set to "{lang}", expected English',
                'slide': 'All'
            })
    except:
        pass

    # 2. Heading/Title Checks
    total_slides = len(prs.slides)
    slides_with_titles = 0
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        has_title = False
        if slide.shapes.title and slide.shapes.title.text and slide.shapes.title.text.strip():
            has_title = True
            slides_with_titles += 1
        else:
            issues.append({
                'issue_type': 'missing_slide_title',
                'severity': ISSUE_SEVERITY['missing_slide_title'],
                'description': 'Slide does not have a title',
                'slide': slide_idx
            })

        # 3. Slide Content Checks
        for shape in slide.shapes:
            if shape.is_placeholder and not shape.has_text_frame:
                continue

            # Alt Text
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # "Mark as decorative" in PowerPoint: screen readers skip it. An empty
                # description alone still counts as missing (legacy files often have them).
                if shape._element.xpath('.//p:cNvPr//*[local-name()="decorative" and (@val="1" or @val="true")]'):
                    continue
                alt_text = get_alt_text_from_shape(shape)
                if not alt_text or not alt_text.strip():
                    issues.append({
                        'issue_type': 'missing_alt_text',
                        'severity': ISSUE_SEVERITY['missing_alt_text'],
                        'description': f'Image "{shape.name}" has no alternative description',
                        'slide': slide_idx
                    })

            # Tables
            if shape.has_table:
                tbl = shape.table
                try:
                    has_header_opt = False
                    if hasattr(tbl, '_tbl') and hasattr(tbl._tbl, 'tblPr'):
                         if tbl._tbl.tblPr.firstRow:
                             has_header_opt = True
                    
                    if not has_header_opt:
                        issues.append({
                            'issue_type': 'table_no_header',
                            'severity': ISSUE_SEVERITY['table_no_header'],
                            'description': 'Table does not have "Header Row" property set',
                            'slide': slide_idx
                        })
                except:
                    pass

            # Contrast & Text
            if shape.has_text_frame:
                check_contrast(slide, shape, slide_idx, issues)

    # Global Heading Check
    if total_slides > 0 and slides_with_titles == 0:
        issues.append({
            'issue_type': 'no_headings',
            'severity': ISSUE_SEVERITY['no_headings'],
            'description': 'The presentation does not have any headings',
            'slide': 'All'
        })

    return issues

def process_single_zip(zip_key, last_modified_str, bucket_name, aws_region='us-east-1'):
    """
    Worker function to process a single ZIP file.
    Downloads the zip, iterates through PPTX files, and returns a list of result rows.
    """
    issue_columns = sorted(ISSUE_SEVERITY.keys())
    results = []
    
    try:
        # Create a new session/client per process for thread safety
        session = boto3.session.Session()
        s3 = session.client('s3', region_name=aws_region)
        
        # Download ZIP to memory
        zip_obj = s3.get_object(Bucket=bucket_name, Key=zip_key)
        zip_content = zip_obj['Body'].read()
        
        with zipfile.ZipFile(io.BytesIO(zip_content)) as zf:
            # Sample a title from the first PPTX/DOCX in this ZIP for mapping verification
            sampled_title, sample_source = sample_title_from_zip(zf)

            # Filter for PPTX files, excluding hidden/metadata files
            pptx_files = [
                f for f in zf.namelist() 
                if f.lower().endswith('.pptx') 
                and not f.startswith('~$') 
                and not f.startswith('._') 
                and '__MACOSX' not in f
                and '/._' not in f
            ]
            
            if not pptx_files:
                return [], sampled_title, sample_source
            
            for pptx_file in pptx_files:
                # Initialize counters for this file
                file_counts = {col: 0 for col in issue_columns}
                other_issues_count = 0
                
                try:
                    file_content = zf.read(pptx_file)
                    file_issues = audit_pptx_bytes(file_content, pptx_file)
                    
                    for issue in file_issues:
                        itype = issue['issue_type']
                        if itype in file_counts:
                            file_counts[itype] += 1
                        else:
                            other_issues_count += 1
                            
                    # Construct Row
                    # Format: Zip File, PPTX File, Last Modified, [Issues...], Other
                    row = [zip_key, pptx_file, last_modified_str]
                    for col in issue_columns:
                        row.append(file_counts[col])
                    row.append(other_issues_count)
                    results.append(row)
                    
                except Exception as exc:
                    # Error row
                    error_row = [zip_key, pptx_file, last_modified_str]
                    for col in issue_columns:
                        if col == 'malformed_presentation':
                            error_row.append(1)
                        else:
                            error_row.append(0)
                    error_row.append(0) # Other
                    results.append(error_row)
                    
    except Exception as e:
        # Zip level error
        error_row = [zip_key, 'ZIP_READ_ERROR', last_modified_str] + (['ERROR'] * len(issue_columns)) + [str(e)]
        results.append(error_row)
        
    return results

def sample_title_from_zip(zf):
    """
    Extracts a title string from the first PPTX (preferred) or DOCX in an open ZipFile.
    For PPTX: reads the title placeholder on slide 1; falls back to the first text shape.
    For DOCX: reads the first non-empty paragraph.
    Returns (title_text, source_label) where source_label is '<filename>:slide1' or
    '<filename>:para1', or ('', '') if nothing could be extracted.
    """
    names = zf.namelist()

    def _valid(f, ext):
        return (
            f.lower().endswith(ext)
            and not f.startswith('~$')
            and not f.startswith('._')
            and '__MACOSX' not in f
            and '/._' not in f
        )

    pptx_files = [f for f in names if _valid(f, '.pptx')]
    docx_files = [f for f in names if _valid(f, '.docx')]

    # Try PPTX first
    for pptx_file in pptx_files[:1]:
        try:
            content = zf.read(pptx_file)
            prs = Presentation(io.BytesIO(content))
            if not prs.slides:
                continue
            slide = prs.slides[0]
            title = ''
            if slide.shapes.title and slide.shapes.title.text and slide.shapes.title.text.strip():
                title = slide.shapes.title.text.strip()
            else:
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text_frame.text.strip():
                        title = shape.text_frame.text.strip()[:200]
                        break
            if title:
                return title[:200], f'{pptx_file}:slide1'
        except Exception:
            continue

    # Fall back to DOCX
    for docx_file in docx_files[:1]:
        try:
            content = zf.read(docx_file)
            doc = DocxDocument(io.BytesIO(content))
            for para in doc.paragraphs:
                if para.text.strip():
                    return para.text.strip()[:200], f'{docx_file}:para1'
        except Exception:
            continue

    return '', ''


def extract_s3_prefix(zip_key):
    """
    Extracts the title-level folder name from an S3 ZIP key.
    e.g. 'nrlassets/secure/21ast5/21ast5_artppt_ch01.zip' -> '21ast5'
    Returns the parent directory of the ZIP file (parts[-2]).
    """
    parts = zip_key.rstrip('/').split('/')
    if len(parts) >= 2:
        return parts[-2]
    return ''


def update_mapping_sampled_titles(mapping_csv, title_samples):
    """
    Updates title_s3_mapping.csv with sampled title data collected during a scan.

    title_samples: dict of {candidate_s3_prefix: (sampled_title, sample_source)}

    Adds 'sampled_title' and 'sample_source' columns if they don't exist, then
    fills in values for every row whose candidate_s3_prefix appears in title_samples.
    Rows with no match are left with empty strings in those columns.
    """
    if not title_samples or not os.path.exists(mapping_csv):
        return

    rows = []
    fieldnames = []
    with open(mapping_csv, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        fieldnames = list(reader.fieldnames or [])
        rows = list(reader)

    for col in ('sampled_title', 'sample_source'):
        if col not in fieldnames:
            fieldnames.append(col)

    updated = 0
    for row in rows:
        prefix = row.get('candidate_s3_prefix', '').strip()
        if prefix and prefix in title_samples:
            sampled_title, sample_source = title_samples[prefix]
            row['sampled_title'] = sampled_title
            row['sample_source'] = sample_source
            updated += 1
        else:
            row.setdefault('sampled_title', row.get('sampled_title', ''))
            row.setdefault('sample_source', row.get('sample_source', ''))

    with open(mapping_csv, 'w', encoding='utf-8', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Mapping updated: {updated} rows enriched with sampled titles -> {mapping_csv}")


def load_processed_files(output_csv):
    """
    Reads the existing CSV to identify processed files and their Last Modified dates.
    Returns a dict: {zip_key: last_modified_str}
    """
    processed = {}
    if not os.path.exists(output_csv):
        return processed
        
    try:
        with open(output_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            if not header or 'Last Modified' not in header:
                return processed
            
            # Find indices
            try:
                zip_idx = header.index('Zip File')
                mod_idx = header.index('Last Modified')
            except ValueError:
                return processed
                
            for row in reader:
                if len(row) > max(zip_idx, mod_idx):
                    zip_key = row[zip_idx]
                    last_mod = row[mod_idx]
                    processed[zip_key] = last_mod
    except Exception:
        # If any error reading (e.g. corrupt CSV), return empty to re-process safely
        pass
    return processed

def process_s3_bucket(bucket_name, prefix, output_csv, aws_region='us-east-1', max_workers=4, match_criteria=''):
    s3 = boto3.client('s3', region_name=aws_region)
    
    print(f"Listing objects in s3://{bucket_name}/{prefix} (recursively)...")
    
    if prefix and not prefix.endswith('/'):
         prefix += '/'

    paginator = s3.get_paginator('list_objects_v2')
    pages = paginator.paginate(Bucket=bucket_name, Prefix=prefix)
    
    # Identify ZIP files and extract LastModified
    # zip_files list of dicts: {'Key': str, 'LastModified': str}
    zip_files_to_scan = []
    
    match_list = [m.lower() for m in match_criteria.split()] if match_criteria else []
    
    for page in pages:
        if 'Contents' in page:
            for obj in page['Contents']:
                key = obj['Key']
                if key.lower().endswith('.zip'):
                    if match_list and not any(m in key.lower() for m in match_list):
                        continue
                    # Convert datetime to string for CSV comparison compatibility
                    last_mod = obj['LastModified'].isoformat()
                    zip_files_to_scan.append({'Key': key, 'LastModified': last_mod})
    
    print(f"Found {len(zip_files_to_scan)} potential ZIP files.")

    # Load existing progress for incremental scan
    processed_map = load_processed_files(output_csv)
    if processed_map:
        print(f"Found existing report with {len(processed_map)} entries. Checking for updates...")
    
    # Filter for new/modified files
    final_queue = []
    skipped_count = 0
    
    for zf in zip_files_to_scan:
        key = zf['Key']
        s3_mod = zf['LastModified']
        
        # Check if processed and timestamp matches
        if key in processed_map and processed_map[key] == s3_mod:
            skipped_count += 1
            continue
            
        final_queue.append(zf)
        
    print(f"Skipping {skipped_count} unchanged files.")
    print(f"Queueing {len(final_queue)} files for processing.")
    
    if not final_queue:
        print("No new content to process.")
        return

    # Define Issue Columns
    issue_columns = sorted(ISSUE_SEVERITY.keys())
    
    # Write CSV Header if new file
    file_exists = os.path.exists(output_csv)
    mode = 'a' if file_exists else 'w'
    
    with open(output_csv, mode, newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists:
            header = ['Zip File', 'PPTX File', 'Last Modified'] + issue_columns + ['Other Issues']
            writer.writerow(header)
            
    print(f"Starting parallel processing with {max_workers} workers...")
    
    # Process ZIPs in parallel
    with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks
        future_to_zip = {
            executor.submit(process_single_zip, item['Key'], item['LastModified'], bucket_name, aws_region): item['Key']
            for item in final_queue
        }
        
        for future in tqdm(concurrent.futures.as_completed(future_to_zip), total=len(final_queue), desc="Processing ZIPs"):
            zip_key = future_to_zip[future]
            try:
                results = future.result()
                if results:
                    with open(output_csv, 'a', newline='', encoding='utf-8') as f:
                        writer = csv.writer(f)
                        writer.writerows(results)
            except Exception as exc:
                print(f"Critical error processing ZIP {zip_key}: {exc}")

def main():
    parser = argparse.ArgumentParser(description='S3 PowerPoint Accessibility Scanner')
    parser.add_argument('s3_bucket', help='S3 Bucket Name or URL')
    parser.add_argument('--prefix', default='', help='S3 Prefix (Folder)')
    parser.add_argument('--output', default='accessibility_report.csv', help='Output CSV file path')
    parser.add_argument('--region', default='us-east-1', help='AWS Region')
    parser.add_argument('--workers', type=int, default=4, help='Number of concurrent worker processes (default: 4)')
    parser.add_argument('--match', default='', help='Space-separated search criteria for filename, e.g. "artppt art_ppt"')
    
    args = parser.parse_args()
    
    # Support s3:// style URLs
    bucket = args.s3_bucket
    prefix = args.prefix

    if bucket.startswith('s3://'):
        from urllib.parse import urlparse
        parsed = urlparse(bucket)
        bucket = parsed.netloc
        url_prefix = parsed.path.lstrip('/')
        if url_prefix:
            if prefix:
                prefix = f"{url_prefix}/{prefix}"
            else:
                prefix = url_prefix

    try:
        process_s3_bucket(bucket, prefix, args.output, args.region, args.workers, args.match)
        print(f"\nDone! Report saved to {args.output}")
    except Exception as e:
        print(f"\nCritical Error: {e}")
        # traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    # Fix for multiprocessing on some platforms
    multiprocessing.freeze_support()
    main()