# PowerPoint accessibility scanner

Checks PowerPoint files against WCAG 2.2 AA-style criteria in bulk and reports what to fix, slide by slide. It is built for the case where you have hundreds of decks and need to know which are worst before you open any of them.

## Two scanners

**`pptx_scanner.py`: local audit with rich reports.** Point it at a folder of `.pptx` files (searched recursively). It runs these checks, in `lenient` or `strict` mode:

- missing or poor-quality alt text, including large images that look like scanned pages;
- images of text (optional OCR with Tesseract);
- slides without titles, and a deck-level heading check;
- table header rows;
- link text that doesn't describe the destination;
- text contrast, using WCAG's 4.5:1 and 3:1 (large text) thresholds, with theme colours flagged for manual review;
- reading order, as a position-based heuristic;
- auto-advance transitions and auto-playing or excessive animation;
- document language.

It scores each deck and writes HTML reports (per-file detail and a rollup with D3 charts) plus a CSV. It screenshots only slides that have issues, caches results by file hash so unchanged decks are skipped, and removes screenshots no report references.

**`s3-bulk-scan/s3_pptx_scanner.py`: in-memory bulk scan.** A leaner variant for a large store of resources. It reads ZIP files from an S3 bucket, extracts the decks in memory, checks each one, and writes a single CSV. ZIPs are processed in parallel, and ones unchanged since the last scan (by S3 modified time) are skipped. It does no screenshots or PDF conversion.

## Usage

```bash
pip install python-pptx Pillow pdf2image boto3
brew install poppler libreoffice          # screenshots (local scanner)

python3 pptx_scanner.py /path/to/decks --mode strict --output-dir reports
python3 s3-bulk-scan/s3_pptx_scanner.py my-bucket --prefix resources/ --output report.csv
```

Useful local options: `--ocr`, `--no-screenshots`, `--expected-language en-US`, `--verbose`. The S3 scanner uses your normal AWS credential chain (environment variables, SSO or `~/.aws`). Nothing is stored in the repo.

## Limits

- Automated checks find a subset of problems. Contrast and reading order are heuristics, and theme colours can't always be resolved. A clean scan does not mean a deck is accessible.
- The checks are not tied to a formal conformance test suite. Treat results as triage.
- `index.html` is a simple reports landing page. It expects the generated report folders next to it.

## Licence

MIT. See [LICENSE](LICENSE).
