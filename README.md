# PowerPoint accessibility scanner

Checks PowerPoint files against WCAG 2.2 AA-style criteria in bulk and reports what to fix, slide by slide. It is built for the case where you have hundreds of decks and need to know which are worst before you open any of them.

## Two scanners

**`pptx_scanner.py`: local audit with rich reports.** Point it at a folder of `.pptx` files (searched recursively). It runs these checks, in `lenient` or `strict` mode:

- missing alt text: none, empty, or just a file name ("chart.png"). Images marked with PowerPoint's "Mark as decorative" are skipped, but an empty description on its own is reported, because older production processes often left meaningful images empty;
- in `strict` mode, alt text that's too short, generic ("image") or a default shape name ("Picture 3");
- large images that look like scanned pages;
- images of text (optional OCR with Tesseract);
- slides without titles, and a deck-level heading check;
- tables without PowerPoint's "Header Row" option set;
- link text that doesn't describe the destination;
- text contrast, using WCAG's 4.5:1 and 3:1 (large text) thresholds, with theme colours flagged for manual review;
- reading order, as a position-based heuristic;
- auto-advance transitions and auto-playing or excessive animation;
- document language.

It scores each deck and writes HTML reports (per-file detail and a rollup with D3 charts) plus a CSV. It screenshots only slides that have issues, caches results by file hash so unchanged decks are skipped, and removes screenshots no report references.

**Optional AI review of alt text (`--ai-alt-check`).** Rule-based checks can tell that an image has alt text, not whether it describes the image. With this option, each image that has alt text (and passed the checks above) is sent with its alt text and the slide's text to a local vision model through [Ollama](https://ollama.com), which answers "adequate" or "needs review" with a one-line reason. Flags appear in the reports as "AI Review: Alt Text". They are prompts for a person, not findings: they don't change scores or count as errors or warnings, and the model can be wrong in both directions. Results are cached by image, alt text and model (`--ai-cache`), so rescans only ask about what changed. The model runs on your machine, so no images are sent anywhere.

```bash
ollama pull gemma3        # or another vision model: --ai-model llava
python3 pptx_scanner.py /path/to/decks --mode strict --ai-alt-check
```

Before relying on it, run it on a sample you've already judged yourself and see how often it agrees with you.

**`s3-bulk-scan/s3_pptx_scanner.py`: in-memory bulk scan.** A leaner variant for a large store of resources. It reads ZIP files from an S3 bucket, extracts the decks in memory, checks each one, and writes a single CSV. ZIPs are processed in parallel, and ones unchanged since the last scan (by S3 modified time) are skipped. It does no screenshots or PDF conversion, and has no AI review. Its alt text checks match the local scanner's, including decorative images.

## Usage

```bash
pip install python-pptx Pillow pdf2image boto3
brew install poppler libreoffice          # screenshots (local scanner)

python3 pptx_scanner.py /path/to/decks --mode strict --output-dir reports
python3 s3-bulk-scan/s3_pptx_scanner.py my-bucket --prefix resources/ --output report.csv
```

Useful local options: `--ocr`, `--no-screenshots`, `--expected-language en-US`, `--verbose`, and `--ai-alt-check` with `--ai-model`, `--ai-host` (default `http://localhost:11434`, or `OLLAMA_HOST`) and `--ai-cache`. The S3 scanner uses your normal AWS credential chain (environment variables, SSO or `~/.aws`). Nothing is stored in the repo.

## Limits

- Automated checks find a subset of problems. Contrast and reading order are heuristics, and theme colours can't always be resolved. A clean scan does not mean a deck is accessible.
- The checks are not tied to a formal conformance test suite. Treat results as triage.
- `index.html` is a simple reports landing page. It expects the generated report folders next to it.

## Licence

MIT. See [LICENSE](LICENSE).
