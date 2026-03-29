# PDF Accessibility Auto Tool

Best-effort accessibility cleanup for already tagged PDFs.

## Files

- `pdf_accessibility_auto.py`
- `requirements.txt`

## What the script does

The script works on PDFs that already contain a `/StructTreeRoot` and attempts to improve common accessibility gaps by:

- adding `/Alt` text to tagged `Figure`, `Table`, and `Formula` elements
- adding `/Summary` text to tagged tables
- removing nested `/Alt`, `/ActualText`, and `/Summary` values under those elements
- promoting missing table headers from `TD` to `TH`
- promoting likely headings from `P` to `H1`, `H2`, and `H3`
- setting PDF title metadata
- setting PDF language metadata
- optionally deleting selected pages
- optionally removing visible running page numbers and overlapping annotations

## What it does not do

This tool does not build a full accessibility tag tree from an untagged PDF.

If a PDF does not already have a valid `/StructTreeRoot`, this script is not the right starting point.

## Requirements

- Python 3.10+
- PyMuPDF
- pypdf

Install dependencies:

```bash
pip install -r requirements.txt
```

## Basic usage

```bash
python pdf_accessibility_auto.py input.pdf -o output.pdf
```

Set title and language:

```bash
python pdf_accessibility_auto.py input.pdf -o output.pdf --title "Document Title" --lang en-US
```

Delete one or more pages before processing:

```bash
python pdf_accessibility_auto.py input.pdf -o output.pdf --delete-page 2 --delete-page 5
```

Remove visible running page numbers:

```bash
python pdf_accessibility_auto.py input.pdf -o output.pdf --clean-running-page-numbers
```

Disable some heuristic transformations:

```bash
python pdf_accessibility_auto.py input.pdf -o output.pdf --disable-heading-promotion --disable-table-header-promotion
```

## Output

The script prints a JSON summary with:

- missing alt-text counts before and after
- number of promoted heading tags
- number of promoted table header cells
- number of running page-number regions redacted
- metadata values applied

## Notes

- The heading and table-header fixes are heuristic.
- The page-number cleanup removes visible page numbers when enabled.
- This tool is intended as a practical post-processing step, not a PDF/UA certification pipeline.
