# mrconvert

Convert PDFs and Word (DOCX) to machine-readable **TXT/MD/JSON**.
Optional OCR: uses **ocrmypdf** (if installed) or **pytesseract+pymupdf** fallback.

## Install (Python ≥ 3.11)
```bash
pip install -e ".[ocr]"
# or without OCR
pip install -e .
```

## CLI

### Text Extraction Mode (Default)
```bash
mrconvert INPUT_PATH --out OUT_DIR --format txt md json --tables csv --ocr auto --lang kor+eng
```

### Bidirectional Conversion Mode
```bash
mrconvert INPUT_PATH --to-docx    # PDF → DOCX
mrconvert INPUT_PATH --to-pdf     # DOCX → PDF
```

### Examples

#### Text Extraction
```bash
# 1) Convert a single PDF to Markdown + JSON with tables as CSV
mrconvert sample.pdf --out out --format md json --tables csv

# 2) Bulk convert a folder, OCR when needed
mrconvert ./incoming --out ./out --format txt --ocr auto --lang kor+eng

# 3) Force OCR (e.g., scanned PDF)
mrconvert scan.pdf --out out --format txt --ocr force
```

#### Bidirectional Conversion
```bash
# 4) Convert PDF to DOCX
mrconvert document.pdf --to-docx --out ./converted

# 5) Convert DOCX to PDF
mrconvert document.docx --to-pdf

# 6) Batch convert multiple files
mrconvert ./pdfs --to-docx --out ./docx_output
```

## Output
For `--format json`, schema:
```json
{
  "meta": {
    "source": "<path>",
    "type": "pdf|docx",
    "pages": 10,
    "parsed_at": "YYYY-MM-DDTHH:MM:SSZ",
    "ocr": {"used": true, "engine": "ocrmypdf|pytesseract|none", "lang": "kor+eng"}
  },
  "text": "...plain text...",
  "markdown": "...optional markdown...",
  "tables": [
    {"page": 1, "index": 0, "rows": [["A","B"],["1","2"]]}
  ]
}
```

## Notes
- **.doc** (legacy) not supported directly. Use LibreOffice to convert to .docx:
  `soffice --headless --convert-to docx file.doc`
- OCR quality depends on the engine and language packs installed.

MIT License.
