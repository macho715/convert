from __future__ import annotations

import io
import os
import shutil
import subprocess
from pathlib import Path
from typing import Dict, Any, List

import pdfplumber

try:
    import fitz  # PyMuPDF, optional
except Exception:  # pragma: no cover
    fitz = None

try:
    import pytesseract
    from PIL import Image
except Exception:  # pragma: no cover
    pytesseract = None
    Image = None

from .utils import OCRConfig, now_iso


def _ensure_tessdata_prefix() -> None:
    """Set TESSDATA_PREFIX to CONVERT/out/tessdata if unset and eng.traineddata exists."""
    if os.environ.get("TESSDATA_PREFIX"):
        return
    # mrconvert_v1/src/mrconvert/pdf_converter.py -> parent^4 = repo root (CONVERT)
    _root = Path(__file__).resolve().parent.parent.parent.parent
    _tessdata = _root / "out" / "tessdata"
    if (_tessdata / "eng.traineddata").exists():
        os.environ.setdefault("TESSDATA_PREFIX", str(_tessdata))


def _page_text(page, keep_layout: bool) -> str | None:
    # pdfplumber's extract_text has options; we use simple mode, keep_layout implies x_tolerance/y_tolerance tweaks
    text = page.extract_text(x_tolerance=1.5 if keep_layout else 3.0, y_tolerance=1.5 if keep_layout else 3.0)
    return text

def _needs_ocr(text: str | None, force: bool) -> bool:
    if force:
        return True
    if text is None:
        return True
    t = text.strip()
    return len(t) < 8  # too little text → probably scanned

def _ocr_with_ocrmypdf(pdf_path: Path, tmp_dir: Path, lang: str | None) -> Path | None:
    if shutil.which("ocrmypdf") is None:
        return None
    out_pdf = tmp_dir / (pdf_path.stem + ".ocr.pdf")
    cmd = ["ocrmypdf", "--skip-text", "--quiet"]
    if lang:
        cmd += ["-l", lang]
    cmd += [str(pdf_path), str(out_pdf)]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return out_pdf if out_pdf.exists() else None
    except Exception:
        return None

def _ocr_page_with_pymupdf(page, lang: str | None) -> str | None:
    # Render page to image then OCR with pytesseract
    if (fitz is None) or (pytesseract is None) or (Image is None):
        return None
    _ensure_tessdata_prefix()
    pix = page.to_pixmap(dpi=300)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    conf = {}
    if lang:
        conf["lang"] = lang
    try:
        return pytesseract.image_to_string(img, **conf)
    except Exception:
        return None

def extract_pdf(path: Path, keep_layout: bool = False, ocr: OCRConfig | None = None) -> Dict[str, Any]:
    ocr = ocr or OCRConfig()
    meta: Dict[str, Any] = {
        "source": str(path),
        "type": "pdf",
        "pages": 0,
        "parsed_at": now_iso(),
        "ocr": {"used": False, "engine": "none", "lang": ocr.lang},
    }
    text_parts: List[str] = []
    tables: List[Dict[str, Any]] = []

    tmp_dir = path.parent / ".mrconvert_tmp"
    tmp_dir.mkdir(exist_ok=True)

    pdf_to_open = path

    # Pass 0: if user forced OCR and ocrmypdf exists, pre-OCR the whole file for better text flow.
    if ocr.mode == "force":
        ocr_pdf = _ocr_with_ocrmypdf(path, tmp_dir, ocr.lang)
        if ocr_pdf:
            pdf_to_open = ocr_pdf
            meta["ocr"].update({"used": True, "engine": "ocrmypdf"})

    with pdfplumber.open(pdf_to_open) as pdf:
        meta["pages"] = len(pdf.pages)
        for i, page in enumerate(pdf.pages, start=1):
            raw = _page_text(page, keep_layout=keep_layout)
            # If text is empty and OCR mode allows, try OCR per-page
            if _needs_ocr(raw, force=(ocr.mode == "force")) and ocr.mode in {"auto", "force"}:
                # Try per-page OCR with PyMuPDF+pytesseract for this page
                if fitz and pytesseract:
                    # Render with PyMuPDF if original PDF specified
                    try:
                        if str(pdf_to_open) != str(path):
                            # We already opened an OCRed PDF — try text again before OCRing image
                            raw = raw or ""
                        else:
                            # Render page with PyMuPDF
                            with fitz.open(str(path)) as doc2:
                                page2 = doc2.load_page(i - 1)
                                ocr_text = _ocr_page_with_pymupdf(page2, ocr.lang)
                                if ocr_text and len(ocr_text.strip()) > 0:
                                    raw = (raw or "") + "\n" + ocr_text
                                    meta["ocr"].update({"used": True, "engine": "pytesseract"})
                    except Exception:
                        pass
                # If still empty and ocrmypdf available, as a fallback OCR the whole pdf once (auto case)
                if (not raw or len(raw.strip()) == 0) and ocr.mode == "auto":
                    ocr_pdf = _ocr_with_ocrmypdf(path, tmp_dir, ocr.lang)
                    if ocr_pdf:
                        try:
                            with pdfplumber.open(ocr_pdf) as pdf2:
                                page2 = pdf2.pages[i - 1]
                                raw = _page_text(page2, keep_layout=keep_layout) or ""
                                meta["ocr"].update({"used": True, "engine": "ocrmypdf"})
                        except Exception:
                            pass
            text_parts.append(raw or "")

            # Tables (best-effort)
            try:
                tbs = page.extract_tables()
                for ti, tbl in enumerate(tbs or []):
                    tables.append({"page": i, "index": ti, "rows": tbl})
            except Exception:
                # graceful degrade
                pass

    # Cleanup temp dir
    try:
        if tmp_dir.exists():
            for p in tmp_dir.glob("*"):
                p.unlink(missing_ok=True)
            tmp_dir.rmdir()
    except Exception:
        pass

    return {
        "meta": meta,
        "text": "\n\n".join(text_parts).strip(),
        "markdown": None,  # We don't generate MD from PDF by default; can be added later
        "tables": tables,
    }
