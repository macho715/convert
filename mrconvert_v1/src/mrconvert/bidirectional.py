from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional

from .types import (
    ConversionResult,
    ConversionError,
    UnsupportedFormatError,
    EngineNotFoundError,
)


def convert_file(
    in_path: str | Path,
    out_path: str | Path | None = None,
    *,
    target: Optional[str] = None,
) -> ConversionResult:
    """PDF<->DOCX 단방향 자동 판별 변환기."""
    src = Path(in_path).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(src)

    suffix = src.suffix.lower()
    if target:
        target = target.lower().lstrip(".")

    match suffix:
        case ".pdf":
            dst = Path(out_path) if out_path else src.with_suffix(".docx")
            return pdf_to_docx(src, dst)
        case ".docx":
            dst = Path(out_path) if out_path else src.with_suffix(".pdf")
            return docx_to_pdf(src, dst)
        case _:
            raise UnsupportedFormatError(
                f"Unsupported input type: {suffix} (use .pdf or .docx)"
            )


def pdf_to_docx(src: Path, dst: Path) -> ConversionResult:
    """Convert PDF to DOCX using pdf2docx"""
    from pdf2docx import Converter, parse  # type: ignore

    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        cv = Converter(str(src))
        try:
            cv.convert(str(dst))  # 모든 페이지
        finally:
            cv.close()
    except Exception:
        # 보조 경로: parse() 한 방
        parse(str(src), str(dst))

    if not dst.exists():
        raise ConversionError(f"pdf2docx failed to create {dst}")
    return ConversionResult(src, dst, "pdf2docx")


def docx_to_pdf(src: Path, dst: Path) -> ConversionResult:
    """Convert DOCX to PDF using docx2pdf (Windows/macOS) or soffice (Linux)"""
    # 1) Word가 있는 Windows/macOS면 docx2pdf 우선
    try:
        from docx2pdf import convert as d2p_convert  # type: ignore

        if sys.platform in {"win32", "darwin"}:
            dst.parent.mkdir(parents=True, exist_ok=True)
            d2p_convert(str(src), str(dst))
            if dst.exists():
                return ConversionResult(src, dst, "docx2pdf")
    except Exception:
        # Linux 등 환경에선 예외가 날 수 있음 → 폴백
        pass

    # 2) 리눅스/서버 폴백: LibreOffice 'soffice'
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise EngineNotFoundError(
            "docx2pdf unavailable and LibreOffice (soffice) not found in PATH."
        )

    outdir = dst.parent
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(outdir),
        str(src),
    ]
    proc = subprocess.run(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
    )
    if proc.returncode != 0:
        raise ConversionError(f"LibreOffice failed: {proc.stderr or proc.stdout}")

    produced = outdir / (src.with_suffix(".pdf").name)
    if produced != dst and produced.exists():
        produced.rename(dst)

    if not dst.exists():
        raise ConversionError(f"Expected {dst} but not found after conversion.")
    return ConversionResult(src, dst, "soffice")
