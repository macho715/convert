from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List

from rich.console import Console
from rich.progress import track

from . import pdf_converter, docx_converter, bidirectional, markdown_to_docx, docx_to_msg, markdown_to_xlsx
from .utils import ensure_dir, write_text, write_json, walk_inputs, OCRConfig

console = Console()


def _process_file(
    path: Path,
    out_dir: Path,
    formats: List[str],
    tables: str,
    keep_layout: bool,
    ocr_mode: str,
    lang: str | None,
):
    if path.suffix.lower() == ".pdf":
        data = pdf_converter.extract_pdf(
            path, keep_layout=keep_layout, ocr=OCRConfig(mode=ocr_mode, lang=lang)
        )
    elif path.suffix.lower() == ".docx":
        data = docx_converter.extract_docx(path)
    else:
        console.print(f"[yellow]Skip unsupported file:[/] {path}")
        return

    stem = path.stem
    # Write formats
    if "txt" in formats:
        write_text(out_dir, stem, data.get("text") or "", "txt")
    if "md" in formats:
        md = data.get("markdown")
        if md is None:
            # fallback to text if no markdown available (e.g., PDF)
            md = data.get("text") or ""
        write_text(out_dir, stem, md, "md")
    if "json" in formats:
        write_json(out_dir, stem, data)

    # Tables
    if tables in {"csv", "json"} and data.get("tables"):
        import csv, json as _json

        for t in data["tables"]:
            page = t.get("page")
            idx = t.get("index")
            rows = t.get("rows") or []
            if tables == "csv":
                csv_path = out_dir / f"{stem}.table-{page or 'NA'}-{idx or 0}.csv"
                with csv_path.open("w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerows(rows)
            else:
                json_path = out_dir / f"{stem}.table-{page or 'NA'}-{idx or 0}.json"
                json_path.write_text(
                    _json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8"
                )


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="mrconvert",
        description="Convert PDF/DOCX to machine-readable TXT/MD/JSON or bidirectional PDF↔DOCX conversion.",
    )
    p.add_argument("input", help="File or folder containing .pdf/.docx")
    p.add_argument(
        "--out",
        dest="out",
        default="./mr_out",
        help="Output directory (default: ./mr_out)",
    )

    # Text extraction mode (existing)
    text_group = p.add_argument_group("Text extraction mode")
    text_group.add_argument(
        "--format",
        dest="formats",
        nargs="+",
        choices=["txt", "md", "json"],
        default=["txt"],
        help="Output formats",
    )
    text_group.add_argument(
        "--tables",
        dest="tables",
        choices=["none", "csv", "json"],
        default="csv",
        help="Export detected tables",
    )
    text_group.add_argument(
        "--keep-layout",
        action="store_true",
        help="Preserve approximate layout for PDF text extraction",
    )
    text_group.add_argument(
        "--ocr",
        choices=["off", "auto", "force"],
        default="auto",
        help="OCR policy for PDFs",
    )
    text_group.add_argument(
        "--lang", default=None, help="OCR language (e.g., 'kor+eng')"
    )

    # Bidirectional conversion mode (new)
    convert_group = p.add_argument_group("Bidirectional conversion mode")
    convert_group.add_argument(
        "--to-docx", action="store_true", help="Convert PDF/MD to DOCX"
    )
    convert_group.add_argument(
        "--to-pdf", action="store_true", help="Convert DOCX to PDF"
    )
    convert_group.add_argument(
        "--to-msg", action="store_true", help="Convert MD/DOCX to Outlook MSG (Windows only)"
    )
    convert_group.add_argument(
        "--to-xlsx", action="store_true", help="Convert MD to Excel XLSX format"
    )

    return p


def run(argv: list[str] | None = None) -> int:
    argv = argv or sys.argv[1:]
    args = build_parser().parse_args(argv)

    # Validate arguments
    if args.to_docx and args.to_pdf:
        console.print("[red]Error:[/] Cannot use both --to-docx and --to-pdf")
        raise SystemExit(1)

    if (args.to_docx or args.to_pdf or args.to_msg or args.to_xlsx) and (
        len(args.formats) > 1
        or "md" in args.formats
        or "json" in args.formats
        or args.tables != "csv"
    ):
        console.print(
            "[red]Error:[/] --to-docx/--to-pdf/--to-msg/--to-xlsx cannot be used with --format/--tables options"
        )
        raise SystemExit(1)

    in_path = Path(args.input).expanduser().resolve()
    out_dir = Path(args.out).expanduser().resolve()
    ensure_dir(out_dir)

    files = list(walk_inputs(in_path))
    if not files:
        console.print(f"[red]No files found under:[/] {in_path}")
        return 2

    # Check if we're in bidirectional conversion mode
    if args.to_docx or args.to_pdf or args.to_msg or args.to_xlsx:
        return _run_bidirectional_conversion(files, out_dir, args.to_docx, args.to_pdf, args.to_msg, args.to_xlsx)
    else:
        return _run_text_extraction(files, out_dir, args)


def _run_bidirectional_conversion(
    files: List[Path], out_dir: Path, to_docx: bool, to_pdf: bool, to_msg: bool, to_xlsx: bool
) -> int:
    """Run bidirectional PDF↔DOCX conversion, MD→DOCX, MD/DOCX→MSG, or MD→XLSX"""
    console.print(
        f"[bold]mrconvert[/] · Bidirectional conversion · {len(files)} file(s) → {out_dir}"
    )

    for f in track(files, description="Converting"):
        try:
            if to_docx and f.suffix.lower() == ".pdf":
                dst = out_dir / f"{f.stem}.docx"
                result = bidirectional.pdf_to_docx(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            elif to_docx and f.suffix.lower() == ".md":
                dst = out_dir / f"{f.stem}.docx"
                result = markdown_to_docx.markdown_to_docx(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            elif to_pdf and f.suffix.lower() == ".docx":
                dst = out_dir / f"{f.stem}.pdf"
                result = bidirectional.docx_to_pdf(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            elif to_msg and f.suffix.lower() == ".md":
                dst = out_dir / f"{f.stem}.msg"
                result = docx_to_msg.markdown_to_msg(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            elif to_msg and f.suffix.lower() == ".docx":
                dst = out_dir / f"{f.stem}.msg"
                result = docx_to_msg.docx_to_msg(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            elif to_xlsx and f.suffix.lower() == ".md":
                dst = out_dir / f"{f.stem}.xlsx"
                result = markdown_to_xlsx.markdown_to_xlsx(f, dst)
                console.print(
                    f"[green][{result.engine}][/] {f.name} → {result.output.name}"
                )
            else:
                console.print(
                    f"[yellow]Skip:[/] {f.name} (wrong file type for conversion)"
                )
        except Exception as e:
            console.print(f"[red]Error:[/] {f} — {e}")

    console.print("[green]Done.[/]")
    return 0


def _run_text_extraction(files: List[Path], out_dir: Path, args) -> int:
    """Run text extraction mode (existing functionality)"""
    console.print(
        f"[bold]mrconvert[/] · Text extraction · {len(files)} file(s) → {out_dir}"
    )

    for f in track(files, description="Converting"):
        try:
            _process_file(
                f,
                out_dir,
                formats=args.formats,
                tables=args.tables,
                keep_layout=args.keep_layout,
                ocr_mode=args.ocr,
                lang=args.lang,
            )
        except Exception as e:
            console.print(f"[red]Error:[/] {f} — {e}")

    console.print("[green]Done.[/]")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
