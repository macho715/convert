"""
Weather folder PDF/JPG parser for AGI schedule daily update.
Sets TESSDATA_PREFIX to CONVERT/out/tessdata so JPG OCR works without system tessdata.
Usage: python scripts/weather_parse.py [weather_folder]
       python scripts/weather_parse.py "AGI TR 1-6 Transportation Master Gantt Chart/weather/20260128"
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# CONVERT root = parent of scripts/
_CONVERT_ROOT = Path(__file__).resolve().parent.parent
_TESSDATA = _CONVERT_ROOT / "out" / "tessdata"
if (_TESSDATA / "eng.traineddata").exists():
    os.environ["TESSDATA_PREFIX"] = str(_TESSDATA)

import pdfplumber

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None


def parse_pdf(path: Path) -> tuple[bool, str]:
    """Extract text from PDF. Returns (success, text)."""
    try:
        parts = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return True, "\n".join(parts) if parts else ""
    except Exception as e:
        return False, str(e)


def parse_image(path: Path) -> tuple[bool, str]:
    """OCR image (JPG/PNG). Returns (success, text)."""
    if pytesseract is None or Image is None:
        return False, "pytesseract or PIL not installed"
    try:
        img = Image.open(path)
        text = pytesseract.image_to_string(img)
        return True, text or ""
    except Exception as e:
        return False, str(e)


def parse_weather_folder(folder: Path, out_dir: Path | None = None) -> dict[str, dict]:
    """Parse all PDF and image files in folder. Returns {filename: {ok, len, text, sample}}.
    If out_dir is set, writes full text to out_dir/<stem>.txt per file."""
    results = {}
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        suf = f.suffix.lower()
        if suf == ".pdf":
            ok, text = parse_pdf(f)
        elif suf in (".jpg", ".jpeg", ".png"):
            ok, text = parse_image(f)
        else:
            continue
        text = text or ""
        stem = f.stem
        if out_dir and (ok or text):
            out_dir.mkdir(parents=True, exist_ok=True)
            (out_dir / f"{stem}.txt").write_text(text, encoding="utf-8")
        results[f.name] = {
            "ok": ok,
            "len": len(text),
            "text": text,
            "sample": (text[:300].strip() if text else "")[:300],
        }
    return results


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out_dir = None
    for a in sys.argv[1:]:
        if a == "--out" and sys.argv.index(a) + 1 < len(sys.argv):
            out_dir = Path(sys.argv[sys.argv.index(a) + 1])
            break
    folder = Path(args[0]) if args else (
        _CONVERT_ROOT / "AGI TR 1-6 Transportation Master Gantt Chart" / "weather" / "20260128"
    )
    if not folder.is_dir():
        print(f"Not a directory: {folder}", file=sys.stderr)
        return 1
    if out_dir:
        out_dir = out_dir if out_dir.is_absolute() else (_CONVERT_ROOT / out_dir)
    print(f"TESSDATA_PREFIX={os.environ.get('TESSDATA_PREFIX', '(not set)')}")
    print(f"Parsing: {folder}")
    if out_dir:
        print(f"Output: {out_dir}")
    results = parse_weather_folder(folder, out_dir=out_dir)
    for name, r in results.items():
        status = "OK" if r["ok"] else "FAIL"
        print(f"  {name}: {status} (len={r['len']})")
        if r.get("sample"):
            sample = r["sample"].replace("\r\n", " ").replace("\n", " ")[:120]
            safe = "".join(c if ord(c) < 128 else "?" for c in sample)
            print(f"    sample: {safe}...")
    return 0 if all(r["ok"] for r in results.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
