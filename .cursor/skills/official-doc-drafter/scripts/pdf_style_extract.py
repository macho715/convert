#!/usr/bin/env python3
"""
PyMuPDF 기반 PDF 스타일 분석
"""

import argparse
import json
import sys
from collections import Counter

import fitz  # PyMuPDF


def pts_to_mm(val: float) -> float:
    return round(val * 0.352778, 2)


def analyze(pdf_path: str, out_json: str) -> None:
    doc = fitz.open(pdf_path)
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    data = page.get_text("dict")
    spans = []
    for block in data.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue
                spans.append(span)
    if not spans:
        print("No text found", file=sys.stderr)
        return

    sizes = [round(s["size"], 1) for s in spans]
    fonts = [s["font"] for s in spans]
    size_counts = Counter(sizes)
    font_counts = Counter(fonts)
    body_size = size_counts.most_common(1)[0][0]
    body_font = font_counts.most_common(1)[0][0]
    x0s = [s["bbox"][0] for s in spans if round(s["size"], 1) == body_size]
    x1s = [s["bbox"][2] for s in spans if round(s["size"], 1) == body_size]
    left_margin = min(x0s) if x0s else 0
    right_margin = w - max(x1s) if x1s else 0

    profile = {
        "page_size_mm": {"width": pts_to_mm(w), "height": pts_to_mm(h)},
        "margins_mm": {
            "left": pts_to_mm(left_margin),
            "right": pts_to_mm(right_margin),
        },
        "body_font": {"name": body_font, "size_pt": body_size},
        "font_counts": font_counts.most_common(10),
        "size_counts": size_counts.most_common(10),
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", "-p", required=True)
    parser.add_argument("--out-json", "-o", required=True)
    args = parser.parse_args()
    analyze(args.pdf, args.out_json)
