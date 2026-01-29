#!/usr/bin/env python3
"""
DOCX 템플릿 스타일 요약 JSON 출력
"""

import argparse
import json
from docx import Document
from docx.shared import Mm


def _length_mm(val):
    """Length 객체를 mm(float)로 변환"""
    if val is None:
        return None
    if hasattr(val, "mm"):
        return round(val.mm, 2)
    if hasattr(val, "inches"):
        return round(val.inches * 25.4, 2)
    try:
        return round(int(val) / 914400 * 25.4, 2)
    except (TypeError, ValueError):
        return None


def inspect(template_path: str, out_json: str) -> None:
    doc = Document(template_path)
    sec = doc.sections[0]
    summary = {
        "page_width_mm": _length_mm(sec.page_width),
        "page_height_mm": _length_mm(sec.page_height),
        "margins_mm": {
            "top": _length_mm(sec.top_margin),
            "bottom": _length_mm(sec.bottom_margin),
            "left": _length_mm(sec.left_margin),
            "right": _length_mm(sec.right_margin),
        },
        "default_font": {
            "name": doc.styles["Normal"].font.name,
            "size_pt": (
                doc.styles["Normal"].font.size.pt
                if doc.styles["Normal"].font.size
                else None
            ),
        },
        "first_paragraph_style": [
            {"text": par.text[:200], "style": par.style.name}
            for par in doc.paragraphs[:5]
        ],
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", "-t", required=True)
    parser.add_argument("--out-json", "-o", required=True)
    args = parser.parse_args()
    inspect(args.template, args.out_json)
