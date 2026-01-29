#!/usr/bin/env python3
"""
docxtpl(Jinja2) 기반 .docx 템플릿 렌더링.
템플릿에 {{ var }}, {% for %} 등 Jinja2 문법이 있을 때 사용.
"""

import argparse
import json
from pathlib import Path

from docxtpl import DocxTemplate


def fill_docx_docxtpl(template_path: str, context_path: str, out_path: str) -> None:
    doc = DocxTemplate(template_path)
    with open(context_path, "r", encoding="utf-8") as f:
        context = json.load(f)
    doc.render(context)
    doc.save(out_path)
    print("Saved:", out_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="docxtpl: Jinja2 템플릿으로 .docx 생성"
    )
    parser.add_argument("--template", "-t", required=True, help=".docx 템플릿 경로")
    parser.add_argument("--context", "-c", required=True, help="context JSON 경로")
    parser.add_argument("--out", "-o", required=True, help="출력 .docx 경로")
    args = parser.parse_args()
    fill_docx_docxtpl(args.template, args.context, args.out)
