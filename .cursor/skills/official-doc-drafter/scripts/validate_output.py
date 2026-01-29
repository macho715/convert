#!/usr/bin/env python3
"""
생성된 문서가 스타일 요약과 일치하는지 검사
"""

import argparse
import json
import sys

from inspect_docx_template import inspect


def validate(docx_path: str, expected_style_json: str) -> list[str]:
    summary_json = docx_path + ".summary.json"
    inspect(docx_path, summary_json)
    with open(summary_json, encoding="utf-8") as f:
        summary = json.load(f)
    with open(expected_style_json, encoding="utf-8") as f:
        expected = json.load(f)
    issues = []
    em = summary.get("margins_mm", {})
    ex = expected.get("margins_mm", {})
    for side in ("left", "right"):
        s_val = em.get(side)
        e_val = ex.get(side)
        if s_val is not None and e_val is not None:
            if abs(float(s_val) - float(e_val)) > 2:
                issues.append(
                    f"Margin diff {side}: got {s_val} mm, expected ~{e_val} mm"
                )
    return issues


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--docx", "-d", required=True)
    p.add_argument("--expected-style", "-s", required=True)
    args = p.parse_args()
    errs = validate(args.docx, args.expected_style)
    if errs:
        for e in errs:
            print("Issue:", e, file=sys.stderr)
        sys.exit(1)
    print("OK")
