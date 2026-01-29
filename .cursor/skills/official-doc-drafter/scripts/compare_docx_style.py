#!/usr/bin/env python3
"""
Compare two DOCX files (template vs candidate) using heuristic profiles and
write a Markdown + JSON report with PASS/WARN/FAIL.

Usage:
  python compare_docx_style.py --template TEMPLATE.docx --candidate OUTPUT.docx \
    --out-md style-report.md --out-json style-report.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

# Allow importing docx_style_profile when run from project root
_scripts_dir = Path(__file__).resolve().parent
if str(_scripts_dir) not in sys.path:
    sys.path.insert(0, str(_scripts_dir))
from docx_style_profile import profile_docx


@dataclass
class Finding:
    level: str  # PASS/WARN/FAIL
    title: str
    detail: str
    suggestion: Optional[str] = None


def _absdiff(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return abs(a - b)


def _cmp_mm(name: str, a: Optional[float], b: Optional[float], tol: float) -> Finding:
    d = _absdiff(a, b)
    if d is None:
        return Finding("WARN", name, f"Cannot compare (missing value). a={a}, b={b}")
    if d <= tol:
        return Finding("PASS", name, f"Within tolerance (±{tol}mm). diff={d:.2f}mm")
    return Finding(
        "FAIL",
        name,
        f"Out of tolerance (±{tol}mm). template={a:.2f}mm candidate={b:.2f}mm diff={d:.2f}mm",
        suggestion="Adjust section margins/page setup in the generated document.",
    )


def _cmp_pt(name: str, a: Optional[float], b: Optional[float], tol: float) -> Finding:
    d = _absdiff(a, b)
    if d is None:
        return Finding("WARN", name, f"Cannot compare (missing value). a={a}, b={b}")
    if d <= tol:
        return Finding("PASS", name, f"Within tolerance (±{tol}pt). diff={d:.2f}pt")
    return Finding(
        "FAIL",
        name,
        f"Out of tolerance (±{tol}pt). template={a:.1f}pt candidate={b:.1f}pt diff={d:.2f}pt",
        suggestion="Ensure Normal style + run fonts/sizes inherit correctly from the template.",
    )


def _cmp_str(name: str, a: Optional[str], b: Optional[str]) -> Finding:
    if not a or not b:
        return Finding("WARN", name, f"Cannot compare (missing value). a={a}, b={b}")
    if a == b:
        return Finding("PASS", name, f"Match: {a}")
    return Finding(
        "WARN",
        name,
        f"Mismatch. template={a} candidate={b}",
        suggestion="If the template embeds fonts or uses localized font names, map/alias fonts explicitly.",
    )


def _find_level(findings: List[Finding]) -> str:
    if any(f.level == "FAIL" for f in findings):
        return "FAIL"
    if any(f.level == "WARN" for f in findings):
        return "WARN"
    return "PASS"


def _md_escape(s: str) -> str:
    return s.replace("\n", "<br>")


def compare(
    template_path: str,
    candidate_path: str,
    tol_margin_mm: float,
    tol_page_mm: float,
    tol_font_pt: float,
) -> Dict[str, Any]:
    t = profile_docx(template_path)
    c = profile_docx(candidate_path)

    findings: List[Finding] = []

    # section 0 only (practical baseline)
    t0 = (t.get("sections") or [{}])[0]
    c0 = (c.get("sections") or [{}])[0]

    # page size
    findings.append(
        _cmp_mm(
            "Page width", t0.get("page_width_mm"), c0.get("page_width_mm"), tol_page_mm
        )
    )
    findings.append(
        _cmp_mm(
            "Page height",
            t0.get("page_height_mm"),
            c0.get("page_height_mm"),
            tol_page_mm,
        )
    )

    # margins
    tm = t0.get("margins_mm") or {}
    cm = c0.get("margins_mm") or {}
    findings.append(_cmp_mm("Margin top", tm.get("top"), cm.get("top"), tol_margin_mm))
    findings.append(
        _cmp_mm("Margin bottom", tm.get("bottom"), cm.get("bottom"), tol_margin_mm)
    )
    findings.append(
        _cmp_mm("Margin left", tm.get("left"), cm.get("left"), tol_margin_mm)
    )
    findings.append(
        _cmp_mm("Margin right", tm.get("right"), cm.get("right"), tol_margin_mm)
    )

    # typography
    findings.append(
        _cmp_str("Dominant font", t.get("dominant_font"), c.get("dominant_font"))
    )
    findings.append(
        _cmp_pt(
            "Dominant font size",
            t.get("dominant_font_size_pt"),
            c.get("dominant_font_size_pt"),
            tol_font_pt,
        )
    )

    # Normal style (soft check)
    nf_t = t.get("normal_font") or {}
    nf_c = c.get("normal_font") or {}
    findings.append(_cmp_str("Normal style font", nf_t.get("name"), nf_c.get("name")))
    findings.append(
        _cmp_pt(
            "Normal style font size",
            nf_t.get("size_pt"),
            nf_c.get("size_pt"),
            tol_font_pt,
        )
    )

    # placeholders
    out_ph = {k: v for k, v in (c.get("placeholders") or [])}
    out_ph_total = sum(out_ph.values())

    if out_ph_total == 0:
        findings.append(
            Finding(
                "PASS",
                "Placeholders remaining",
                "No {{TOKEN}} placeholders found in candidate.",
            )
        )
    else:
        top = sorted(out_ph.items(), key=lambda kv: kv[1], reverse=True)[:10]
        detail = "Remaining placeholders detected: " + ", ".join(
            [f"{k}×{v}" for k, v in top]
        )
        findings.append(
            Finding(
                "FAIL",
                "Placeholders remaining",
                detail,
                suggestion="Ensure fill/replace step covers headers, tables, and all paragraphs; then re-run verification.",
            )
        )

    verdict = _find_level(findings)

    return {
        "verdict": verdict,
        "template": t,
        "candidate": c,
        "tolerances": {
            "page_mm": tol_page_mm,
            "margin_mm": tol_margin_mm,
            "font_pt": tol_font_pt,
        },
        "findings": [f.__dict__ for f in findings],
    }


def write_md(report: Dict[str, Any], out_md: str):
    verdict = report["verdict"]
    findings = report["findings"]

    lines = []
    lines.append("# DOCX Style Verification Report")
    lines.append("")
    lines.append(f"## Verdict: **{verdict}**")
    lines.append("")
    lines.append("## Findings")
    lines.append("")
    lines.append("| Level | Check | Detail | Suggestion |")
    lines.append("|---|---|---|---|")
    for f in findings:
        lines.append(
            f"| {f['level']} | {_md_escape(f['title'])} | {_md_escape(f['detail'])} | {_md_escape(f.get('suggestion') or '')} |"
        )

    lines.append("")
    lines.append("## Notes")
    lines.append(
        "- This is heuristic (python-docx observable properties). For text boxes/shapes/anchored objects, do a manual visual QA if needed."
    )

    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, help="Template .docx")
    ap.add_argument("--candidate", required=True, help="Generated .docx")
    ap.add_argument("--out-md", required=True, help="Output Markdown report path")
    ap.add_argument("--out-json", required=True, help="Output JSON report path")
    ap.add_argument("--tol-margin-mm", type=float, default=2.0)
    ap.add_argument("--tol-page-mm", type=float, default=1.0)
    ap.add_argument("--tol-font-pt", type=float, default=0.5)
    args = ap.parse_args()

    report = compare(
        template_path=args.template,
        candidate_path=args.candidate,
        tol_margin_mm=args.tol_margin_mm,
        tol_page_mm=args.tol_page_mm,
        tol_font_pt=args.tol_font_pt,
    )

    write_md(report, args.out_md)
    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print(f"Wrote: {args.out_md}")
    print(f"Wrote: {args.out_json}")


if __name__ == "__main__":
    main()
