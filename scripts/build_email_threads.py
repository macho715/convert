#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build derived fields and thread metadata for Outlook Excel exports.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

import pandas as pd

from email_thread_tracker_v2_enhanced import EmailThreadTrackerV2Enhanced


def write_report(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build email threads with derived fields")
    parser.add_argument("excel_file", help="Input Excel file")
    parser.add_argument("--sheet", help="Sheet name (default: first sheet)")
    parser.add_argument("--out", help="Output .xlsx or .csv path")
    parser.add_argument("--report", help="Run report JSON path")
    parser.add_argument("--max-rows", type=int, help="Limit rows for quick runs")
    args = parser.parse_args()

    start = datetime.now()
    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    sheet = args.sheet
    if not sheet:
        sheet = pd.ExcelFile(excel_path).sheet_names[0]

    df = pd.read_excel(excel_path, sheet_name=sheet)
    if args.max_rows:
        df = df.head(args.max_rows)

    tracker = EmailThreadTrackerV2Enhanced(df)
    result_df = tracker.df

    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        if out_path.suffix.lower() == ".csv":
            result_df.to_csv(out_path, index=False)
        else:
            result_df.to_excel(out_path, index=False, engine="openpyxl")

    elapsed = (datetime.now() - start).total_seconds()
    report = {
        "task": "build_email_threads",
        "input_file": str(excel_path.resolve()),
        "sheet": sheet,
        "rows_in": int(len(df)),
        "rows_out": int(len(result_df)),
        "threads": int(len(tracker.threads)),
        "elapsed_seconds": round(elapsed, 3),
        "generated_at": datetime.now().isoformat(),
        "output_file": str(Path(args.out).resolve()) if args.out else None,
    }
    if args.report:
        write_report(Path(args.report), report)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
