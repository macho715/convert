#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Full export run for 21k rows with progress reporting.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any


def _import_cli_helpers():
    try:
        import export_email_threads_cli as cli
        return cli
    except ImportError:
        scripts_dir = Path(__file__).parent
        sys.path.insert(0, str(scripts_dir))
        import export_email_threads_cli as cli
        return cli


def main() -> int:
    parser = argparse.ArgumentParser(description="Full export run")
    parser.add_argument("--excel", required=True, help="Excel file path")
    parser.add_argument("--sheet", default="전체_데이터", help="Sheet name")
    parser.add_argument("--out", required=True, help="Output directory")
    parser.add_argument("--query", default="", help="Search query (optional)")
    parser.add_argument("--max-results", type=int, default=200, help="Max search rows")
    parser.add_argument("--tz", default="Asia/Dubai", help="Timezone")
    parser.add_argument("--lookback-k", type=int, default=50, help="Edges lookback")
    parser.add_argument("--window-days", type=int, default=14, help="Edges window days")
    parser.add_argument("--parent-min-conf", type=float, default=0.35, help="Min parent confidence")
    parser.add_argument(
        "--flag-below-threshold",
        action="store_true",
        help="Add below_threshold flag to edges.csv",
    )
    parser.add_argument(
        "--filter-below-threshold",
        action="store_true",
        help="Exclude edges below parent_min_conf",
    )
    parser.add_argument(
        "--assume-local-time",
        action="store_true",
        help="Assume DeliveryTime is already local time",
    )

    args = parser.parse_args()

    cli = _import_cli_helpers()
    excel_path = Path(args.excel).expanduser().resolve()
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("Full export run")
    print("=" * 60)
    print(f"Input file : {excel_path}")
    print(f"Sheet      : {args.sheet}")
    print(f"Out dir    : {out_dir}")
    print()

    print("[1/4] Loading Excel...")
    start_load = time.time()
    df = cli._read_excel(excel_path, args.sheet, max_rows=None)
    cli._validate_columns(df)
    elapsed_load = time.time() - start_load
    print(f"  Done: {len(df):,} rows in {elapsed_load:.2f}s")
    print()

    print("[2/4] Initializing thread tracker...")
    start_init = time.time()
    EmailThreadTrackerV3 = cli._import_tracker()
    tracker = EmailThreadTrackerV3(df)
    elapsed_init = time.time() - start_init
    print(f"  Done: {len(tracker.thread_meta):,} threads in {elapsed_init:.2f}s")
    print(f"  Speed: {len(df) / max(elapsed_init, 0.01):.0f} rows/s")
    print()

    print("[3/4] Exporting outputs...")
    start_export = time.time()

    threads_json_path = out_dir / "threads.json"
    cli._export_threads_json(tracker, threads_json_path, tz=args.tz)

    edges_csv_path = out_dir / "edges.csv"
    edges_df = cli._build_edges(
        tracker,
        tz=args.tz,
        lookback_k=args.lookback_k,
        window_days=args.window_days,
        parent_min_conf=args.parent_min_conf,
        flag_below_threshold=args.flag_below_threshold,
        filter_below_threshold=args.filter_below_threshold,
        assume_local_time=args.assume_local_time,
    )
    edges_df.to_csv(edges_csv_path, index=False, encoding="utf-8-sig")

    search_csv_path = out_dir / "search_result.csv"
    n_rows, context = cli._export_search_result_csv(
        tracker,
        args.query.strip(),
        search_csv_path,
        max_results=args.max_results,
        tz=args.tz,
        assume_local_time=args.assume_local_time,
    )

    elapsed_export = time.time() - start_export
    print(f"  Done in {elapsed_export:.2f}s")
    print()

    print("[4/4] Writing report...")
    total_elapsed = time.time() - start_load

    thread_sizes = [len(meta.members) for meta in tracker.thread_meta.values()]
    confidences = [meta.confidence for meta in tracker.thread_meta.values()]

    report: dict[str, Any] = {
        "task": "full_export",
        "input_file": str(excel_path),
        "sheet": args.sheet,
        "rows_processed": int(len(df)),
        "threads_found": int(len(tracker.thread_meta)),
        "edges_found": int(len(edges_df)),
        "search_results": int(n_rows),
        "elapsed_seconds": {
            "load": round(elapsed_load, 2),
            "init": round(elapsed_init, 2),
            "export": round(elapsed_export, 2),
            "total": round(total_elapsed, 2),
        },
        "performance": {
            "rows_per_second": round(len(df) / max(elapsed_init, 0.01), 0),
            "threads_per_second": round(len(tracker.thread_meta) / max(elapsed_init, 0.01), 2),
        },
        "thread_stats": {
            "total": int(len(tracker.thread_meta)),
            "with_multiple": int(sum(1 for s in thread_sizes if s > 1)),
            "size_min": int(min(thread_sizes)) if thread_sizes else 0,
            "size_max": int(max(thread_sizes)) if thread_sizes else 0,
            "size_avg": round(sum(thread_sizes) / len(thread_sizes), 2) if thread_sizes else 0,
            "confidence_min": round(min(confidences), 3) if confidences else 0,
            "confidence_max": round(max(confidences), 3) if confidences else 0,
            "confidence_avg": round(sum(confidences) / len(confidences), 3) if confidences else 0,
            "low_confidence_count": int(sum(1 for c in confidences if c < 0.60)),
        },
        "output_files": {
            "threads_json": str(threads_json_path),
            "edges_csv": str(edges_csv_path),
            "search_result_csv": str(search_csv_path),
        },
        "options": {
            "flag_below_threshold": args.flag_below_threshold,
            "filter_below_threshold": args.filter_below_threshold,
            "assume_local_time": args.assume_local_time,
        },
        "search_context": context if args.query else None,
        "generated_at": datetime.now().isoformat(),
    }

    report_path = out_dir / "_run_report_full.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    print("=" * 60)
    print("Summary")
    print("=" * 60)
    print(f"Rows processed : {report['rows_processed']:,}")
    print(f"Threads        : {report['threads_found']:,}")
    print(f"Avg size       : {report['thread_stats']['size_avg']:.1f}")
    print(f"Avg confidence : {report['thread_stats']['confidence_avg']:.3f}")
    print(f"Edges          : {report['edges_found']:,}")
    print(f"Total time     : {report['elapsed_seconds']['total']:.2f}s")
    print()
    print("Outputs:")
    print(f"  {threads_json_path}")
    print(f"  {edges_csv_path}")
    print(f"  {search_csv_path}")
    print(f"  {report_path}")

    if args.query:
        print()
        print("Search context:")
        print(f"  direct matches : {context.get('total_found', 0)}")
        print(f"  with context   : {context.get('total_with_context', 0)}")
        print(f"  threads        : {context.get('threads_included', 0)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
