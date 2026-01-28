#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export:
  (1) threads.json
  (2) edges.csv
  (3) search_result.csv
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


def _import_tracker() -> type:
    try:
        from outlook_thread_tracker_v3 import EmailThreadTrackerV3

        return EmailThreadTrackerV3
    except ImportError:
        scripts_dir = Path(__file__).parent
        if (scripts_dir / "outlook_thread_tracker_v3.py").exists():
            sys.path.insert(0, str(scripts_dir))
            from outlook_thread_tracker_v3 import EmailThreadTrackerV3

            return EmailThreadTrackerV3
        parent_dir = scripts_dir.parent
        sys.path.insert(0, str(parent_dir))
        try:
            from outlook_thread_tracker_v3 import EmailThreadTrackerV3

            return EmailThreadTrackerV3
        except ImportError as exc:
            raise SystemExit(
                "Error: outlook_thread_tracker_v3.py not found.\n"
                f"Checked: {scripts_dir / 'outlook_thread_tracker_v3.py'}"
            ) from exc


def _read_excel(excel_path: Path, sheet: str, max_rows: Optional[int]) -> pd.DataFrame:
    if max_rows:
        return pd.read_excel(excel_path, sheet_name=sheet, nrows=max_rows)
    return pd.read_excel(excel_path, sheet_name=sheet)


def _validate_columns(df: pd.DataFrame) -> None:
    required = ["Subject", "DeliveryTime"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")


def _safe_tz_convert(dt: pd.Timestamp, tz: str, assume_local: bool = False) -> str:
    if dt is None or pd.isna(dt):
        return ""
    if dt.tzinfo is None:
        if assume_local:
            dt = dt.tz_localize(tz)
        else:
            dt = dt.tz_localize("UTC")
    try:
        return dt.tz_convert(tz).isoformat()
    except Exception:
        return dt.isoformat()


def _export_threads_json(tracker, output_path: Path, tz: str) -> None:
    threads = tracker.export_threads()
    for item in threads:
        if item.get("start_dt"):
            item["start_dt"] = _safe_tz_convert(pd.to_datetime(item["start_dt"], utc=True), tz)
        if item.get("end_dt"):
            item["end_dt"] = _safe_tz_convert(pd.to_datetime(item["end_dt"], utc=True), tz)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(threads, indent=2, ensure_ascii=False), encoding="utf-8")


def _build_edges(
    tracker,
    tz: str = "Asia/Dubai",
    lookback_k: int = 50,
    window_days: int = 14,
    parent_min_conf: float = 0.35,
    flag_below_threshold: bool = True,
    filter_below_threshold: bool = False,
    assume_local_time: bool = False,
) -> pd.DataFrame:
    rows: List[Dict] = []
    df = tracker.df

    for tid, meta in tracker.thread_meta.items():
        members = list(meta.members)
        if len(members) <= 1:
            continue

        temp = df.loc[members].copy()
        if assume_local_time:
            temp["_dt"] = pd.to_datetime(temp.get("DeliveryTime", None), errors="coerce")
        else:
            temp["_dt"] = pd.to_datetime(temp.get("DeliveryTime", None), errors="coerce", utc=True)
        temp = temp.sort_values(["_dt"], ascending=True)
        ordered = list(temp.index)

        for pos in range(1, len(ordered)):
            child = ordered[pos]
            if assume_local_time:
                child_dt = pd.to_datetime(
                    df.loc[child].get("DeliveryTime", None), errors="coerce"
                )
            else:
                child_dt = pd.to_datetime(
                    df.loc[child].get("DeliveryTime", None), errors="coerce", utc=True
                )

            start = max(0, pos - lookback_k)
            candidates = ordered[start:pos]

            best_parent = candidates[-1]
            best_conf = 0.0

            for p in reversed(candidates):
                if assume_local_time:
                    p_dt = pd.to_datetime(
                        df.loc[p].get("DeliveryTime", None), errors="coerce"
                    )
                else:
                    p_dt = pd.to_datetime(
                        df.loc[p].get("DeliveryTime", None), errors="coerce", utc=True
                    )
                if pd.notna(child_dt) and pd.notna(p_dt):
                    if abs((child_dt - p_dt).days) > window_days:
                        continue

                try:
                    if hasattr(tracker, "get_pair_confidence"):
                        conf = float(tracker.get_pair_confidence(p, child))
                    else:
                        conf = float(tracker._pair_confidence(p, child))
                except Exception:
                    conf = 0.0

                if conf >= best_conf:
                    best_conf = conf
                    best_parent = p

            below_threshold = best_conf < parent_min_conf
            if below_threshold:
                best_parent = candidates[-1]
                try:
                    if hasattr(tracker, "get_pair_confidence"):
                        best_conf = float(tracker.get_pair_confidence(best_parent, child))
                    else:
                        best_conf = float(tracker._pair_confidence(best_parent, child))
                except Exception:
                    best_conf = 0.0
            if filter_below_threshold and below_threshold:
                continue

            rows.append(
                {
                    "thread_id": tid,
                    "relation_type": "heuristic",
                    "confidence": round(best_conf, 2),
                    "parent_row": int(best_parent),
                    "child_row": int(child),
                    "parent_no": str(df.loc[best_parent].get("no", "")),
                    "child_no": str(df.loc[child].get("no", "")),
                    "parent_delivery_time": _safe_tz_convert(
                        pd.to_datetime(
                            df.loc[best_parent].get("DeliveryTime", None),
                            errors="coerce",
                            utc=not assume_local_time,
                        ),
                        tz,
                        assume_local=assume_local_time,
                    ),
                    "child_delivery_time": _safe_tz_convert(
                        pd.to_datetime(
                            df.loc[child].get("DeliveryTime", None),
                            errors="coerce",
                            utc=not assume_local_time,
                        ),
                        tz,
                        assume_local=assume_local_time,
                    ),
                    "subject_norm": str(df.loc[child].get("_subject_norm", "")),
                }
            )
            if flag_below_threshold:
                rows[-1]["below_threshold"] = below_threshold

    return pd.DataFrame(rows)


def _export_search_result_csv(
    tracker,
    query: str,
    output_path: Path,
    max_results: int,
    tz: str,
    assume_local_time: bool = False,
) -> Tuple[int, Dict]:
    if not query:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_csv(output_path, index=False, encoding="utf-8-sig")
        return 0, {}

    results, context = tracker.search_with_context(query, max_results=max_results)
    if "DeliveryTime" in results.columns:
        if assume_local_time:
            results["DeliveryTime"] = pd.to_datetime(results["DeliveryTime"], errors="coerce")
        else:
            results["DeliveryTime"] = pd.to_datetime(results["DeliveryTime"], errors="coerce", utc=True)
        results["DeliveryTime"] = results["DeliveryTime"].apply(
            lambda dt: _safe_tz_convert(dt, tz, assume_local=assume_local_time)
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    results.to_csv(output_path, index=False, encoding="utf-8-sig")
    return int(len(results)), context


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export email threads (threads.json), edges (edges.csv), and search_result.csv",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--excel", required=True, help="Input Excel file")
    parser.add_argument("--sheet", default="전체_데이터", help="Sheet name")
    parser.add_argument("--out", default="out_email_threads", help="Output directory")
    parser.add_argument("--tz", default="Asia/Dubai", help="Timezone for output")
    parser.add_argument("--query", default="", help="Search query")
    parser.add_argument("--max-results", type=int, default=200, help="Max search rows")
    parser.add_argument("--max-rows", type=int, help="Limit rows for quick runs")
    parser.add_argument("--lookback-k", type=int, default=50, help="Lookback window size")
    parser.add_argument("--window-days", type=int, default=14, help="Parent time window (days)")
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

    try:
        excel_path = Path(args.excel).expanduser().resolve()
        if not excel_path.exists():
            print(f"Error: Excel file not found: {excel_path}", file=sys.stderr)
            return 1

        out_dir = Path(args.out).expanduser().resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        print(f"[Load] {excel_path}")
        df = _read_excel(excel_path, args.sheet, args.max_rows)
        _validate_columns(df)
        print(f"[Load] rows={len(df)}")

        EmailThreadTrackerV3 = _import_tracker()
        print("[Init] Thread tracker...")
        tracker = EmailThreadTrackerV3(df)
        print(f"[Init] threads={len(tracker.thread_meta)}")

        threads_json_path = out_dir / "threads.json"
        _export_threads_json(tracker, threads_json_path, tz=args.tz)
        print(f"[Export] {threads_json_path}")

        edges_csv_path = out_dir / "edges.csv"
        edges_df = _build_edges(
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
        print(f"[Export] {edges_csv_path} (edges={len(edges_df)})")

        search_csv_path = out_dir / "search_result.csv"
        n_rows, context = _export_search_result_csv(
            tracker,
            args.query.strip(),
            search_csv_path,
            max_results=args.max_results,
            tz=args.tz,
            assume_local_time=args.assume_local_time,
        )
        print(f"[Export] {search_csv_path} (rows={n_rows})")

        if args.query:
            print("\n[Search Context]")
            print(json.dumps(context, ensure_ascii=False, indent=2))

        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
