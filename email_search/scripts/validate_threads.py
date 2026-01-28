#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validate threading outputs (sampling-based).
"""

from __future__ import annotations

import argparse
import json
import random
from pathlib import Path
from typing import Dict, List

import pandas as pd


def validate_threads(threads_json_path: Path, sample_size: int = 50) -> Dict:
    threads = json.loads(threads_json_path.read_text(encoding="utf-8"))
    sample = random.sample(threads, min(sample_size, len(threads)))

    issues: List[Dict] = []
    stats = {
        "total_threads": len(threads),
        "sampled": len(sample),
        "issues_found": 0,
    }

    for thread in sample:
        thread_id = thread.get("thread_id", "")
        members = thread.get("members", [])
        confidence = thread.get("confidence", 0.0)
        subject_norm = thread.get("subject_norm", "")

        if len(members) == 0:
            issues.append(
                {"thread_id": thread_id, "issue": "empty_thread", "severity": "high"}
            )

        if len(members) > 100:
            issues.append(
                {
                    "thread_id": thread_id,
                    "issue": "oversized_thread",
                    "size": len(members),
                    "severity": "medium",
                }
            )

        if confidence < 0.30:
            issues.append(
                {
                    "thread_id": thread_id,
                    "issue": "low_confidence",
                    "confidence": confidence,
                    "severity": "medium",
                }
            )

        if not subject_norm and len(members) > 1:
            issues.append(
                {
                    "thread_id": thread_id,
                    "issue": "missing_subject_norm",
                    "severity": "low",
                }
            )

    stats["issues_found"] = len(issues)
    stats["issues"] = issues[:20]
    return stats


def validate_edges(edges_csv_path: Path) -> Dict:
    df = pd.read_csv(edges_csv_path)

    issues: List[Dict] = []
    duplicates = df.duplicated(subset=["parent_row", "child_row"])
    if duplicates.any():
        issues.append(
            {
                "issue": "duplicate_edges",
                "count": int(duplicates.sum()),
                "severity": "medium",
            }
        )

    parent_set = set(df["parent_row"])
    child_set = set(df["child_row"])
    cycles = parent_set & child_set
    if cycles:
        issues.append(
            {"issue": "potential_cycles", "count": len(cycles), "severity": "low"}
        )

    conf_stats = df["confidence"].describe()

    return {
        "total_edges": int(len(df)),
        "issues_found": len(issues),
        "issues": issues,
        "confidence_stats": {
            "min": float(conf_stats["min"]),
            "max": float(conf_stats["max"]),
            "mean": float(conf_stats["mean"]),
            "median": float(df["confidence"].median()),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate threading outputs")
    parser.add_argument("--threads-json", required=True, help="threads.json path")
    parser.add_argument("--edges-csv", required=True, help="edges.csv path")
    parser.add_argument("--sample-size", type=int, default=50, help="Sample size")
    parser.add_argument("--out", help="Validation report JSON path")

    args = parser.parse_args()

    threads_path = Path(args.threads_json)
    edges_path = Path(args.edges_csv)

    print("=" * 60)
    print("Validation")
    print("=" * 60)

    print("\n[1/2] Threads validation...")
    thread_stats = validate_threads(threads_path, args.sample_size)
    print(f"  Total threads: {thread_stats['total_threads']:,}")
    print(f"  Sampled      : {thread_stats['sampled']}")
    print(f"  Issues       : {thread_stats['issues_found']}")

    print("\n[2/2] Edges validation...")
    edge_stats = validate_edges(edges_path)
    print(f"  Total edges : {edge_stats['total_edges']:,}")
    print(f"  Issues      : {edge_stats['issues_found']}")
    print(
        "  Confidence  : "
        f"{edge_stats['confidence_stats']['min']:.3f} ~ "
        f"{edge_stats['confidence_stats']['max']:.3f} "
        f"(mean {edge_stats['confidence_stats']['mean']:.3f})"
    )

    if args.out:
        report = {
            "threads_validation": thread_stats,
            "edges_validation": edge_stats,
        }
        Path(args.out).write_text(
            json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print(f"\n[Export] Report saved: {args.out}")

    print("\n" + "=" * 60)
    print("Validation complete")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
