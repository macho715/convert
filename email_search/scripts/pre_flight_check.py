#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pre-flight checks before full run.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd


def check_dependencies() -> bool:
    try:
        import pandas  # noqa: F401
        import openpyxl  # noqa: F401
        print("OK: pandas, openpyxl installed")
        return True
    except ImportError as exc:
        print(f"FAIL: Missing dependency: {exc}")
        return False


def check_input_file(excel_path: Path) -> bool:
    if not excel_path.exists():
        print(f"FAIL: Excel file missing: {excel_path}")
        return False
    try:
        xl = pd.ExcelFile(excel_path)
        print(f"OK: Excel file OK ({len(xl.sheet_names)} sheets)")
        return True
    except Exception as exc:
        print(f"FAIL: Excel read failed: {exc}")
        return False


def check_disk_space(excel_path: Path, required_gb: float = 0.5) -> bool:
    import shutil

    stat = shutil.disk_usage(excel_path.parent)
    free_gb = stat.free / (1024**3)
    if free_gb < required_gb:
        print(f"FAIL: Low disk space: {free_gb:.2f}GB (need {required_gb}GB)")
        return False
    print(f"OK: Disk space ({free_gb:.2f}GB free)")
    return True


def main() -> int:
    excel_path = Path("email_search/data/OUTLOOK_HVDC_ALL_rev.xlsx")
    print("=" * 60)
    print("Pre-flight check")
    print("=" * 60)

    checks = [
        ("Dependencies", check_dependencies),
        ("Input file", lambda: check_input_file(excel_path)),
        ("Disk space", lambda: check_disk_space(excel_path)),
    ]

    all_ok = True
    for name, func in checks:
        print(f"\n[{name}]")
        if not func():
            all_ok = False

    print("\n" + "=" * 60)
    if all_ok:
        print("OK: Ready for full run")
        return 0
    print("FAIL: Pre-flight check failed")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
