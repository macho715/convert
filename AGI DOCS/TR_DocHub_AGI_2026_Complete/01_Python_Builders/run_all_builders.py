#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TR_DocHub integrated builder runner

Runs four builders by scenario and writes outputs into ../05_Templates.
"""
from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path


def run_builder(script: Path, args: list[str], description: str) -> bool:
    try:
        cmd = [sys.executable, str(script)] + args
        print("\n" + "=" * 70)
        print(f"{description}")
        print("=" * 70)
        print("Command:", " ".join(cmd))
        print("=" * 70)
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        if result.stdout:
            print(result.stdout)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        return True
    except subprocess.CalledProcessError as exc:
        print("ERROR: builder failed", file=sys.stderr)
        if exc.stderr:
            print(exc.stderr, file=sys.stderr)
        return False
    except FileNotFoundError:
        print(f"ERROR: script not found: {script}", file=sys.stderr)
        return False


def main() -> None:
    base_dir = Path(__file__).parent
    output_dir = base_dir.parent / "05_Templates"
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    print("=" * 70)
    print("TR_DocHub AGI 2026 - Integrated Builder Runner")
    print("=" * 70)
    print("Select scenario:")
    print("  1) Normalized model (통합빌더.py)")
    print("  2) Legacy model (create_tr_document_tracker_v2.py)")
    print("  3) Legacy + DocGap v3.1 operational patch")
    print("  4) DocGap v2 -> v3 full options")
    print("  5) DocGap v3.1 operational patch (existing file)")

    choice = input("Choice (1-5): ").strip()

    if choice == "1":
        output = output_dir / f"TR_DocHub_AGI_2026_Normalized_{timestamp}.xlsx"
        run_builder(
            base_dir / "통합빌더.py",
            ["--output", str(output)],
            "Normalized model template",
        )

    elif choice == "2":
        output = output_dir / f"TR_Document_Tracker_v2_{timestamp}.xlsx"
        run_builder(
            base_dir / "create_tr_document_tracker_v2.py",
            ["--output", str(output)],
            "Legacy TR Document Tracker template",
        )

    elif choice == "3":
        tr_template = output_dir / f"TR_Tracker_Template_{timestamp}.xlsx"
        final_output = output_dir / f"TR_DocHub_AGI_2026_Integrated_{timestamp}.xlsx"

        if not run_builder(
            base_dir / "create_tr_document_tracker_v2.py",
            ["--output", str(tr_template)],
            "Step 1/2 - build TR template",
        ):
            return

        run_builder(
            base_dir / "build_docgap_v3_1_operational.py",
            ["--in", str(tr_template), "--out", str(final_output)],
            "Step 2/2 - apply DocGap v3.1 operational patch",
        )

    elif choice == "4":
        src = input("Path to DocGap v2 source file: ").strip()
        if not src or not Path(src).exists():
            print("ERROR: source file not found.")
            return

        out_xlsx = output_dir / f"OFCO_AGI_TR1_DocGap_Tracker_v3_FULLOPTIONS_{timestamp}.xlsx"
        out_xlsm = output_dir / f"OFCO_AGI_TR1_DocGap_Tracker_v3_FULLOPTIONS_{timestamp}.xlsm"
        run_builder(
            base_dir / "build_docgap_v3_fulloptions.py",
            ["--src", src, "--out_xlsx", str(out_xlsx), "--out_xlsm", str(out_xlsm)],
            "DocGap v2 -> v3 full options",
        )

    elif choice == "5":
        src = input("Path to file to patch: ").strip()
        if not src or not Path(src).exists():
            print("ERROR: source file not found.")
            return

        output = output_dir / f"TR_DocHub_AGI_2026_Patched_{timestamp}.xlsx"
        run_builder(
            base_dir / "build_docgap_v3_1_operational.py",
            ["--in", src, "--out", str(output)],
            "DocGap v3.1 operational patch",
        )
    else:
        print("ERROR: invalid choice.")


if __name__ == "__main__":
    main()
