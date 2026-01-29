#!/usr/bin/env python3
"""Shift schedule dates: pivot_date and later get +delta_days. YYYY-MM-DD only."""
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

PIVOT = "2026-01-29"
DELTA_DAYS = 1


def shift_date(s: str) -> str:
    if not s or not re.match(r"\d{4}-\d{2}-\d{2}", s):
        return s
    d = datetime.strptime(s[:10], "%Y-%m-%d").date()
    if d < datetime.strptime(PIVOT, "%Y-%m-%d").date():
        return s
    new_d = d + timedelta(days=DELTA_DAYS)
    return new_d.strftime("%Y-%m-%d") + (s[10:] if len(s) > 10 else "")


def shift_json(path: Path, dry_run: bool = False, backup: bool = False) -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    for act in data.get("activities", []):
        for key in ("planned_start", "planned_finish"):
            if key in act and act[key]:
                act[key] = shift_date(act[key])
    dr = data.get("summary", {}).get("date_range", {})
    for key in ("start", "finish"):
        if key in dr and dr[key]:
            dr[key] = shift_date(dr[key])
    if dry_run:
        print("DRY RUN: would write", path)
        return
    if backup:
        backup_path = (
            path.parent
            / f"{path.stem}_backup_{datetime.now().strftime('%Y%m%d')}{path.suffix}"
        )
        with open(backup_path, "w", encoding="utf-8") as f:
            json.dump(
                json.load(open(path, "r", encoding="utf-8")),
                f,
                indent=2,
                ensure_ascii=False,
            )
        print("Backup:", backup_path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print("Shifted JSON:", path)


def shift_html(path: Path, dry_run: bool = False) -> None:
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    pat = re.compile(r"\b(2026-\d{2}-\d{2})\b")
    pivot_dt = datetime.strptime(PIVOT, "%Y-%m-%d").date()

    def repl(m):
        s = m.group(1)
        d = datetime.strptime(s[:10], "%Y-%m-%d").date()
        if d < pivot_dt:
            return s
        new_d = d + timedelta(days=DELTA_DAYS)
        return new_d.strftime("%Y-%m-%d")

    new_text = pat.sub(repl, text)
    if new_text == text:
        print("No date changes in", path)
        return
    if dry_run:
        print("DRY RUN: would write", path)
        return
    with open(path, "w", encoding="utf-8") as f:
        f.write(new_text)
    print("Shifted HTML:", path)


def main():
    dry = "--dry-run" in sys.argv
    backup = "--backup" in sys.argv
    base = Path(__file__).resolve().parent.parent
    json_path = base / "agi tr final schedule.json"
    if json_path.exists():
        shift_json(json_path, dry_run=dry, backup=backup)
    for name in [
        "AGI TR Unit 1 Schedule_20260126.html",
        "AGI TR Unit 1 Schedule_20260128.html",
    ]:
        p = base / name
        if p.exists():
            shift_html(p, dry_run=dry)
    files_dir = base / "files"
    for name in [
        "AGI20Unit20Schedule_20260126_redesigned_v2.html",
        "AGI20Unit20Schedule_20260128_redesigned_v2.html",
        "AGI20Unit20Schedule_20260129_redesigned_v2.html",
    ]:
        p = files_dir / name
        if p.exists():
            shift_html(p, dry_run=dry)


if __name__ == "__main__":
    main()
