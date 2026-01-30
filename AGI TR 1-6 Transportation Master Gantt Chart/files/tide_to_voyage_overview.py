# -*- coding: utf-8 -*-
"""
WATER TIDE.csv 기반: 주간(6:00~17:00) 최고 물때 상위 3시간대를 Voyage Overview tide-table에 연동.
files/ 전용. 실행: files/ 폴더에서 python tide_to_voyage_overview.py [--dry-run]
"""
from __future__ import annotations

import csv
import re
from datetime import datetime
from pathlib import Path

FILES_DIR = Path(__file__).resolve().parent
TIDE_CSV = FILES_DIR / "WATER TIDE.csv"
SCHEDULE_GLOB = "AGI TR SCHEDULE_*.html"

# 6:00 ~ 17:00 컬럼명 (CSV 헤더와 일치)
HOUR_COLS = [f"{h}:00" for h in range(6, 18)]


def parse_tide_csv(path: Path) -> list[tuple[str, dict[str, float]]]:
    """Parse WATER TIDE.csv -> [(date_str, {hour: height_m}), ...]."""
    rows: list[tuple[str, dict[str, float]]] = []
    with open(path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        header = [h.strip() for h in (reader.fieldnames or [])]
        # first column is date (날짜 or BOM+날짜)
        date_key = header[0] if header else "날짜"
        for row in reader:
            date_str = (row.get(date_key) or row.get("날짜") or "").strip()
            if not date_str or not date_str[0].isdigit():
                continue
            heights: dict[str, float] = {}
            for col in HOUR_COLS:
                val = (row.get(col) or "").strip().replace(" ", "")
                try:
                    heights[col] = float(val) if val else 0.0
                except ValueError:
                    heights[col] = 0.0
            rows.append((date_str, heights))
    return rows


def voyage_cards_from_html(html_path: Path) -> list[tuple[int, str, str]]:
    """Extract (voyage_num, data-start, data-end) from voyage-card divs."""
    text = html_path.read_text(encoding="utf-8")
    cards: list[tuple[int, str, str]] = []
    # <div class="voyage-card" data-voyage="1" data-start="2026-01-31" data-end="2026-02-09">
    pattern = re.compile(
        r'<div\s+class="voyage-card"[^>]*\s+data-voyage="(\d+)"[^>]*\s+data-start="([^"]+)"[^>]*\s+data-end="([^"]+)"',
        re.I,
    )
    for m in pattern.finditer(text):
        cards.append((int(m.group(1)), m.group(2), m.group(3)))
    return cards


def top3_tide_for_range(
    tide_rows: list[tuple[str, dict[str, float]]], start: str, end: str
) -> list[tuple[str, float]]:
    """For dates in [start, end], compute max height per hour (6~17), return top 3 (time, height)."""
    try:
        d_start = datetime.strptime(start, "%Y-%m-%d")
        d_end = datetime.strptime(end, "%Y-%m-%d")
    except ValueError:
        return []

    hour_max: dict[str, float] = {h: 0.0 for h in HOUR_COLS}
    for date_str, heights in tide_rows:
        try:
            d = datetime.strptime(date_str.strip(), "%Y-%m-%d")
        except ValueError:
            continue
        if d_start <= d <= d_end:
            for h in HOUR_COLS:
                hour_max[h] = max(hour_max.get(h, 0), heights.get(h, 0))

    sorted_hours = sorted(
        [(t, hour_max[t]) for t in HOUR_COLS],
        key=lambda x: -x[1],
    )
    top3 = sorted_hours[:3]
    return [(t, round(h, 2)) for t, h in top3]


def replace_tide_table_in_html(
    html_path: Path,
    voyage_top3: list[tuple[int, list[tuple[str, float]]]],
    dry_run: bool,
) -> bool:
    """Replace each voyage-card's table.tide-table tbody with 3 rows from voyage_top3.
    Process from last voyage to first so string indices stay valid."""
    html = html_path.read_text(encoding="utf-8")
    changed = False
    for voyage_num, rows in reversed(voyage_top3):
        # Match voyage-card (has data-start), not voyage-chip button
        card_start = f'data-voyage="{voyage_num}" data-start'
        idx = html.find(card_start)
        if idx == -1:
            continue
        tbody_start = html.find("<tbody>", idx)
        tbody_end = html.find("</tbody>", tbody_start)
        if tbody_start == -1 or tbody_end == -1 or tbody_end < tbody_start:
            continue
        block = html[max(0, tbody_start - 200) : tbody_start]
        if "tide-table" not in block:
            continue
        new_tbody = "<tbody>\n"
        for time_str, height in rows:
            new_tbody += f"                                <tr>\n                                    <td>{time_str}</td>\n                                    <td>{height}m</td>\n                                </tr>\n"
        new_tbody += "                            </tbody>"
        old_tbody = html[tbody_start : tbody_end + len("</tbody>")]
        if old_tbody != new_tbody:
            html = html[:tbody_start] + new_tbody + html[tbody_end + len("</tbody>") :]
            changed = True
    if changed and not dry_run:
        html_path.write_text(html, encoding="utf-8")
    return changed


def main() -> None:
    import sys

    dry_run = "--dry-run" in sys.argv

    if not TIDE_CSV.is_file():
        print(f"SKIP: {TIDE_CSV.name} not found")
        return

    tide_rows = parse_tide_csv(TIDE_CSV)
    if not tide_rows:
        print("SKIP: no tide rows parsed")
        return

    for html_path in sorted(FILES_DIR.glob(SCHEDULE_GLOB)):
        cards = voyage_cards_from_html(html_path)
        if not cards:
            continue
        print(f"Processing {html_path.name} ({len(cards)} voyages)")
        voyage_top3: list[tuple[int, list[tuple[str, float]]]] = []
        for v_num, start, end in cards:
            top3 = top3_tide_for_range(tide_rows, start, end)
            if not top3:
                # fallback: use first 3 hours from HOUR_COLS with zero or first day
                top3 = [(HOUR_COLS[i], 0.0) for i in range(3)]
            voyage_top3.append((v_num, top3))
            if dry_run:
                print(f"Voyage {v_num} [{start} ~ {end}]: {top3}")

        if dry_run:
            print(f"[dry-run] Would update {html_path.name}")
            continue
        if replace_tide_table_in_html(html_path, voyage_top3, dry_run=False):
            print(f"Updated tide tables: {html_path.name}")


if __name__ == "__main__":
    main()
