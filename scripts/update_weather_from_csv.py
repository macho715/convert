#!/usr/bin/env python3
"""
Update weather JSON using the latest CSV data.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_CSV = (
    ROOT_DIR
    / "AGI TR 1-6 Transportation Master Gantt Chart"
    / "Date,Wind Max (kn),Gust Max (kn),Wind Di.csv"
)
DEFAULT_JSON = (
    ROOT_DIR / "AGI TR 1-6 Transportation Master Gantt Chart" / "weather_data_20260106.json"
)


def parse_wind_dir(wind_dir_str: str) -> Optional[float]:
    """Extract numeric degrees from strings like '315 (NW)'."""
    if not wind_dir_str:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", str(wind_dir_str))
    if match:
        return float(match.group(1))
    return None


def parse_risk_level(risk_str: str) -> str:
    """Normalize risk level strings."""
    if not risk_str:
        return "MEDIUM"
    return str(risk_str).strip().upper()


def parse_bool(value) -> bool:
    """Parse common boolean representations."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "y", "t"}
    if value is None:
        return False
    return bool(value)


def safe_float(value) -> Optional[float]:
    """Convert values to float safely."""
    if value is None:
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def read_csv_weather(csv_path: Path) -> Dict[str, Dict]:
    """
    Read CSV and return a date-indexed dict of weather records.
    """
    print(f"Reading CSV: {csv_path}")

    weather_dict: Dict[str, Dict] = {}
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_str = (row.get("Date") or "").strip()
            if not date_str:
                continue

            wind_dir_deg = parse_wind_dir(row.get("Wind Dir (deg)", ""))
            wave_dir_deg = parse_wind_dir(row.get("Wave Dir (deg)", ""))

            record = {
                "date": date_str,
                "wind_max_kn": safe_float(row.get("Wind Max (kn)")),
                "gust_max_kn": safe_float(row.get("Gust Max (kn)")),
                "wind_dir_deg": wind_dir_deg,
                "wave_max_m": safe_float(row.get("Wave Max (m)")),
                "wave_period_s": safe_float(row.get("Wave Period (s)")),
                "wave_dir_deg": wave_dir_deg,
                "visibility_km": safe_float(row.get("Visibility (km)")),
                "source": "CSV Updated",
                "notes": (row.get("Notes") or "").strip(),
                "risk_level": parse_risk_level(row.get("Risk Level", "MEDIUM")),
                "is_shamal": parse_bool(row.get("Is Shamal", False)),
            }

            weather_dict[date_str] = record

    print(f"Records read: {len(weather_dict)}")
    if weather_dict:
        dates = sorted(weather_dict.keys())
        print(f"Date range: {dates[0]} to {dates[-1]}")

    return weather_dict


def update_json_from_csv(csv_path: Path, json_path: Path) -> Dict:
    """
    Update a JSON file with records from CSV.
    Dates present in CSV overwrite JSON, other dates remain.
    """
    csv_weather = read_csv_weather(csv_path)

    if json_path.exists():
        print(f"Loading JSON: {json_path}")
        with json_path.open("r", encoding="utf-8") as f:
            json_data = json.load(f)
        weather_records = list(json_data.get("weather_records", []))
        print(f"Existing records: {len(weather_records)}")
    else:
        print("JSON file not found; creating a new one.")
        json_data = {
            "source": "AGI TR Weather Data (CSV updated)",
            "generated_at": datetime.utcnow().isoformat() + "Z",
            "location": {
                "name": "Mina Zayed Port / AGI Site",
                "latitude": 24.12,
                "longitude": 52.53,
            },
            "weather_records": [],
        }
        weather_records = []

    existing_records_dict = {}
    for i, rec in enumerate(weather_records):
        date_str = rec.get("date")
        if isinstance(date_str, str) and date_str:
            existing_records_dict[date_str] = i

    updated_count = 0
    added_count = 0
    for date_str, csv_record in csv_weather.items():
        if date_str in existing_records_dict:
            idx = existing_records_dict[date_str]
            weather_records[idx] = csv_record
            updated_count += 1
        else:
            weather_records.append(csv_record)
            added_count += 1

    weather_records.sort(key=lambda r: r.get("date", ""))

    json_data["source"] = "AGI TR Weather Data (CSV updated)"
    json_data["generated_at"] = datetime.utcnow().isoformat() + "Z"
    json_data["weather_records"] = weather_records

    print(f"Writing JSON: {json_path}")
    json_path.parent.mkdir(parents=True, exist_ok=True)
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, ensure_ascii=True)

    shamal_count = sum(1 for r in weather_records if r.get("is_shamal", False))
    high_risk_count = sum(1 for r in weather_records if r.get("risk_level") == "HIGH")

    print("Update summary:")
    print(f"  Total records: {len(weather_records)}")
    print(f"  Updated records: {updated_count}")
    print(f"  Added records: {added_count}")
    print(f"  Shamal days: {shamal_count}")
    print(f"  HIGH risk days: {high_risk_count}")
    if weather_records:
        dates = sorted([r.get("date", "") for r in weather_records if r.get("date")])
        if dates:
            print(f"  Date range: {dates[0]} to {dates[-1]}")

    print("Done.")
    return json_data


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Update weather JSON using the latest CSV data."
    )
    parser.add_argument(
        "--csv",
        default=str(DEFAULT_CSV),
        help="Path to the CSV file (default: project weather CSV).",
    )
    parser.add_argument(
        "--json",
        default=str(DEFAULT_JSON),
        help="Path to the JSON file to update (default: weather_data_20260106.json).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    csv_path = Path(args.csv)
    json_path = Path(args.json)

    if not csv_path.is_absolute():
        csv_path = Path.cwd() / csv_path
    if not json_path.is_absolute():
        json_path = Path.cwd() / json_path

    if not csv_path.exists():
        print(f"CSV not found: {csv_path}")
        return 1

    update_json_from_csv(csv_path, json_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
