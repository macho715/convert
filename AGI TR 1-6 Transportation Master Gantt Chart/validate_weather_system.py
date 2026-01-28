#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Validate weather JSON data and request file presence.
"""

from __future__ import annotations

import json
import os
from datetime import date


def parse_request_file(path):
    params = {}
    if not os.path.exists(path):
        return params
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            params[k.strip()] = v.strip()
    return params


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    req_path = os.path.join(script_dir, "weather_data_requests.txt")
    json_path = os.path.join(script_dir, "weather_data_manual.json")

    print("[CHECK] weather_data_requests.txt:", "OK" if os.path.exists(req_path) else "MISSING")
    req = parse_request_file(req_path)

    start_date = req.get("start_date")
    end_date = req.get("end_date")
    start_d = date.fromisoformat(start_date) if start_date else None
    end_d = date.fromisoformat(end_date) if end_date else None

    if not os.path.exists(json_path):
        print("[ERROR] weather_data_manual.json not found")
        return 1

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERROR] JSON load failed: {e}")
        return 1

    records = data.get("weather_records", []) if isinstance(data, dict) else []
    print(f"[INFO] weather_records: {len(records)}")

    out_of_range = 0
    bad_dates = 0
    duplicates = 0
    seen = set()

    for rec in records:
        d_str = rec.get("date")
        if not d_str:
            bad_dates += 1
            continue
        try:
            d = date.fromisoformat(d_str)
        except Exception:
            bad_dates += 1
            continue

        if d_str in seen:
            duplicates += 1
        seen.add(d_str)

        if start_d and end_d and (d < start_d or d > end_d):
            out_of_range += 1

    if bad_dates:
        print(f"[WARN] invalid date entries: {bad_dates}")
    if duplicates:
        print(f"[WARN] duplicate dates: {duplicates}")
    if out_of_range:
        print(f"[WARN] out-of-range dates: {out_of_range}")
    else:
        print("[OK] all dates within requested range")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
