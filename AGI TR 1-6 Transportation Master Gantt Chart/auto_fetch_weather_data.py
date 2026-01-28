#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto-fetch weather data from Open-Meteo and write JSON for manual pipeline.
Reads weather_data_requests.txt (key=value) for parameters.
"""

from __future__ import annotations

import json
import os
import requests
from datetime import date, datetime, timedelta


FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
MARINE_URL = "https://marine-api.open-meteo.com/v1/marine"


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


def group_visibility_by_day(times, visibility_m):
    vis_by_day = {}
    for t_str, v in zip(times, visibility_m):
        try:
            d = date.fromisoformat(t_str[:10])
        except Exception:
            continue
        if v is None:
            continue
        cur = vis_by_day.get(d)
        if cur is None or v < cur:
            vis_by_day[d] = v
    return vis_by_day


def fetch_weather(lat, lon, tz, start_date, end_date):
    params = {
        "latitude": lat,
        "longitude": lon,
        "timezone": tz,
        "wind_speed_unit": "kn",
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "daily": ",".join(
            [
                "wind_speed_10m_max",
                "wind_gusts_10m_max",
                "wind_direction_10m_dominant",
            ]
        ),
        "hourly": "visibility",
    }
    r = requests.get(FORECAST_URL, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_marine(lat, lon, tz, start_date, end_date):
    params = {
        "latitude": lat,
        "longitude": lon,
        "timezone": tz,
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "daily": "wave_height_max",
    }
    r = requests.get(MARINE_URL, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    req_path = os.path.join(script_dir, "weather_data_requests.txt")
    req = parse_request_file(req_path)

    start_date = date.fromisoformat(req.get("start_date", "2026-01-06"))
    end_date = date.fromisoformat(req.get("end_date", "2026-02-21"))
    lat = float(req.get("lat", "24.12"))
    lon = float(req.get("lon", "52.53"))
    tz = req.get("timezone", "Asia/Dubai")
    output_json = req.get("output_json", "weather_data_manual.json")
    output_path = os.path.join(script_dir, output_json)

    weather = fetch_weather(lat, lon, tz, start_date, end_date)
    marine = fetch_marine(lat, lon, tz, start_date, end_date)

    daily_dates = weather.get("daily", {}).get("time", [])
    wind_max = weather.get("daily", {}).get("wind_speed_10m_max", [])
    gust_max = weather.get("daily", {}).get("wind_gusts_10m_max", [])
    wind_dir = weather.get("daily", {}).get("wind_direction_10m_dominant", [])

    vis_times = weather.get("hourly", {}).get("time", [])
    vis_vals = weather.get("hourly", {}).get("visibility", [])
    vis_by_day = group_visibility_by_day(vis_times, vis_vals)

    wave_dates = marine.get("daily", {}).get("time", [])
    wave_max = marine.get("daily", {}).get("wave_height_max", [])
    wave_map = {d: w for d, w in zip(wave_dates, wave_max)}

    weather_records = []
    for i, d_str in enumerate(daily_dates):
        try:
            d = date.fromisoformat(d_str)
        except Exception:
            continue
        rec = {
            "date": d_str,
            "wind_max_kn": wind_max[i] if i < len(wind_max) else None,
            "gust_max_kn": gust_max[i] if i < len(gust_max) else None,
            "wind_dir_deg": wind_dir[i] if i < len(wind_dir) else None,
            "wave_max_m": wave_map.get(d_str),
            "visibility_km": None,
            "source": "AUTO_OPEN_METEO",
            "notes": "",
        }
        if d in vis_by_day:
            rec["visibility_km"] = vis_by_day[d] / 1000.0
        weather_records.append(rec)

    payload = {
        "source": "Open-Meteo Auto Fetch",
        "generated_at": datetime.now().isoformat(),
        "location": {"lat": lat, "lon": lon},
        "weather_records": weather_records,
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)

    print(f"[OK] Wrote {len(weather_records)} records to {output_path}")


if __name__ == "__main__":
    main()
