# -*- coding: utf-8 -*-
"""
AGI TR Transportation - Weather Risk Heatmap (Jan-Feb 2026)  [PATCH v2]
- No random / deterministic
- Multi-source ensemble: Open-Meteo forecast + gfs + ecmwf + dwd-icon + gem (>=5 feeds)
- Wave: Open-Meteo Marine API (wave_height_max) when available; fallback approximation otherwise
- Past dates: Open-Meteo Historical Weather API (archive-api) when available
- Far-future gap fill: Open-Meteo Climate API (climate-api) [marked as assumption]
- Emoji removed to avoid matplotlib glyph warnings

External weather-site cross-check basis (qualitative):
- UAE NCM monthly outlook: Shamal + rough sea + fog/mist risk
- WAM/NCM: NW wind up to 50 km/h, dust -> reduced visibility, rough sea
- Windy / Meteoblue / Windy.app: wind+gust+wave presentation & multi-model comparison
References: see citations in your report, not embedded here.
"""

from __future__ import annotations
import math
import json
import os
import requests
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.offsetbox import AnchoredText
from datetime import datetime, timedelta, date

# -----------------------------
# USER CONFIG
# -----------------------------
TZ = "Asia/Dubai"

# Assumption: coordinates are AGI/route representative point (recommend replacing with user actual AGI coordinates)
# If you have exact AGI berth/ramp coordinate, replace these:
LAT = 24.12
LON = 52.53

# Schedule 4-day mode: update date + 4 days (D..D+3), output schedule PNG
SCHEDULE_4DAY_MODE = True
if SCHEDULE_4DAY_MODE:
    _update = date.today()
    START_DATE = _update
    END_DATE = _update + timedelta(days=3)  # inclusive, 4 days
    _script_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    _out_dir = os.path.join(_script_dir, "out")
    os.makedirs(_out_dir, exist_ok=True)
    OUTPUT_PATH = os.path.join(_out_dir, "weather_4day_heatmap.png")
else:
    START_DATE = date(2026, 1, 15)
    END_DATE = date(2026, 2, 15)  # inclusive
    OUTPUT_PATH = "AGI_TR_Weather_Risk_Heatmap_v3.png"

# Weather data source settings
USE_MANUAL_JSON = True  # True: Use JSON file, False: Use API
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
WEATHER_JSON_PATH = os.path.join(SCRIPT_DIR, "weather_data_20260106.json")
WEATHER_REQUEST_PATH = os.path.join(SCRIPT_DIR, "weather_data_requests.txt")
# Pipeline: PDF parse -> parsed_to_weather_json.py -> this file. If 4-day mode and parsed JSON exists, use it.
if SCHEDULE_4DAY_MODE:
    _convert_root = os.path.dirname(SCRIPT_DIR)
    _parsed_candidate = os.path.join(
        _convert_root, "out", "weather_parsed",
        _update.strftime("%Y%m%d"), "weather_for_weather_py.json"
    )
    if os.path.exists(_parsed_candidate):
        WEATHER_JSON_PATH = _parsed_candidate

# Voyage overlay (optional): put only those within Jan-Feb window
# If you want schedule overlay, edit these to exact dates.
VOYAGES = [
    {
        "name": "V1",
        "start": date(2026, 1, 18),
        "end": date(2026, 1, 22),
        "label": "TR1",
        "type": "transport",
    },
    {
        "name": "V2",
        "start": date(2026, 1, 24),
        "end": date(2026, 2, 6),
        "label": "TR2+JD-1",
        "type": "jackdown",
    },
    {
        "name": "V3",
        "start": date(2026, 2, 8),
        "end": date(2026, 2, 12),
        "label": "TR3",
        "type": "transport",
    },
    {
        "name": "V4",
        "start": date(2026, 2, 14),
        "end": date(2026, 2, 15),  # Adjusted to fit new END_DATE
        "label": "TR4+JD-2",
        "type": "jackdown",
    },
]

MANUAL_SHAMAL_PERIODS = [
    (date(2026, 2, 5), date(2026, 2, 14)),
]

THEME = {
    "cmap": [
        "#F7FBFF",
        "#DEEbf7",
        "#C6DBEF",
        "#9ECAE1",
        "#6BAED6",
        "#4292C6",
        "#2171B5",
        "#08519C",
        "#08306B",
    ],  # Changed to a more sophisticated blue gradient for high-end feel
    "status": {
        "GO": "#1B5E20",
        "HOLD": "#F57F17",
        "NO-GO": "#B71C1C",
    },
    "risk_band": {
        "GO": "#E8F5E9",
        "HOLD": "#FFF8E1",
        "NO-GO": "#FFEBEE",
    },
    "shamal": "#8B0000",
    "voyage": {
        "transport": "#1565C0",
        "jackdown": "#4527A0",
        "default": "#424242",
    },
}

# -----------------------------
# API ENDPOINTS (Open-Meteo)
# -----------------------------
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

MODEL_URLS = {
    "forecast": FORECAST_URL,
    "gfs": "https://api.open-meteo.com/v1/gfs",
    "ecmwf": "https://api.open-meteo.com/v1/ecmwf",
    "dwd": "https://api.open-meteo.com/v1/dwd-icon",
    "gem": "https://api.open-meteo.com/v1/gem",
}

ARCHIVE_URL = "https://archive-api.open-meteo.com/v1/archive"

MARINE_URL = "https://marine-api.open-meteo.com/v1/marine"

CLIMATE_URL = "https://climate-api.open-meteo.com/v1/climate"
CLIMATE_MODELS = ["EC_Earth3P_HR", "MRI_AGCM3_2_S", "MPI_ESM1_2_XR"]


# -----------------------------
# HELPERS
# -----------------------------
def daterange(d0: date, d1: date) -> list[date]:
    n = (d1 - d0).days
    return [d0 + timedelta(days=i) for i in range(n + 1)]


def ensure_weather_json(json_path):
    if os.path.exists(json_path):
        return

    empty_json = {
        "source": "Manual Weather Data Entry",
        "generated_at": datetime.now().isoformat(),
        "location": {"lat": LAT, "lon": LON},
        "weather_records": [],
    }
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(empty_json, f, indent=2, ensure_ascii=False)
        print(f"[WARN] Created empty JSON file: {json_path}")
    except Exception as e:
        print(f"[ERROR] Failed to create empty JSON file: {e}")


def load_weather_data_from_json(json_path, start_date=None, end_date=None):
    """
    Load manually entered weather data from JSON file
    Function used instead of API
    """
    if not os.path.exists(json_path):
        print(f"[WARN] Weather data JSON file not found: {json_path}")
        return []

    if hasattr(start_date, "date"):
        start_date = start_date.date()
    if hasattr(end_date, "date"):
        end_date = end_date.date()

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERROR] Weather data JSON load error: {e}")
        return []

    weather_records = []
    raw_records = data.get("weather_records", []) if isinstance(data, dict) else []
    for record in raw_records:
        date_str = record.get("date", "")
        if not date_str:
            continue
        try:
            d = date.fromisoformat(date_str)
        except Exception:
            continue

        if start_date and end_date and (d < start_date or d > end_date):
            print(f"[WARN] Ignoring data outside date range: {date_str}")
            continue

        weather_records.append(
            {
                "date": date_str,
                "wind_max_kn": record.get("wind_max_kn"),
                "gust_max_kn": record.get("gust_max_kn"),
                "wind_dir_deg": record.get("wind_dir_deg"),
                "wave_max_m": record.get("wave_max_m"),
                "wave_period_s": record.get("wave_period_s"),
                "wave_dir_deg": record.get("wave_dir_deg"),
                "visibility_km": record.get("visibility_km"),
                "source": record.get("source", "MANUAL"),
                "notes": record.get("notes", ""),
                "risk_level": record.get("risk_level"),
                "is_shamal": record.get("is_shamal"),
            }
        )

    print(f"[OK] Loaded {len(weather_records)} weather records from JSON")
    return weather_records


def to_idx_map(days: list[date]) -> dict[date, int]:
    return {d: i for i, d in enumerate(days)}


def safe_get(dct, *keys, default=None):
    cur = dct
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


def request_json(url: str, params: dict, timeout=30) -> dict:
    try:
        r = requests.get(url, params=params, timeout=timeout)
        r.raise_for_status()
        return r.json()
    except:
        return {}


def parse_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in ("true", "1", "yes", "y", "t")
    return bool(value)


# -----------------------------
# FETCH WEATHER (DAILY + VISIBILITY MIN from HOURLY)
# -----------------------------
def fetch_weather_model(model_name: str, base_url: str, d0: date, d1: date) -> dict:
    """
    Returns dict with keys:
      dates (list[str]), wind_max_kn, gust_max_kn, wind_dir_deg, vis_min_km
    Notes:
      - wind_speed_unit supports 'kn' for many endpoints
      - visibility is hourly in many models; we compute daily min km
    """
    params = {
        "latitude": LAT,
        "longitude": LON,
        "timezone": TZ,
        "wind_speed_unit": "kn",
        "start_date": d0.isoformat(),
        "end_date": d1.isoformat(),
        "daily": ",".join(
            [
                "wind_speed_10m_max",
                "wind_gusts_10m_max",
                "wind_direction_10m_dominant",
            ]
        ),
        "hourly": ",".join(
            [
                "visibility",  # meters
            ]
        ),
        "forecast_days": 16,  # will be ignored/limited by some endpoints
    }
    j = request_json(base_url, params)

    daily_time = safe_get(j, "daily", "time", default=[])
    wind_max = np.array(
        safe_get(j, "daily", "wind_speed_10m_max", default=[]), dtype=float
    )
    gust_max = np.array(
        safe_get(j, "daily", "wind_gusts_10m_max", default=[]), dtype=float
    )
    wind_dir = np.array(
        safe_get(j, "daily", "wind_direction_10m_dominant", default=[]), dtype=float
    )

    hourly_time = safe_get(j, "hourly", "time", default=[])
    hourly_vis_m = np.array(
        safe_get(j, "hourly", "visibility", default=[]), dtype=float
    )  # meters

    # Compute daily min visibility (km) from hourly
    # Map hourly timestamps (ISO8601) to date
    vis_min_km = np.full(len(daily_time), np.nan, dtype=float)
    if len(hourly_time) == len(hourly_vis_m) and len(daily_time) > 0:
        day_to_vals = {t: [] for t in daily_time}
        for ts, vm in zip(hourly_time, hourly_vis_m):
            day = ts[:10]  # 'YYYY-MM-DD'
            if day in day_to_vals and not np.isnan(vm):
                day_to_vals[day].append(vm / 1000.0)
        for i, day in enumerate(daily_time):
            vals = day_to_vals.get(day, [])
            if vals:
                vis_min_km[i] = float(np.nanmin(vals))

    return {
        "model": model_name,
        "dates": daily_time,
        "wind_max_kn": wind_max,
        "gust_max_kn": gust_max,
        "wind_dir_deg": wind_dir,
        "vis_min_km": vis_min_km,
    }


# -----------------------------
# FETCH HISTORICAL (ARCHIVE) for past days
# -----------------------------
def fetch_archive(d0: date, d1: date) -> dict:
    """
    Uses Historical Weather API /v1/archive (archive-api.open-meteo.com)
    """
    params = {
        "latitude": LAT,
        "longitude": LON,
        "timezone": TZ,
        "wind_speed_unit": "kn",
        "start_date": d0.isoformat(),
        "end_date": d1.isoformat(),
        "daily": ",".join(
            [
                "wind_speed_10m_max",
                "wind_gusts_10m_max",
                "wind_direction_10m_dominant",
            ]
        ),
        "hourly": "visibility",
    }
    j = request_json(ARCHIVE_URL, params)

    daily_time = safe_get(j, "daily", "time", default=[])
    wind_max = np.array(
        safe_get(j, "daily", "wind_speed_10m_max", default=[]), dtype=float
    )
    gust_max = np.array(
        safe_get(j, "daily", "wind_gusts_10m_max", default=[]), dtype=float
    )
    wind_dir = np.array(
        safe_get(j, "daily", "wind_direction_10m_dominant", default=[]), dtype=float
    )

    hourly_time = safe_get(j, "hourly", "time", default=[])
    hourly_vis_m = np.array(
        safe_get(j, "hourly", "visibility", default=[]), dtype=float
    )

    vis_min_km = np.full(len(daily_time), np.nan, dtype=float)
    if len(hourly_time) == len(hourly_vis_m) and len(daily_time) > 0:
        day_to_vals = {t: [] for t in daily_time}
        for ts, vm in zip(hourly_time, hourly_vis_m):
            day = ts[:10]
            if day in day_to_vals and not np.isnan(vm):
                day_to_vals[day].append(vm / 1000.0)
        for i, day in enumerate(daily_time):
            vals = day_to_vals.get(day, [])
            if vals:
                vis_min_km[i] = float(np.nanmin(vals))

    return {
        "model": "archive",
        "dates": daily_time,
        "wind_max_kn": wind_max,
        "gust_max_kn": gust_max,
        "wind_dir_deg": wind_dir,
        "vis_min_km": vis_min_km,
    }


# -----------------------------
# FETCH MARINE WAVES (DAILY wave_height_max)
# -----------------------------
def fetch_marine_waves(d0: date, d1: date) -> dict:
    params = {
        "latitude": LAT,
        "longitude": LON,
        "timezone": TZ,
        "start_date": d0.isoformat(),
        "end_date": d1.isoformat(),
        "daily": "wave_height_max",
        "cell_selection": "sea",
    }
    j = request_json(MARINE_URL, params)
    daily_time = safe_get(j, "daily", "time", default=[])
    wave_max = np.array(
        safe_get(j, "daily", "wave_height_max", default=[]), dtype=float
    )
    return {"dates": daily_time, "wave_max_m": wave_max}


# -----------------------------
# CLIMATE GAP FILL (wind_speed_10m_max)
# -----------------------------
def fetch_climate_wind_max(d0: date, d1: date) -> dict:
    params = {
        "latitude": LAT,
        "longitude": LON,
        "start_date": d0.isoformat(),
        "end_date": d1.isoformat(),
        "models": ",".join(CLIMATE_MODELS),
        "daily": "wind_speed_10m_max",
        "wind_speed_unit": "kn",
    }
    j = request_json(CLIMATE_URL, params)
    daily_time = safe_get(j, "daily", "time", default=[])
    # Climate API returns per-model arrays; Open-Meteo packs them depending on request.
    # If it's a single combined, handle both.
    w = safe_get(j, "daily", "wind_speed_10m_max", default=None)
    if w is None:
        # fallback: try model-specific layout (rare)
        return {"dates": daily_time, "wind_max_kn": np.full(len(daily_time), np.nan)}
    wind = np.array(w, dtype=float)
    return {"dates": daily_time, "wind_max_kn": wind}


# -----------------------------
# RISK MODEL (tuned to Shamal signals: strong NW wind, dust -> vis drop, rough seas)
# -----------------------------
def calc_risk_score(wind_kn, gust_kn, wave_m, vis_km):
    """
    Scoring 0..100 approx (clipped).
    """
    wind_risk = np.clip((wind_kn - 12.0) * 4.0, 0.0, 40.0)  # >12kt starts
    gust_risk = np.clip((gust_kn - 18.0) * 2.5, 0.0, 30.0)  # >18kt gusts
    wave_risk = np.clip((wave_m - 0.80) * 35.0, 0.0, 25.0)  # >0.8m waves
    vis_risk = np.clip((6.0 - vis_km) * 6.0, 0.0, 25.0)  # <6km
    score = wind_risk + gust_risk + wave_risk + vis_risk
    return np.clip(score, 0.0, 100.0)


def op_status_from_score(score):
    if score < 30.0:
        return "GO"
    elif score < 60.0:
        return "HOLD"
    return "NO-GO"


def is_shamal_day(wind_dir_deg, wind_kn, gust_kn):
    """
    NW sector + strong wind/gust.
    """
    if np.isnan(wind_dir_deg) or np.isnan(wind_kn) or np.isnan(gust_kn):
        return False
    nw = 285.0 <= wind_dir_deg <= 345.0
    strong = (wind_kn >= 18.0) or (gust_kn >= 22.0)
    return bool(nw and strong)


# -----------------------------
# MAIN PIPELINE
# -----------------------------
def main():
    days = daterange(START_DATE, END_DATE)
    idx = to_idx_map(days)
    n = len(days)

    # Arrays
    wind_kn = np.full(n, np.nan)
    gust_kn = np.full(n, np.nan)
    wdir_deg = np.full(n, np.nan)
    vis_km = np.full(n, np.nan)
    wave_m = np.full(n, np.nan)
    risk_level_override = [None] * n
    shamal_override = [None] * n

    # Source coverage tracker
    coverage = np.array([""] * n, dtype=object)

    # ============================================
    # Load weather data from manual input JSON file (instead of API)
    # ============================================
    if USE_MANUAL_JSON:
        ensure_weather_json(WEATHER_JSON_PATH)
        weather_records = load_weather_data_from_json(
            WEATHER_JSON_PATH, start_date=START_DATE, end_date=END_DATE
        )

        if weather_records:
            # Fill arrays with JSON data
            for record in weather_records:
                try:
                    d = date.fromisoformat(record["date"])
                    if d in idx:
                        i = idx[d]
                        if record.get("wind_max_kn") is not None:
                            wind_kn[i] = record["wind_max_kn"]
                            coverage[i] = record.get("source", "MANUAL")
                        if record.get("gust_max_kn") is not None:
                            gust_kn[i] = record["gust_max_kn"]
                        if record.get("wind_dir_deg") is not None:
                            wdir_deg[i] = record["wind_dir_deg"]
                        if record.get("wave_max_m") is not None:
                            wave_m[i] = record["wave_max_m"]
                        if record.get("visibility_km") is not None:
                            vis_km[i] = record["visibility_km"]
                        if record.get("risk_level"):
                            risk_level_override[i] = str(record["risk_level"]).upper()
                        if record.get("is_shamal") is not None:
                            shamal_override[i] = parse_bool(record["is_shamal"])
                except Exception as e:
                    print(
                        f"[WARN] Weather data processing error ({record.get('date', 'Unknown')}): {e}"
                    )
        else:
            print("[WARN] Unable to load weather data.")
            print(
                "   Input data in weather_data_template.csv and run convert_weather_csv_to_json.py"
            )
            print("   Or set USE_MANUAL_JSON = False to use API mode")

    # ============================================
    # API mode (used only if USE_MANUAL_JSON = False)
    # ============================================
    if not USE_MANUAL_JSON:
        # 1) Archive for past dates (2-day delay may apply)
        today = datetime.now().date()
        archive_end = min(END_DATE, today - timedelta(days=2))
        if archive_end >= START_DATE:
            try:
                arc = fetch_archive(START_DATE, archive_end)
                for d_str, w, g, wd, v in zip(
                    arc["dates"],
                    arc["wind_max_kn"],
                    arc["gust_max_kn"],
                    arc["wind_dir_deg"],
                    arc["vis_min_km"],
                ):
                    d = date.fromisoformat(d_str)
                    if d in idx:
                        i = idx[d]
                        wind_kn[i], gust_kn[i], wdir_deg[i], vis_km[i] = w, g, wd, v
                        coverage[i] = "ARCHIVE"
            except Exception:
                pass

        # 2) Multi-model forecast ensemble
        remaining_start = max(START_DATE, today - timedelta(days=1))
        if remaining_start <= END_DATE:
            model_payloads = []
            for name, url in MODEL_URLS.items():
                try:
                    model_payloads.append(
                        fetch_weather_model(name, url, remaining_start, END_DATE)
                    )
                except Exception:
                    pass

            for d in days:
                if d < remaining_start:
                    continue
                i = idx[d]
                d_str = d.isoformat()

                w_list, g_list, wd_list, v_list = [], [], [], []
                for p in model_payloads:
                    if d_str in p["dates"]:
                        k = p["dates"].index(d_str)
                        w_list.append(p["wind_max_kn"][k])
                        g_list.append(p["gust_max_kn"][k])
                        wd_list.append(p["wind_dir_deg"][k])
                        v_list.append(p["vis_min_km"][k])

                if coverage[i] != "ARCHIVE":
                    if w_list:
                        wind_kn[i] = float(np.nanmean(w_list))
                        gust_kn[i] = float(np.nanmean(g_list)) if g_list else np.nan
                        wdir_deg[i] = float(np.nanmean(wd_list)) if wd_list else np.nan
                        vis_km[i] = float(np.nanmean(v_list)) if v_list else np.nan
                        coverage[i] = "FORECAST_ENSEMBLE"

        # 3) Marine waves
        try:
            mw = fetch_marine_waves(START_DATE, END_DATE)
            for d_str, wv in zip(mw["dates"], mw["wave_max_m"]):
                d = date.fromisoformat(d_str)
                if d in idx:
                    wave_m[idx[d]] = wv
        except Exception:
            pass

        # 4) Climate gap fill
        missing = np.isnan(wind_kn)
        if missing.any():
            try:
                clim = fetch_climate_wind_max(START_DATE, END_DATE)
                clim_map = {
                    date.fromisoformat(t): v
                    for t, v in zip(clim["dates"], clim["wind_max_kn"])
                }
                for d in days:
                    i = idx[d]
                    if np.isnan(wind_kn[i]) and d in clim_map:
                        wind_kn[i] = float(clim_map[d])
                        coverage[i] = "CLIMATE_FILL"
            except Exception:
                pass

    # ============================================
    # Gap fill: Interpolate missing data (common)
    # ============================================
    print("\n[INFO] Checking data completeness...")
    missing_count = np.sum(np.isnan(wind_kn))
    if missing_count > 0:
        print(f"   [WARN] {missing_count} days of data are missing.")
        print("   Applying default or interpolation to missing dates")

        # Simple interpolation (linear interpolation)
        for i in range(n):
            if np.isnan(wind_kn[i]):
                # Try interpolation with previous/next values
                prev_val = (
                    wind_kn[i - 1] if i > 0 and not np.isnan(wind_kn[i - 1]) else None
                )
                next_val = (
                    wind_kn[i + 1]
                    if i < n - 1 and not np.isnan(wind_kn[i + 1])
                    else None
                )

                if prev_val is not None and next_val is not None:
                    wind_kn[i] = (prev_val + next_val) / 2
                    gust_kn[i] = (
                        wind_kn[i] * 1.3 if np.isnan(gust_kn[i]) else gust_kn[i]
                    )
                    coverage[i] = "INTERPOLATED"
                elif prev_val is not None:
                    wind_kn[i] = prev_val
                    gust_kn[i] = (
                        wind_kn[i] * 1.3 if np.isnan(gust_kn[i]) else gust_kn[i]
                    )
                    coverage[i] = "INTERPOLATED"
                elif next_val is not None:
                    wind_kn[i] = next_val
                    gust_kn[i] = (
                        wind_kn[i] * 1.3 if np.isnan(gust_kn[i]) else gust_kn[i]
                    )
                    coverage[i] = "INTERPOLATED"
                else:
                    # Default values
                    wind_kn[i] = 12.0
                    gust_kn[i] = 15.0
                    coverage[i] = "DEFAULT"

            # Wave interpolation
            if np.isnan(wave_m[i]) and not np.isnan(wind_kn[i]):
                wave_m[i] = float(np.clip(wind_kn[i] * 0.04, 0.30, 2.50))

            # Visibility default
            if np.isnan(vis_km[i]):
                vis_km[i] = 8.00

            # Gust interpolation
            if np.isnan(gust_kn[i]) and not np.isnan(wind_kn[i]):
                gust_kn[i] = wind_kn[i] * 1.30

    # 6) Risk + status
    risk = calc_risk_score(wind_kn, gust_kn, wave_m, vis_km)
    status = [op_status_from_score(s) for s in risk]
    shamal = np.array(
        [is_shamal_day(wdir_deg[i], wind_kn[i], gust_kn[i]) for i in range(n)],
        dtype=bool,
    )
    risk_map = {"LOW": 20.0, "MEDIUM": 45.0, "HIGH": 75.0}
    for i, level in enumerate(risk_level_override):
        if level is None:
            continue
        level_key = str(level).strip().upper()
        if level_key in risk_map:
            risk[i] = risk_map[level_key]
            status[i] = op_status_from_score(risk[i])
    for i, val in enumerate(shamal_override):
        if val is not None:
            shamal[i] = parse_bool(val)
    if MANUAL_SHAMAL_PERIODS:
        for s, e in MANUAL_SHAMAL_PERIODS:
            for d in days:
                if s <= d <= e:
                    shamal[idx[d]] = True

    # -----------------------------
    # VISUALIZATION
    # -----------------------------
    mpl.rcParams.update(
        {
            "font.family": "serif",
            "axes.unicode_minus": False,
            "figure.dpi": 200,
            "axes.titlesize": 14,
            "axes.labelsize": 12,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
            "axes.grid": True,
            "grid.alpha": 0.2,
            "grid.linewidth": 0.5,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )

    # Heatmap rows (fixed normalization ranges)
    params = [
        "Wind (kt)",
        "Gust (kt)",
        "Wave (m)",
        "Vis (km)",
        "Dir (deg)",
        "Risk (0-100)",
    ]
    data_matrix = np.vstack(
        [
            wind_kn,
            gust_kn,
            wave_m,
            vis_km,
            wdir_deg,
            risk,
        ]
    )

    # Fixed ranges for normalization (stable colors)
    ranges = [
        (0.0, 35.0),  # wind kt
        (0.0, 45.0),  # gust kt
        (0.0, 2.5),  # wave m
        (0.0, 10.0),  # vis km
        (0.0, 360.0),  # dir deg
        (0.0, 100.0),  # risk
    ]

    data_norm = np.zeros_like(data_matrix, dtype=float)
    for r, (mn, mx) in enumerate(ranges):
        data_norm[r] = np.clip((data_matrix[r] - mn) / (mx - mn + 1e-9), 0.0, 1.0)

    cmap = LinearSegmentedColormap.from_list("risk", THEME["cmap"], N=256)

    fig = plt.figure(figsize=(20, 14))  # Slightly reduced width for better layout
    fig.patch.set_facecolor("none")
    fig.patch.set_alpha(0)
    gs = fig.add_gridspec(
        3,
        2,
        height_ratios=[2.2, 1.0, 0.6],
        width_ratios=[1, 0.05],
        hspace=0.3,
        wspace=0.05,
    )

    ax1 = fig.add_subplot(gs[0, 0])
    ax2 = fig.add_subplot(gs[1, 0], sharex=ax1)
    ax3 = fig.add_subplot(gs[2, 0], sharex=ax1)
    ax_cbar = fig.add_subplot(gs[0, 1])
    for ax in (ax1, ax2, ax3, ax_cbar):
        ax.set_facecolor("none")
        ax.patch.set_alpha(0)

    date_labels = [d.strftime("%m/%d") for d in days]
    x_limits = (-0.5, n - 0.5)
    tick_step = max(1, n // 15)  # Adjusted for fewer days
    x_ticks = list(range(0, n, tick_step))
    x_tick_labels = [date_labels[i] for i in x_ticks]

    # Heatmap
    im = ax1.imshow(
        data_norm,
        aspect="auto",
        cmap=cmap,
        interpolation="nearest",
        extent=[-0.5, n - 0.5, -0.5, len(params) - 0.5],
    )

    ax1.set_yticks(range(len(params)))
    ax1.set_yticklabels(params, fontsize=12, fontweight="bold")
    ax1.set_xlim(x_limits)
    ax1.set_xticks(x_ticks)
    ax1.tick_params(labelbottom=False)
    ax1.grid(False)

    # annotate every tick_step
    for r in range(len(params)):
        for c in range(0, n, tick_step):
            val = data_matrix[r, c]
            if np.isnan(val):
                continue
            if r in [0, 1, 5]:
                txt = f"{val:.0f}"
            elif r == 2:
                txt = f"{val:.1f}"
            elif r == 3:
                txt = f"{val:.1f}"
            elif r == 4:
                txt = f"{val:.0f}"
            else:
                txt = f"{val:.0f}"
            color_txt = "white" if data_norm[r, c] > 0.5 else "black"
            ax1.text(c, r, txt, ha="center", va="center", fontsize=8, color=color_txt)

    ax1.set_title(
        "AGI TR Transportation - Weather Risk Heatmap (Jan-Feb 2026, Multi-Source)",
        fontsize=16,
        fontweight="bold",
        pad=15,
    )
    cbar = fig.colorbar(im, cax=ax_cbar, orientation="vertical", aspect=30)
    cbar.set_label("Normalized (fixed ranges)", fontsize=12)

    # Risk timeline
    ax2.fill_between(range(n), risk, alpha=0.15)
    ax2.plot(range(n), risk, "o-", linewidth=2, markersize=4)
    ax2.axhspan(0, 30, color=THEME["risk_band"]["GO"], alpha=0.3, zorder=0)
    ax2.axhspan(30, 60, color=THEME["risk_band"]["HOLD"], alpha=0.3, zorder=0)
    ax2.axhspan(60, 100, color=THEME["risk_band"]["NO-GO"], alpha=0.3, zorder=0)
    ax2.axhline(y=30, linestyle="--", linewidth=1.5, label="GO Threshold (30)")
    ax2.axhline(y=60, linestyle="--", linewidth=1.5, label="NO-GO Threshold (60)")

    # Voyage overlays
    voyage_colors = THEME["voyage"]
    for v in VOYAGES:
        if v["end"] < START_DATE or v["start"] > END_DATE:
            continue
        s = max(v["start"], START_DATE)
        e = min(v["end"], END_DATE)
        xs = idx[s]
        xe = idx[e]
        color = voyage_colors.get(v["type"], voyage_colors["default"])
        ax2.axvspan(xs, xe, alpha=0.1, color=color, zorder=0)
        mid = (xs + xe) / 2
        ax2.text(
            mid,
            92,
            f'{v["name"]}\n{v["label"]}',
            ha="center",
            va="top",
            fontsize=10,
            fontweight="bold",
            color=color,
            bbox=dict(
                boxstyle="round,pad=0.4", facecolor="white", alpha=0.95, edgecolor=color
            ),
        )

    ax2.set_xlim(x_limits)
    ax2.set_ylim(0, 100)
    ax2.set_xticks(x_ticks)
    ax2.tick_params(labelbottom=False)
    ax2.set_ylabel("Risk Score (0-100)", fontsize=12, fontweight="bold")
    ax2.set_title(
        "Composite Weather Risk Score (Ensemble + Marine + Archive/Climate)",
        fontsize=14,
        fontweight="bold",
        pad=15,
    )
    ax2.legend(loc="upper right", fontsize=10, framealpha=0.95)
    ax2.grid(True, alpha=0.25)

    # Operation status bar
    status_colors = THEME["status"]
    ax3.bar(
        range(n),
        [1] * n,
        color=[status_colors[s] for s in status],
        edgecolor="white",
        linewidth=0.5,
    )

    ax3.set_xlim(x_limits)
    ax3.set_ylim(0, 1.4)
    ax3.set_xticks(x_ticks)
    ax3.set_xticklabels(x_tick_labels, rotation=45, ha="right", fontsize=10)
    ax3.tick_params(axis="x", pad=10)
    ax3.set_yticks([])
    ax3.set_title(
        "Daily Operation Status (GO / HOLD / NO-GO)",
        fontsize=14,
        fontweight="bold",
        pad=15,
    )
    ax3.grid(False)

    go_patch = mpatches.Patch(color=status_colors["GO"], label="GO (Risk < 30)")
    hold_patch = mpatches.Patch(color=status_colors["HOLD"], label="HOLD (30-60)")
    nogo_patch = mpatches.Patch(color=status_colors["NO-GO"], label="NO-GO (>=60)")
    ax3.legend(
        handles=[go_patch, hold_patch, nogo_patch],
        loc="upper left",
        ncol=3,
        fontsize=10,
        framealpha=0.95,
    )

    # Summary box (no emoji -> avoid glyph warnings)
    go_n = status.count("GO")
    hold_n = status.count("HOLD")
    nogo_n = status.count("NO-GO")
    shamal_n = int(shamal.sum())

    stats_text = (
        "Weather Analysis Summary\n"
        "------------------------\n"
        f"Period: {START_DATE.isoformat()} to {END_DATE.isoformat()} ({n} days)\n"
        f"GO Days: {go_n} ({go_n/n*100:.2f}%)\n"
        f"HOLD Days: {hold_n} ({hold_n/n*100:.2f}%)\n"
        f"NO-GO Days: {nogo_n} ({nogo_n/n*100:.2f}%)\n"
        f"Shamal Detected Days (NW+Strong): {shamal_n}\n"
        f"Max Gust (kt): {np.nanmax(gust_kn):.2f}\n"
        f"Max Wave (m): {np.nanmax(wave_m):.2f}\n"
    )
    cov_counts = {k: int(np.sum(coverage == k)) for k in np.unique(coverage) if k}
    cov_lines = "\n".join([f"{k}: {v}" for k, v in cov_counts.items()])
    cov_text = (
        "Data Coverage\n"
        "-------------\n"
        f"{cov_lines}\n"
        "\nNote: CLIMATE_FILL is modelled baseline, not actual measurement."
    )
    stats_box = AnchoredText(
        stats_text,
        loc="lower left",
        prop={"size": 10, "family": "monospace", "weight": "bold"},
        pad=1.0,
        borderpad=1.0,
        frameon=True,
    )
    stats_box.patch.set_facecolor("#F5F5F5")
    stats_box.patch.set_alpha(0.3)
    stats_box.patch.set_edgecolor("#4D4D4D")
    ax2.add_artist(stats_box)

    cov_box = AnchoredText(
        cov_text,
        loc="lower right",
        prop={"size": 9, "family": "monospace"},
        pad=1.0,
        borderpad=1.0,
        frameon=True,
    )
    cov_box.patch.set_facecolor("#F0F4F8")
    cov_box.patch.set_alpha(0.9)
    cov_box.patch.set_edgecolor("blue")
    ax1.add_artist(cov_box)

    # Shamal highlight (detected)
    for i in range(n):
        if shamal[i]:
            for ax in (ax1, ax2, ax3):
                ax.axvspan(i - 0.5, i + 0.5, alpha=0.1, color=THEME["shamal"], zorder=0)

    plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.15)

    plt.savefig(OUTPUT_PATH, dpi=200, bbox_inches="tight", facecolor="none", transparent=True)
    plt.close()

    print(f"OK: Heatmap generated -> {OUTPUT_PATH}")
    print(f"GO/HOLD/NO-GO: {go_n}/{hold_n}/{nogo_n} (days)")
    print(f"Detected Shamal days: {shamal_n}")
    print("Coverage:", cov_counts)


if __name__ == "__main__":
    main()
