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

# 가정: 좌표는 AGI/항로 대표점(사용자 실제 AGI 좌표로 교체 권장)
# If you have exact AGI berth/ramp coordinate, replace these:
LAT = 24.12
LON = 52.53

# Update date range to focus on mid‑January through mid‑February.
# With this change the heatmap will only display data between 2026‑01‑15 and
# 2026‑02‑15 (inclusive). Anything outside this window is filtered out so
# that the visualization remains concise and relevant.
START_DATE = date(2026, 1, 15)
END_DATE = date(2026, 2, 15)  # inclusive

OUTPUT_PATH = "AGI_TR_Weather_Risk_Heatmap_v2.png"

# 날씨 데이터 소스 설정
USE_MANUAL_JSON = True  # True: JSON 파일 사용, False: API 사용
SCRIPT_DIR = (
    os.path.dirname(os.path.abspath(__file__))
    if "__file__" in globals()
    else os.getcwd()
)
WEATHER_JSON_PATH = os.path.join(SCRIPT_DIR, "weather_data_20260106.json")
WEATHER_REQUEST_PATH = os.path.join(SCRIPT_DIR, "weather_data_requests.txt")

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
        "end": date(2026, 2, 27),
        "label": "TR4+JD-2",
        "type": "jackdown",
    },
]

MANUAL_SHAMAL_PERIODS = [
    (date(2026, 2, 5), date(2026, 2, 14)),
]

THEME = {
    # For a more polished and luxurious appearance, the heatmap uses a smooth
    # sequential colour palette that transitions from light yellow (low risk) to
    # dark brown/red (high risk). This choice aids interpretation of risk
    # magnitudes and looks more refined than the previous flat palette.
    "cmap": [
        "#ffffe5",
        "#fff7bc",
        "#fee391",
        "#fec44f",
        "#fe9929",
        "#ec7014",
        "#cc4c02",
        "#993404",
        "#662506",
    ],
    "status": {
        "GO": "#2F6B4F",
        "HOLD": "#C8A64B",
        "NO-GO": "#A23B3B",
    },
    "risk_band": {
        "GO": "#E6F2E7",
        "HOLD": "#FFF4D6",
        "NO-GO": "#F6D6D6",
    },
    "shamal": "#7F1D1D",
    "voyage": {
        "transport": "#2C5D7A",
        "jackdown": "#6B4C9A",
        "default": "#7A7A7A",
    },
}

# -----------------------------
# API ENDPOINTS (Open-Meteo)
# -----------------------------
# Forecast API endpoint /v1/forecast (up to 16 days)  (models auto-selected)  :contentReference[oaicite:9]{index=9}
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

# Provider-specific endpoints (model feeds)
# /v1/gfs  :contentReference[oaicite:10]{index=10}
# /v1/ecmwf :contentReference[oaicite:11]{index=11}
# /v1/dwd-icon :contentReference[oaicite:12]{index=12}
# /v1/gem :contentReference[oaicite:13]{index=13}
MODEL_URLS = {
    "forecast": FORECAST_URL,
    "gfs": "https://api.open-meteo.com/v1/gfs",
    "ecmwf": "https://api.open-meteo.com/v1/ecmwf",
    "dwd": "https://api.open-meteo.com/v1/dwd-icon",
    "gem": "https://api.open-meteo.com/v1/gem",
}

# Historical Weather API (archive-api)  :contentReference[oaicite:14]{index=14}
ARCHIVE_URL = "https://archive-api.open-meteo.com/v1/archive"

# Marine Weather API endpoint /v1/marine  :contentReference[oaicite:15]{index=15}
MARINE_URL = "https://marine-api.open-meteo.com/v1/marine"

# Climate API endpoint /v1/climate (gap fill; not “actual measurements”)  :contentReference[oaicite:16]{index=16}
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
        print(f"[WARN] 빈 JSON 파일 생성: {json_path}")
    except Exception as e:
        print(f"[ERROR] 빈 JSON 파일 생성 실패: {e}")


def load_weather_data_from_json(json_path, start_date=None, end_date=None):
    """
    수동 입력된 날씨 데이터 JSON 파일 로드
    API 대신 사용하는 함수
    """
    if not os.path.exists(json_path):
        print(f"[WARN] 날씨 데이터 JSON 파일을 찾을 수 없습니다: {json_path}")
        return []

    if hasattr(start_date, "date"):
        start_date = start_date.date()
    if hasattr(end_date, "date"):
        end_date = end_date.date()

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"[ERROR] 날씨 데이터 JSON 로드 오류: {e}")
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
            print(f"[WARN] 날짜 범위 밖 데이터 무시: {date_str}")
            continue

        weather_records.append(
            {
                "date": date_str,
                "wind_max_kn": record.get("wind_max_kn"),
                "gust_max_kn": record.get("gust_max_kn"),
                "wind_dir_deg": record.get("wind_dir_deg"),
                "wave_max_m": record.get("wave_max_m"),
                "visibility_km": record.get("visibility_km"),
                "source": record.get("source", "MANUAL"),
                "notes": record.get("notes", ""),
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
    r = requests.get(url, params=params, timeout=timeout)
    r.raise_for_status()
    return r.json()


# -----------------------------
# FETCH WEATHER (DAILY + VISIBILITY MIN from HOURLY)
# -----------------------------
def fetch_weather_model(model_name: str, base_url: str, d0: date, d1: date) -> dict:
    """
    Returns dict with keys:
      dates (list[str]), wind_max_kn, gust_max_kn, wind_dir_deg, vis_min_km
    Notes:
      - wind_speed_unit supports 'kn' for many endpoints  :contentReference[oaicite:17]{index=17}
      - visibility is hourly in meters in many models; we compute daily min km
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
    Uses Historical Weather API /v1/archive (archive-api.open-meteo.com)  :contentReference[oaicite:18]{index=18}
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

    # Source coverage tracker
    coverage = np.array([""] * n, dtype=object)

    # ============================================
    # 수동 입력 JSON 파일에서 날씨 데이터 로드 (API 대신)
    # ============================================
    if USE_MANUAL_JSON:
        ensure_weather_json(WEATHER_JSON_PATH)
        weather_records = load_weather_data_from_json(
            WEATHER_JSON_PATH, start_date=START_DATE, end_date=END_DATE
        )

        if weather_records:
            # JSON 데이터를 배열에 채우기
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
                except Exception as e:
                    print(
                        f"[WARN] 날씨 데이터 처리 오류 ({record.get('date', 'Unknown')}): {e}"
                    )
        else:
            print("[WARN] 날씨 데이터를 로드할 수 없습니다.")
            print(
                "   weather_data_template.csv에 데이터 입력 후 convert_weather_csv_to_json.py 실행"
            )
            print("   또는 USE_MANUAL_JSON = False로 설정하여 API 모드 사용")

    # ============================================
    # API 모드 (USE_MANUAL_JSON = False일 때만 사용)
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
    # 갭 필: 누락된 데이터 보간 (공통)
    # ============================================
    print("\n[INFO] 데이터 완전성 확인 중...")
    missing_count = np.sum(np.isnan(wind_kn))
    if missing_count > 0:
        print(f"   [WARN] {missing_count}일의 데이터가 누락되었습니다.")
        print("   누락된 날짜에 대해 기본값 또는 보간 적용")

        # 간단한 보간 (선형 보간)
        for i in range(n):
            if np.isnan(wind_kn[i]):
                # 이전/다음 값으로 보간 시도
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
                    # 기본값
                    wind_kn[i] = 12.0
                    gust_kn[i] = 15.0
                    coverage[i] = "DEFAULT"

            # 파고 보간
            if np.isnan(wave_m[i]) and not np.isnan(wind_kn[i]):
                wave_m[i] = float(np.clip(wind_kn[i] * 0.04, 0.30, 2.50))

            # 가시거리 기본값
            if np.isnan(vis_km[i]):
                vis_km[i] = 8.00

            # Gust 보간
            if np.isnan(gust_kn[i]) and not np.isnan(wind_kn[i]):
                gust_kn[i] = wind_kn[i] * 1.30

    # 6) Risk + status
    risk = calc_risk_score(wind_kn, gust_kn, wave_m, vis_km)
    status = [op_status_from_score(s) for s in risk]
    shamal = np.array(
        [is_shamal_day(wdir_deg[i], wind_kn[i], gust_kn[i]) for i in range(n)],
        dtype=bool,
    )
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
            "font.family": "DejaVu Sans",
            "axes.unicode_minus": False,
            "figure.dpi": 150,
            "axes.titlesize": 13,
            "axes.labelsize": 11,
            "xtick.labelsize": 9,
            "ytick.labelsize": 9,
            "axes.grid": True,
            "grid.alpha": 0.25,
            "grid.linewidth": 0.6,
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

    fig, (ax1, ax2, ax3) = plt.subplots(
        3,
        1,
        figsize=(22, 14),
        sharex=True,
        gridspec_kw={"height_ratios": [2.2, 1.0, 0.6], "hspace": 0.25},
    )

    date_labels = [d.strftime("%m/%d") for d in days]
    x_limits = (-0.5, n - 0.5)
    tick_step = max(2, n // 20)
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
    ax1.set_yticklabels(params, fontsize=11, fontweight="bold")
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
            color_txt = "white" if data_norm[r, c] > 0.60 else "black"
            ax1.text(c, r, txt, ha="center", va="center", fontsize=7, color=color_txt)

    ax1.set_title(
        "AGI TR Transportation - Weather Risk Heatmap (Jan-Feb 2026, Multi-Source)",
        fontsize=16,
        fontweight="bold",
        pad=10,
    )
    cbar = plt.colorbar(im, ax=ax1, orientation="vertical", pad=0.02, aspect=30)
    cbar.set_label("Normalized (fixed ranges)", fontsize=10)

    # Colorbar 추가 후 모든 서브플롯의 x좌표와 너비를 동일하게 설정
    # ax1의 x좌표와 너비를 기준으로 ax2, ax3의 위치를 동기화
    pos1 = ax1.get_position()
    # ax2와 ax3의 x좌표(x0)와 너비(width)만 ax1과 동일하게 설정
    # y좌표(y0)와 높이(height)는 각 서브플롯의 고유 값을 유지
    for ax in [ax2, ax3]:
        pos = ax.get_position()
        # x0와 width만 ax1과 동일하게, y0와 height는 원래 값 유지
        ax.set_position([pos1.x0, pos.y0, pos1.width, pos.height])

    # Risk timeline
    ax2.fill_between(range(n), risk, alpha=0.18)
    ax2.plot(range(n), risk, "o-", linewidth=1.8, markersize=3.5)
    ax2.axhspan(0, 30, color=THEME["risk_band"]["GO"], alpha=0.35, zorder=0)
    ax2.axhspan(30, 60, color=THEME["risk_band"]["HOLD"], alpha=0.35, zorder=0)
    ax2.axhspan(60, 100, color=THEME["risk_band"]["NO-GO"], alpha=0.35, zorder=0)
    ax2.axhline(y=30, linestyle="--", linewidth=1.6, label="GO Threshold (30)")
    ax2.axhline(y=60, linestyle="--", linewidth=1.6, label="NO-GO Threshold (60)")

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
        ax2.axvspan(xs, xe, alpha=0.12, color=color, zorder=0)
        mid = (xs + xe) / 2
        ax2.text(
            mid,
            88,
            f'{v["name"]}\n{v["label"]}',
            ha="center",
            va="top",
            fontsize=9,
            fontweight="bold",
            color=color,
            bbox=dict(
                boxstyle="round,pad=0.3", facecolor="white", alpha=0.9, edgecolor=color
            ),
        )

    ax2.set_xlim(x_limits)
    ax2.set_ylim(0, 100)
    ax2.set_xticks(x_ticks)
    ax2.tick_params(labelbottom=False)
    ax2.set_ylabel("Risk Score (0-100)", fontsize=11, fontweight="bold")
    ax2.set_title(
        "Composite Weather Risk Score (Ensemble + Marine + Archive/Climate)",
        fontsize=13,
        fontweight="bold",
        pad=10,
    )
    ax2.legend(loc="upper right", fontsize=9, framealpha=0.9)
    ax2.grid(True, alpha=0.3)

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
    ax3.set_xticklabels(x_tick_labels, rotation=45, ha="right", fontsize=9)
    ax3.tick_params(axis="x", pad=8)
    ax3.set_yticks([])
    ax3.set_title(
        "Daily Operation Status (GO / HOLD / NO-GO)",
        fontsize=13,
        fontweight="bold",
        pad=10,
    )
    ax3.grid(False)

    go_patch = mpatches.Patch(color=status_colors["GO"], label="GO (Risk < 30)")
    hold_patch = mpatches.Patch(color=status_colors["HOLD"], label="HOLD (30-60)")
    nogo_patch = mpatches.Patch(color=status_colors["NO-GO"], label="NO-GO (>=60)")
    ax3.legend(
        handles=[go_patch, hold_patch, nogo_patch],
        loc="upper left",
        ncol=3,
        fontsize=9,
        framealpha=0.9,
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
    # Summary box와 Coverage box 위치를 겹치지 않도록 조정
    # stats_box는 ax2의 lower left, cov_box는 ax1의 lower right로 분리
    stats_box = AnchoredText(
        stats_text,
        loc="lower left",
        prop={"size": 9, "family": "monospace"},
        pad=0.8,  # 패딩 증가로 여백 확보
        frameon=True,
    )
    stats_box.patch.set_facecolor("wheat")
    stats_box.patch.set_alpha(0.90)
    stats_box.patch.set_edgecolor("black")
    ax2.add_artist(stats_box)

    # Coverage box를 ax1의 lower right로 이동하여 겹침 방지
    cov_box = AnchoredText(
        cov_text,
        loc="lower right",
        prop={"size": 8, "family": "monospace"},
        pad=0.8,  # 패딩 증가로 여백 확보
        frameon=True,
    )
    cov_box.patch.set_facecolor("#E8F5E9")
    cov_box.patch.set_alpha(0.90)
    cov_box.patch.set_edgecolor("green")
    ax1.add_artist(cov_box)  # ax2에서 ax1으로 변경, lower right 위치

    # Shamal highlight (detected)
    for i in range(n):
        if shamal[i]:
            for ax in (ax1, ax2, ax3):
                ax.axvspan(
                    i - 0.5, i + 0.5, alpha=0.12, color=THEME["shamal"], zorder=0
                )

    # 모든 서브플롯의 위치를 명시적으로 동일하게 설정
    # colorbar를 추가한 후 ax1의 위치를 기준으로 ax2, ax3 동기화
    pos1 = ax1.get_position()

    # ax2와 ax3의 x좌표(x0)와 너비(width)만 ax1과 동일하게 설정
    # y좌표(y0)와 높이(height)는 각 서브플롯의 고유 값을 유지
    for ax in [ax2, ax3]:
        pos = ax.get_position()
        # x0와 width만 ax1과 동일하게, y0와 height는 원래 값 유지
        ax.set_position([pos1.x0, pos.y0, pos1.width, pos.height])

    # subplots_adjust 호출 - 전체 레이아웃 조정
    # colorbar 공간을 고려하여 right 값 설정
    plt.subplots_adjust(
        left=0.07,
        right=0.92,  # colorbar 공간 확보
        top=0.93,
        bottom=0.18,
        hspace=0.25,
        wspace=0.0,
    )

    # subplots_adjust가 서브플롯 위치를 변경할 수 있으므로
    # 호출 후 즉시 위치를 강제로 동기화
    pos1_after = ax1.get_position()
    for ax in [ax2, ax3]:
        pos = ax.get_position()
        # x0와 width만 ax1과 동일하게 강제 설정, y0와 height는 원래 값 유지
        ax.set_position([pos1_after.x0, pos.y0, pos1_after.width, pos.height])

    # 최종 위치 동기화 확인 및 재설정
    # savefig 직전에 한 번 더 동기화하여 위치가 정확히 동일한지 보장
    pos1_final = ax1.get_position()
    for ax in [ax2, ax3]:
        pos = ax.get_position()
        # x0와 width가 ax1과 정확히 같은지 확인하고 동기화 (허용 오차 0.0001)
        if (
            abs(pos.x0 - pos1_final.x0) > 0.0001
            or abs(pos.width - pos1_final.width) > 0.0001
        ):
            ax.set_position([pos1_final.x0, pos.y0, pos1_final.width, pos.height])

    # bbox_inches="tight" 사용하지 않고 명시적인 여백 설정
    # tight는 서브플롯 위치를 변경할 수 있으므로 제거
    plt.savefig(
        OUTPUT_PATH, dpi=150, bbox_inches=None, facecolor="white", pad_inches=0.1
    )
    plt.close()

    print(f"OK: Heatmap generated -> {OUTPUT_PATH}")
    print(f"GO/HOLD/NO-GO: {go_n}/{hold_n}/{nogo_n} (days)")
    print(f"Detected Shamal days: {shamal_n}")
    print("Coverage:", cov_counts)


if __name__ == "__main__":
    import sys

    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except:
            pass
    main()
