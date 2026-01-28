## 1) Executive Summary (패치 범위)

* 기존 **가짜(패턴/난수) Wind/Wave** 대신, **외부 API(Open-Meteo) 실데이터**로 **최소 5개 지점**(포트/해상) 일별 지표를 생성하도록 교체했습니다. ([Open-Meteo][1])
* **Forecast는 최대 16일** 제약이 있어, **Hybrid 모드(최근/예보 + 과거동기간 Climatology)**를 기본값으로 넣었습니다. ([Open-Meteo][2])
* Marine 파라미터는 **Open-Meteo Marine API(wave_height)** 사용, 폰트 경고는 **이모지 제거**로 해결했습니다. ([Open-Meteo][3])

ENG-KR 1-liner: “Synthetic → Open-Meteo (Forecast+Archive+Marine) with ≥5 locations, Hybrid fill beyond 16-day forecast.”

---

## 2) Visual-first (5+ 지점 정의)

| Site (UN/LOCODE/Ref)             |   Lat |   Lon | Note                                    |
| -------------------------------- | ----: | ----: | --------------------------------------- |
| Mina Zayed / Abu Dhabi (AEMZD)   | 24.52 | 54.37 | Port 기준점 ([Vocabulary UNCEFACT][4])     |
| Mina Khalifa / Abu Dhabi (AEKHL) | 24.83 | 54.67 | Khalifa Port ([Vocabulary UNCEFACT][4]) |
| Ruwais Port (AERWP)              | 24.10 | 52.72 | 서부 해역 대표 ([Vocabulary UNCEFACT][4])     |
| Al Dhannah (Ruwais 인근)           | 24.10 | 52.58 | 연안/항만권 대표 ([Wikipedia][5])              |
| Das Island (해상)                  | 25.15 | 52.87 | Offshore 대표 ([Wikipedia][6])            |

---

## 3) Options (운영용 선택지)

1. **Option A — HYBRID (추천)**

   * Pros: 예보(최대 16일) + 과거동기간 평균으로 **Jan–Feb 전체 커버** ([Open-Meteo][2])
   * Cons: 16일 이후는 “예보”가 아니라 “동기간 평균(Proxy)”
   * Cost: API free(키 불필요) ([Open-Meteo][1])
   * Risk/Time: 낮음 / 즉시

2. **Option B — FORECAST ONLY**

   * Pros: 순수 예보 기반
   * Cons: 16일 이후 공백 ([Open-Meteo][2])
   * Risk: 일정 전체 히트맵엔 부적합

3. **Option C — 상용 Marine/Met API (예: Xweather 등)**

   * Pros: 장기/해상 파라미터 옵션 확대 가능
   * Cons: 유료/계약/키관리 필요 ([Xweather][7])

---

## 4) Patch Code (외부 데이터 ≥5지점 조회 + 히트맵 재생성)

아래 파일 하나로 대체하세요: `weather_risk_heatmap_external.py`

* **외부 호출 포함** → 인터넷 되는 PC에서 실행 필요
* 기본: `--mode hybrid` (예보 16일 + 과거동기간 평균으로 나머지 채움)

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGI TR Transportation - Weather Risk Heatmap (External Data Patch)
- Replaces synthetic/random weather with Open-Meteo external data (>=5 locations)
- Hybrid mode: forecast (up to 16 days) + climatology proxy (historical mean) for remaining horizon
Refs:
- Open-Meteo Forecast API: up to 16 days, supports past_days (archived forecasts)  https://open-meteo.com/en/docs
- Open-Meteo Historical Weather API: reanalysis-based archive             https://open-meteo.com/en/docs/historical-weather-api
- Open-Meteo Marine API: wave_height                                      https://open-meteo.com/en/docs/marine-weather-api
"""

from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import requests

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap

# ----------------------------
# Config
# ----------------------------
TZ = "Asia/Dubai"
FORECAST_MAX_DAYS = 16  # Open-Meteo forecast horizon (best practice)
CACHE_DIR = ".cache_openmeteo"
OUT_DIR_DEFAULT = "./out"

# Font settings (remove emojis to avoid glyph warnings)
plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.unicode_minus"] = False

@dataclass(frozen=True)
class Site:
    name: str
    lat: float
    lon: float
    tag: str  # short ref

# >=5 locations (user can edit/add)
SITES: List[Site] = [
    Site("Mina Zayed / Abu Dhabi (AEMZD)", 24.516666, 54.366665, "MZD"),
    Site("Mina Khalifa / Abu Dhabi (AEKHL)", 24.833334, 54.666668, "KHL"),
    Site("Ruwais Port (AERWP)", 24.100000, 52.716667, "RWP"),
    Site("Al Dhannah (Ruwais area)", 24.10333, 52.58361, "DHN"),
    Site("Das Island (Offshore)", 25.15139, 52.87361, "DAS"),
    Site("Upper Zakum Field (Offshore Abu Dhabi)", 24.844434, 53.653288, "UZK"),
]

# ----------------------------
# Open-Meteo helpers
# ----------------------------
def _cache_path(key: str) -> str:
    os.makedirs(CACHE_DIR, exist_ok=True)
    safe = key.replace("/", "_").replace("?", "_").replace("&", "_").replace("=", "_")
    return os.path.join(CACHE_DIR, f"{safe}.json")

def http_get_json(url: str, params: Dict, cache_key: str, timeout_s: int = 30) -> Dict:
    path = _cache_path(cache_key)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    r = requests.get(url, params=params, timeout=timeout_s)
    r.raise_for_status()
    data = r.json()

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    return data

def fetch_forecast_hourly(site: Site, start: date, end: date, past_days: int = 7, forecast_days: int = 16) -> pd.DataFrame:
    """
    Forecast API (also can return past_days of archived forecast) for wind/gust/visibility.
    """
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": site.lat,
        "longitude": site.lon,
        "timezone": TZ,
        "wind_speed_unit": "kn",  # knots
        "hourly": "wind_speed_10m,wind_gusts_10m,visibility",
        "past_days": past_days,
        "forecast_days": forecast_days,
    }
    key = f"forecast_hourly_{site.tag}_{past_days}_{forecast_days}"
    data = http_get_json(url, params, key)

    hourly = data.get("hourly", {})
    df = pd.DataFrame({
        "time": pd.to_datetime(hourly.get("time", [])),
        "wind_kt": hourly.get("wind_speed_10m", []),
        "gust_kt": hourly.get("wind_gusts_10m", []),
        # visibility is typically meters
        "vis_m": hourly.get("visibility", []),
    })
    if df.empty:
        return df

    df = df.set_index("time").sort_index()
    df = df.loc[(df.index.date >= start) & (df.index.date <= end)]
    return df

def fetch_archive_hourly(site: Site, start: date, end: date) -> pd.DataFrame:
    """
    Historical Weather API (reanalysis) for wind/gust/visibility.
    """
    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {
        "latitude": site.lat,
        "longitude": site.lon,
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "timezone": TZ,
        "wind_speed_unit": "kn",
        "hourly": "wind_speed_10m,wind_gusts_10m,visibility",
    }
    key = f"archive_hourly_{site.tag}_{start}_{end}"
    data = http_get_json(url, params, key)

    hourly = data.get("hourly", {})
    df = pd.DataFrame({
        "time": pd.to_datetime(hourly.get("time", [])),
        "wind_kt": hourly.get("wind_speed_10m", []),
        "gust_kt": hourly.get("wind_gusts_10m", []),
        "vis_m": hourly.get("visibility", []),
    })
    if df.empty:
        return df
    return df.set_index("time").sort_index()

def fetch_marine_hourly(site: Site, start: date, end: date) -> pd.DataFrame:
    """
    Marine API for wave_height (m).
    """
    url = "https://marine-api.open-meteo.com/v1/marine"
    params = {
        "latitude": site.lat,
        "longitude": site.lon,
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "timezone": TZ,
        "hourly": "wave_height",
    }
    key = f"marine_hourly_{site.tag}_{start}_{end}"
    data = http_get_json(url, params, key)

    hourly = data.get("hourly", {})
    df = pd.DataFrame({
        "time": pd.to_datetime(hourly.get("time", [])),
        "wave_m": hourly.get("wave_height", []),
    })
    if df.empty:
        return df
    return df.set_index("time").sort_index()

# ----------------------------
# Aggregation & Risk
# ----------------------------
def daily_from_hourly(wx: pd.DataFrame, marine: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    Build daily metrics:
      - wind_max (kt), gust_max (kt), wave_max (m), vis_min (km)
    """
    if wx.empty:
        return pd.DataFrame()

    d = wx.copy()
    d["vis_km"] = d["vis_m"] / 1000.0

    daily = pd.DataFrame(index=pd.to_datetime(sorted(set(d.index.date))))
    daily.index.name = "date"
    daily["wind_kt_max"] = d["wind_kt"].groupby(d.index.date).max().values
    daily["gust_kt_max"] = d["gust_kt"].groupby(d.index.date).max().values
    daily["vis_km_min"]  = d["vis_km"].groupby(d.index.date).min().values

    if marine is not None and not marine.empty:
        m = marine.copy()
        wave = m["wave_m"].groupby(m.index.date).max()
        # Align by date
        daily["wave_m_max"] = [wave.get(dt, np.nan) for dt in daily.index.date]
    else:
        daily["wave_m_max"] = np.nan

    # Fallback: if wave missing, approximate from wind (same spirit as original)
    # (kept explicit as an internal fallback; adjust coefficient if you have site-specific calibration)
    missing_wave = daily["wave_m_max"].isna()
    daily.loc[missing_wave, "wave_m_max"] = np.clip(daily.loc[missing_wave, "wind_kt_max"] * 0.045, 0.30, 1.80)

    # Sea state proxy (kept consistent with original logic)
    daily["sea_state"] = np.clip(daily["wind_kt_max"] / 5.0, 1.0, 6.0)

    return daily

def calc_risk_score(wind_kt: np.ndarray, gust_kt: np.ndarray, wave_m: np.ndarray, vis_km: np.ndarray) -> np.ndarray:
    wind_risk = np.clip((wind_kt - 10.0) * 4.0, 0.0, 40.0)
    gust_risk = np.clip((gust_kt - 15.0) * 2.5, 0.0, 25.0)
    wave_risk = np.clip((wave_m - 0.5) * 40.0, 0.0, 20.0)
    vis_risk  = np.clip((8.0 - vis_km) * 3.0, 0.0, 15.0)
    return wind_risk + gust_risk + wave_risk + vis_risk

def op_status_from_risk(r: np.ndarray) -> List[str]:
    out = []
    for x in r:
        if x < 30.0:
            out.append("GO")
        elif x < 60.0:
            out.append("HOLD")
        else:
            out.append("NO-GO")
    return out

def detect_shamal_windows(wind_kt_max: np.ndarray, risk: np.ndarray,
                          wind_thr: float = 20.0, risk_thr: float = 75.0, min_len: int = 2) -> List[Tuple[int,int]]:
    """
    Simple detection: day is "shamal-like" if wind>=wind_thr OR risk>=risk_thr.
    Returns list of (start_idx, end_idx) inclusive.
    """
    flag = (wind_kt_max >= wind_thr) | (risk >= risk_thr)
    windows = []
    i = 0
    n = len(flag)
    while i < n:
        if not flag[i]:
            i += 1
            continue
        j = i
        while j < n and flag[j]:
            j += 1
        if (j - i) >= min_len:
            windows.append((i, j-1))
        i = j
    return windows

# ----------------------------
# Data build modes
# ----------------------------
def build_daily_site(site: Site, start: date, days: int, mode: str, climo_years: int = 5) -> pd.DataFrame:
    """
    mode:
      - forecast: forecast API only (limited horizon)
      - climo: historical mean of previous years for the target window
      - hybrid: forecast (incl past_days) for near window, climo for remaining
    """
    end = start + timedelta(days=days-1)
    today = datetime.now().date()

    if mode == "climo":
        # Average same calendar window over last N years (proxy)
        frames = []
        for y in range(today.year - climo_years, today.year):
            s = date(y, start.month, start.day)
            e = date(y, end.month, end.day)
            wx = fetch_archive_hourly(site, s, e)
            marine = fetch_marine_hourly(site, s, e)
            d = daily_from_hourly(wx, marine)
            d.index = pd.to_datetime([date(2000, dt.month, dt.day) for dt in d.index.date])  # align by month/day
            frames.append(d)

        if not frames:
            return pd.DataFrame()

        merged = pd.concat(frames).groupby(level=0).mean(numeric_only=True)
        # Map aligned month/day back to target year
        target_dates = [start + timedelta(days=i) for i in range(days)]
        idx = pd.to_datetime([date(2000, d.month, d.day) for d in target_dates])
        out = merged.reindex(idx).copy()
        out.index = pd.to_datetime(target_dates)
        out["source"] = "CLIMO"
        return out

    # forecast or hybrid
    # 1) Forecast window available (past_days + forecast_days)
    wx_fc = fetch_forecast_hourly(site, start, end, past_days=7, forecast_days=FORECAST_MAX_DAYS)
    # Marine forecast does not support forecast_days on same endpoint; use marine API with date bounds (may return limited future)
    # We call marine API for the full window; if it returns partial, fallback wave approximation handles gaps.
    marine_all = fetch_marine_hourly(site, start, end)

    daily_fc = daily_from_hourly(wx_fc, marine_all)
    if daily_fc.empty:
        daily_fc = pd.DataFrame(index=pd.to_datetime([start + timedelta(days=i) for i in range(days)]))

    daily_fc.index = pd.to_datetime(daily_fc.index)
    daily_fc = daily_fc.reindex(pd.to_datetime([start + timedelta(days=i) for i in range(days)]))

    # Mark which days are actually covered by forecast response
    daily_fc["source"] = np.where(daily_fc["wind_kt_max"].notna(), "FORECAST", "MISSING")

    if mode == "forecast":
        # keep missing days as-is (user sees gaps or fallback wave only)
        return daily_fc

    # 2) Hybrid fill beyond forecast with climatology proxy
    # Identify missing wind/gust/vis (forecast horizon overflow)
    need_fill = daily_fc["wind_kt_max"].isna() | daily_fc["gust_kt_max"].isna() | daily_fc["vis_km_min"].isna()

    if need_fill.any():
        climo = build_daily_site(site, start, days, mode="climo", climo_years=climo_years)
        for col in ["wind_kt_max", "gust_kt_max", "vis_km_min", "wave_m_max", "sea_state"]:
            if col in climo.columns:
                daily_fc.loc[need_fill, col] = climo.loc[need_fill, col].values
        daily_fc.loc[need_fill, "source"] = "CLIMO"

    return daily_fc

def aggregate_sites(daily_sites: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Returns:
      - agg: worst-case daily metrics across sites
      - site_risk: per-site risk scores (rows=site, cols=day index)
    """
    # Ensure aligned index
    any_df = next(iter(daily_sites.values()))
    idx = any_df.index

    wind = []
    gust = []
    wave = []
    vis  = []
    site_names = []

    for name, df in daily_sites.items():
        site_names.append(name)
        wind.append(df["wind_kt_max"].to_numpy())
        gust.append(df["gust_kt_max"].to_numpy())
        wave.append(df["wave_m_max"].to_numpy())
        vis.append(df["vis_km_min"].to_numpy())

    wind = np.vstack(wind)
    gust = np.vstack(gust)
    wave = np.vstack(wave)
    vis  = np.vstack(vis)

    # Worst-case rule: max wind/gust/wave, min visibility
    agg = pd.DataFrame(index=idx)
    agg["wind_kt_max"] = np.nanmax(wind, axis=0)
    agg["gust_kt_max"] = np.nanmax(gust, axis=0)
    agg["wave_m_max"]  = np.nanmax(wave, axis=0)
    agg["vis_km_min"]  = np.nanmin(vis, axis=0)
    agg["sea_state"]   = np.clip(agg["wind_kt_max"] / 5.0, 1.0, 6.0)

    risk_sites = calc_risk_score(wind, gust, wave, vis)  # vectorized per-site
    site_risk = pd.DataFrame(risk_sites, index=site_names, columns=range(len(idx)))

    return agg, site_risk

# ----------------------------
# Plot
# ----------------------------
def plot_heatmap(start: date, days: int, voyages: List[Dict], mode: str, out_path: str):
    end = start + timedelta(days=days-1)
    dates = [start + timedelta(days=i) for i in range(days)]
    date_labels = [d.strftime("%m/%d") for d in dates]

    # Build per-site daily
    daily_sites: Dict[str, pd.DataFrame] = {}
    for s in SITES:
        d = build_daily_site(s, start, days, mode=mode, climo_years=5)
        daily_sites[s.tag] = d

    # Aggregate
    agg, site_risk = aggregate_sites(daily_sites)

    # Risk score (aggregate worst-case)
    risk = calc_risk_score(
        agg["wind_kt_max"].to_numpy(),
        agg["gust_kt_max"].to_numpy(),
        agg["wave_m_max"].to_numpy(),
        agg["vis_km_min"].to_numpy()
    )
    status = op_status_from_risk(risk)

    # Detect shamal-like windows based on data
    shamal_windows = detect_shamal_windows(agg["wind_kt_max"].to_numpy(), risk)

    # Figure layout: add site-risk heatmap (>=5 rows)
    fig = plt.figure(figsize=(22, 18))
    ax1 = plt.subplot2grid((5, 1), (0, 0), rowspan=2)  # main param heatmap
    axS = plt.subplot2grid((5, 1), (2, 0), rowspan=1)  # site risk heatmap
    ax2 = plt.subplot2grid((5, 1), (3, 0), rowspan=1)  # risk timeline
    ax3 = plt.subplot2grid((5, 1), (4, 0), rowspan=1)  # status bar

    params = ["Wind (kt, max)", "Gust (kt, max)", "Wave (m, max)", "Visibility (km, min)", "Sea State", "Risk Score"]
    data_matrix = np.array([
        agg["wind_kt_max"].to_numpy(),
        agg["gust_kt_max"].to_numpy(),
        agg["wave_m_max"].to_numpy(),
        agg["vis_km_min"].to_numpy(),
        agg["sea_state"].to_numpy(),
        risk
    ])

    colors = ["#2E7D32","#4CAF50","#8BC34A","#CDDC39","#FFEB3B","#FFC107","#FF9800","#FF5722","#D32F2F"]
    cmap = LinearSegmentedColormap.from_list("risk", colors, N=100)

    # Normalize rows for heatmap color
    data_norm = np.zeros_like(data_matrix, dtype=float)
    for i in range(len(params)):
        row = data_matrix[i]
        rmin, rmax = np.nanmin(row), np.nanmax(row)
        data_norm[i] = (row - rmin) / (rmax - rmin + 1e-6)

    im = ax1.imshow(data_norm, aspect="auto", cmap=cmap, interpolation="nearest")
    ax1.set_yticks(range(len(params)))
    ax1.set_yticklabels(params, fontsize=11, fontweight="bold")

    ax1.set_xticks(range(0, days, 2))
    ax1.set_xticklabels([date_labels[i] for i in range(0, days, 2)], rotation=45, ha="right", fontsize=9)

    # Annotate (every 2 days) with real values
    for i in range(len(params)):
        for j in range(days):
            if j % 2 != 0:
                continue
            if i == 0: val = f"{data_matrix[i, j]:.0f}"
            elif i == 1: val = f"{data_matrix[i, j]:.0f}"
            elif i == 2: val = f"{data_matrix[i, j]:.1f}"
            elif i == 3: val = f"{data_matrix[i, j]:.1f}"
            elif i == 4: val = f"{data_matrix[i, j]:.1f}"
            else: val = f"{data_matrix[i, j]:.0f}"
            color = "white" if data_norm[i, j] > 0.6 else "black"
            ax1.text(j, i, val, ha="center", va="center", fontsize=7, color=color)

    # Highlight shamal windows
    for (a, b) in shamal_windows:
        ax1.axvspan(a - 0.5, b + 0.5, alpha=0.25, color="red", zorder=0)
    ax1.set_title(f"AGI TR Transportation - Weather Risk Heatmap (External Data, mode={mode})\n"
                  f"{start.isoformat()} ~ {end.isoformat()} | Sites={len(SITES)}",
                  fontsize=15, fontweight="bold", pad=16)

    cbar = plt.colorbar(im, ax=ax1, orientation="vertical", pad=0.02, aspect=30)
    cbar.set_label("Risk Level (Row-normalized)", fontsize=10)

    # --- Site risk heatmap (rows=sites, cols=days)
    site_names = list(site_risk.index)
    sr = site_risk.to_numpy()
    # normalize by row
    sr_norm = (sr - sr.min(axis=1, keepdims=True)) / (sr.max(axis=1, keepdims=True) - sr.min(axis=1, keepdims=True) + 1e-6)
    axS.imshow(sr_norm, aspect="auto", cmap=cmap, interpolation="nearest")
    axS.set_yticks(range(len(site_names)))
    axS.set_yticklabels(site_names, fontsize=10, fontweight="bold")
    axS.set_xticks(range(0, days, 2))
    axS.set_xticklabels([date_labels[i] for i in range(0, days, 2)], rotation=45, ha="right", fontsize=9)
    axS.set_title("Per-site Risk (Row-normalized) | Worst-case aggregation drives main chart", fontsize=12, fontweight="bold")

    # --- Risk timeline
    ax2.fill_between(range(days), risk, alpha=0.3, color="orange")
    ax2.plot(range(days), risk, "o-", color="#FF5722", linewidth=2, markersize=4)
    ax2.axhline(y=30, color="green", linestyle="--", linewidth=2, label="GO Threshold (30)")
    ax2.axhline(y=60, color="red", linestyle="--", linewidth=2, label="NO-GO Threshold (60)")
    for (a, b) in shamal_windows:
        ax2.axvspan(a, b, alpha=0.15, color="red")

    voyage_colors = {"transport": "#2196F3", "jackdown": "#9C27B0"}
    for v in voyages:
        color = voyage_colors.get(v["type"], "#999999")
        ax2.axvspan(v["start"], v["end"], alpha=0.12, color=color, zorder=0)
        mid = (v["start"] + v["end"]) / 2
        ax2.text(mid, 98, f"{v['name']}\n{v['tr']}", ha="center", va="top", fontsize=9,
                 fontweight="bold", color=color,
                 bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.85))

    ax2.set_xlim(-0.5, days - 0.5)
    ax2.set_ylim(0, 110)
    ax2.set_xticks(range(0, days, 2))
    ax2.set_xticklabels([date_labels[i] for i in range(0, days, 2)], rotation=45, ha="right", fontsize=9)
    ax2.set_ylabel("Risk Score (0-100)", fontsize=11, fontweight="bold")
    ax2.set_title("Composite Weather Risk Score with Voyage Schedule", fontsize=12, fontweight="bold")
    ax2.legend(loc="upper right", fontsize=9)
    ax2.grid(True, alpha=0.3)

    # --- Status bar
    status_colors = {"GO": "#4CAF50", "HOLD": "#FFC107", "NO-GO": "#F44336"}
    ax3.bar(range(days), [1] * days, color=[status_colors[s] for s in status], edgecolor="white", linewidth=0.5)

    for v in voyages:
        ax3.axvline(x=v["start"], color="blue", linestyle="-", linewidth=2, alpha=0.7)
        ax3.text(v["start"], 1.15, v["name"], fontsize=10, fontweight="bold", ha="center", color="blue")

    ax3.set_xlim(-0.5, days - 0.5)
    ax3.set_ylim(0, 1.4)
    ax3.set_xticks(range(0, days, 2))
    ax3.set_xticklabels([date_labels[i] for i in range(0, days, 2)], rotation=45, ha="right", fontsize=9)
    ax3.set_yticks([])
    ax3.set_title("Daily Operation Status (GO / HOLD / NO-GO)", fontsize=12, fontweight="bold")

    go_patch = mpatches.Patch(color=status_colors["GO"], label="GO (<30)")
    hold_patch = mpatches.Patch(color=status_colors["HOLD"], label="HOLD (30-60)")
    nogo_patch = mpatches.Patch(color=status_colors["NO-GO"], label="NO-GO (>=60)")
    ax3.legend(handles=[go_patch, hold_patch, nogo_patch], loc="upper right", ncol=3, fontsize=9)

    # --- Summary box (NO emojis)
    go_n = status.count("GO")
    hold_n = status.count("HOLD")
    nogo_n = status.count("NO-GO")
    stats_text = (
        "Weather Analysis Summary\n"
        "------------------------\n"
        f"Period: {start.isoformat()} ~ {end.isoformat()} ({days} days)\n"
        f"Sites: {len(SITES)} (worst-case aggregation)\n"
        f"GO Days: {go_n} ({go_n/days*100:.0f}%)\n"
        f"HOLD Days: {hold_n} ({hold_n/days*100:.0f}%)\n"
        f"NO-GO Days: {nogo_n} ({nogo_n/days*100:.0f}%)\n"
        "------------------------\n"
        f"Max Gust: {np.nanmax(agg['gust_kt_max']):.1f} kt\n"
        f"Max Wave: {np.nanmax(agg['wave_m_max']):.1f} m\n"
        f"Avg Risk: {np.nanmean(risk):.0f}/100\n"
        f"Detected Shamal-like windows: {len(shamal_windows)}"
    )
    fig.text(0.02, 0.01, stats_text, fontsize=9, fontfamily="DejaVu Sans",
             verticalalignment="bottom",
             bbox=dict(boxstyle="round", facecolor="wheat", alpha=0.9, edgecolor="black"))

    plt.tight_layout()
    plt.subplots_adjust(bottom=0.12, hspace=0.55)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    plt.savefig(out_path, dpi=160, bbox_inches="tight", facecolor="white")
    plt.close()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default="2026-01-06", help="YYYY-MM-DD")
    ap.add_argument("--days", type=int, default=47)
    ap.add_argument("--mode", choices=["hybrid","forecast","climo"], default="hybrid")
    ap.add_argument("--out", default=os.path.join(OUT_DIR_DEFAULT, "AGI_TR_Weather_Risk_Heatmap.png"))
    args = ap.parse_args()

    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    days = args.days

    # Keep your voyage blocks; indices are relative to start date
    voyages = [
        {"name": "V1", "start": 3,  "end": 7,  "tr": "TR1",     "type": "transport"},
        {"name": "V2", "start": 7,  "end": 20, "tr": "TR2+JD",  "type": "jackdown"},
        {"name": "V3", "start": 16, "end": 20, "tr": "TR3",     "type": "transport"},
        {"name": "V4", "start": 21, "end": 34, "tr": "TR4+JD",  "type": "jackdown"},
        {"name": "V5", "start": 27, "end": 31, "tr": "TR5",     "type": "transport"},
        {"name": "V6", "start": 32, "end": 45, "tr": "TR6+JD",  "type": "jackdown"},
        {"name": "V7", "start": 67, "end": 77, "tr": "TR7+JD-4", "type": "jackdown"},
    ]

    plot_heatmap(start, days, voyages, mode=args.mode, out_path=args.out)
    print(f"OK: Weather Risk Heatmap generated: {args.out}")

if __name__ == "__main__":
    main()
```

### requirements.txt

```txt
numpy
pandas
matplotlib
requests
```

### 실행 명령

```bash
pip install -r requirements.txt
python weather_risk_heatmap_external.py --mode hybrid --start 2026-01-06 --days 47 --out ./out/AGI_TR_Weather_Risk_Heatmap.png
```

---

## 5) Roadmap (Prepare→Pilot→Build→Operate→Scale + KPI)

* Prepare: 좌표(특히 **AGI 실좌표**) 확정, 운영 임계치(Go/Hold/No-Go) 합의
* Pilot: 1주 운영(예보 7~14일 구간) → 실제 Sea Condition/Harbour Master 판단과 비교
* Build: /logi-master weather-tie 모듈에 CSV export + 자동 첨부(메일/Teams)
* Operate: 매일 06:00 GST 자동 갱신 + “NO-GO/SHAMAL 윈도우” TG Alert
* Scale: 6개 현장/6개 항로 포인트 템플릿화
  KPI 예: “NO-GO 예측 적중률”, “Marine ops cancel 감소(%)”, “DEM/DET 회피(AED/day)”

---

## 6) Automation Hooks (RPA/LLM/Sheets/TG)

* (Sheets) `daily_risk.csv` 저장 → 간트/항차 일정과 조인
* (TG) `NO-GO ≥2days` 감지 시 Ops 채널 자동 알림
* (LLM) “윈도우 추천” 요약: next 72h GO window + berth/tide 체크리스트 생성

---

## 7) QA / Recheck (핵심만)

* 좌표: 포트/현장 기준점이 “작업해역(정박/접안)”과 일치하는지
* Marine wave: 일부 연안 좌표는 파라미터 결측 가능 → 지금 코드는 **결측 시 wind→wave fallback** 동작(로그 필요)
* Threshold: Risk 수식/Go-NoGo는 현장 기준으로 재보정 필요(특히 Jack-down vs Marine transit)

---

## 8) Accuracy (data-only / 가정)

* **가정:** `--mode hybrid`에서 **16일 이후 구간은 예보가 아니라 과거 동기간 평균(Climatology proxy)**로 채웁니다. ([Open-Meteo][2])
* **데이터 소스:** Open-Meteo Forecast/Archive(재분석) + Marine(wave_height). ([Open-Meteo][2])

---

## 9) CmdRec

* `/switch_mode ORACLE + /logi-master weather-tie --deep`
* `/visualize_data --type=heatmap <daily_risk.csv>`

원하시면, **AGI(Al Ghallan Island) 정확 좌표**만 주시면 6번째 포인트로 넣고(현장 최우선), “Route-weighted(항로가중 평균)” 집계로도 바꿔드리겠습니다.

[1]: https://open-meteo.com/?utm_source=chatgpt.com "Open-Meteo.com: 🌤️ Free Open-Source Weather API"
[2]: https://open-meteo.com/en/docs "️ Docs | Open-Meteo.com"
[3]: https://open-meteo.com/en/docs/marine-weather-api?utm_source=chatgpt.com "Marine Weather API"
[4]: https://vocabulary.uncefact.org/unlocode?country=ae "UN/LOCODE"
[5]: https://en.wikipedia.org/wiki/Al_Dhannah "Al Dhannah - Wikipedia"
[6]: https://en.wikipedia.org/wiki/Das_Island "Das Island - Wikipedia"
[7]: https://www.xweather.com/docs/weather-api/endpoints/maritime?utm_source=chatgpt.com "Maritime – Weather API"
