"""
Generate a simulation image of the Weather & Marine Risk chart (4-day wind + wave).
Uses parsed summary data; saves PNG to out/weather_parsed/ or out/.
"""
from __future__ import annotations

import sys
from pathlib import Path

# CONVERT root
_ROOT = Path(__file__).resolve().parent.parent
_OUT = _ROOT / "out" / "weather_parsed"
_OUT.mkdir(parents=True, exist_ok=True)

# 4-day data from ADNOC/NCM summary (28-31 Jan 2026)
DATES = ["28 Jan", "29 Jan", "30 Jan", "31 Jan"]
WIND_KT = [11, 10.5, 14, 14]   # representative kt (7-15, 7-14, 8-16/20)
WAVE_FT = [2.75, 2.5, 4.0, 4.0]  # ft (2-3/4, 2-3, 2-4/6)
# Theme from schedule HTML: accent-amber
COLOR_BAR = "rgba(245, 158, 11, 0.85)"
COLOR_BAR2 = "rgba(6, 182, 212, 0.75)"
PAPER_BG = "#ffffff"
PLOT_BG = (0.97, 0.98, 0.99, 1.0)  # light gray for matplotlib
FONT_COLOR = "#374151"
GRID_COLOR = "rgba(0, 0, 0, 0.06)"


def _save_plotly(out_path: Path) -> bool:
    try:
        import plotly.graph_objects as go
        from plotly.subplots import make_subplots
    except ImportError:
        return False
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("Wind Speed (kt)", "Wave Height (ft)"),
        horizontal_spacing=0.12,
        column_widths=[0.5, 0.5],
    )
    fig.add_trace(
        go.Bar(x=DATES, y=WIND_KT, name="Wind (kt)", marker_color=COLOR_BAR),
        row=1, col=1,
    )
    fig.add_trace(
        go.Bar(x=DATES, y=WAVE_FT, name="Wave (ft)", marker_color=COLOR_BAR2),
        row=1, col=2,
    )
    fig.update_layout(
        title_text="Weather & Marine Risk Update (Mina Zayed Port) — 4-day outlook",
        title_font_size=14,
        title_x=0.5,
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=PLOT_BG,
        font=dict(color=FONT_COLOR, size=11),
        margin=dict(t=56, b=44, l=48, r=32),
        showlegend=False,
        height=320,
    )
    fig.update_xaxes(gridcolor=GRID_COLOR, linecolor=GRID_COLOR)
    fig.update_yaxes(gridcolor=GRID_COLOR, linecolor=GRID_COLOR)
    try:
        fig.write_image(str(out_path), scale=2)
        return True
    except Exception:
        return False


def _save_matplotlib(out_path: Path) -> bool:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np
    except ImportError:
        return False
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9, 3.6))
    x = np.arange(len(DATES))
    w = 0.36
    ax1.bar(x - w/2, WIND_KT, w, color="#f59e0b", alpha=0.9, edgecolor="none")
    ax1.set_xticks(x)
    ax1.set_xticklabels(DATES)
    ax1.set_ylabel("Wind (kt)")
    ax1.set_title("Wind Speed (kt)")
    ax1.set_facecolor(PLOT_BG)
    ax2.bar(x - w/2, WAVE_FT, w, color="#06b6d4", alpha=0.85, edgecolor="none")
    ax2.set_xticks(x)
    ax2.set_xticklabels(DATES)
    ax2.set_ylabel("Wave (ft)")
    ax2.set_title("Wave Height (ft)")
    ax2.set_facecolor(PLOT_BG)
    fig.suptitle("Weather & Marine Risk Update (Mina Zayed Port) — 4-day outlook", fontsize=12, y=1.02)
    fig.patch.set_facecolor("#ffffff")
    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close()
    return True


def main() -> int:
    out_path = _OUT / "weather_chart_simulation.png"
    # Prefer matplotlib for reliable PNG output (no kaleido dependency)
    if _save_matplotlib(out_path):
        print(f"Saved: {out_path}")
        return 0
    if _save_plotly(out_path):
        print(f"Saved (Plotly): {out_path}")
        return 0
    print("Install matplotlib: pip install matplotlib", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
