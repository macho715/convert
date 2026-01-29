# -*- coding: utf-8 -*-
"""Embed heatmap PNG as Base64 in HTML img src (files/ only). Single HTML portable."""
import base64
import os
from pathlib import Path

FILES_DIR = Path(__file__).resolve().parent

# 1) files/out/weather_4day_heatmap.png
png_out = FILES_DIR / "out" / "weather_4day_heatmap.png"
if png_out.is_file():
    with open(png_out, "rb") as f:
        b64 = base64.standard_b64encode(f.read()).decode("ascii")
    data_uri = "data:image/png;base64," + b64
    for path in sorted(FILES_DIR.glob("AGI TR SCHEDULE_*.html")):
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
        if 'src="out/weather_4day_heatmap.png"' in html:
            html = html.replace(
                'src="out/weather_4day_heatmap.png"', 'src="' + data_uri + '"'
            )
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            print("Updated (out/heatmap)", path.name)
else:
    print("Skipped: files/out/weather_4day_heatmap.png not found")

# 2) files/weather_4day_heatmap_dashboard.png
png_dash = FILES_DIR / "weather_4day_heatmap_dashboard.png"
if png_dash.is_file():
    with open(png_dash, "rb") as f:
        b64_2 = base64.standard_b64encode(f.read()).decode("ascii")
    data_uri_2 = "data:image/png;base64," + b64_2
    for path in sorted(FILES_DIR.glob("AGI TR SCHEDULE_*.html")):
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
        if 'src="weather_4day_heatmap_dashboard.png"' in html:
            html = html.replace(
                'src="weather_4day_heatmap_dashboard.png"', 'src="' + data_uri_2 + '"'
            )
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            print("Updated (dashboard)", path.name)
