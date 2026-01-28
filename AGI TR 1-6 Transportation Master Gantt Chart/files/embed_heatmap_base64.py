# -*- coding: utf-8 -*-
"""Embed heatmap PNG as Base64 in HTML img src for mobile/portable use."""
import base64
import os

gantt = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 1) out/weather_4day_heatmap.png for Schedule_20260126 and Schedule_20260128
png1 = os.path.join(gantt, "out", "weather_4day_heatmap.png")
with open(png1, "rb") as f:
    b64_1 = base64.standard_b64encode(f.read()).decode("ascii")
data_uri_1 = "data:image/png;base64," + b64_1

for name in ["AGI TR Unit 1 Schedule_20260126.html", "AGI TR Unit 1 Schedule_20260128.html"]:
    path = os.path.join(gantt, name)
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace('src="out/weather_4day_heatmap.png"', "src=\"" + data_uri_1 + "\"")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print("Updated", name)

# 2) files/weather_4day_heatmap_dashboard.png for AGI20Unit20Schedule_20260126_redesigned_v2.html
png2 = os.path.join(gantt, "files", "weather_4day_heatmap_dashboard.png")
with open(png2, "rb") as f:
    b64_2 = base64.standard_b64encode(f.read()).decode("ascii")
data_uri_2 = "data:image/png;base64," + b64_2

path2 = os.path.join(gantt, "files", "AGI20Unit20Schedule_20260126_redesigned_v2.html")
with open(path2, "r", encoding="utf-8") as f:
    html2 = f.read()
html2 = html2.replace('src="weather_4day_heatmap_dashboard.png"', "src=\"" + data_uri_2 + "\"")
with open(path2, "w", encoding="utf-8") as f:
    f.write(html2)
print("Updated files/AGI20Unit20Schedule_20260126_redesigned_v2.html")
