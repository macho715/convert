# -*- coding: utf-8 -*-
"""Replace Base64-embedded weather img with file ref (files/ only)."""
import re
from pathlib import Path

FILES_DIR = Path(__file__).resolve().parent
REPLACEMENT = '<img src="weather_4day_heatmap_dashboard.png" alt="Composite Weather Risk 4-day" style="max-width:100%;height:auto;">'
PATTERN = re.compile(r'<img src="data:image/png;base64,[^"]+"[^>]*>')

for path in sorted(FILES_DIR.glob("AGI TR SCHEDULE_*.html")):
    try:
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
        if PATTERN.search(html):
            html = PATTERN.sub(REPLACEMENT, html)
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            print("Replaced img in", path.name)
    except FileNotFoundError:
        pass
