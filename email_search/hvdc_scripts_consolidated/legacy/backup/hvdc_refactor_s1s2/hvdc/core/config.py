
from __future__ import annotations
from pathlib import Path

# 중앙 설정(기존 하드코딩 흡수 예정)
EMAIL_ROOT = Path("EMAIL")                 
ALLOWED_EXT = {".eml", ".txt", ".html", ".csv"}  
EXCEL_OUTDIR = Path("out")
EXCEL_FILENAME = "hvdc_email_report.xlsx"

# 실행 파라미터
MAX_FILES = None  # None = 제한 없음
ENCODING_FALLBACKS = ["utf-8", "cp1252", "latin-1"]

# 로깅
LOG_LEVEL = "INFO"
