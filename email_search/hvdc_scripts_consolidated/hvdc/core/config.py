from __future__ import annotations
from pathlib import Path

# 중앙 설정(기존 하드코딩 흡수 예정) — 지금은 기본값만
EMAIL_ROOT = Path("C:/Users/SAMSUNG/Documents/EMAIL")  # 실제 EMAIL 폴더 경로
ALLOWED_EXT = {".eml", ".txt", ".html", ".csv"}  # 스캐너 화이트리스트
EXCEL_OUTDIR = Path("out")
EXCEL_FILENAME = "hvdc_email_report.xlsx"

# 실행 파라미터(샘플 제한 등)
MAX_FILES = None  # None = 제한 없음
ENCODING_FALLBACKS = ["utf-8", "cp1252", "latin-1"]

# 로깅(단순 프리셋; 실제 로거는 각 엔트리에서 구성)
LOG_LEVEL = "INFO"

# Outlook 스캔 설정 (추가)
OUTLOOK_ENABLED = False  # 기본값: 비활성화
OUTLOOK_MAX_EMAILS = None  # None = 전체
OUTLOOK_DATE_RANGE = None  # None = 전체 (365 = 최근 1년)
OUTLOOK_DEFAULT_FOLDERS = ['Inbox', 'Sent Items']  # 기본 폴더
