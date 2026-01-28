#!/usr/bin/env python3
"""
날씨 데이터 수동 입력용 CSV 템플릿 생성
웹 검색 후 데이터를 이 템플릿에 입력
"""

import csv
from datetime import date, timedelta
import os
import sys

def create_weather_template(start_date, end_date, output_path="weather_data_template.csv"):
    """
    날씨 데이터 입력용 CSV 템플릿 생성

    컬럼:
    - Date: YYYY-MM-DD
    - Wind_Max_kn: 최대 풍속 (knots)
    - Gust_Max_kn: 최대 돌풍 (knots)
    - Wind_Dir_deg: 풍향 (0-360도, NW=315도)
    - Wave_Max_m: 최대 파고 (meters)
    - Visibility_km: 가시거리 (km)
    - Source: 데이터 출처 (예: "UAE NCM", "Windy.com", "Meteoblue")
    - Notes: 비고 (예: "Shamal detected")
    """
    headers = [
        "Date",
        "Wind_Max_kn",
        "Gust_Max_kn",
        "Wind_Dir_deg",
        "Wave_Max_m",
        "Visibility_km",
        "Source",
        "Notes"
    ]

    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    full_path = os.path.join(script_dir, output_path)

    with open(full_path, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)

        # 날짜별 빈 행 생성
        current_date = start_date
        while current_date <= end_date:
            writer.writerow([
                current_date.isoformat(),
                "",  # Wind_Max_kn
                "",  # Gust_Max_kn
                "",  # Wind_Dir_deg
                "",  # Wave_Max_m
                "",  # Visibility_km
                "",  # Source
                ""   # Notes
            ])
            current_date += timedelta(days=1)

    print(f"✅ 템플릿 생성 완료: {full_path}")
    print(f"   날짜 범위: {start_date.isoformat()} ~ {end_date.isoformat()}")
    print(f"   총 {(end_date - start_date).days + 1}일")
    print("\n📋 사용 방법:")
    print("1. 웹 검색으로 날씨 데이터 수집 (UAE NCM, Windy.com, Meteoblue 등)")
    print("2. 이 CSV 파일에 데이터 입력")
    print("3. convert_weather_csv_to_json.py 실행하여 JSON 변환")
    print("4. UntitSSSed-1.py 실행하여 히트맵 생성")

if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    start_date = date(2026, 1, 6)
    end_date = date(2026, 2, 21)
    create_weather_template(start_date, end_date)

