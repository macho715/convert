#!/usr/bin/env python3
"""
날씨 데이터 CSV를 JSON 형식으로 변환
조수 데이터 변환 스크립트와 동일한 구조
"""

import csv
import json
import os
import sys
from datetime import datetime
from typing import List, Dict

def calculate_risk_level(wind_kn, gust_kn, wave_m, vis_km):
    """위험도 계산 (LOW/MEDIUM/HIGH)"""
    if wind_kn is None and gust_kn is None:
        return "UNKNOWN"

    wind_val = wind_kn or 0
    gust_val = gust_kn or (wind_val * 1.3 if wind_val else 0)
    wave_val = wave_m or 0
    vis_val = vis_km or 10

    # 위험도 점수 계산
    risk_score = 0
    if wind_val > 18:
        risk_score += 1
    if gust_val > 22:
        risk_score += 1
    if wave_val > 0.8:
        risk_score += 1
    if vis_val < 6:
        risk_score += 1

    if risk_score >= 3:
        return "HIGH"
    elif risk_score >= 2:
        return "MEDIUM"
    else:
        return "LOW"

def is_shamal_day(wind_dir_deg, wind_kn, gust_kn):
    """샤말 바람 감지"""
    if wind_dir_deg is None or wind_kn is None or gust_kn is None:
        return False

    # NW 방향 (285-345도)
    nw_sector = 285.0 <= wind_dir_deg <= 345.0
    # 강한 바람
    strong_wind = (wind_kn >= 18.0) or (gust_kn >= 22.0)

    return bool(nw_sector and strong_wind)

def convert_weather_csv_to_json(csv_path, json_path=None):
    """
    날씨 데이터 CSV를 JSON으로 변환

    Args:
        csv_path: 입력 CSV 파일 경로
        json_path: 출력 JSON 파일 경로 (None이면 자동 생성)
    """
    if json_path is None:
        json_path = csv_path.replace('.csv', '_manual.json')

    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    csv_full_path = os.path.join(script_dir, csv_path) if not os.path.isabs(csv_path) else csv_path
    json_full_path = os.path.join(script_dir, json_path) if not os.path.isabs(json_path) else json_path

    if not os.path.exists(csv_full_path):
        print(f"❌ 파일을 찾을 수 없습니다: {csv_full_path}")
        return None

    weather_data = {
        "source": "Manual Weather Data Entry",
        "generated_at": datetime.now().isoformat(),
        "location": {
            "name": "Mina Zayed Port / AGI Site",
            "latitude": 24.12,
            "longitude": 52.53
        },
        "weather_records": []
    }

    with open(csv_full_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if not row.get('Date'):
                continue

            # 데이터 파싱 (빈 값은 None 처리)
            try:
                record = {
                    "date": row['Date'].strip(),
                    "wind_max_kn": float(row['Wind_Max_kn']) if row.get('Wind_Max_kn') and row['Wind_Max_kn'].strip() else None,
                    "gust_max_kn": float(row['Gust_Max_kn']) if row.get('Gust_Max_kn') and row['Gust_Max_kn'].strip() else None,
                    "wind_dir_deg": float(row['Wind_Dir_deg']) if row.get('Wind_Dir_deg') and row['Wind_Dir_deg'].strip() else None,
                    "wave_max_m": float(row['Wave_Max_m']) if row.get('Wave_Max_m') and row['Wave_Max_m'].strip() else None,
                    "visibility_km": float(row['Visibility_km']) if row.get('Visibility_km') and row['Visibility_km'].strip() else None,
                    "source": row.get('Source', '').strip(),
                    "notes": row.get('Notes', '').strip()
                }

                # 위험도 계산
                record["risk_level"] = calculate_risk_level(
                    record["wind_max_kn"],
                    record["gust_max_kn"],
                    record["wave_max_m"],
                    record["visibility_km"]
                )

                # 샤말 감지
                record["is_shamal"] = is_shamal_day(
                    record["wind_dir_deg"],
                    record["wind_max_kn"],
                    record["gust_max_kn"]
                )

                weather_data["weather_records"].append(record)
            except Exception as e:
                print(f"⚠️ 레코드 파싱 오류 ({row.get('Date', 'Unknown')}): {e}")
                continue

    # JSON 파일로 저장
    with open(json_full_path, 'w', encoding='utf-8') as f:
        json.dump(weather_data, f, indent=2, ensure_ascii=False)

    print(f"✅ 변환 완료: {json_full_path}")
    print(f"   총 {len(weather_data['weather_records'])}개 레코드")

    # 통계 출력
    filled_records = [r for r in weather_data['weather_records'] if r.get('wind_max_kn') is not None]
    print(f"   데이터 입력된 레코드: {len(filled_records)}개")

    if filled_records:
        shamal_count = sum(1 for r in filled_records if r.get('is_shamal'))
        high_risk_count = sum(1 for r in filled_records if r.get('risk_level') == 'HIGH')
        print(f"   샤말 감지: {shamal_count}일")
        print(f"   고위험일: {high_risk_count}일")

    return weather_data

if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    csv_file = "weather_data_template.csv"
    if len(sys.argv) > 1:
        csv_file = sys.argv[1]

    if not os.path.exists(csv_file):
        print(f"❌ 파일을 찾을 수 없습니다: {csv_file}")
        print("먼저 create_weather_data_template.py를 실행하여 템플릿을 생성하세요.")
    else:
        convert_weather_csv_to_json(csv_file, "weather_data_manual.json")

