#!/usr/bin/env python3
"""
MINA ZAYED PORT WATER TIDE CSV to JSON Converter
"""

import csv
import json
import os
from datetime import datetime

def convert_tide_csv_to_json(csv_path, json_path=None):
    """
    CSV 파일을 JSON 형식으로 변환합니다.

    Args:
        csv_path: 입력 CSV 파일 경로
        json_path: 출력 JSON 파일 경로 (None이면 자동 생성)
    """
    if json_path is None:
        json_path = csv_path.replace('.csv', '.json')

    tide_data = {
        "source": "MINA ZAYED PORT WATER TIDE",
        "generated_at": datetime.now().isoformat(),
        "tide_records": []
    }

    with open(csv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 첫 번째 줄은 제목, 두 번째 줄은 헤더
    if len(lines) < 2:
        raise ValueError("CSV 파일 형식이 올바르지 않습니다.")

    # 헤더 파싱 (탭 구분)
    headers = lines[1].strip().split('\t')

    # 데이터 행 파싱
    for line in lines[2:]:
        line = line.strip()
        if not line:  # 빈 줄 건너뛰기
            continue

        values = line.split('\t')
        if len(values) < 4:
            continue

        record = {
            "date": values[0].strip(),
            "high_tide_window": values[1].strip(),
            "max_height_m": float(values[2].strip()) if values[2].strip() else None,
            "risk_level": values[3].strip()
        }

        tide_data["tide_records"].append(record)

    # JSON 파일로 저장
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(tide_data, f, indent=2, ensure_ascii=False)

    print(f"✅ 변환 완료: {json_path}")
    print(f"   총 {len(tide_data['tide_records'])}개 레코드")

    return tide_data

if __name__ == "__main__":
    import sys

    # Windows 콘솔 UTF-8 인코딩 설정
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    csv_file = "MINA ZAYED PORT WATER TIDE.csv"

    # 현재 스크립트와 같은 디렉토리에서 파일 찾기
    script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
    csv_path = os.path.join(script_dir, csv_file)

    if not os.path.exists(csv_path):
        # 상대 경로로 시도
        csv_path = csv_file

    print(f"Converting {csv_path} to JSON...")
    tide_data = convert_tide_csv_to_json(csv_path)

    # 콘솔에 샘플 출력
    print("\n📋 JSON 샘플 (처음 3개 레코드):")
    print(json.dumps(tide_data["tide_records"][:3], indent=2, ensure_ascii=False))

