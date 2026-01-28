#!/usr/bin/env python3
"""
MINA ZAYED PORT WATER TIDE 데이터 통합 및 정규화 스크립트
- 잘못된 형식의 CSV 파일 정리
- TSV 파일과 통합
- 정규화된 TSV 및 JSON 파일 생성
"""

import csv
import json
import os
from datetime import datetime
from typing import List, Dict, Tuple

def parse_malformed_csv(csv_path: str) -> List[Dict[str, str]]:
    """
    잘못된 형식의 CSV 파일 파싱 (따옴표로 감싸진 탭 구분 데이터)
    """
    records = []

    with open(csv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 첫 번째 줄은 제목, 두 번째 줄은 헤더
    if len(lines) < 2:
        return records

    # 헤더 파싱 (따옴표 제거 후 탭 구분)
    header_line = lines[1].strip().strip('"')
    headers = [h.strip() for h in header_line.split('\t')]

    # 데이터 행 파싱
    for line in lines[2:]:
        line = line.strip()
        if not line:
            continue

        # 따옴표 제거 후 탭 구분
        clean_line = line.strip('"')
        values = [v.strip() for v in clean_line.split('\t')]

        if len(values) < 4:
            continue

        record = {
            'Date': values[0],
            'High Tide Window': values[1] if values[1] else '',
            'Max Height (m)': values[2],
            'Risk Level': values[3]
        }
        records.append(record)

    return records

def parse_tsv(tsv_path: str) -> List[Dict[str, str]]:
    """
    TSV 파일 파싱
    """
    records = []

    with open(tsv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f, delimiter='\t')
        for row in reader:
            if row.get('Date'):
                records.append(row)

    return records

def merge_tide_data(records1: List[Dict], records2: List[Dict]) -> List[Dict]:
    """
    두 데이터셋을 날짜순으로 병합
    """
    all_records = records1 + records2

    # 날짜로 정렬
    def get_date(record):
        try:
            return datetime.strptime(record['Date'], '%Y-%m-%d')
        except:
            return datetime.min

    all_records.sort(key=get_date)

    # 중복 제거 (같은 날짜가 있으면 첫 번째 것 유지)
    seen_dates = set()
    unique_records = []
    for record in all_records:
        date = record.get('Date', '')
        if date and date not in seen_dates:
            seen_dates.add(date)
            unique_records.append(record)

    return unique_records

def save_tsv(records: List[Dict], output_path: str):
    """
    TSV 파일로 저장
    """
    if not records:
        return

    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        fieldnames = ['Date', 'High Tide Window', 'Max Height (m)', 'Risk Level']
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter='\t')
        writer.writeheader()
        writer.writerows(records)

    print(f"✅ TSV 파일 저장 완료: {output_path}")
    print(f"   총 {len(records)}개 레코드")

def save_json(records: List[Dict], output_path: str):
    """
    JSON 파일로 저장
    """
    tide_data = {
        "source": "MINA ZAYED PORT WATER TIDE",
        "generated_at": datetime.now().isoformat(),
        "date_range": {
            "start": records[0]['Date'] if records else None,
            "end": records[-1]['Date'] if records else None
        },
        "total_records": len(records),
        "tide_records": []
    }

    for record in records:
        try:
            tide_record = {
                "date": record['Date'],
                "high_tide_window": record.get('High Tide Window', '').strip(),
                "max_height_m": float(record.get('Max Height (m)', '0')) if record.get('Max Height (m)') else None,
                "risk_level": record.get('Risk Level', 'LOW').strip()
            }
            tide_data["tide_records"].append(tide_record)
        except Exception as e:
            print(f"⚠️ 레코드 파싱 오류: {record.get('Date', 'Unknown')} - {e}")
            continue

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(tide_data, f, indent=2, ensure_ascii=False)

    print(f"✅ JSON 파일 저장 완료: {output_path}")
    print(f"   총 {len(tide_data['tide_records'])}개 레코드")

def main():
    """메인 실행 함수"""
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 입력 파일 경로
    malformed_csv = os.path.join(script_dir, "MINA ZAYED PORT WATER TIDEㅇㅇㅇ.csv")
    tsv_file = os.path.join(script_dir, "Date High Tide Window Max Height (m) Ris.tsv")

    # 출력 파일 경로
    output_tsv = os.path.join(script_dir, "MINA ZAYED PORT WATER TIDE_MERGED.tsv")
    output_json = os.path.join(script_dir, "MINA ZAYED PORT WATER TIDE_MERGED.json")

    print("=" * 60)
    print("MINA ZAYED PORT WATER TIDE 데이터 통합 및 정규화")
    print("=" * 60)

    # 1. 잘못된 형식의 CSV 파일 파싱
    print("\n1️⃣ 잘못된 형식의 CSV 파일 파싱 중...")
    if os.path.exists(malformed_csv):
        records_csv = parse_malformed_csv(malformed_csv)
        print(f"   ✅ {len(records_csv)}개 레코드 파싱 완료 (2026-03 데이터)")
    else:
        print(f"   ⚠️ 파일을 찾을 수 없습니다: {malformed_csv}")
        records_csv = []

    # 2. TSV 파일 파싱
    print("\n2️⃣ TSV 파일 파싱 중...")
    if os.path.exists(tsv_file):
        records_tsv = parse_tsv(tsv_file)
        print(f"   ✅ {len(records_tsv)}개 레코드 파싱 완료 (2026-01~02 데이터)")
    else:
        print(f"   ⚠️ 파일을 찾을 수 없습니다: {tsv_file}")
        records_tsv = []

    # 3. 데이터 병합
    print("\n3️⃣ 데이터 병합 중...")
    merged_records = merge_tide_data(records_tsv, records_csv)
    print(f"   ✅ 총 {len(merged_records)}개 레코드 병합 완료")

    if merged_records:
        date_range = f"{merged_records[0]['Date']} ~ {merged_records[-1]['Date']}"
        print(f"   📅 날짜 범위: {date_range}")

    # 4. TSV 파일 저장
    print("\n4️⃣ 정규화된 TSV 파일 저장 중...")
    save_tsv(merged_records, output_tsv)

    # 5. JSON 파일 저장
    print("\n5️⃣ JSON 파일 저장 중...")
    save_json(merged_records, output_json)

    print("\n" + "=" * 60)
    print("✅ 모든 작업 완료!")
    print("=" * 60)
    print(f"\n📁 생성된 파일:")
    print(f"   - {output_tsv}")
    print(f"   - {output_json}")

if __name__ == "__main__":
    import sys

    # Windows 콘솔 UTF-8 인코딩 설정
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass

    main()

