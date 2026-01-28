#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""2025년 10-11월 스캔 결과 확인"""

import pandas as pd
from pathlib import Path
from datetime import datetime

def check_scan_result():
    """2025-10-01 ~ 2025-11-20 스캔 결과 확인"""
    file_path = Path("results/OUTLOOK_202510_20251121.xlsx")
    
    print("="*60)
    print("2025-10-01 ~ 2025-11-20 스캔 결과 확인")
    print("="*60)
    
    if not file_path.exists():
        print(f"❌ 파일 없음: {file_path}")
        print("\n생성된 파일 검색 중...")
        files = list(Path("results").glob("OUTLOOK_2025*.xlsx"))
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        print(f"\n최근 생성된 파일 ({len(files)}개):")
        for f in files[:5]:
            print(f"  - {f.name} ({datetime.fromtimestamp(f.stat().st_mtime)})")
        return
    
    print(f"✓ 파일 존재: {file_path}")
    print(f"  파일 크기: {file_path.stat().st_size:,} bytes")
    print(f"  수정 시간: {datetime.fromtimestamp(file_path.stat().st_mtime)}")
    
    try:
        excel_file = pd.ExcelFile(file_path)
        print(f"\n시트 목록: {excel_file.sheet_names}")
        
        if '전체_이메일' in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name='전체_이메일')
            print(f"\n전체_이메일 시트:")
            print(f"  총 이메일 수: {len(df):,}개")
            
            if len(df) > 0:
                print(f"  컬럼 수: {len(df.columns)}개")
                
                date_cols = [col for col in df.columns if 'date' in col.lower() or 'time' in col.lower() or '날짜' in col.lower() or 'Delivery' in col]
                if date_cols:
                    date_col = date_cols[0]
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    
                    print(f"\n날짜 정보 ({date_col}):")
                    print(f"  최소: {df[date_col].min()}")
                    print(f"  최대: {df[date_col].max()}")
                    
                    # 10월 이메일
                    oct_count = df[(df[date_col].dt.year == 2025) & (df[date_col].dt.month == 10)]
                    print(f"  2025년 10월: {len(oct_count)}개")
                    
                    # 11월 이메일 (11월 20일까지)
                    nov_count = df[(df[date_col].dt.year == 2025) & 
                                   (df[date_col].dt.month == 11) & 
                                   (df[date_col].dt.day <= 20)]
                    print(f"  2025년 11월 1-20일: {len(nov_count)}개")
                    
                    # 전체 기간 (2025-10-01 ~ 2025-11-20)
                    period_count = df[(df[date_col] >= pd.Timestamp('2025-10-01')) & 
                                     (df[date_col] <= pd.Timestamp('2025-11-20 23:59:59'))]
                    print(f"  2025-10-01 ~ 2025-11-20: {len(period_count)}개")
            else:
                print("  ⚠️ 데이터 없음 (빈 파일)")
        
        if '폴더별_통계' in excel_file.sheet_names:
            folder_stats = pd.read_excel(file_path, sheet_name='폴더별_통계')
            print(f"\n폴더별_통계 시트:")
            print(f"  총 폴더 수: {len(folder_stats)}개")
            if len(folder_stats) > 0:
                print(f"\n상위 5개 폴더:")
                print(folder_stats.head().to_string(index=False))
                
    except Exception as e:
        print(f"❌ 오류: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_scan_result()

