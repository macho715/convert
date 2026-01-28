#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""스캔 결과 확인 스크립트"""

import pandas as pd
from pathlib import Path
from datetime import datetime

def check_scan_result():
    """2025년 10월 11일 스캔 결과 확인"""
    file_path = Path("results/OUTLOOK_202510.xlsx")
    
    print("="*60)
    print("2025년 10월 11일 스캔 결과 확인")
    print("="*60)
    
    if not file_path.exists():
        print(f"❌ 파일 없음: {file_path}")
        return
    
    print(f"✓ 파일 존재: {file_path}")
    print(f"  파일 크기: {file_path.stat().st_size:,} bytes")
    print(f"  수정 시간: {datetime.fromtimestamp(file_path.stat().st_mtime)}")
    
    try:
        # 시트 목록 확인
        excel_file = pd.ExcelFile(file_path)
        print(f"\n시트 목록: {excel_file.sheet_names}")
        
        # 전체 이메일 데이터 확인
        if '전체_이메일' in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name='전체_이메일')
            print(f"\n전체_이메일 시트:")
            print(f"  총 이메일 수: {len(df):,}개")
            
            if len(df) > 0:
                print(f"  컬럼 수: {len(df.columns)}개")
                print(f"  주요 컬럼: {list(df.columns[:10])}")
                
                # 날짜 컬럼 확인
                date_cols = [col for col in df.columns if 'date' in col.lower() or '날짜' in col.lower() or 'time' in col.lower()]
                if date_cols:
                    date_col = date_cols[0]
                    print(f"\n날짜 정보 ({date_col}):")
                    print(f"  최소: {df[date_col].min()}")
                    print(f"  최대: {df[date_col].max()}")
                    
                    # 2025-10-11에 해당하는 이메일 확인
                    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                    oct_11 = df[df[date_col].dt.date == pd.Timestamp('2025-10-11').date()]
                    print(f"  2025-10-11 이메일: {len(oct_11)}개")
            else:
                print("  ⚠️ 데이터 없음 (빈 파일)")
        
        # 폴더별 통계 확인
        if '폴더별_통계' in excel_file.sheet_names:
            folder_stats = pd.read_excel(file_path, sheet_name='폴더별_통계')
            print(f"\n폴더별_통계 시트:")
            print(f"  총 폴더 수: {len(folder_stats)}개")
            if len(folder_stats) > 0:
                print(f"  상위 5개 폴더:")
                print(folder_stats.head().to_string(index=False))
        
    except Exception as e:
        print(f"❌ 오류: {e}")

if __name__ == "__main__":
    check_scan_result()

