#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 파일 검증 스크립트"""

import pandas as pd
from pathlib import Path
import sys

file_path = Path("results/OUTLOOK_HVDC_ALL_rev.xlsx")

print("="*70)
print(f"Excel 파일 검증: {file_path}")
print("="*70)

try:
    # 파일 존재 확인
    if not file_path.exists():
        print(f"❌ 파일이 존재하지 않습니다: {file_path}")
        sys.exit(1)
    
    print(f"✓ 파일 존재: {file_path.stat().st_size / (1024*1024):.2f} MB")
    
    # Excel 파일 열기 테스트
    print("\n[Excel 파일 열기 테스트...]")
    try:
        xl = pd.ExcelFile(file_path, engine='openpyxl')
        print(f"✓ 파일 열기 성공")
        print(f"  시트 수: {len(xl.sheet_names)}")
        print(f"  시트 목록: {xl.sheet_names}")
    except Exception as e:
        print(f"❌ 파일 열기 실패: {e}")
        sys.exit(1)
    
    # 각 시트 읽기 테스트
    print("\n[시트별 읽기 테스트...]")
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', nrows=1)
            rows = len(pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl'))
            cols = len(df.columns)
            print(f"  ✓ {sheet_name:20s}: {rows:>6,}행 x {cols:>3}컬럼")
        except Exception as e:
            print(f"  ❌ {sheet_name:20s}: 오류 - {e}")
    
    # 첫 번째 시트 상세 확인
    print("\n[전체_데이터 시트 상세 확인...]")
    try:
        df = pd.read_excel(file_path, sheet_name='전체_데이터', engine='openpyxl', nrows=10)
        print(f"  ✓ 샘플 데이터 읽기 성공")
        print(f"  컬럼 수: {len(df.columns)}")
        print(f"  컬럼 목록 (첫 10개):")
        for i, col in enumerate(df.columns[:10], 1):
            print(f"    {i:2d}. {col}")
        print(f"\n  첫 3행 샘플:")
        print(df[['no', 'Month', 'Subject']].head(3).to_string(index=False))
    except Exception as e:
        print(f"  ❌ 전체_데이터 시트 읽기 실패: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "="*70)
    print("✓ 파일 검증 완료: 정상")
    print("="*70)
    
except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)


