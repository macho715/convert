#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""사용자가 수정한 포맷 확인 스크립트"""

import pandas as pd
from pathlib import Path

file_path = Path("results/OUTLOOK_HVDC_202501_rev.xlsx")

xl = pd.ExcelFile(file_path, engine='openpyxl')
print("="*70)
print(f"파일: {file_path}")
print(f"시트 목록: {xl.sheet_names}")
print("="*70)

# 첫 번째 시트 (전체_데이터) 확인
df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0], engine='openpyxl')
print(f"\n[시트: {xl.sheet_names[0]}]")
print(f"행 수: {len(df):,}개")
print(f"컬럼 수: {len(df.columns)}개")
print("\n컬럼 목록:")
print("-" * 70)
for i, col in enumerate(df.columns, 1):
    non_null = df[col].notna().sum()
    null = df[col].isna().sum()
    print(f"{i:2d}. {col:30s} (비어있음: {null:>5,}개, 값있음: {non_null:>5,}개)")

# 컬럼 순서 저장
print("\n" + "="*70)
print("컬럼 순서 (Python 리스트):")
print("-" * 70)
print("column_order = [")
for col in df.columns:
    print(f"    '{col}',")
print("]")

# 다른 시트도 확인
for sheet_name in xl.sheet_names[1:]:
    print(f"\n[시트: {sheet_name}]")
    df_sheet = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    print(f"행 수: {len(df_sheet):,}개")
    print(f"컬럼: {list(df_sheet.columns)}")


