#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OUTLOOK_HVDC_202508_rev.xlsx 파일 포맷 확인"""

import pandas as pd

xl = pd.ExcelFile('results/OUTLOOK_HVDC_202508_rev.xlsx')

print("="*70)
print("OUTLOOK_HVDC_202508_rev.xlsx 파일 구조 분석")
print("="*70)

# 전체 데이터 확인
df_main = pd.read_excel(xl, sheet_name='전체_데이터', nrows=10)
print(f"\n[전체_데이터] 시트:")
print(f"  총 행 수: {len(pd.read_excel(xl, sheet_name='전체_데이터')):,}개")
print(f"  컬럼 수: {len(df_main.columns)}개")
print(f"  컬럼 목록:")
for i, col in enumerate(df_main.columns, 1):
    non_null = df_main[col].notna().sum()
    print(f"    {i:2d}. {col:20s} (샘플에서 {non_null}개 값 존재)")

print(f"\n  주요 컬럼 샘플:")
sample_cols = ['Subject', 'SenderEmail', 'site', 'hvdc_cases', 'primary_case', 
               'sites', 'primary_site', 'lpo_numbers', 'stage']
for col in sample_cols:
    if col in df_main.columns:
        val = df_main[col].iloc[0] if len(df_main) > 0 else None
        print(f"    {col}: {str(val)[:50] if pd.notna(val) else 'NaN'}")

# 통계 시트 확인
print(f"\n[통계 시트들]:")
for sheet in xl.sheet_names[1:]:
    df_stat = pd.read_excel(xl, sheet_name=sheet)
    print(f"\n  [{sheet}]")
    print(f"    행 수: {len(df_stat):,}개")
    print(f"    컬럼: {list(df_stat.columns)}")
    print(f"    상위 3개:")
    print(df_stat.head(3).to_string())


