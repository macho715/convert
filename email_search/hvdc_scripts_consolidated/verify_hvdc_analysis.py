#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HVDC 분석 결과 확인"""

import pandas as pd
from pathlib import Path

def verify_hvdc_analysis(file_path):
    """HVDC 분석 결과 확인"""
    print("="*70)
    print("HVDC 분석 결과 확인")
    print("="*70)
    
    file = Path(file_path)
    if not file.exists():
        print(f"❌ 파일 없음: {file}")
        return
    
    print(f"\n생성된 파일: {file.name}")
    print(f"파일 크기: {file.stat().st_size:,} bytes")
    
    xl = pd.ExcelFile(file)
    print(f"\n시트 목록: {xl.sheet_names}")
    
    # 분석 결과 확인
    df = pd.read_excel(file, sheet_name='analysis')
    print(f"\n분석 결과: {len(df):,}개 이메일")
    print(f"\n컬럼: {list(df.columns)}")
    
    # 추출 통계
    print("\n추출 통계:")
    hvdc_cases = df['hvdc_cases'].notna().sum()
    sites = df['sites'].notna().sum()
    lpos = df['lpo_numbers'].notna().sum()
    stages = df['stage'].notna().sum()
    
    print(f"  - 추출된 케이스: {hvdc_cases}개 ({hvdc_cases/len(df)*100:.1f}%)")
    print(f"  - 추출된 사이트: {sites}개 ({sites/len(df)*100:.1f}%)")
    print(f"  - 추출된 LPO: {lpos}개 ({lpos/len(df)*100:.1f}%)")
    print(f"  - 단계 분류: {stages}개 ({stages/len(df)*100:.1f}%)")
    
    # 단계별 요약
    if 'summary_by_stage' in xl.sheet_names:
        df_stage = pd.read_excel(file, sheet_name='summary_by_stage')
        print("\n단계별 요약:")
        print(df_stage.to_string(index=False))
    
    # 사이트별 요약
    if 'summary_by_site' in xl.sheet_names:
        df_site = pd.read_excel(file, sheet_name='summary_by_site')
        print("\n사이트별 요약 (상위 10개):")
        print(df_site.head(10).to_string(index=False))

if __name__ == "__main__":
    verify_hvdc_analysis("results/OUTLOOK_HVDC_ANALYSIS_20251121_1250.xlsx")

