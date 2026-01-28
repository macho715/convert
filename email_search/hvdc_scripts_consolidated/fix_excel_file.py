#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 파일 호환성 개선 스크립트
기존 파일을 Excel에서 확실히 열 수 있도록 재생성
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import re

def sanitize_text(text):
    """Excel 호환성을 위해 텍스트 정리"""
    if pd.isna(text):
        return None
    
    text_str = str(text)
    
    # Excel 셀 제한 (32,767 문자)
    if len(text_str) > 32767:
        text_str = text_str[:32767]
    
    # 제어 문자 제거 (탭, 줄바꿈은 유지)
    text_str = ''.join(c if ord(c) >= 32 or c in ['\n', '\r', '\t'] else ' ' for c in text_str)
    
    return text_str

def fix_excel_file(input_file, output_file):
    """Excel 파일 호환성 개선"""
    print("="*70)
    print("Excel 파일 호환성 개선")
    print(f"입력: {input_file}")
    print(f"출력: {output_file}")
    print("="*70)
    
    # 원본 파일 읽기
    print("\n[원본 파일 읽기...]")
    xl = pd.ExcelFile(input_file, engine='openpyxl')
    print(f"  시트 수: {len(xl.sheet_names)}")
    
    # 각 시트 처리
    print("\n[시트별 처리...]")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name in xl.sheet_names:
            print(f"\n  처리 중: {sheet_name}")
            df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
            print(f"    원본: {len(df):,}행 x {len(df.columns)}컬럼")
            
            # 텍스트 컬럼 정리
            for col in df.columns:
                if df[col].dtype == 'object':
                    # Excel 호환성 개선
                    df[col] = df[col].apply(sanitize_text)
            
            # NaN 컬럼 처리 (전체가 NaN인 경우 빈 문자열로)
            for col in df.columns:
                if df[col].isna().all():
                    df[col] = ''
            
            # 데이터 타입 최적화
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 빈 문자열과 None 정규화
                    df[col] = df[col].fillna('')
            
            # 날짜 컬럼 처리
            date_cols = ['DeliveryTime', 'CreationTime']
            for col in date_cols:
                if col in df.columns:
                    try:
                        df[col] = pd.to_datetime(df[col], errors='coerce')
                    except:
                        pass
            
            print(f"    처리 후: {len(df):,}행 저장")
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("\n" + "="*70)
    print("✓ Excel 호환성 개선 완료")
    print(f"  출력 파일: {output_file}")
    print("  이 파일로 Excel에서 열어보세요.")
    print("="*70)

if __name__ == "__main__":
    input_file = Path("results/OUTLOOK_HVDC_ALL_rev.xlsx")
    output_file = Path("results/OUTLOOK_HVDC_ALL_rev_FIXED.xlsx")
    
    if not input_file.exists():
        print(f"❌ 입력 파일이 없습니다: {input_file}")
        exit(1)
    
    fix_excel_file(input_file, output_file)


