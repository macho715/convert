#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 파일 심층 검증 - Excel 열기 문제 진단"""

import pandas as pd
from pathlib import Path
import sys
import warnings
warnings.filterwarnings('ignore')

file_path = Path("results/OUTLOOK_HVDC_ALL_rev.xlsx")

print("="*70)
print(f"Excel 파일 심층 검증 (Excel 열기 문제 진단)")
print("="*70)

try:
    # 모든 시트 읽기
    print("\n[전체_데이터 시트 상세 분석...]")
    df = pd.read_excel(file_path, sheet_name='전체_데이터', engine='openpyxl')
    
    print(f"  총 행 수: {len(df):,}")
    print(f"  총 컬럼 수: {len(df.columns)}")
    
    # 데이터 타입 확인
    print("\n[데이터 타입 확인...]")
    dtype_issues = []
    for col in df.columns:
        dtype = df[col].dtype
        null_count = df[col].isna().sum()
        
        # 문제가 될 수 있는 데이터 타입 체크
        if dtype == 'object':
            # 매우 긴 텍스트 확인
            max_len = df[col].astype(str).str.len().max()
            if max_len > 32767:  # Excel 셀 제한
                dtype_issues.append(f"{col}: 최대 길이 {max_len} (Excel 제한 초과)")
        
        # NaN 비율이 너무 높은 경우
        if null_count / len(df) > 0.99:
            dtype_issues.append(f"{col}: NaN {null_count}/{len(df)} ({null_count/len(df)*100:.1f}%)")
    
    if dtype_issues:
        print("  ⚠️  잠재적 문제:")
        for issue in dtype_issues:
            print(f"    - {issue}")
    else:
        print("  ✓ 데이터 타입 정상")
    
    # 특수 문자 확인
    print("\n[특수 문자 확인...]")
    problematic_chars = []
    text_cols = ['Subject', 'SenderName', 'PlainTextBody']
    for col in text_cols:
        if col in df.columns:
            sample = df[col].dropna().head(1000).astype(str)
            # 제어 문자나 특수 문자 확인
            for idx, val in sample.items():
                if any(ord(c) < 32 and c not in ['\n', '\r', '\t'] for c in str(val)):
                    problematic_chars.append(f"{col}[{idx}]: 제어 문자 포함")
                    break
    
    if problematic_chars:
        print("  ⚠️  문제 가능성:")
        for char in problematic_chars[:10]:
            print(f"    - {char}")
    else:
        print("  ✓ 특수 문자 정상")
    
    # 파일 재저장 테스트 (Excel 호환성 개선)
    print("\n[Excel 호환성 재저장 테스트...]")
    try:
        test_file = Path("results/OUTLOOK_HVDC_ALL_rev_TEST.xlsx")
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            # 전체_데이터 시트 (샘플 1000행만)
            df_sample = df.head(1000).copy()
            # 문제가 될 수 있는 매우 긴 텍스트 자르기
            for col in df_sample.columns:
                if df_sample[col].dtype == 'object':
                    df_sample[col] = df_sample[col].astype(str).str[:32767]
            df_sample.to_excel(writer, sheet_name='전체_데이터', index=False)
            
            # 다른 시트들도 읽어서 저장
            xl = pd.ExcelFile(file_path, engine='openpyxl')
            for sheet in xl.sheet_names:
                if sheet != '전체_데이터':
                    df_sheet = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
                    df_sheet.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"  ✓ 테스트 파일 생성: {test_file}")
        print(f"  이 파일로 Excel에서 열어보세요.")
        
    except Exception as e:
        print(f"  ❌ 재저장 실패: {e}")
    
    # Excel 열기 권장사항
    print("\n" + "="*70)
    print("진단 완료")
    print("="*70)
    print("\n[Excel에서 파일을 열 수 없는 경우 시도해볼 방법:]")
    print("1. Excel에서 '파일 > 열기 > 이 파일을 복구' 시도")
    print("2. Excel 2016 이상 버전 사용 (openpyxl 호환성)")
    print("3. 파일을 다른 위치로 복사 후 열어보기")
    print("4. Excel을 재시작한 후 다시 시도")
    print("5. 파일 크기가 7.65MB이므로 로딩 시간이 다소 걸릴 수 있음")
    print(f"\n테스트 파일도 생성했습니다: {test_file}")
    print("="*70)
    
except Exception as e:
    print(f"\n❌ 오류 발생: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

