#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""중복 제거 결과 확인 스크립트"""

import pandas as pd
from pathlib import Path

def verify_dedup_result(original_file, dedup_file):
    """중복 제거 결과 확인"""
    print("="*70)
    print("중복 제거 결과 확인")
    print("="*70)
    
    print(f"\n원본 파일: {original_file}")
    print(f"중복 제거 파일: {dedup_file}")
    
    # 원본 파일 읽기
    df_original = pd.read_excel(original_file, sheet_name='전체_이메일')
    df_dedup = pd.read_excel(dedup_file, sheet_name='전체_이메일')
    
    print(f"\n원본: {len(df_original):,}개")
    print(f"중복 제거 후: {len(df_dedup):,}개")
    print(f"제거됨: {len(df_original) - len(df_dedup)}개")
    
    # 중복 확인
    dup_cols = ['Subject', 'SenderEmail', 'DeliveryTime']
    
    print(f"\n중복 확인 기준: {', '.join(dup_cols)}")
    
    original_dup = df_original[dup_cols].duplicated().sum()
    dedup_dup = df_dedup[dup_cols].duplicated().sum()
    
    print(f"원본 중복 수: {original_dup}개")
    print(f"중복 제거 후 중복 수: {dedup_dup}개")
    
    if dedup_dup == 0:
        print("\n✅ 중복 제거 성공 - 중복 없음")
    else:
        print(f"\n⚠️ 중복 제거 불완전 - {dedup_dup}개 중복 남음")
    
    # 샘플 비교
    if len(df_original) > len(df_dedup):
        print("\n제거된 항목 샘플 확인:")
        dup_mask = df_original[dup_cols].duplicated(keep=False)
        dup_rows = df_original[dup_mask]
        
        if len(dup_rows) > 0:
            print(f"\n중복 그룹 수: {len(dup_rows) // 2}개")
            sample = dup_rows.head(2)
            print("\n제거된 중복 항목 샘플:")
            for idx, row in sample.iterrows():
                print(f"  - Subject: {row['Subject'][:60]}...")
                print(f"    Sender: {row['SenderEmail']}")
                print(f"    Date: {row['DeliveryTime']}")

if __name__ == "__main__":
    original = "results/OUTLOOK_202510_20251121.xlsx"
    dedup = "results/OUTLOOK_202510_20251121_dedup.xlsx"
    
    verify_dedup_result(original, dedup)

