#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""중복 이메일 확인 스크립트"""

import pandas as pd
from pathlib import Path

def check_duplicates(file_path):
    """중복 이메일 확인"""
    print("="*60)
    print("중복 이메일 확인")
    print("="*60)
    
    df = pd.read_excel(file_path, sheet_name='전체_이메일')
    print(f"\n총 이메일 수: {len(df):,}개")
    
    # 다양한 기준으로 중복 확인
    print("\n중복 확인 기준:")
    
    # 1. Subject 기준
    subject_dup = df['Subject'].duplicated().sum()
    print(f"  - Subject 기준: {subject_dup}개 중복")
    
    # 2. Subject + SenderEmail 기준
    if 'SenderEmail' in df.columns:
        subj_sender_dup = df[['Subject', 'SenderEmail']].duplicated().sum()
        print(f"  - Subject + SenderEmail 기준: {subj_sender_dup}개 중복")
    
    # 3. Subject + DeliveryTime 기준
    if 'DeliveryTime' in df.columns:
        subj_date_dup = df[['Subject', 'DeliveryTime']].duplicated().sum()
        print(f"  - Subject + DeliveryTime 기준: {subj_date_dup}개 중복")
    
    # 4. Subject + SenderEmail + DeliveryTime 기준
    if 'SenderEmail' in df.columns and 'DeliveryTime' in df.columns:
        subj_sender_date_dup = df[['Subject', 'SenderEmail', 'DeliveryTime']].duplicated().sum()
        print(f"  - Subject + SenderEmail + DeliveryTime 기준: {subj_sender_date_dup}개 중복")
    
    # 5. 완전 중복 행
    complete_dup = df.duplicated().sum()
    print(f"\n완전 중복 행 (모든 컬럼 동일): {complete_dup}개")
    
    # 중복 예시 확인
    if subject_dup > 0:
        print("\n중복 Subject 예시 (상위 5개):")
        dup_subjects = df[df['Subject'].duplicated(keep=False)]['Subject'].value_counts().head(5)
        for subj, count in dup_subjects.items():
            print(f"  - {subj[:80]}... ({count}개)")
    
    return df

if __name__ == "__main__":
    file_path = "results/OUTLOOK_202510_20251121.xlsx"
    df = check_duplicates(file_path)

