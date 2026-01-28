#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중복 이메일 제거 스크립트

Outlook 문제로 동일한 이메일이 중복으로 들어간 경우를 제거합니다.
기준: Subject + SenderEmail + DeliveryTime 조합
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import sys

def remove_duplicates(input_file, output_file=None, keep='first'):
    """
    중복 이메일 제거
    
    Args:
        input_file: 입력 Excel 파일 경로
        output_file: 출력 Excel 파일 경로 (None이면 원본 덮어쓰기)
        keep: 중복 제거 시 유지할 항목 ('first', 'last', False)
    """
    print("="*70)
    print("중복 이메일 제거")
    print("="*70)
    
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"❌ 입력 파일이 없습니다: {input_path}")
        return False
    
    print(f"\n입력 파일: {input_path}")
    
    # Excel 파일 읽기
    print("\n[파일 읽기 중...]")
    xl = pd.ExcelFile(input_path, engine='openpyxl')
    print(f"  시트 수: {len(xl.sheet_names)}")
    
    results = {}
    
    for sheet_name in xl.sheet_names:
        print(f"\n[시트 처리: {sheet_name}]")
        df = pd.read_excel(input_path, sheet_name=sheet_name, engine='openpyxl')
        original_count = len(df)
        print(f"  원본 행 수: {original_count:,}개")
        
        # 중복 확인 기준 컬럼 존재 여부 확인
        required_cols = ['Subject', 'SenderEmail', 'DeliveryTime']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            print(f"  ⚠️ 필수 컬럼 누락: {missing_cols}")
            print(f"  사용 가능한 컬럼: {list(df.columns)}")
            
            # 컬럼명 정규화 시도
            col_mapping = {}
            df_lower = df.columns.str.lower()
            for req_col in missing_cols:
                req_lower = req_col.lower()
                if req_lower in df_lower:
                    idx = list(df_lower).index(req_lower)
                    col_mapping[req_col] = df.columns[idx]
                    print(f"  → {req_col} → {df.columns[idx]} 매핑됨")
            
            if len(col_mapping) == len(required_cols):
                df = df.rename(columns=col_mapping)
            else:
                print(f"  ❌ 필수 컬럼을 찾을 수 없습니다. 스킵합니다.")
                results[sheet_name] = df
                continue
        
        # 중복 확인
        dup_cols = ['Subject', 'SenderEmail', 'DeliveryTime']
        
        # 각 기준별 중복 수 확인
        print(f"\n  중복 확인 기준: {', '.join(dup_cols)}")
        
        before_dedup = df[dup_cols].duplicated().sum()
        print(f"  중복 행 수 (제거 전): {before_dedup}개")
        
        if before_dedup == 0:
            print(f"  ✓ 중복 없음")
            results[sheet_name] = df
            continue
        
        # 중복 제거
        print(f"\n  [중복 제거 중...]")
        df_dedup = df.drop_duplicates(subset=dup_cols, keep=keep)
        after_count = len(df_dedup)
        removed_count = original_count - after_count
        
        print(f"  제거 전: {original_count:,}개")
        print(f"  제거 후: {after_count:,}개")
        print(f"  제거됨: {removed_count}개")
        
        results[sheet_name] = df_dedup
    
    # 결과 저장
    if not output_file:
        # 원본 파일 백업
        backup_path = input_path.parent / f"{input_path.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{input_path.suffix}"
        print(f"\n[원본 파일 백업 중...]")
        import shutil
        shutil.copy2(input_path, backup_path)
        print(f"  백업 파일: {backup_path}")
        output_file = input_path
    else:
        output_file = Path(output_file)
    
    print(f"\n[결과 저장 중...]")
    print(f"  출력 파일: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  ✓ {sheet_name}: {len(df):,}행 저장")
    
    print("\n" + "="*70)
    print("✅ 중복 제거 완료")
    print(f"  출력 파일: {output_file}")
    print("="*70)
    
    return True

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='중복 이메일 제거')
    parser.add_argument('--input', '-i', required=True, help='입력 Excel 파일')
    parser.add_argument('--output', '-o', default=None, help='출력 Excel 파일 (기본: 원본 덮어쓰기)')
    parser.add_argument('--keep', choices=['first', 'last'], default='first', 
                       help='중복 제거 시 유지할 항목 (기본: first)')
    
    args = parser.parse_args()
    
    success = remove_duplicates(args.input, args.output, args.keep)
    sys.exit(0 if success else 1)

