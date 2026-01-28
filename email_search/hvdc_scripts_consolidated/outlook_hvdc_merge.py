#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook HVDC 통합 스크립트
모든 월별 OUTLOOK_HVDC_*_rev.xlsx 파일을 하나의 통합 파일로 합치기

기능:
- 모든 월별 _rev 파일 자동 탐색
- 데이터 결합 및 전체 중복 제거
- 통합 통계 생성 (월별 요약, 케이스별, 사이트별, LPO별, 단계별)

출력:
- OUTLOOK_HVDC_ALL_rev.xlsx (통합 파일)
- 시트: 전체_데이터, 월별_요약, 케이스별_통계, 사이트별_통계, LPO별_통계, 단계별_통계

빠른 실행:
  python outlook_hvdc_merge.py
"""

import pandas as pd
import glob
from pathlib import Path
from datetime import datetime
from typing import List, Tuple
import re

# 중복 제거 함수 (outlook_hvdc_analyzer.py와 동일)
def remove_duplicates(df: pd.DataFrame, keep='last', use_body=False) -> Tuple[pd.DataFrame, dict]:
    """
    중복 메시지 제거 (강화된 로직)
    """
    df_work = df.copy()
    
    # Subject 정규화 강화
    df_work['subject_norm'] = (
        df_work['Subject'].fillna('')
        .str.lower()
        .str.strip()
        .str.replace(r'^(re:|fwd?:|fw:|reply:|답변:)\s*', '', regex=True)
        .str.replace(r'\s+', ' ', regex=True)
        .str.replace(r'[^\w\s\-]', '', regex=True)
        .str.strip()
    )
    
    # Sender 정규화
    df_work['sender_norm'] = df_work['SenderEmail'].fillna('').str.lower().str.strip()
    
    # 날짜 정규화
    if 'DeliveryTime' in df_work.columns:
        df_work['date_str'] = pd.to_datetime(df_work['DeliveryTime'], errors='coerce').dt.date.astype(str)
    elif 'CreationTime' in df_work.columns:
        df_work['date_str'] = pd.to_datetime(df_work['CreationTime'], errors='coerce').dt.date.astype(str)
    else:
        df_work['date_str'] = ''
    
    # Body 일부 비교 (옵션)
    if use_body and 'PlainTextBody' in df_work.columns:
        df_work['body_snippet'] = (
            df_work['PlainTextBody'].fillna('')
            .str[:100]
            .str.lower()
            .str.strip()
            .str.replace(r'\s+', ' ', regex=True)
        )
    else:
        df_work['body_snippet'] = ''
    
    # 중복 키 생성
    if use_body and 'body_snippet' in df_work.columns:
        df_work['duplicate_key'] = (
            df_work['subject_norm'] + '|' + 
            df_work['sender_norm'] + '|' + 
            df_work['date_str'] + '|' +
            df_work['body_snippet'].astype(str)
        )
    else:
        df_work['duplicate_key'] = (
            df_work['subject_norm'] + '|' + 
            df_work['sender_norm'] + '|' + 
            df_work['date_str']
        )
    
    # 중복 제거
    df_clean = df_work.drop_duplicates(subset=['duplicate_key'], keep=keep)
    
    # 중복 패턴 분석
    duplicate_counts = df_work.groupby('duplicate_key').size()
    duplicates_only = duplicate_counts[duplicate_counts > 1]
    
    # 통계
    stats = {
        'original': len(df),
        'deduplicated': len(df_clean),
        'removed': len(df) - len(df_clean),
        'ratio': (len(df) - len(df_clean)) / len(df) * 100 if len(df) > 0 else 0,
        'duplicate_groups': len(duplicates_only),
        'max_duplicates': int(duplicates_only.max()) if len(duplicates_only) > 0 else 1
    }
    
    # 임시 컬럼 제거
    cols_to_remove = ['subject_norm', 'sender_norm', 'date_str', 'duplicate_key', 'body_snippet']
    df_clean = df_clean[[col for col in df_clean.columns if col not in cols_to_remove]]
    
    return df_clean, stats

def find_all_hvdc_rev_files() -> List[Path]:
    """모든 OUTLOOK_HVDC_*_rev.xlsx 파일 찾기"""
    pattern = "results/OUTLOOK_HVDC_*_rev.xlsx"
    files = glob.glob(pattern)
    
    # 통합 파일(ALL)과 타임스탬프가 있는 파일은 제외
    files = [f for f in files 
             if 'OUTLOOK_HVDC_ALL' not in f 
             and not re.search(r'_rev_\d{8}(_\d{6})?\.xlsx$', f)]
    
    files = [Path(f) for f in files]
    files.sort(key=lambda f: f.name)
    
    return files

def extract_year_month_from_filename(filename: str) -> str:
    """파일명에서 YYYYMM 형식 추출"""
    match = re.search(r'OUTLOOK_HVDC_(\d{6})_rev', filename)
    if match:
        return match.group(1)
    return None

def load_monthly_data(file_path: Path) -> Tuple[pd.DataFrame, dict]:
    """월별 _rev 파일 로드"""
    print(f"  로드 중: {file_path.name}")
    
    xl = pd.ExcelFile(file_path, engine='openpyxl')
    
    # 전체_데이터 시트 찾기
    data_sheet = None
    for sheet in xl.sheet_names:
        if sheet in ['전체_데이터', '전체 데이터']:
            data_sheet = sheet
            break
    
    if not data_sheet:
        data_sheet = xl.sheet_names[0]
        print(f"    경고: '전체_데이터' 시트를 찾지 못함. 첫 번째 시트 사용: {data_sheet}")
    
    df = pd.read_excel(file_path, sheet_name=data_sheet, engine='openpyxl')
    
    # Month 컬럼이 없으면 파일명에서 추출
    year_month = extract_year_month_from_filename(file_path.name)
    if 'Month' not in df.columns and year_month:
        df['Month'] = str(year_month)
    elif 'Month' in df.columns:
        # 이미 있는 경우: 다양한 형식 처리
        if df['Month'].dtype in ['int64', 'int32', 'float64', 'float32']:
            # 숫자 형식인 경우 직접 YYYYMM 문자열로 변환
            df['Month'] = df['Month'].astype(str).str.zfill(6)
        else:
            # 문자열 또는 날짜 형식인 경우
            month_parsed = pd.to_datetime(df['Month'], errors='coerce')
            month_mask = month_parsed.notna()
            if month_mask.any():
                # 날짜 형식인 행만 YYYYMM으로 변환
                df.loc[month_mask, 'Month'] = month_parsed[month_mask].dt.strftime('%Y%m')
                # 날짜 형식이 아닌 행은 원본 문자열 유지 (이미 YYYYMM 형식일 수 있음)
                df.loc[~month_mask, 'Month'] = df.loc[~month_mask, 'Month'].astype(str)
            else:
                # 모두 날짜 형식이 아니면 문자열로 유지
                df['Month'] = df['Month'].astype(str)
    else:
        # DeliveryTime에서 추출
        if 'DeliveryTime' in df.columns:
            df['Month'] = pd.to_datetime(df['DeliveryTime'], errors='coerce').dt.strftime('%Y%m')
            df['Month'] = df['Month'].fillna('UNKNOWN')
        else:
            df['Month'] = 'UNKNOWN'
    
    # 통계 정보
    stats = {
        'file': file_path.name,
        'year_month': year_month or 'UNKNOWN',
        'rows': len(df),
        'has_cases': df['case_numbers'].notna().sum() if 'case_numbers' in df.columns else 0,
        'has_sites': df['site'].notna().sum() if 'site' in df.columns else 0,
        'has_lpo': df['lpo'].notna().sum() if 'lpo' in df.columns else 0,
        'has_phase': df['phase'].notna().sum() if 'phase' in df.columns else 0
    }
    
    return df, stats

def merge_all_monthly_data(use_body=False) -> Tuple[pd.DataFrame, List[dict]]:
    """모든 월별 데이터 결합"""
    print("\n[월별 파일 탐색 중...]")
    files = find_all_hvdc_rev_files()
    
    if not files:
        print("❌ OUTLOOK_HVDC_*_rev.xlsx 파일을 찾을 수 없습니다")
        return None, []
    
    print(f"  발견: {len(files)}개 파일")
    
    print("\n[월별 데이터 로드 중...]")
    all_dataframes = []
    monthly_stats = []
    
    for file_path in files:
        try:
            df, stats = load_monthly_data(file_path)
            all_dataframes.append(df)
            monthly_stats.append(stats)
            print(f"    ✓ {stats['year_month']}: {stats['rows']:,}행")
        except Exception as e:
            print(f"    ✗ {file_path.name}: 오류 - {e}")
            continue
    
    if not all_dataframes:
        print("❌ 로드된 데이터가 없습니다")
        return None, []
    
    print(f"\n[데이터 결합 중...]")
    df_combined = pd.concat(all_dataframes, ignore_index=True)
    print(f"  결합 전: {len(df_combined):,}행")
    
    # 전체 중복 제거
    print(f"\n[전체 중복 제거 중...] (기준: Subject+Sender+Date{'+Body' if use_body else ''})")
    df_clean, dup_stats = remove_duplicates(df_combined, keep='last', use_body=use_body)
    print(f"  원본: {dup_stats['original']:,}개")
    print(f"  정리: {dup_stats['deduplicated']:,}개")
    print(f"  제거: {dup_stats['removed']:,}개 ({dup_stats['ratio']:.1f}%)")
    print(f"  중복 그룹: {dup_stats['duplicate_groups']:,}개")
    if dup_stats['max_duplicates'] > 1:
        print(f"  최대 중복 횟수: {dup_stats['max_duplicates']}회")
    
    # no 컬럼 재정렬 (1부터 시작)
    if 'no' in df_clean.columns:
        df_clean['no'] = pd.Series(range(1, len(df_clean) + 1), index=df_clean.index)
    else:
        df_clean['no'] = pd.Series(range(1, len(df_clean) + 1), index=df_clean.index)
    
    return df_clean, monthly_stats

def generate_monthly_summary(df: pd.DataFrame, monthly_stats: List[dict]) -> pd.DataFrame:
    """월별 요약 통계 생성"""
    summary_data = []
    
    # Month 컬럼을 문자열로 통일
    if 'Month' in df.columns:
        df['Month'] = df['Month'].astype(str)
    
    for stats in monthly_stats:
        month_str = str(stats['year_month'])
        month_data = df[df['Month'].astype(str) == month_str]
        
        summary_data.append({
            'Month': stats['year_month'],
            '파일명': stats['file'],
            '원본_이메일수': stats['rows'],
            '통합후_이메일수': len(month_data),
            '케이스_수': month_data['case_numbers'].notna().sum() if 'case_numbers' in month_data.columns else 0,
            '사이트_수': month_data['site'].notna().sum() if 'site' in month_data.columns else 0,
            'LPO_수': month_data['lpo'].notna().sum() if 'lpo' in month_data.columns else 0,
            '단계_수': month_data['phase'].notna().sum() if 'phase' in month_data.columns else 0
        })
    
    return pd.DataFrame(summary_data)

def standardize_column_order(df: pd.DataFrame) -> pd.DataFrame:
    """컬럼 순서 표준화 (사용자 수정 포맷 기준 - PlainTextBody는 마지막)"""
    column_order = [
        'no', 'Month', 'Subject', 'SenderName', 'SenderEmail', 'RecipientTo',
        'DeliveryTime', 'CreationTime',
        'site', 'lpo', 'phase',
        'hvdc_cases', 'primary_case', 'sites', 'primary_site', 'lpo_numbers', 'stage', 'stage_hits'
    ]
    
    # PlainTextBody를 별도로 처리 (항상 마지막)
    ordered_columns = [col for col in column_order if col in df.columns]
    extra_columns = [col for col in df.columns if col not in column_order and col != 'PlainTextBody']
    
    # PlainTextBody가 있으면 마지막에 추가
    if 'PlainTextBody' in df.columns:
        final_columns = ordered_columns + extra_columns + ['PlainTextBody']
    else:
        final_columns = ordered_columns + extra_columns
    
    return df[final_columns]

def save_merged_report(df: pd.DataFrame, monthly_stats: List[dict], output_path: Path):
    """통합 보고서 저장"""
    print(f"\n[통합 보고서 저장 중...]")
    print(f"  출력 파일: {output_path}")
    
    # 컬럼 순서 표준화
    df = standardize_column_order(df)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 시트 1: 전체_데이터
        df.to_excel(writer, sheet_name='전체_데이터', index=False)
        print(f"  ✓ 전체_데이터: {len(df):,}행")
        
        # 시트 2: 월별_요약
        monthly_summary = generate_monthly_summary(df, monthly_stats)
        monthly_summary.to_excel(writer, sheet_name='월별_요약', index=False)
        print(f"  ✓ 월별_요약: {len(monthly_summary)}행")
        
        # 시트 3: 케이스별_통계
        if 'case_numbers' in df.columns and df['case_numbers'].notna().any():
            case_stats = df[df['case_numbers'].notna()].groupby('case_numbers').size().reset_index(name='count')
            case_stats = case_stats.sort_values('count', ascending=False)
            case_stats.to_excel(writer, sheet_name='케이스별_통계', index=False)
            print(f"  ✓ 케이스별_통계: {len(case_stats)}행")
        
        # 시트 4: 사이트별_통계
        if 'site' in df.columns and df['site'].notna().any():
            site_stats = df[df['site'].notna()].groupby('site').size().reset_index(name='count')
            site_stats = site_stats.sort_values('count', ascending=False)
            site_stats.to_excel(writer, sheet_name='사이트별_통계', index=False)
            print(f"  ✓ 사이트별_통계: {len(site_stats)}행")
        
        # 시트 5: LPO별_통계
        if 'lpo' in df.columns and df['lpo'].notna().any():
            lpo_stats = df[df['lpo'].notna()].groupby('lpo').size().reset_index(name='count')
            lpo_stats = lpo_stats.sort_values('count', ascending=False)
            lpo_stats.to_excel(writer, sheet_name='LPO별_통계', index=False)
            print(f"  ✓ LPO별_통계: {len(lpo_stats)}행")
        
        # 시트 6: 단계별_통계
        if 'phase' in df.columns and df['phase'].notna().any():
            phase_stats = df[df['phase'].notna()].groupby('phase').size().reset_index(name='count')
            phase_stats = phase_stats.sort_values('count', ascending=False)
            phase_stats.to_excel(writer, sheet_name='단계별_통계', index=False)
            print(f"  ✓ 단계별_통계: {len(phase_stats)}행")
    
    print(f"\n[완료] 통합 보고서 저장 완료: {output_path}")

def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Outlook HVDC 통합 스크립트 - 모든 월별 _rev 파일을 하나로 합치기',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--use-body', action='store_true',
                       help='Body 일부도 중복 판별에 사용 (기본값: Subject+Sender+Date만)')
    parser.add_argument('--output', type=str, default=None,
                       help='출력 파일명 (기본값: OUTLOOK_HVDC_ALL_rev.xlsx)')
    
    args = parser.parse_args()
    
    print("="*70)
    print("  Outlook HVDC 통합 스크립트")
    print("  (모든 월별 _rev 파일 → 통합 파일)")
    if args.use_body:
        print("  [중복 제거: +Body 포함]")
    print("="*70)
    
    # 데이터 결합
    df_merged, monthly_stats = merge_all_monthly_data(use_body=args.use_body)
    
    if df_merged is None:
        print("\n❌ 통합 실패")
        return
    
    # 출력 파일 경로
    if args.output:
        output_path = Path("results") / args.output
    else:
        output_path = Path("results") / "OUTLOOK_HVDC_ALL_rev.xlsx"
    
    # 기존 파일이 있으면 타임스탬프 추가
    if output_path.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = Path("results") / f"OUTLOOK_HVDC_ALL_rev_{timestamp}.xlsx"
    
    # 통합 보고서 저장
    save_merged_report(df_merged, monthly_stats, output_path)
    
    print("\n" + "="*70)
    print(f"통합 완료!")
    print(f"  총 데이터: {len(df_merged):,}행")
    print(f"  월별 파일: {len(monthly_stats)}개")
    print(f"  출력 파일: {output_path}")
    print("="*70)

if __name__ == "__main__":
    main()

