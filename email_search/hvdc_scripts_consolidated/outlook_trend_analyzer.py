#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
월별 HVDC 온톨로지 트렌드 분석 및 통합 보고서 생성

기능:
- 모든 월 분석 파일 통합
- 시계열 트렌드 분석
- 데이터 품질 평가
- Excel 다중 시트 보고서 생성

사용:
  python outlook_trend_analyzer.py
  python outlook_trend_analyzer.py --output HVDC_TREND_REPORT.xlsx
"""

import pandas as pd
import glob
import re
from pathlib import Path
from datetime import datetime
import sys
from typing import List, Dict, Tuple

def extract_month_from_filename(filename: str) -> str:
    """파일명에서 YYYYMM 추출"""
    match = re.search(r'(\d{6})', filename)
    if match:
        return match.group(1)
    return None

def load_all_monthly_data() -> pd.DataFrame:
    """모든 월별 분석 파일 로드"""
    pattern = "results/OUTLOOK_HVDC_ONTOLOGY_*.xlsx"
    files = sorted(glob.glob(pattern))
    
    if not files:
        print("오류: HVDC 분석 파일을 찾을 수 없습니다")
        print(f"검색 경로: {pattern}")
        sys.exit(1)
    
    print(f"\n발견된 파일 ({len(files)}개):")
    all_data = []
    
    for file in files:
        month = extract_month_from_filename(file)
        if not month:
            continue
        
        try:
            # V1 파일 구조 시도
            df = pd.read_excel(file, sheet_name='전체_데이터', engine='openpyxl')
        except:
            try:
                # V2 파일 구조 시도
                df = pd.read_excel(file, sheet_name='analysis', engine='openpyxl')
            except Exception as e:
                print(f"  ⚠️ {file}: 읽기 실패 ({e})")
                continue
        
        df['Month'] = month
        df['YearMonth'] = f"{month[:4]}-{month[4:]}"
        all_data.append(df)
        print(f"  ✅ {file}: {len(df):,}개 ({month[:4]}-{month[4:]})")
    
    if not all_data:
        print("오류: 로드된 데이터가 없습니다")
        sys.exit(1)
    
    combined = pd.concat(all_data, ignore_index=True)
    print(f"\n총 {len(combined):,}개 이메일 로드 완료")
    return combined

def generate_monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    """월별 요약 통계"""
    summary = df.groupby('YearMonth').agg({
        'Subject': 'count',
        'case_numbers': lambda x: x.notna().sum(),
        'site': lambda x: x.notna().sum(),
        'lpo': lambda x: x.notna().sum(),
        'phase': lambda x: x.notna().sum()
    }).reset_index()
    
    summary.columns = ['월', '총 이메일', '케이스 추출', '사이트 식별', 'LPO 추출', '단계 분류']
    
    # 추출률 계산
    summary['케이스 추출률(%)'] = (summary['케이스 추출'] / summary['총 이메일'] * 100).round(1)
    summary['사이트 식별률(%)'] = (summary['사이트 식별'] / summary['총 이메일'] * 100).round(1)
    summary['LPO 추출률(%)'] = (summary['LPO 추출'] / summary['총 이메일'] * 100).round(1)
    
    return summary

def analyze_case_trends(df: pd.DataFrame) -> pd.DataFrame:
    """케이스별 트렌드 분석"""
    # 케이스가 있는 데이터만
    df_cases = df[df['case_numbers'].notna()].copy()
    
    # 케이스 분리 (쉼표 구분)
    df_cases['case_list'] = df_cases['case_numbers'].str.split(', ')
    df_exploded = df_cases.explode('case_list')
    
    # 월별 케이스 집계
    case_trend = df_exploded.groupby(['YearMonth', 'case_list']).size().reset_index(name='count')
    case_pivot = case_trend.pivot_table(index='case_list', columns='YearMonth', values='count', fill_value=0)
    case_pivot['총계'] = case_pivot.sum(axis=1)
    case_pivot = case_pivot.sort_values('총계', ascending=False)
    
    return case_pivot.reset_index()

def analyze_site_trends(df: pd.DataFrame) -> pd.DataFrame:
    """사이트별 트렌드 분석"""
    # 사이트가 있는 데이터만
    df_sites = df[df['site'].notna()].copy()
    
    # 사이트 분리 (쉼표 구분)
    df_sites['site_list'] = df_sites['site'].str.split(', ')
    df_exploded = df_sites.explode('site_list')
    
    # 월별 사이트 집계
    site_trend = df_exploded.groupby(['YearMonth', 'site_list']).size().reset_index(name='count')
    site_pivot = site_trend.pivot_table(index='site_list', columns='YearMonth', values='count', fill_value=0)
    site_pivot['총계'] = site_pivot.sum(axis=1)
    site_pivot = site_pivot.sort_values('총계', ascending=False)
    
    return site_pivot.reset_index()

def analyze_lpo_trends(df: pd.DataFrame) -> pd.DataFrame:
    """LPO별 트렌드 분석"""
    # LPO가 있는 데이터만
    df_lpo = df[df['lpo'].notna()].copy()
    
    # LPO 분리 (쉼표 구분)
    df_lpo['lpo_list'] = df_lpo['lpo'].str.split(', ')
    df_exploded = df_lpo.explode('lpo_list')
    
    # 월별 LPO 집계
    lpo_trend = df_exploded.groupby(['YearMonth', 'lpo_list']).size().reset_index(name='count')
    lpo_pivot = lpo_trend.pivot_table(index='lpo_list', columns='YearMonth', values='count', fill_value=0)
    lpo_pivot['총계'] = lpo_pivot.sum(axis=1)
    lpo_pivot = lpo_pivot.sort_values('총계', ascending=False).head(50)  # 상위 50개만
    
    return lpo_pivot.reset_index()

def analyze_phase_trends(df: pd.DataFrame) -> pd.DataFrame:
    """단계별 트렌드 분석"""
    # 단계가 있는 데이터만
    df_phase = df[df['phase'].notna()].copy()
    
    # 단계 분리 (쉼표 구분)
    df_phase['phase_list'] = df_phase['phase'].str.split(', ')
    df_exploded = df_phase.explode('phase_list')
    
    # 월별 단계 집계
    phase_trend = df_exploded.groupby(['YearMonth', 'phase_list']).size().reset_index(name='count')
    phase_pivot = phase_trend.pivot_table(index='phase_list', columns='YearMonth', values='count', fill_value=0)
    phase_pivot['총계'] = phase_pivot.sum(axis=1)
    phase_pivot = phase_pivot.sort_values('총계', ascending=False)
    
    return phase_pivot.reset_index()

def quality_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """데이터 품질 분석"""
    quality = []
    
    for month in sorted(df['YearMonth'].unique()):
        df_month = df[df['YearMonth'] == month]
        
        quality.append({
            '월': month,
            '총 이메일': len(df_month),
            '케이스 추출': df_month['case_numbers'].notna().sum(),
            '케이스 추출률(%)': round(df_month['case_numbers'].notna().sum() / len(df_month) * 100, 1),
            '사이트 식별': df_month['site'].notna().sum(),
            '사이트 식별률(%)': round(df_month['site'].notna().sum() / len(df_month) * 100, 1),
            'LPO 추출': df_month['lpo'].notna().sum(),
            'LPO 추출률(%)': round(df_month['lpo'].notna().sum() / len(df_month) * 100, 1),
            '단계 분류': df_month['phase'].notna().sum(),
            '단계 분류률(%)': round(df_month['phase'].notna().sum() / len(df_month) * 100, 1),
            'Subject 누락': df_month['Subject'].isna().sum(),
            'SenderEmail 누락': df_month['SenderEmail'].isna().sum() if 'SenderEmail' in df_month.columns else 0,
        })
    
    return pd.DataFrame(quality)

def generate_overall_summary(df: pd.DataFrame) -> Dict:
    """전체 요약 통계"""
    total_emails = len(df)
    months = sorted(df['YearMonth'].unique())
    
    # 케이스 수집
    all_cases = set()
    for cases in df['case_numbers'].dropna():
        all_cases.update([c.strip() for c in str(cases).split(',')])
    
    # 사이트 수집
    all_sites = set()
    for sites in df['site'].dropna():
        all_sites.update([s.strip() for s in str(sites).split(',')])
    
    # LPO 수집
    all_lpos = set()
    for lpos in df['lpo'].dropna():
        all_lpos.update([l.strip() for l in str(lpos).split(',')])
    
    return {
        '분석 기간': f"{months[0]} ~ {months[-1]}",
        '분석 월 수': len(months),
        '총 이메일': total_emails,
        '고유 케이스 수': len(all_cases),
        '고유 사이트 수': len(all_sites),
        '고유 LPO 수': len(all_lpos),
        '평균 월별 이메일': int(total_emails / len(months)),
        '케이스 추출률(%)': round(df['case_numbers'].notna().sum() / total_emails * 100, 1),
        '사이트 식별률(%)': round(df['site'].notna().sum() / total_emails * 100, 1),
        'LPO 추출률(%)': round(df['lpo'].notna().sum() / total_emails * 100, 1),
    }

def save_trend_report(df: pd.DataFrame, output_path: str):
    """트렌드 보고서 Excel 저장"""
    print("\n보고서 생성 중...")
    
    # 1. 전체 요약
    overall = generate_overall_summary(df)
    overall_df = pd.DataFrame([overall]).T
    overall_df.columns = ['값']
    
    # 2. 월별 요약
    monthly_summary = generate_monthly_summary(df)
    
    # 3. 케이스 트렌드
    print("  - 케이스 트렌드 분석")
    case_trend = analyze_case_trends(df)
    
    # 4. 사이트 트렌드
    print("  - 사이트 트렌드 분석")
    site_trend = analyze_site_trends(df)
    
    # 5. LPO 트렌드
    print("  - LPO 트렌드 분석")
    lpo_trend = analyze_lpo_trends(df)
    
    # 6. 단계 트렌드
    print("  - 단계 트렌드 분석")
    phase_trend = analyze_phase_trends(df)
    
    # 7. 데이터 품질
    print("  - 데이터 품질 분석")
    quality = quality_analysis(df)
    
    # Excel 저장
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        overall_df.to_excel(writer, sheet_name='전체_요약')
        monthly_summary.to_excel(writer, sheet_name='월별_요약', index=False)
        case_trend.to_excel(writer, sheet_name='케이스_트렌드', index=False)
        site_trend.to_excel(writer, sheet_name='사이트_트렌드', index=False)
        lpo_trend.to_excel(writer, sheet_name='LPO_트렌드', index=False)
        phase_trend.to_excel(writer, sheet_name='단계_트렌드', index=False)
        quality.to_excel(writer, sheet_name='데이터_품질', index=False)
    
    print(f"\n✅ 보고서 저장: {output_path}")
    print(f"\n포함된 시트:")
    print(f"  1. 전체_요약 - 전체 통계")
    print(f"  2. 월별_요약 - 월별 집계")
    print(f"  3. 케이스_트렌드 - 케이스별 추이")
    print(f"  4. 사이트_트렌드 - 사이트별 추이")
    print(f"  5. LPO_트렌드 - LPO별 추이")
    print(f"  6. 단계_트렌드 - 단계별 추이")
    print(f"  7. 데이터_품질 - 품질 지표")

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='월별 HVDC 트렌드 분석 및 통합 보고서')
    parser.add_argument('--output', '-o', default=None, 
                       help='출력 파일 경로 (기본: HVDC_TREND_REPORT_YYYYMMDD.xlsx)')
    
    args = parser.parse_args()
    
    print("="*70)
    print("  HVDC 월별 트렌드 분석 및 통합 보고서")
    print("="*70)
    
    # 모든 월 데이터 로드
    all_data = load_all_monthly_data()
    
    # 출력 파일명
    if args.output:
        output_path = args.output
    else:
        timestamp = datetime.now().strftime("%Y%m%d")
        output_path = f"results/HVDC_TREND_REPORT_{timestamp}.xlsx"
    
    # 보고서 생성
    save_trend_report(all_data, output_path)
    
    print(f"\n✅ 완료!")

