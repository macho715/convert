#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook HVDC Analyzer (PST → HVDC 온톨로지 통합 분석)
Legacy 패턴 + HVDC 케이스/사이트/LPO/단계 추출

기능:
- HVDC 케이스 번호 추출 (다양한 패턴 지원)
- 사이트 식별 (DAS/AGI/MIR/MIRFA/GHALLAN)
- LPO 번호 추출
- 프로젝트 단계 분류 (procurement/shipping/customs/logistics/installation/testing/certification)
- 중복 제거 (기본값: 활성화, Subject+Sender+Date 기준, Body 비교 옵션)

입력:
- OUTLOOK_YYYYMM.xlsx (outlook_pst_scanner.py 출력)
- 시트: 전체_이메일

출력:
- OUTLOOK_HVDC_YYYYMM_rev.xlsx (표준 포맷)
- 시트: 전체_데이터, 케이스별_통계, 사이트별_통계, LPO별_통계, 단계별_통계
- 컬럼: V1 형식(site, lpo, phase) + V2 형식(hvdc_cases, primary_case, sites, primary_site, lpo_numbers, stage, stage_hits)

빠른 실행:
  python outlook_hvdc_analyzer.py                    # 기본 (중복 제거 활성화)
  python outlook_hvdc_analyzer.py --use-body        # Body도 비교
  python outlook_hvdc_analyzer.py --no-deduplicate  # 중복 제거 비활성화
  
자동으로 results/ 폴더에서 최신 OUTLOOK_*.xlsx 파일을 찾아 분석합니다
"""

import pandas as pd
import glob
import re
from datetime import datetime
from pathlib import Path
import sys
from typing import Tuple, Dict

# ===== 중복 제거 함수 =====

def remove_duplicates(df: pd.DataFrame, keep='last', use_body=False) -> Tuple[pd.DataFrame, Dict]:
    """
    중복 메시지 제거 (강화된 로직)
    
    Args:
        df: 입력 데이터프레임
        keep: 'first' (첫 번째), 'last' (최신), False (모두 제거)
        use_body: True면 Body 일부도 비교에 사용 (기본: False, Subject+Sender+Date만)
    
    Returns:
        (정리된 데이터프레임, 중복 통계)
    """
    df_work = df.copy()
    
    # Subject 정규화 강화 (공백, 대소문자, 특수문자, RE: FWD: 등 접두사 제거)
    df_work['subject_norm'] = (
        df_work['Subject'].fillna('')
        .str.lower()
        .str.strip()
        .str.replace(r'^(re:|fwd?:|fw:|reply:|답변:)\s*', '', regex=True)  # 접두사 제거
        .str.replace(r'\s+', ' ', regex=True)  # 연속 공백 통일
        .str.replace(r'[^\w\s\-]', '', regex=True)  # 특수문자 제거 (하이픈 제외)
        .str.strip()
    )
    
    # Sender 정규화 (이메일 주소만 추출, 도메인 정규화)
    df_work['sender_norm'] = df_work['SenderEmail'].fillna('').str.lower().str.strip()
    
    # 날짜 정규화 (날짜만 사용, 시간 제외)
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
            .str[:100]  # 첫 100자만
            .str.lower()
            .str.strip()
            .str.replace(r'\s+', ' ', regex=True)
        )
    else:
        df_work['body_snippet'] = ''
    
    # 중복 키 생성 (Subject + Sender + Date + Body(옵션))
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

# ===== Legacy 패턴 통합 =====

def extract_case_numbers_enhanced(subject: str):
    """케이스 번호 추출 (legacy 패턴)"""
    case_numbers = []
    subject_str = str(subject)
    
    # 패턴 1: HVDC-ADOPT-XXX-XXXX
    pattern1 = r'HVDC-ADOPT-([A-Z]+)-([A-Z0-9\-]+)'
    matches1 = re.findall(pattern1, subject_str, re.IGNORECASE)
    for match in matches1:
        case_numbers.append(f"HVDC-ADOPT-{match[0]}-{match[1]}".upper())
    
    # 패턴 2: HVDC-XXX-XXX-XXXX
    pattern2 = r'HVDC-([A-Z]+)-([A-Z]+)-([A-Z0-9\-]+)'
    matches2 = re.findall(pattern2, subject_str, re.IGNORECASE)
    for match in matches2:
        full_case = f"HVDC-{match[0]}-{match[1]}-{match[2]}".upper()
        if full_case not in case_numbers:
            case_numbers.append(full_case)
    
    # 패턴 3: 괄호 안의 약식 (HE-XXXX)
    pattern3_outer = r'\(([^\)]+)\)'
    outer_matches = re.findall(pattern3_outer, subject_str)
    
    for outer_match in outer_matches:
        pattern3_inner = r'([A-Z]+)-([0-9]+(?:-[0-9A-Z]+)?)'
        inner_matches = re.findall(pattern3_inner, outer_match, re.IGNORECASE)
        
        for match in inner_matches:
            vendor_code = match[0].upper()
            case_num = match[1]
            full_case = f"HVDC-ADOPT-{vendor_code}-{case_num}"
            if full_case not in case_numbers:
                case_numbers.append(full_case)
    
    # 패턴 4: JPTW-XX / GRM-XXX
    pattern4 = r'\[HVDC-AGI\].*?(JPTW-(\d+))\s*/\s*(GRM-(\d+))'
    matches4 = re.findall(pattern4, subject_str, re.IGNORECASE)
    for match in matches4:
        jptw_num = match[1]
        grm_num = match[3]
        full_case = f"HVDC-AGI-JPTW{jptw_num}-GRM{grm_num}".upper()
        if full_case not in case_numbers:
            case_numbers.append(full_case)
    
    # 패턴 5: 콜론 뒤 완성된 케이스 번호
    pattern5 = r':\s*([A-Z]+-[A-Z]+-[A-Z]+\d+-[A-Z]+\d+)'
    matches5 = re.findall(pattern5, subject_str, re.IGNORECASE)
    for match in matches5:
        clean_case = re.sub(r'\(.*?\)', '', match).strip().upper()
        if clean_case not in case_numbers:
            case_numbers.append(clean_case)
    
    return ', '.join(case_numbers) if case_numbers else None

def extract_site(subject: str):
    """사이트 추출"""
    match = re.search(r'\b(DAS|AGI|MIR|MIRFA|GHALLAN)\b', str(subject), re.IGNORECASE)
    return match.group(1).upper() if match else None

def extract_lpo(subject: str):
    """LPO 번호 추출"""
    matches = re.findall(r'LPO[-\s]?(\d+)', str(subject), re.IGNORECASE)
    return ', '.join([f"LPO-{lpo}" for lpo in matches]) if matches else None

def extract_phase(subject: str):
    """프로젝트 단계 추출"""
    phases = {
        'procurement': r'\b(LPO|PO|Purchase Order|Procurement|Order)\b',
        'shipping': r'\b(Shipping|Delivery|Container|CNTR|LCT|Vessel)\b',
        'customs': r'\b(Customs|Clearance|Import|Export|Duty)\b',
        'logistics': r'\b(Logistics|Transport|Freight|Cargo|Material)\b',
        'installation': r'\b(Installation|Install|Mounting|Assembly)\b',
        'testing': r'\b(Test|Testing|Commissioning|Startup)\b',
        'certification': r'\b(Certificate|Cert|MTC|COC|Quality)\b'
    }
    
    detected_phases = []
    for phase, pattern in phases.items():
        if re.search(pattern, str(subject), re.IGNORECASE):
            detected_phases.append(phase)
    
    return ', '.join(detected_phases) if detected_phases else None

# ===== 메인 로직 =====

def extract_year_month_from_filename(filename):
    """
    파일명에서 YYYYMM 형식 추출
    예: OUTLOOK_202508.xlsx → 202508
    예: pst_folder_select_20250501_to_20250531_*.xlsx → 202505
    """
    # OUTLOOK_YYYYMM 패턴
    match = re.search(r'OUTLOOK_(\d{6})', filename)
    if match:
        return match.group(1)
    
    # pst_folder_select_YYYYMMDD 패턴
    match = re.search(r'(\d{4})(\d{2})\d{2}_to_', filename)
    if match:
        return match.group(1) + match.group(2)
    
    # pst_202YYYYMM 패턴
    match = re.search(r'pst_(\d{6})', filename)
    if match:
        return match.group(1)
    
    return None

def find_all_pst_files():
    """모든 PST 스캔 파일 찾기"""
    patterns = [
        "OUTLOOK_*.xlsx",
        "results/OUTLOOK_*.xlsx",
        "pst_folder_select_*.xlsx",
        "pst_202*.xlsx",
        "pst_optimized_*.xlsx",
        "pst_analysis_*.xlsx",
        "pst_sample_*.xlsx"
    ]
    
    all_files = []
    for pattern in patterns:
        all_files.extend(glob.glob(pattern))
    
    unique_files = list(set(all_files))
    unique_files.sort(key=lambda f: Path(f).stat().st_mtime, reverse=True)
    
    return unique_files

def select_pst_file(files):
    """사용자에게 파일 선택 제공"""
    if not files:
        return None
    
    print(f"\n📁 발견된 PST 스캔 파일 ({len(files)}개):")
    for i, f in enumerate(files, 1):
        file_path = Path(f)
        size_mb = file_path.stat().st_size / (1024 * 1024)
        mod_time = datetime.fromtimestamp(file_path.stat().st_mtime)
        print(f"  [{i}] {f}")
        print(f"      크기: {size_mb:.2f} MB | 수정: {mod_time.strftime('%Y-%m-%d %H:%M')}")
    
    if len(sys.argv) > 1:
        try:
            choice = int(sys.argv[1])
            if 1 <= choice <= len(files):
                return files[choice - 1]
        except ValueError:
            pass
    
    while True:
        choice = input(f"\n선택 (1-{len(files)}, Enter=최신): ").strip()
        if not choice:
            return files[0]
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
        except ValueError:
            pass
        print("❌ 잘못된 선택입니다")

def detect_data_sheet(xl_file):
    """데이터 시트 자동 감지"""
    sheet_names = xl_file.sheet_names
    for candidate in ['전체_이메일', '전체 데이터', '전체_데이터']:
        if candidate in sheet_names:
            return candidate
    return sheet_names[0]

def analyze_and_create_hvdc_report(pst_file, deduplicate=True, keep='last', use_body=False):
    """PST 파일 분석 및 HVDC 온톨로지 통합 보고서 생성"""
    print(f"\n[HVDC 온톨로지 분석 시작: {pst_file}]")
    
    xl = pd.ExcelFile(pst_file, engine='openpyxl')
    print(f"   시트: {xl.sheet_names}")
    
    data_sheet = detect_data_sheet(xl)
    print(f"   데이터 시트: '{data_sheet}'")
    
    df = pd.read_excel(pst_file, sheet_name=data_sheet, engine='openpyxl')
    print(f"   총 이메일: {len(df):,}개")
    
    # 중복 제거 (기본값: 활성화)
    if deduplicate:
        print(f"\n[중복 제거 중...] (기준: Subject+Sender+Date{'+Body' if use_body else ''})")
        df, dup_stats = remove_duplicates(df, keep=keep, use_body=use_body)
        print(f"   원본: {dup_stats['original']:,}개")
        print(f"   정리: {dup_stats['deduplicated']:,}개")
        print(f"   제거: {dup_stats['removed']:,}개 ({dup_stats['ratio']:.1f}%)")
        print(f"   중복 그룹: {dup_stats['duplicate_groups']:,}개")
        if dup_stats['max_duplicates'] > 1:
            print(f"   최대 중복 횟수: {dup_stats['max_duplicates']}회")
    else:
        print(f"\n[중복 제거: 비활성화]")
    
    # HVDC 온톨로지 메타데이터 추출
    print(f"\n[HVDC 온톨로지 메타데이터 추출 중...]")
    
    # V1 형식 추출
    df['case_numbers'] = df['Subject'].apply(extract_case_numbers_enhanced)
    df['site'] = df['Subject'].apply(extract_site)
    df['lpo'] = df['Subject'].apply(extract_lpo)
    df['phase'] = df['Subject'].apply(extract_phase)
    
    # V2 형식 컬럼 추가 (OUTLOOK_HVDC_rev 포맷)
    df['hvdc_cases'] = df['case_numbers']  # 동일
    df['primary_case'] = df['case_numbers'].apply(
        lambda x: x.split(',')[0].strip() if pd.notna(x) and x else None
    )
    df['sites'] = df['site']  # 동일
    df['primary_site'] = df['site']  # 동일
    df['lpo_numbers'] = df['lpo']  # 동일
    df['stage'] = df['phase']  # 동일
    df['stage_hits'] = None  # 빈 값 (필요시 키워드 매핑 추가 가능)
    
    # 사용자 수정 포맷에 맞춘 컬럼 추가
    # no: 행 번호 (1부터 시작)
    df['no'] = pd.Series(range(1, len(df) + 1), index=df.index)
    
    # Month: YYYYMM 형식 (파일명에서 추출, 아래에서도 사용)
    year_month = extract_year_month_from_filename(pst_file)
    if not year_month:
        # DeliveryTime에서 추출 시도
        if 'DeliveryTime' in df.columns:
            df['Month'] = pd.to_datetime(df['DeliveryTime'], errors='coerce').dt.strftime('%Y%m')
            # 첫 번째 유효한 값을 year_month로 사용
            valid_months = df['Month'].dropna()
            if len(valid_months) > 0:
                year_month = valid_months.iloc[0]
            else:
                year_month = datetime.now().strftime("%Y%m")
            df['Month'] = df['Month'].fillna(year_month)
        else:
            year_month = datetime.now().strftime("%Y%m")
            df['Month'] = year_month
    else:
        df['Month'] = year_month
    
    # 컬럼 순서 표준화 (사용자 수정 포맷 기준 - PlainTextBody는 마지막)
    column_order = [
        'no',                           # 1. 행 번호
        'Month',                        # 2. 월
        'Subject',                      # 3. 제목
        'SenderName',                   # 4. 발신자 이름
        'SenderEmail',                  # 5. 발신자 이메일
        'RecipientTo',                  # 6. 수신자
        'DeliveryTime',                 # 7. 배송 시간
        'CreationTime',                 # 8. 생성 시간
        # V1 형식 메타데이터
        'site',                         # 9. 사이트
        'lpo',                          # 10. LPO
        'phase',                        # 11. 단계
        # V2 형식 메타데이터
        'hvdc_cases',                   # 12. HVDC 케이스들
        'primary_case',                 # 13. 주요 케이스
        'sites',                        # 14. 사이트들
        'primary_site',                 # 15. 주요 사이트
        'lpo_numbers',                  # 16. LPO 번호들
        'stage',                        # 17. 단계
        'stage_hits'                     # 18. 단계 히트
    ]
    
    # PlainTextBody를 별도로 처리 (항상 마지막)
    ordered_columns = [col for col in column_order if col in df.columns]
    extra_columns = [col for col in df.columns if col not in column_order and col != 'PlainTextBody']
    
    # PlainTextBody가 있으면 마지막에 추가
    if 'PlainTextBody' in df.columns:
        final_columns = ordered_columns + extra_columns + ['PlainTextBody']
    else:
        final_columns = ordered_columns + extra_columns
    
    df = df[final_columns]
    
    # 통계 출력
    print(f"\n[추출 통계]")
    print(f"   케이스 번호: {df['case_numbers'].notna().sum():,}개 ({df['case_numbers'].notna().sum()/len(df)*100:.1f}%)")
    print(f"   사이트: {df['site'].notna().sum():,}개 ({df['site'].notna().sum()/len(df)*100:.1f}%)")
    print(f"   LPO: {df['lpo'].notna().sum():,}개 ({df['lpo'].notna().sum()/len(df)*100:.1f}%)")
    print(f"   단계: {df['phase'].notna().sum():,}개 ({df['phase'].notna().sum()/len(df)*100:.1f}%)")
    
    # Excel 저장 (OUTLOOK_HVDC_YYYYMM_rev 형식)
    # year_month는 위에서 이미 추출됨 (Month 컬럼 추가 시)
    if not year_month:
        year_month = datetime.now().strftime("%Y%m")  # fallback (이미 처리되어야 하지만 안전장치)
    
    base_name = f"OUTLOOK_HVDC_{year_month}_rev"
    output_path = Path("results") / f"{base_name}.xlsx"
    
    # 충돌 방지: 기존 파일이 있으면 타임스탬프 추가
    if output_path.exists():
        timestamp = datetime.now().strftime("%Y%m%d")
        output_path = Path("results") / f"OUTLOOK_HVDC_{year_month}_rev_{timestamp}.xlsx"
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 시트 1: 전체 데이터 (확장된 컬럼)
        df.to_excel(writer, sheet_name='전체_데이터', index=False)
        
        # 시트 2: 케이스별 통계
        if df['case_numbers'].notna().any():
            case_stats = df[df['case_numbers'].notna()].groupby('case_numbers').size().reset_index(name='count')
            case_stats = case_stats.sort_values('count', ascending=False)
            case_stats.to_excel(writer, sheet_name='케이스별_통계', index=False)
        
        # 시트 3: 사이트별 통계
        if df['site'].notna().any():
            site_stats = df[df['site'].notna()].groupby('site').size().reset_index(name='count')
            site_stats = site_stats.sort_values('count', ascending=False)
            site_stats.to_excel(writer, sheet_name='사이트별_통계', index=False)
        
        # 시트 4: LPO별 통계
        if df['lpo'].notna().any():
            lpo_stats = df[df['lpo'].notna()].groupby('lpo').size().reset_index(name='count')
            lpo_stats = lpo_stats.sort_values('count', ascending=False)
            lpo_stats.to_excel(writer, sheet_name='LPO별_통계', index=False)
        
        # 시트 5: 단계별 통계
        if df['phase'].notna().any():
            phase_stats = df[df['phase'].notna()].groupby('phase').size().reset_index(name='count')
            phase_stats = phase_stats.sort_values('count', ascending=False)
            phase_stats.to_excel(writer, sheet_name='단계별_통계', index=False)
    
    print(f"\n[완료] HVDC 온톨로지 보고서: {output_path}")
    print(f"   포맷: OUTLOOK_HVDC_rev (표준)")
    print(f"   - 전체_데이터 (V1 + V2 컬럼)")
    print(f"   - 케이스별_통계")
    print(f"   - 사이트별_통계")
    print(f"   - LPO별_통계")
    print(f"   - 단계별_통계")
    
    return output_path

if __name__ == "__main__":
    import argparse
    
    # CLI 인자 파싱
    parser = argparse.ArgumentParser(
        description='PST → HVDC 온톨로지 통합 분석기 (기본값: 중복 제거 활성화)',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('--no-deduplicate', action='store_true',
                       help='중복 제거 비활성화 (기본값: 활성화)')
    parser.add_argument('--use-body', action='store_true',
                       help='Body 일부도 중복 판별에 사용 (기본값: Subject+Sender+Date만)')
    parser.add_argument('--keep', choices=['first', 'last'], default='last',
                       help='중복 시 유지할 메시지 (first=첫번째, last=최신, 기본=last)')
    parser.add_argument('file', nargs='?', help='분석할 파일 경로 (선택, 없으면 대화형 모드)')
    
    args = parser.parse_args()
    
    # 중복 제거 기본값: True (--no-deduplicate가 있으면 False)
    deduplicate = not args.no_deduplicate
    
    print("="*70)
    print("  PST → HVDC 온톨로지 통합 분석기")
    print("  (케이스/사이트/LPO/벤더/단계 추출)")
    if deduplicate:
        print(f"  [중복 제거: ON (keep={args.keep}{', +Body' if args.use_body else ''})]")
    else:
        print("  [중복 제거: OFF]")
    print("="*70)
    
    # 파일 선택
    if args.file:
        pst_file = args.file
    else:
        files = find_all_pst_files()
        
        if not files:
            print("\n❌ PST 스캔 파일을 찾을 수 없습니다")
            sys.exit(1)
        
        pst_file = select_pst_file(files)
    
    if pst_file:
        report = analyze_and_create_hvdc_report(pst_file, 
                                                deduplicate=deduplicate,
                                                keep=args.keep,
                                                use_body=args.use_body)
        print(f"\n[완료]")
    else:
        print("\n[오류] 파일이 선택되지 않았습니다")
        sys.exit(1)



