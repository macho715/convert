"""
MAMMOET 인력 데이터 검증 스크립트 (개선 버전)
- 이름 유사도 기반 매칭 (S.N. 무시)
- 폴더명 정규화 강화 (축약 이름 허용)
- Excel 빈 행 필터링
"""

import pandas as pd
import os
import sys
from pathlib import Path
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple
import json
from datetime import datetime
import re

# Windows 콘솔 UTF-8 인코딩 설정
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def normalize_name_advanced(name: str) -> str:
    """이름 정규화 (고급 버전: 축약 이름, 중간 이름 처리)"""
    if pd.isna(name) or name is None:
        return ""
    
    name_str = str(name).strip()
    
    # 여러 공백을 하나로
    while "  " in name_str:
        name_str = name_str.replace("  ", " ")
    
    # 대문자 변환
    name_str = name_str.upper()
    
    # 특수문자 제거 (하이픈, 점 등은 유지하되 정규화)
    name_str = re.sub(r'[^\w\s\-\.]', '', name_str)
    
    # 중간 이름 축약 처리 (예: "Muhammad Nasir" -> "MUHAMMAD NASIR")
    # "Bin", "Bint", "Al", "Abu" 등의 아랍어 접두사 정규화
    arabic_prefixes = ['BIN', 'BINT', 'AL', 'ABU', 'ABUL', 'IBN']
    parts = name_str.split()
    normalized_parts = []
    
    for part in parts:
        # 접두사는 유지하되 정규화
        if part in arabic_prefixes:
            normalized_parts.append(part)
        else:
            # 일반 이름 부분은 그대로 유지
            normalized_parts.append(part)
    
    return ' '.join(normalized_parts)

def extract_first_last_name(name: str) -> Tuple[str, str]:
    """이름에서 첫 이름과 마지막 이름 추출"""
    normalized = normalize_name_advanced(name)
    parts = normalized.split()
    
    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], ""
    else:
        # 첫 이름과 마지막 이름
        first = parts[0]
        last = parts[-1]
        return first, last

def similarity_advanced(name1: str, name2: str) -> float:
    """고급 유사도 계산 (전체 이름 + 첫/마지막 이름 조합)"""
    if not name1 or not name2:
        return 0.0
    
    norm1 = normalize_name_advanced(name1)
    norm2 = normalize_name_advanced(name2)
    
    # 전체 이름 유사도
    full_sim = SequenceMatcher(None, norm1, norm2).ratio()
    
    # 첫 이름 + 마지막 이름 유사도
    first1, last1 = extract_first_last_name(name1)
    first2, last2 = extract_first_last_name(name2)
    
    first_sim = SequenceMatcher(None, first1, first2).ratio() if first1 and first2 else 0.0
    last_sim = SequenceMatcher(None, last1, last2).ratio() if last1 and last2 else 0.0
    
    # 가중 평균 (전체 50%, 첫이름 25%, 마지막이름 25%)
    combined_sim = (full_sim * 0.5) + (first_sim * 0.25) + (last_sim * 0.25)
    
    return max(full_sim, combined_sim)

def normalize_folder_name(folder_name: str) -> str:
    """폴더명 정규화 (직책 제거, 축약 이름 처리)"""
    # 번호 제거 (예: "1. ", "10. ")
    folder_name = re.sub(r'^\d+\.\s*', '', folder_name)
    
    # 직책 제거 (예: "SPMT SV - ", "ENGINEER - " 등)
    folder_name = re.sub(r'^[A-Z\s]+-\s*', '', folder_name)
    
    # 추가 정보 제거 (예: "- new visa", "- old visa and eid")
    folder_name = re.sub(r'\s*-\s*new\s+visa.*$', '', folder_name, flags=re.IGNORECASE)
    folder_name = re.sub(r'\s*-\s*old\s+visa.*$', '', folder_name, flags=re.IGNORECASE)
    folder_name = re.sub(r'\s*-\s*new\s+visa\s*&\s*eid.*$', '', folder_name, flags=re.IGNORECASE)
    folder_name = re.sub(r'\s*-\s*old\s+visa\s+and\s+eid.*$', '', folder_name, flags=re.IGNORECASE)
    
    return normalize_name_advanced(folder_name.strip())

def get_folder_mapping(base_folder: str) -> Dict[int, Dict]:
    """폴더명에서 이름 추출하여 매핑 (개선 버전)"""
    folder_mapping = {}
    base_path = Path(base_folder)
    
    if not base_path.exists():
        return folder_mapping
    
    for folder in sorted(base_path.iterdir()):
        if folder.is_dir() and folder.name[0].isdigit():
            try:
                # 다양한 폴더명 형식 처리
                # 예: "1. SPMT SV - NOR ASEAN BIN ATAN"
                # 예: "6. SPMT RIGGER - JOSEPH MALIEKKAL - old visa and eid"
                
                # 번호 추출
                match = re.match(r'^(\d+)\.', folder.name)
                if not match:
                    continue
                
                folder_num = int(match.group(1))
                
                # 이름 부분 추출 (첫 번째 " - " 이후)
                if ' - ' in folder.name:
                    parts = folder.name.split(' - ', 1)
                    folder_name = parts[1].strip()
                    
                    # 추가 정보 제거 (예: "- old visa and eid")
                    folder_name = re.sub(r'\s*-\s*(new|old)\s+visa.*$', '', folder_name, flags=re.IGNORECASE)
                    folder_name = folder_name.strip()
                else:
                    # " - "가 없으면 전체 이름 사용
                    folder_name = re.sub(r'^\d+\.\s*', '', folder.name).strip()
                
                files = list(folder.glob('*'))
                pdf_files = [f for f in files if f.suffix.lower() == '.pdf']
                img_files = [f for f in files if f.suffix.lower() in ['.jpg', '.jpeg', '.png']]
                
                folder_mapping[folder_num] = {
                    'folder_name': folder_name,
                    'folder_path': str(folder),
                    'pdf_count': len(pdf_files),
                    'img_count': len(img_files),
                    'total_files': len(files)
                }
            except (ValueError, IndexError) as e:
                continue
    
    return folder_mapping

def load_tsv_data(tsv_path: str) -> pd.DataFrame:
    """TSV 파일 로드"""
    try:
        df = pd.read_csv(tsv_path, sep='\t', encoding='utf-8')
        print(f"   ✓ TSV 파일 로드 성공: {len(df)}행")
        return df
    except Exception as e:
        print(f"   ❌ TSV 파일 읽기 오류: {e}")
        sys.exit(1)

def load_excel_data_filtered(excel_path: str) -> Optional[pd.DataFrame]:
    """Excel 파일 로드 (빈 행 필터링 + Gate Pass 형식 지원)"""
    try:
        if not os.path.exists(excel_path):
            print(f"   ⚠️  Excel 파일을 찾을 수 없습니다: {excel_path}")
            return None
        
        xls_file = pd.ExcelFile(excel_path, engine='openpyxl')
        print(f"   ✓ Excel 파일 열기 성공: {len(xls_file.sheet_names)}개 시트")
        
        sheet_name = xls_file.sheet_names[0]
        if 'Sheet1' in xls_file.sheet_names:
            sheet_name = 'Sheet1'
        
        # 전체 데이터 로드 (header=None으로 로드하여 Gate Pass 형식 확인)
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        print(f"   ✓ 원본 데이터: {len(df_raw)}행 x {len(df_raw.columns)}열")
        
        # Gate Pass 형식 확인 (열 23에 "Full Name" 패턴이 있는지)
        is_gatepass_format = False
        col_23_idx = 22  # 열 23 (0-based)
        for row_idx in range(min(50, len(df_raw))):
            cell_value = df_raw.iloc[row_idx, col_23_idx] if col_23_idx < len(df_raw.columns) else None
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                if 'full name' in cell_str.lower() and len(cell_str) > 10:
                    is_gatepass_format = True
                    break
        
        if is_gatepass_format:
            print(f"   ✓ Gate Pass 형식 감지됨")
            # Gate Pass 형식 파싱
            names = []
            for row_idx in range(len(df_raw)):
                cell_value = df_raw.iloc[row_idx, col_23_idx] if col_23_idx < len(df_raw.columns) else None
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    if 'full name' in cell_str.lower():
                        # "Full Name" 이후의 이름 추출
                        match = re.search(r'full\s+name\s+(.+)', cell_str, re.IGNORECASE)
                        if match:
                            name = match.group(1).strip()
                            # 줄바꿈이나 특수문자 제거
                            name = re.sub(r'\s+', ' ', name)
                            name = name.replace('\n', ' ').replace('\xa0', ' ')
                            name = ' '.join(name.split())
                            if len(name) > 3:
                                names.append({
                                    'S.N.': len(names) + 1,
                                    'Name': name,
                                    'Excel_Row': row_idx + 1
                                })
            
            if names:
                df_filtered = pd.DataFrame(names)
                print(f"   ✓ Gate Pass 형식에서 {len(df_filtered)}명 추출")
                return df_filtered
            else:
                print(f"   ⚠️  Gate Pass 형식에서 이름을 추출할 수 없습니다")
                return None
        
        # 일반 테이블 형식 처리
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
        
        # 빈 행 필터링
        df_filtered = df.dropna(how='all')
        
        # 이름 컬럼이 있는 경우, 이름이 비어있는 행도 제거
        name_col = find_excel_name_column(df_filtered)
        if name_col:
            df_filtered = df_filtered[df_filtered[name_col].notna()]
            df_filtered = df_filtered[df_filtered[name_col].astype(str).str.strip() != '']
        
        # S.N. 컬럼이 숫자가 아닌 행 제거 (헤더, 빈 행 등)
        if 'S.N.' in df_filtered.columns:
            df_filtered = df_filtered[pd.to_numeric(df_filtered['S.N.'], errors='coerce').notna()]
        
        print(f"   ✓ 필터링 후: {len(df_filtered)}행 (제거: {len(df) - len(df_filtered)}행)")
        
        return df_filtered.reset_index(drop=True)
    except Exception as e:
        print(f"   ⚠️  Excel 파일 읽기 오류: {e}")
        import traceback
        traceback.print_exc()
        return None

def find_excel_name_column(df: pd.DataFrame) -> Optional[str]:
    """Excel DataFrame에서 이름 컬럼 찾기"""
    if df is None or df.empty:
        return None
    
    name_patterns = [
        'name', 'employee name', 'full name', '이름',
        'employee', 'staff name', 'personnel name'
    ]
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        for pattern in name_patterns:
            if pattern in col_lower:
                return col
    
    # 패턴 매칭 실패 시 첫 번째 텍스트 컬럼 반환
    for col in df.columns:
        if df[col].dtype == 'object':
            non_null = df[col].dropna()
            if len(non_null) > 0:
                sample = str(non_null.iloc[0])
                if ' ' in sample and 5 < len(sample) < 50:
                    return col
    
    return None

def match_by_name_similarity(tsv_df: pd.DataFrame, excel_df: Optional[pd.DataFrame], 
                            folder_mapping: Dict[int, Dict]) -> Dict:
    """이름 유사도 기반 매칭 (S.N. 무시)"""
    results = {
        'tsv_to_excel': [],
        'tsv_to_folder': [],
        'excel_to_tsv': [],
        'folder_to_tsv': [],
        'unmatched_tsv': [],
        'unmatched_excel': [],
        'unmatched_folder': []
    }
    
    # TSV 데이터 준비
    tsv_records = []
    for idx, row in tsv_df.iterrows():
        tsv_records.append({
            'S.N.': int(row['S.N.']),
            'Name': str(row['Name']),
            'Name_Normalized': normalize_name_advanced(row['Name']),
            'Position': str(row['Position']),
            'Employee_Number': row.get('Employee Number', ''),
            'EID': row.get('EID Number', ''),
            'Email': row.get('Email address', '')
        })
    
    # Excel 데이터 준비
    excel_records = []
    if excel_df is not None:
        name_col = find_excel_name_column(excel_df)
        if name_col:
            for idx, row in excel_df.iterrows():
                name = str(row[name_col]).strip()
                if name and name.lower() not in ['nan', 'none', '']:
                    excel_records.append({
                        'Excel_Row': idx + 2,  # Excel 행 번호 (헤더 제외)
                        'S.N.': row.get('S.N.', ''),
                        'Name': name,
                        'Name_Normalized': normalize_name_advanced(name),
                        'Raw_Data': row.to_dict()
                    })
    
    # 폴더 데이터 준비
    folder_records = []
    for folder_num, folder_info in folder_mapping.items():
        folder_name = folder_info['folder_name']
        folder_records.append({
            'Folder_Number': folder_num,
            'Folder_Name': folder_name,
            'Folder_Name_Normalized': normalize_folder_name(folder_name),
            'PDF_Count': folder_info['pdf_count'],
            'Image_Count': folder_info['img_count']
        })
    
    # TSV -> Excel 매칭
    for tsv_rec in tsv_records:
        best_match = None
        best_sim = 0.0
        
        for excel_rec in excel_records:
            sim = similarity_advanced(tsv_rec['Name'], excel_rec['Name'])
            if sim > best_sim and sim > 0.7:  # 70% 이상 유사도
                best_sim = sim
                best_match = excel_rec
        
        if best_match:
            results['tsv_to_excel'].append({
                'TSV_S.N.': tsv_rec['S.N.'],
                'TSV_Name': tsv_rec['Name'],
                'Excel_Row': best_match['Excel_Row'],
                'Excel_S.N.': best_match.get('S.N.', ''),
                'Excel_Name': best_match['Name'],
                'Similarity': f"{best_sim:.2%}"
            })
        else:
            results['unmatched_tsv'].append(tsv_rec)
    
    # TSV -> Folder 매칭
    for tsv_rec in tsv_records:
        best_match = None
        best_sim = 0.0
        
        for folder_rec in folder_records:
            sim = similarity_advanced(tsv_rec['Name'], folder_rec['Folder_Name'])
            if sim > best_sim and sim > 0.6:  # 60% 이상 유사도
                best_sim = sim
                best_match = folder_rec
        
        if best_match:
            results['tsv_to_folder'].append({
                'TSV_S.N.': tsv_rec['S.N.'],
                'TSV_Name': tsv_rec['Name'],
                'Folder_Number': best_match['Folder_Number'],
                'Folder_Name': best_match['Folder_Name'],
                'Similarity': f"{best_sim:.2%}",
                'PDF_Count': best_match['PDF_Count'],
                'Image_Count': best_match['Image_Count']
            })
    
    # Excel -> TSV 매칭 (역방향)
    for excel_rec in excel_records:
        best_match = None
        best_sim = 0.0
        
        for tsv_rec in tsv_records:
            sim = similarity_advanced(excel_rec['Name'], tsv_rec['Name'])
            if sim > best_sim and sim > 0.7:
                best_sim = sim
                best_match = tsv_rec
        
        if best_match:
            results['excel_to_tsv'].append({
                'Excel_Row': excel_rec['Excel_Row'],
                'Excel_S.N.': excel_rec.get('S.N.', ''),
                'Excel_Name': excel_rec['Name'],
                'TSV_S.N.': best_match['S.N.'],
                'TSV_Name': best_match['Name'],
                'Similarity': f"{best_sim:.2%}"
            })
        else:
            results['unmatched_excel'].append(excel_rec)
    
    # Folder -> TSV 매칭 (역방향)
    for folder_rec in folder_records:
        best_match = None
        best_sim = 0.0
        
        for tsv_rec in tsv_records:
            sim = similarity_advanced(folder_rec['Folder_Name'], tsv_rec['Name'])
            if sim > best_sim and sim > 0.6:
                best_sim = sim
                best_match = tsv_rec
        
        if not best_match:
            results['unmatched_folder'].append(folder_rec)
    
    return results

def generate_advanced_report(tsv_df: pd.DataFrame, excel_df: Optional[pd.DataFrame],
                           folder_mapping: Dict, match_results: Dict) -> str:
    """고급 검증 리포트 생성"""
    report_lines = []
    report_lines.append("=" * 80)
    report_lines.append("📋 MAMMOET 인력 데이터 검증 리포트 (이름 기반 매칭)")
    report_lines.append(f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("=" * 80)
    
    # 1. 데이터 요약
    report_lines.append("\n[1] 데이터 요약")
    report_lines.append("-" * 80)
    report_lines.append(f"   TSV: {len(tsv_df)}명")
    report_lines.append(f"   Excel: {len(excel_df) if excel_df is not None else 0}명 (필터링 후)")
    report_lines.append(f"   폴더: {len(folder_mapping)}개")
    
    # 2. TSV -> Excel 매칭 결과
    report_lines.append("\n[2] TSV → Excel 매칭 결과")
    report_lines.append("-" * 80)
    report_lines.append(f"   ✓ 매칭 성공: {len(match_results['tsv_to_excel'])}/{len(tsv_df)}명")
    
    if match_results['tsv_to_excel']:
        report_lines.append("\n   매칭된 항목:")
        for match in sorted(match_results['tsv_to_excel'], key=lambda x: x['TSV_S.N.']):
            report_lines.append(f"      TSV S.N.{match['TSV_S.N.']:2d}: {match['TSV_Name']}")
            report_lines.append(f"         → Excel 행 {match['Excel_Row']:2d}: {match['Excel_Name']} (유사도: {match['Similarity']})")
    
    if match_results['unmatched_tsv']:
        report_lines.append(f"\n   ⚠️  매칭 실패: {len(match_results['unmatched_tsv'])}명")
        for unmatched in match_results['unmatched_tsv']:
            report_lines.append(f"      - S.N. {unmatched['S.N.']}: {unmatched['Name']} ({unmatched['Position']})")
    
    # 3. TSV -> Folder 매칭 결과
    report_lines.append("\n[3] TSV → Folder 매칭 결과")
    report_lines.append("-" * 80)
    report_lines.append(f"   ✓ 매칭 성공: {len(match_results['tsv_to_folder'])}/{len(tsv_df)}명")
    
    if match_results['tsv_to_folder']:
        report_lines.append("\n   매칭된 항목:")
        for match in sorted(match_results['tsv_to_folder'], key=lambda x: x['TSV_S.N.']):
            report_lines.append(f"      TSV S.N.{match['TSV_S.N.']:2d}: {match['TSV_Name']}")
            report_lines.append(f"         → 폴더 {match['Folder_Number']:2d}: {match['Folder_Name']} (유사도: {match['Similarity']})")
            report_lines.append(f"            문서: PDF {match['PDF_Count']}개, 이미지 {match['Image_Count']}개")
    
    # 4. Excel -> TSV 역방향 매칭
    report_lines.append("\n[4] Excel → TSV 역방향 매칭")
    report_lines.append("-" * 80)
    report_lines.append(f"   ✓ 매칭 성공: {len(match_results['excel_to_tsv'])}명")
    
    if match_results['unmatched_excel']:
        report_lines.append(f"\n   ⚠️  Excel에만 있는 항목: {len(match_results['unmatched_excel'])}명")
        for unmatched in match_results['unmatched_excel'][:10]:  # 최대 10개만 표시
            report_lines.append(f"      - Excel 행 {unmatched['Excel_Row']}: {unmatched['Name']}")
        if len(match_results['unmatched_excel']) > 10:
            report_lines.append(f"      ... 외 {len(match_results['unmatched_excel']) - 10}개")
    
    # 5. 폴더 매칭 실패
    if match_results['unmatched_folder']:
        report_lines.append("\n[5] 폴더 매칭 실패")
        report_lines.append("-" * 80)
        report_lines.append(f"   ⚠️  매칭 실패 폴더: {len(match_results['unmatched_folder'])}개")
        for unmatched in match_results['unmatched_folder']:
            report_lines.append(f"      - 폴더 {unmatched['Folder_Number']}: {unmatched['Folder_Name']}")
    
    # 6. 불일치 분석
    report_lines.append("\n[6] 불일치 분석")
    report_lines.append("-" * 80)
    
    # S.N. 불일치 찾기
    sn_mismatches = []
    for tsv_match in match_results['tsv_to_excel']:
        tsv_sn = tsv_match['TSV_S.N.']
        excel_sn = tsv_match.get('Excel_S.N.', '')
        if excel_sn and str(excel_sn) != str(tsv_sn):
            sn_mismatches.append({
                'TSV_S.N.': tsv_sn,
                'Excel_S.N.': excel_sn,
                'Name': tsv_match['TSV_Name']
            })
    
    if sn_mismatches:
        report_lines.append(f"   ⚠️  S.N. 불일치: {len(sn_mismatches)}건")
        for mismatch in sn_mismatches:
            report_lines.append(f"      TSV S.N.{mismatch['TSV_S.N.']} ↔ Excel S.N.{mismatch['Excel_S.N.']}: {mismatch['Name']}")
    else:
        report_lines.append("   ✓ S.N. 일치 확인")
    
    report_lines.append("\n" + "=" * 80)
    report_lines.append("✅ 검증 완료")
    report_lines.append("=" * 80)
    
    return "\n".join(report_lines)

def validate_data_advanced(tsv_path: str, excel_path: str, folder_path: str, output_dir: str = None):
    """고급 검증 작업 (이름 기반 매칭)"""
    print("=" * 80)
    print("📋 MAMMOET 인력 데이터 검증 시작 (이름 기반 매칭)")
    print("=" * 80)
    
    if output_dir is None:
        output_dir = Path(tsv_path).parent
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. TSV 로드
    print("\n[1] TSV 파일 로드 중...")
    tsv_df = load_tsv_data(tsv_path)
    
    # 2. Excel 로드 (필터링)
    print("\n[2] Excel 파일 로드 중 (빈 행 필터링)...")
    excel_df = load_excel_data_filtered(excel_path)
    
    # 3. 폴더 매핑
    print("\n[3] 폴더 구조 분석 중...")
    folder_mapping = get_folder_mapping(folder_path)
    print(f"   ✓ 폴더 수: {len(folder_mapping)}개")
    
    # 4. 이름 기반 매칭
    print("\n[4] 이름 유사도 기반 매칭 중...")
    match_results = match_by_name_similarity(tsv_df, excel_df, folder_mapping)
    print(f"   ✓ TSV→Excel 매칭: {len(match_results['tsv_to_excel'])}/{len(tsv_df)}명")
    print(f"   ✓ TSV→Folder 매칭: {len(match_results['tsv_to_folder'])}/{len(tsv_df)}명")
    
    # 5. 리포트 생성
    print("\n[5] 검증 리포트 생성 중...")
    report_text = generate_advanced_report(tsv_df, excel_df, folder_mapping, match_results)
    print("\n" + report_text)
    
    # 리포트 저장
    report_path = output_dir / "mammoet_validation_report_advanced.txt"
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report_text)
    print(f"\n   ✓ 리포트 저장: {report_path}")
    
    # JSON 저장
    json_data = {
        'timestamp': datetime.now().isoformat(),
        'match_results': match_results,
        'summary': {
            'tsv_count': len(tsv_df),
            'excel_count': len(excel_df) if excel_df is not None else 0,
            'folder_count': len(folder_mapping),
            'tsv_to_excel_matched': len(match_results['tsv_to_excel']),
            'tsv_to_folder_matched': len(match_results['tsv_to_folder']),
            'unmatched_tsv': len(match_results['unmatched_tsv']),
            'unmatched_excel': len(match_results['unmatched_excel']),
            'unmatched_folder': len(match_results['unmatched_folder'])
        }
    }
    
    json_path = output_dir / "mammoet_validation_report_advanced.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    print(f"   ✓ JSON 리포트 저장: {json_path}")
    
    print("\n" + "=" * 80)
    print("✅ 검증 완료")
    print("=" * 80)
    
    return match_results

if __name__ == "__main__":
    # 스크립트 파일 위치를 기준으로 경로 설정
    script_dir = Path(__file__).parent.absolute()
    base_dir = script_dir
    
    tsv_path = base_dir / "S.N.tsv"
    
    # Excel 파일 경로 자동 탐색 (우선순위: gatepass > 원본 파일)
    excel_paths = [
        base_dir / "mammoet_gatepass.xlsx",  # Gate Pass 파일 우선
        base_dir / "15111578 - Samsung HVDC - Mina Zayed Manpower - 2026.xlsx",
        base_dir / "Mammoet Mina Zayed Manpower - 2026 - Part 1" / "15111578 - Samsung HVDC - Mina Zayed Manpower - 2026.xlsx"
    ]
    
    excel_path = None
    for path in excel_paths:
        if path.exists():
            excel_path = path
            print(f"   [INFO] Excel 파일 발견: {path}")
            break
    
    if excel_path is None:
        print("[ERROR] Excel 파일을 찾을 수 없습니다.")
        print(f"   검색 경로:")
        for path in excel_paths:
            print(f"     - {path} (존재: {path.exists()})")
        sys.exit(1)
    
    folder_path = base_dir / "Mammoet Mina Zayed Manpower - 2026 - Part 1"
    
    result = validate_data_advanced(
        str(tsv_path),
        str(excel_path),
        str(folder_path),
        output_dir=str(base_dir)
    )
