"""
Gate Pass Excel 파일에서 최종 발급 내역 추출
- Pass Type: Short Term Pass
- Entry Date
- Departure Date
"""
import pandas as pd
from pathlib import Path
import sys
import io
import re
from datetime import datetime

# Windows 콘솔 UTF-8 인코딩 설정
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def extract_gatepass_records(excel_path: Path) -> list:
    """Gate Pass Excel 파일에서 발급 내역 추출"""
    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None, engine='openpyxl')
    
    records = []
    
    # 각 Gate Pass 레코드 찾기 (Full Name이 있는 행부터 시작)
    col_23_idx = 22  # 열 23 (Full Name)
    
    current_record = None
    
    for row_idx in range(len(df_raw)):
        # Full Name 찾기
        if col_23_idx < len(df_raw.columns):
            cell_value = df_raw.iloc[row_idx, col_23_idx]
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                
                # Full Name 패턴 찾기
                if 'full name' in cell_str.lower():
                    match = re.search(r'full\s+name\s+(.+)', cell_str, re.IGNORECASE)
                    if match:
                        name = match.group(1).strip()
                        name = re.sub(r'\s+', ' ', name)
                        name = name.replace('\n', ' ').replace('\xa0', ' ')
                        name = ' '.join(name.split())
                        
                        if len(name) > 3:
                            # 새 레코드 시작
                            if current_record:
                                records.append(current_record)
                            
                            current_record = {
                                'row': row_idx + 1,
                                'Full Name': name,
                                'Pass Type': 'Short Term Pass',
                                'Entry Date': None,
                                'Departure Date': None
                            }
        
        # 현재 레코드가 있으면 날짜 정보 찾기
        if current_record:
            # Entry Date 찾기 (다양한 패턴 시도)
            entry_patterns = ['entry date', 'entry', 'valid from', 'from date', 'arrival date']
            departure_patterns = ['departure date', 'departure', 'valid until', 'to date', 'exit date', 'valid to']
            
            # 같은 행과 다음 몇 행에서 날짜 찾기
            for check_row in range(row_idx, min(row_idx + 20, len(df_raw))):
                for col_idx in range(len(df_raw.columns)):
                    cell_value = df_raw.iloc[check_row, col_idx]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip().lower()
                        
                        # Entry Date 찾기
                        if not current_record['Entry Date']:
                            for pattern in entry_patterns:
                                if pattern in cell_str:
                                    # 다음 셀 또는 같은 행의 다른 셀에서 날짜 찾기
                                    for next_col in range(col_idx, min(col_idx + 3, len(df_raw.columns))):
                                        date_cell = df_raw.iloc[check_row, next_col]
                                        if pd.notna(date_cell):
                                            try:
                                                if isinstance(date_cell, datetime):
                                                    current_record['Entry Date'] = date_cell.strftime('%Y-%m-%d')
                                                    break
                                                elif isinstance(date_cell, str):
                                                    # 날짜 문자열 파싱 시도
                                                    date_str = date_cell.strip()
                                                    if re.match(r'\d{4}-\d{2}-\d{2}', date_str):
                                                        current_record['Entry Date'] = date_str
                                                        break
                                            except:
                                                pass
                        
                        # Departure Date 찾기
                        if not current_record['Departure Date']:
                            for pattern in departure_patterns:
                                if pattern in cell_str:
                                    # 다음 셀 또는 같은 행의 다른 셀에서 날짜 찾기
                                    for next_col in range(col_idx, min(col_idx + 3, len(df_raw.columns))):
                                        date_cell = df_raw.iloc[check_row, next_col]
                                        if pd.notna(date_cell):
                                            try:
                                                if isinstance(date_cell, datetime):
                                                    current_record['Departure Date'] = date_cell.strftime('%Y-%m-%d')
                                                    break
                                                elif isinstance(date_cell, str):
                                                    date_str = date_cell.strip()
                                                    if re.match(r'\d{4}-\d{2}-\d{2}', date_str):
                                                        current_record['Departure Date'] = date_str
                                                        break
                                            except:
                                                pass
        
        # 다음 Full Name을 만나기 전까지 계속 검색
        # 레코드가 너무 길어지면 저장 (다음 Full Name이 50행 이상 떨어져 있으면)
        if current_record and row_idx - current_record['row'] > 50:
            records.append(current_record)
            current_record = None
    
    # 마지막 레코드 추가
    if current_record:
        records.append(current_record)
    
    return records

def extract_gatepass_detailed(excel_path: Path) -> list:
    """Gate Pass Excel 파일에서 상세 정보 추출 (개선 버전)"""
    df_raw = pd.read_excel(excel_path, sheet_name=0, header=None, engine='openpyxl')
    
    records = []
    
    # Full Name이 있는 행 찾기
    col_23_idx = 22
    name_rows = []
    
    for row_idx in range(len(df_raw)):
        if col_23_idx < len(df_raw.columns):
            cell_value = df_raw.iloc[row_idx, col_23_idx]
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                if 'full name' in cell_str.lower():
                    match = re.search(r'full\s+name\s+(.+)', cell_str, re.IGNORECASE)
                    if match:
                        name = match.group(1).strip()
                        name = re.sub(r'\s+', ' ', name)
                        name = name.replace('\n', ' ').replace('\xa0', ' ')
                        name = ' '.join(name.split())
                        if len(name) > 3:
                            name_rows.append((row_idx, name))
    
    # 각 이름에 대해 해당 행 주변에서 날짜 정보 찾기
    for name_row_idx, name in name_rows:
        record = {
            'Full Name': name,
            'Pass Type': 'Short Term Pass',
            'Entry Date': None,
            'Departure Date': None,
            'Row': name_row_idx + 1
        }
        
        # 해당 행부터 다음 이름 행까지 또는 30행까지 검색
        end_row = name_rows[name_rows.index((name_row_idx, name)) + 1][0] if name_rows.index((name_row_idx, name)) + 1 < len(name_rows) else min(name_row_idx + 30, len(df_raw))
        
        # 모든 셀에서 날짜 찾기
        for row_idx in range(name_row_idx, end_row):
            for col_idx in range(len(df_raw.columns)):
                cell_value = df_raw.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    # datetime 객체인 경우
                    if isinstance(cell_value, datetime):
                        date_str = cell_value.strftime('%Y-%m-%d')
                        # Entry Date가 없으면 첫 번째 날짜를 Entry로
                        if not record['Entry Date']:
                            record['Entry Date'] = date_str
                        # 두 번째 날짜를 Departure로
                        elif not record['Departure Date']:
                            record['Departure Date'] = date_str
                    
                    # 문자열인 경우 날짜 패턴 확인
                    elif isinstance(cell_value, str):
                        cell_lower = cell_value.lower().strip()
                        # "Valid Until" 또는 "Valid To" 패턴 찾기
                        if 'valid until' in cell_lower or 'valid to' in cell_lower:
                            # 다음 셀들에서 날짜 찾기
                            for next_col in range(col_idx, min(col_idx + 5, len(df_raw.columns))):
                                next_cell = df_raw.iloc[row_idx, next_col]
                                if pd.notna(next_cell):
                                    if isinstance(next_cell, datetime):
                                        record['Departure Date'] = next_cell.strftime('%Y-%m-%d')
                                        break
                                    elif isinstance(next_cell, str) and re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', next_cell):
                                        record['Departure Date'] = next_cell.strip()
                                        break
                        
                        # "Entry Date" 또는 "From" 패턴 찾기
                        if 'entry' in cell_lower or ('from' in cell_lower and 'date' in cell_lower):
                            for next_col in range(col_idx, min(col_idx + 5, len(df_raw.columns))):
                                next_cell = df_raw.iloc[row_idx, next_col]
                                if pd.notna(next_cell):
                                    if isinstance(next_cell, datetime):
                                        record['Entry Date'] = next_cell.strftime('%Y-%m-%d')
                                        break
                                    elif isinstance(next_cell, str) and re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', next_cell):
                                        record['Entry Date'] = next_cell.strip()
                                        break
        
        # 열 28 (인덱스 27)도 확인 (이전 분석에서 날짜가 있었음)
        if col_23_idx + 5 < len(df_raw.columns):
            for row_idx in range(name_row_idx, min(name_row_idx + 10, len(df_raw))):
                date_cell = df_raw.iloc[row_idx, 27]  # 열 28
                if pd.notna(date_cell) and isinstance(date_cell, datetime):
                    if not record['Entry Date']:
                        record['Entry Date'] = date_cell.strftime('%Y-%m-%d')
                    elif not record['Departure Date']:
                        record['Departure Date'] = date_cell.strftime('%Y-%m-%d')
        
        records.append(record)
    
    return records

# 실행
script_dir = Path(__file__).parent.absolute()
excel_path = script_dir / "mammoet_gatepass.xlsx"

if not excel_path.exists():
    print(f"❌ Excel 파일을 찾을 수 없습니다: {excel_path}")
    sys.exit(1)

print("="*80)
print("📋 Gate Pass 최종 발급 내역 추출")
print("="*80)

records = extract_gatepass_detailed(excel_path)

print(f"\n추출된 레코드: {len(records)}개\n")

# DataFrame으로 변환하여 Excel로 저장
df_output = pd.DataFrame(records)
df_output = df_output[['Full Name', 'Pass Type', 'Entry Date', 'Departure Date']]

# 출력
print("="*80)
print("📊 Gate Pass 발급 내역")
print("="*80)
print(df_output.to_string(index=False))

# Excel 파일로 저장
output_path = script_dir / "mammoet_gatepass_final_issue.xlsx"
df_output.to_excel(output_path, index=False, engine='openpyxl')
print(f"\n✓ Excel 파일 저장: {output_path}")

# CSV로도 저장
csv_path = script_dir / "mammoet_gatepass_final_issue.csv"
df_output.to_csv(csv_path, index=False, encoding='utf-8-sig')
print(f"✓ CSV 파일 저장: {csv_path}")

print("\n" + "="*80)
print("✅ 추출 완료")
print("="*80)
