"""
December Tide Table 2025 PDF 추출 및 구조화 스크립트
조석표 데이터를 CSV/Excel/Markdown 형식으로 변환
"""

import pdfplumber
import pandas as pd
from pathlib import Path
from datetime import datetime
import json
import re

def extract_tide_table(pdf_path: str) -> dict:
    """
    조석표 PDF에서 데이터 추출
    
    Returns:
        dict: {
            'metadata': {...},
            'tables': [...],
            'text': '...',
            'raw_data': [...]
        }
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
    
    result = {
        'source': str(pdf_path),
        'extracted_at': datetime.now().isoformat(),
        'pages': 0,
        'tables': [],
        'text': [],
        'metadata': {}
    }
    
    print(f"📄 PDF 처리 중: {pdf_path.name}")
    
    with pdfplumber.open(str(pdf_path)) as pdf:
        result['pages'] = len(pdf.pages)
        result['metadata'] = {
            'total_pages': len(pdf.pages),
            'title': pdf.metadata.get('Title', ''),
            'author': pdf.metadata.get('Author', ''),
            'subject': pdf.metadata.get('Subject', '')
        }
        
        all_text = []
        all_tables = []
        
        for i, page in enumerate(pdf.pages, start=1):
            print(f"  페이지 {i}/{len(pdf.pages)} 처리 중...")
            
            # 텍스트 추출
            text = page.extract_text()
            if text:
                all_text.append(f"=== Page {i} ===\n{text}\n")
            
            # 테이블 추출
            tables = page.extract_tables()
            if tables:
                for j, table in enumerate(tables):
                    if table and len(table) > 0:
                        all_tables.append({
                            'page': i,
                            'table_index': j,
                            'rows': len(table),
                            'columns': len(table[0]) if table[0] else 0,
                            'data': table
                        })
                        print(f"    ✓ 테이블 {j+1} 발견: {len(table)}행 x {len(table[0]) if table[0] else 0}열")
        
        result['text'] = '\n'.join(all_text)
        result['tables'] = all_tables
    
    return result

def process_tide_data(extracted_data: dict) -> pd.DataFrame:
    """
    추출된 조석표 데이터를 구조화된 DataFrame으로 변환
    구조: 행(세로) = 시간대 (0:00 ~ 23:00), 열(가로) = 날짜 (01-Dec ~ 31-Dec)
    """
    if not extracted_data['tables']:
        print("⚠️  테이블을 찾을 수 없습니다.")
        return pd.DataFrame()
    
    # 가장 큰 테이블 선택
    largest_table = max(extracted_data['tables'], key=lambda t: t['rows'] * t['columns'])
    
    print(f"\n📊 메인 테이블 처리: {largest_table['rows']}행 x {largest_table['columns']}열")
    
    table_data = largest_table['data']
    
    # 1. 첫 번째 행에서 시간대 추출
    if len(table_data) == 0 or len(table_data[0]) < 2:
        print("⚠️  테이블 구조가 올바르지 않습니다.")
        return pd.DataFrame()
    
    # 첫 번째 행의 두 번째 셀에서 시간대 추출
    time_header = table_data[0][1]
    if time_header and isinstance(time_header, str):
        # "0:00 1:00 2:00 ... 23:00" 형태를 리스트로 변환
        hours = [h.strip() for h in time_header.split() if ':' in h]
        print(f"  ✓ 시간대 추출: {len(hours)}개 ({hours[0]} ~ {hours[-1]})")
    else:
        hours = [f"{i:02d}:00" for i in range(24)]
        print(f"  ⚠️  시간대 헤더를 찾을 수 없어 기본값 사용: 0:00 ~ 23:00")
    
    # 2. 날짜와 조석값 추출
    dates = []
    tide_values = {}  # {날짜: {시간대: 값}}
    current_date_group = []
    date_group_start_row = -1
    
    for row_idx in range(1, len(table_data)):
        row = table_data[row_idx]
        if not row or len(row) < 2:
            continue
        
        # 첫 번째 열에서 날짜 추출
        date_cell = row[0]
        
        # 날짜가 있는 행: 새로운 날짜 그룹 시작
        if date_cell and isinstance(date_cell, str) and 'Dec' in date_cell:
            # 여러 날짜가 줄바꿈으로 구분되어 있음
            date_list = [d.strip() for d in date_cell.split('\n') if d.strip() and 'Dec' in d]
            current_date_group = date_list
            date_group_start_row = row_idx
            # 이 날짜들을 dates에 추가
            for date in date_list:
                if date not in dates:
                    dates.append(date)
                    tide_values[date] = {}
        
        # 두 번째 열부터가 조석값들 (시간대 순서대로)
        values = []
        for col_idx in range(1, min(len(row), len(hours) + 1)):
            val = row[col_idx]
            if val is not None and val != '':
                val_str = str(val).strip()
                # 공백으로 구분된 여러 값 처리 (예: "0.93 0.93")
                if ' ' in val_str:
                    val_parts = val_str.split()
                    for v in val_parts:
                        try:
                            float(v)
                            values.append(v)
                        except:
                            pass
                else:
                    try:
                        float(val_str)
                        values.append(val_str)
                    except:
                        pass
        
        # 현재 날짜 그룹이 있고, 값들이 있으면 매핑
        if current_date_group and len(values) >= len(hours):
            # 현재 행이 날짜 그룹의 몇 번째 행인지 계산
            rows_since_group_start = row_idx - date_group_start_row
            if rows_since_group_start < len(current_date_group):
                target_date = current_date_group[rows_since_group_start]
                # 시간대별로 값 할당
                for hour_idx, hour in enumerate(hours):
                    if hour_idx < len(values):
                        try:
                            tide_values[target_date][hour] = float(values[hour_idx])
                        except:
                            pass
    
    print(f"  ✓ 날짜 추출: {len(dates)}개 ({dates[0] if dates else 'N/A'} ~ {dates[-1] if dates else 'N/A'})")
    
    # 3. DataFrame 생성: 행=시간대, 열=날짜
    if not dates or not hours:
        print("⚠️  날짜나 시간대를 찾을 수 없습니다.")
        return pd.DataFrame()
    
    # 데이터 행렬 구성
    data_matrix = []
    for hour in hours:
        row = []
        for date in dates:
            if date in tide_values and hour in tide_values[date]:
                row.append(tide_values[date][hour])
            else:
                row.append(None)
        data_matrix.append(row)
    
    # DataFrame 생성 (행=시간대, 열=날짜)
    df = pd.DataFrame(data_matrix, index=hours, columns=dates)
    
    # 인덱스 이름 설정
    df.index.name = 'Time'
    
    print(f"\n✓ 최종 DataFrame: {len(df)}행(시간대) x {len(df.columns)}열(날짜)")
    print(f"  샘플: {df.iloc[0, 0]}m @ {df.index[0]} on {df.columns[0]}")
    
    return df

def save_results(extracted_data: dict, df: pd.DataFrame, output_dir: str = "tide_extracted"):
    """
    추출 결과를 다양한 형식으로 저장
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    base_name = "December_Tide_Table_2025"
    
    # 1. Markdown 형식
    md_path = output_path / f"{base_name}.md"
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# December Tide Table 2025\n\n")
        f.write(f"**추출 일시:** {extracted_data['extracted_at']}\n")
        f.write(f"**원본 파일:** {extracted_data['source']}\n")
        f.write(f"**총 페이지:** {extracted_data['metadata']['total_pages']}\n\n")
        
        if extracted_data['text']:
            f.write("## 추출된 텍스트\n\n")
            f.write(extracted_data['text'])
            f.write("\n\n")
        
        if not df.empty:
            f.write("## 조석표 데이터\n\n")
            # 간단한 마크다운 테이블 생성 (tabulate 없이)
            f.write("| " + " | ".join(str(col) for col in df.columns) + " |\n")
            f.write("| " + " | ".join(["---"] * len(df.columns)) + " |\n")
            for _, row in df.iterrows():
                f.write("| " + " | ".join(str(val) if pd.notna(val) else "" for val in row) + " |\n")
            f.write("\n")
    
    print(f"✓ Markdown 저장: {md_path}")
    
    # 2. CSV 형식
    if not df.empty:
        csv_path = output_path / f"{base_name}.csv"
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"✓ CSV 저장: {csv_path}")
        
        # Excel 형식
        xlsx_path = output_path / f"{base_name}.xlsx"
        try:
            with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Tide_Table', index=True)
                
                # 추가 시트: 모든 테이블
                if len(extracted_data['tables']) > 1:
                    for idx, table_info in enumerate(extracted_data['tables']):
                        if table_info['data']:
                            table_df = pd.DataFrame(table_info['data'])
                            if len(table_df) > 0:
                                header_row = table_df.iloc[0]
                                table_df.columns = [str(col) if col is not None else f"Column_{i}" for i, col in enumerate(header_row)]
                                table_df = table_df.iloc[1:].reset_index(drop=True)
                            table_df.columns = [str(col) if col is not None else f"Column_{i}" for i, col in enumerate(table_df.columns)]
                            table_df.to_excel(writer, sheet_name=f'Table_{table_info["page"]}_{idx}', index=False)
            
            print(f"✓ Excel 저장: {xlsx_path}")
        except PermissionError:
            print(f"⚠️  Excel 파일이 열려있어 저장할 수 없습니다: {xlsx_path}")
            print(f"   파일을 닫고 다시 실행하세요.")
    
    # 3. JSON 형식 (전체 데이터)
    json_path = output_path / f"{base_name}_full.json"
    json_data = {
        'metadata': extracted_data['metadata'],
        'extracted_at': extracted_data['extracted_at'],
        'source': extracted_data['source'],
        'tables': [
            {
                'page': t['page'],
                'table_index': t['table_index'],
                'data': t['data']
            }
            for t in extracted_data['tables']
        ],
        'text': extracted_data['text']
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ JSON 저장: {json_path}")
    
    # 4. 구조화된 DataFrame JSON
    if not df.empty:
        json_df_path = output_path / f"{base_name}_structured.json"
        df.to_json(json_df_path, orient='records', force_ascii=False, indent=2)
        print(f"✓ 구조화된 JSON 저장: {json_df_path}")

def main():
    pdf_path = "December Tide Table 2025.pdf"
    
    try:
        # 1. PDF에서 데이터 추출
        print("=" * 60)
        print("🌊 조석표 PDF 추출 시작")
        print("=" * 60)
        extracted_data = extract_tide_table(pdf_path)
        
        # 2. 조석표 데이터 구조화
        print("\n" + "=" * 60)
        print("📊 데이터 구조화")
        print("=" * 60)
        df = process_tide_data(extracted_data)
        
        if not df.empty:
            print(f"\n✓ 처리 완료: {len(df)}행의 데이터")
            column_names = [str(col) if col is not None else f"Column_{i}" for i, col in enumerate(df.columns)]
            print(f"\n컬럼: {', '.join(column_names)}")
            print(f"\n샘플 데이터 (처음 5행):")
            print(df.head().to_string())
        
        # 3. 결과 저장
        print("\n" + "=" * 60)
        print("💾 파일 저장")
        print("=" * 60)
        save_results(extracted_data, df)
        
        print("\n" + "=" * 60)
        print("✅ 완료!")
        print("=" * 60)
        print(f"\n결과 파일은 'tide_extracted' 폴더에 저장되었습니다.")
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())

