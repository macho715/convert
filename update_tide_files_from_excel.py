"""
수정된 Excel 파일을 읽어서 다른 형식의 파일들(CSV, Markdown, JSON)을 재생성
"""

import pandas as pd
import json
from pathlib import Path
from datetime import datetime

def read_excel_file(excel_path: str) -> pd.DataFrame:
    """
    Excel 파일을 읽어서 DataFrame으로 반환
    구조: 행(세로) = 날짜, 열(가로) = 시간대
    """
    print(f"📖 Excel 파일 읽기: {excel_path}")
    df = pd.read_excel(excel_path, index_col=0)
    
    # 현재 구조 확인
    print(f"  ✓ 원본 데이터 크기: {df.shape[0]}행 x {df.shape[1]}열")
    
    # 전치 필요 여부 확인
    first_index = str(df.index[0])
    first_col = str(df.columns[0])
    
    # 날짜 형식인지 확인
    is_date_index = any([
        'Dec' in first_index,
        '2025' in first_index,
        isinstance(df.index[0], pd.Timestamp),
        '01-' in first_index or '02-' in first_index
    ])
    
    # 시간 형식인지 확인
    is_time_col = any([
        ':' in first_col,
        '0:00' in first_col,
        '00:00' in first_col
    ])
    
    if is_date_index and is_time_col:
        # 올바른 구조: 행=날짜, 열=시간대 (전치 불필요)
        print(f"  ✓ 구조 확인: 행=날짜, 열=시간대 (전치 불필요)")
    elif ':' in first_index or '0:00' in first_index:
        # 잘못된 구조: 행=시간대, 열=날짜 → 전치 필요
        print(f"  ⚠️  전치 필요: 현재 구조는 행=시간대, 열=날짜")
        df = df.T  # 전치
        print(f"  ✓ 전치 완료")
    else:
        print(f"  ✓ 구조 확인: 행=날짜, 열=시간대 (전치 불필요)")
    
    print(f"  ✓ 최종 데이터 크기: {df.shape[0]}행(날짜) x {df.shape[1]}열(시간대)")
    print(f"  ✓ 날짜 범위: {df.index[0]} ~ {df.index[-1]}")
    print(f"  ✓ 시간대 범위: {df.columns[0]} ~ {df.columns[-1]}")
    return df

def save_csv(df: pd.DataFrame, output_path: Path):
    """CSV 파일로 저장"""
    csv_path = output_path / "December_Tide_Table_2025.csv"
    df.to_csv(csv_path, index=True, encoding='utf-8-sig')
    print(f"✓ CSV 저장: {csv_path}")

def save_markdown(df: pd.DataFrame, output_path: Path, metadata: dict = None):
    """Markdown 파일로 저장"""
    md_path = output_path / "December_Tide_Table_2025.md"
    
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write("# December Tide Table 2025\n\n")
        f.write(f"**추출 일시:** {metadata.get('extracted_at', datetime.now().isoformat()) if metadata else datetime.now().isoformat()}\n")
        f.write(f"**원본 파일:** {metadata.get('source', 'December Tide Table 2025.xlsx') if metadata else 'December Tide Table 2025.xlsx'}\n")
        f.write(f"**데이터 구조:** 행(세로) = 날짜, 열(가로) = 시간대\n\n")
        
        f.write("## 조석표 데이터\n\n")
        f.write("| 날짜 | " + " | ".join(str(col) for col in df.columns) + " |\n")
        f.write("| " + " | ".join(["---"] * (len(df.columns) + 1)) + " |\n")
        
        for date, row in df.iterrows():
            values = [str(date)]
            for val in row:
                if pd.notna(val):
                    values.append(f"{val:.2f}" if isinstance(val, (int, float)) else str(val))
                else:
                    values.append("")
            f.write("| " + " | ".join(values) + " |\n")
        
        f.write("\n")
        f.write("## 데이터 요약\n\n")
        f.write(f"- **총 날짜:** {len(df)}개\n")
        f.write(f"- **총 시간대:** {len(df.columns)}개\n")
        f.write(f"- **최고 조석:** {df.max().max():.2f}m\n")
        f.write(f"- **최저 조석:** {df.min().min():.2f}m\n")
        f.write(f"- **평균 조석:** {df.mean().mean():.2f}m\n")
    
    print(f"✓ Markdown 저장: {md_path}")

def save_json(df: pd.DataFrame, output_path: Path, metadata: dict = None):
    """JSON 파일로 저장"""
    
    # 1. 구조화된 JSON (시간대별, 날짜별 데이터)
    structured_path = output_path / "December_Tide_Table_2025_structured.json"
    structured_data = {
        'metadata': {
            'source': metadata.get('source', 'December Tide Table 2025.xlsx') if metadata else 'December Tide Table 2025.xlsx',
            'extracted_at': metadata.get('extracted_at', datetime.now().isoformat()) if metadata else datetime.now().isoformat(),
            'structure': 'rows=date, columns=time',
            'date_range': [str(df.index[0]), str(df.index[-1])],
            'time_range': [str(df.columns[0]), str(df.columns[-1])],
            'total_dates': len(df),
            'total_times': len(df.columns)
        },
        'data': []
    }
    
    # 날짜별로 데이터 구성
    for date in df.index:
        date_data = {
            'date': str(date),
            'tide_levels': {}
        }
        for time in df.columns:
            val = df.loc[date, time]
            if pd.notna(val):
                date_data['tide_levels'][str(time)] = float(val)
        structured_data['data'].append(date_data)
    
    with open(structured_path, 'w', encoding='utf-8') as f:
        json.dump(structured_data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ 구조화된 JSON 저장: {structured_path}")
    
    # 2. 전체 데이터 JSON (DataFrame 전체를 JSON으로)
    full_path = output_path / "December_Tide_Table_2025_full.json"
    full_data = {
        'metadata': structured_data['metadata'],
        'table': df.to_dict(orient='index')
    }
    
    # JSON 직렬화를 위해 NaN을 None으로 변환
    full_data['table'] = {
        str(k): {str(col): (float(v) if pd.notna(v) else None) for col, v in row.items()}
        for k, row in df.to_dict(orient='index').items()
    }
    
    with open(full_path, 'w', encoding='utf-8') as f:
        json.dump(full_data, f, ensure_ascii=False, indent=2)
    
    print(f"✓ 전체 JSON 저장: {full_path}")

def main():
    excel_path = "tide_extracted/December_Tide_Table_2025.xlsx"
    output_dir = Path("tide_extracted")
    
    print("=" * 60)
    print("🔄 Excel 파일 기반 파일 재생성")
    print("=" * 60)
    
    try:
        # 1. Excel 파일 읽기
        df = read_excel_file(excel_path)
        
        # 메타데이터 (기존 파일에서 읽거나 새로 생성)
        metadata = {
            'source': 'December Tide Table 2025.xlsx',
            'extracted_at': datetime.now().isoformat()
        }
        
        # 2. 다른 형식으로 저장
        print("\n" + "=" * 60)
        print("💾 파일 저장")
        print("=" * 60)
        
        save_csv(df, output_dir)
        save_markdown(df, output_dir, metadata)
        save_json(df, output_dir, metadata)
        
        print("\n" + "=" * 60)
        print("✅ 완료!")
        print("=" * 60)
        print(f"\n모든 파일이 '{output_dir}' 폴더에 저장되었습니다.")
        print(f"\n생성된 파일:")
        print(f"  - December_Tide_Table_2025.csv")
        print(f"  - December_Tide_Table_2025.md")
        print(f"  - December_Tide_Table_2025_structured.json")
        print(f"  - December_Tide_Table_2025_full.json")
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())

