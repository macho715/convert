import pandas as pd
import json
from datetime import datetime
from pathlib import Path

def convert_csv_to_json(input_file: str, output_file: str = None):
    """CSV(TSV) 파일을 JSON으로 변환 - 원본 NO 순서대로 모든 데이터 포함"""
    
    if output_file is None:
        output_file = Path(input_file).stem + '.json'
    
    # 컬럼명 정의
    column_names = [
        'NO',
        'Batch',
        'Loading Date',
        'Sub-Con',
        'ID Number',
        'Description of Material',
        'Size',
        'Delivery Qty in Ton',
        'OFCO INVOICE NO'
    ]
    
    # 파일을 직접 읽어서 처리 (pandas skiprows 문제 회피)
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 헤더 3줄 건너뛰고 데이터 추출
    data_lines = []
    for line in lines[3:]:  # 4번째 줄부터
        stripped = line.strip()
        if stripped:  # 빈 줄이 아닌 경우만
            parts = stripped.split('\t')
            if parts[0].strip():  # NO 컬럼이 있는 경우만
                # 컬럼 수가 부족한 경우 빈 문자열로 채우기
                while len(parts) < len(column_names):
                    parts.append('')
                data_lines.append(parts[:len(column_names)])
    
    # DataFrame 생성
    df = pd.DataFrame(data_lines, columns=column_names)
    
    # NO를 숫자로 변환하여 정렬 (원본 순서 유지)
    df['NO_numeric'] = pd.to_numeric(df['NO'], errors='coerce')
    df = df.sort_values('NO_numeric', na_position='last')
    df = df.drop('NO_numeric', axis=1)
    
    # Delivery Qty in Ton을 숫자로 변환 (필요시)
    df['Delivery Qty in Ton'] = pd.to_numeric(
        df['Delivery Qty in Ton'].str.strip(), 
        errors='coerce'
    )
    
    # 모든 문자열 필드의 앞뒤 공백 제거
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
    
    # JSON 구조 생성
    result = {
        'meta': {
            'source': str(Path(input_file).name),
            'type': 'csv',
            'total_records': len(df),
            'parsed_at': datetime.now().isoformat()
        },
        'data': df.to_dict('records')
    }
    
    # NaN 값을 None으로 변환 (JSON null)
    import numpy as np
    def replace_nan(obj):
        if isinstance(obj, dict):
            return {k: replace_nan(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [replace_nan(item) for item in obj]
        elif isinstance(obj, float) and (pd.isna(obj) or np.isnan(obj)):
            return None
        return obj
    
    result = replace_nan(result)
    
    # JSON 파일 저장
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 변환 완료: {output_file}")
    print(f"  - 총 레코드 수: {len(df)}")
    if len(df) > 0:
        print(f"  - 첫 번째 NO: {df.iloc[0]['NO']}")
        print(f"  - 마지막 NO: {df.iloc[-1]['NO']}")
    print(f"  - 출력 파일: {output_file}")
    
    return output_file

if __name__ == '__main__':
    convert_csv_to_json('JPT_GRM.csv', 'JPT_GRM.json')
