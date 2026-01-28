import csv
import json
from pathlib import Path

def csv_to_json(csv_file_path, json_file_path=None):
    """
    CSV 파일을 JSON 형식으로 변환
    
    Args:
        csv_file_path: 입력 CSV 파일 경로
        json_file_path: 출력 JSON 파일 경로 (None이면 자동 생성)
    """
    # 헤더 정의 (CSV의 헤더가 여러 줄에 걸쳐 있어서 직접 정의)
    headers = [
        "Batch",
        "Loading Date",
        "Sub-Con",
        "ID Number",
        "Description of Material",
        "Size(mm or bag)",
        "Delivery Qty. in Ton",
        "OFCO INVOICE NO"
    ]
    
    data = []
    
    # CSV 파일 읽기 (탭 구분)
    with open(csv_file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
        # 첫 3줄은 헤더이므로 스킵하고 4번째 줄부터 시작
        for line_num, line in enumerate(lines[3:], start=4):
            # 빈 행 스킵
            if not line.strip():
                continue
            
            # 탭으로 분리
            row = line.rstrip('\n').split('\t')
            
            # 빈 셀 정리
            row = [cell.strip() for cell in row]
            
            # 컬럼 수 맞추기 (8개 컬럼)
            while len(row) < 8:
                row.append("")
            
            # 딕셔너리 생성
            record = {}
            for i, header in enumerate(headers):
                value = row[i] if i < len(row) else ""
                
                # 숫자 변환 시도 (Delivery Qty. in Ton)
                if header == "Delivery Qty. in Ton" and value:
                    try:
                        # 공백 제거 후 숫자 변환
                        value = float(value.strip())
                    except (ValueError, AttributeError):
                        # 변환 실패 시 원본 유지
                        pass
                
                record[header] = value
            
            data.append(record)
    
    # JSON 출력
    if json_file_path is None:
        json_file_path = str(Path(csv_file_path).with_suffix('.json'))
    
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 변환 완료: {len(data)}개 레코드")
    print(f"[FILE] 출력 파일: {json_file_path}")
    
    return data, json_file_path

# 실행
if __name__ == "__main__":
    csv_file = "Untitled-1.csv"
    json_data, output_file = csv_to_json(csv_file, "Untitled-1.json")
    
    # 통계 출력
    print(f"\n[STATS] 변환 통계:")
    print(f"   - 총 레코드 수: {len(json_data)}")
    print(f"   - 배치 번호 범위: {min([r['Batch'] for r in json_data if r['Batch']])} ~ {max([r['Batch'] for r in json_data if r['Batch']])}")
    print(f"   - 날짜 범위: {min([r['Loading Date'] for r in json_data if r['Loading Date']])} ~ {max([r['Loading Date'] for r in json_data if r['Loading Date']])}")
    
    # 샘플 출력 (처음 2개만)
    print("\n[SAMPLE] 샘플 데이터 (처음 2개):")
    print(json.dumps(json_data[:2], ensure_ascii=False, indent=2))

