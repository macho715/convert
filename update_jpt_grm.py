import json
from datetime import datetime

def update_json_from_csv(csv_path, json_path):
    """CSV 파일을 파싱하여 JSON 파일 업데이트"""
    
    # CSV 파일 읽기
    with open(csv_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 헤더 3줄 스킵, 데이터 파싱
    data = []
    for line in lines[3:]:  # 4번째 줄부터 시작
        if not line.strip():
            continue
        
        # 탭으로 분리
        parts = line.rstrip('\n').split('\t')
        
        # 필드 추출 (9개 컬럼)
        if len(parts) >= 9:
            # Delivery Qty in Ton 숫자 변환
            try:
                qty = float(parts[7].strip()) if parts[7].strip() else 0.0
            except ValueError:
                qty = 0.0
            
            record = {
                "NO": parts[0].strip(),
                "Batch": parts[1].strip(),
                "Loading Date": parts[2].strip(),
                "Sub-Con": parts[3].strip(),
                "ID Number": parts[4].strip(),
                "Description of Material": parts[5].strip(),
                "Size": parts[6].strip(),
                "Delivery Qty in Ton": qty,
                "OFCO INVOICE NO": parts[8].strip()
            }
            data.append(record)
    
    # JSON 구조 생성
    output = {
        "meta": {
            "source": "JPT71_VOYAGE.csv",
            "type": "csv",
            "total_records": len(data),
            "parsed_at": datetime.now().isoformat()
        },
        "data": data
    }
    
    # JSON 파일 저장
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] JSON file updated: {len(data)} records")
    return output

# 실행
if __name__ == "__main__":
    csv_file = r"c:\Users\SAMSUNG\Downloads\CONVERT\JPT71_VOYAGE.csv"
    json_file = r"c:\Users\SAMSUNG\Downloads\CONVERT\JPT_GRM.json"
    
    update_json_from_csv(csv_file, json_file)
