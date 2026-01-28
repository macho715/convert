import json
from pathlib import Path

def add_new_records(json_file_path, new_data_text):
    """
    JSON 파일에 새로운 레코드 추가
    
    Args:
        json_file_path: JSON 파일 경로
        new_data_text: 새 데이터 텍스트
    """
    # 기존 JSON 파일 읽기
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 새 레코드 파싱
    new_records = []
    lines = new_data_text.strip().split('\n')
    
    for line in lines:
        if not line.strip():
            continue
        
        # 탭으로 분리
        parts = [p.strip() for p in line.split('\t')]
        
        # 빈 행 스킵
        if not any(parts):
            continue
        
        # 레코드 생성
        record = {
            "Batch": parts[0] if len(parts) > 0 else "",
            "Loading Date": parts[1] if len(parts) > 1 else "",
            "Sub-Con": parts[2] if len(parts) > 2 else "",
            "ID Number": parts[3] if len(parts) > 3 else "",
            "Description of Material": parts[4] if len(parts) > 4 else "",
            "Size(mm or bag)": parts[5] if len(parts) > 5 else "",
            "Delivery Qty. in Ton": None,
            "OFCO INVOICE NO": parts[7] if len(parts) > 7 else ""
        }
        
        # Delivery Qty. in Ton 숫자 변환
        if len(parts) > 6 and parts[6]:
            try:
                record["Delivery Qty. in Ton"] = float(parts[6])
            except (ValueError, TypeError):
                record["Delivery Qty. in Ton"] = None
        
        new_records.append(record)
    
    # 기존 데이터와 새 데이터 병합 (중복 제거)
    existing_batches = {r.get("Batch", "") for r in data}
    
    for record in new_records:
        batch = record.get("Batch", "")
        # 같은 Batch가 있으면 업데이트, 없으면 추가
        if batch:
            # 기존 레코드 찾기
            found = False
            for i, existing in enumerate(data):
                if existing.get("Batch") == batch and existing.get("ID Number") == record.get("ID Number"):
                    data[i] = record  # 업데이트
                    found = True
                    break
            
            if not found:
                data.append(record)  # 새 레코드 추가
    
    # Batch 번호로 정렬 (숫자와 문자 혼합 처리)
    def sort_key(r):
        batch = r.get("Batch", "")
        if not batch:
            return (999, "")
        try:
            # 숫자로 시작하는 경우
            if batch[0].isdigit():
                return (0, int(batch) if batch.isdigit() else (int(batch.rstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ')), batch))
            else:
                # 문자로 시작하는 경우 (D1, A1, E1 등)
                return (1, batch)
        except:
            return (2, batch)
    
    data.sort(key=sort_key)
    
    # JSON 파일 저장
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 업데이트 완료: 총 {len(data)}개 레코드")
    print(f"[NEW] 추가/업데이트된 레코드: {len(new_records)}개")
    
    return data

# 새 데이터
new_data = """75	2025-10-09	GRM	HVDC-AGI-GRM-J71-072	AGGREGATE 5MM	5	740.00 	
76	2025-10-14	GRM	HVDC-AGI-GRM-J71-073	AGGREGATE 10MM	10	772.50 	
77	2025-10-19	GRM	HVDC-AGI-GRM-J71-074	AGGREGATE 20MM	20	755.46 	
78	2025-10-21	GRM	HVDC-AGI-GRM-J71-075	AGGREGATE 5MM	5	807.48 	
79	2025-10-24	GRM	HVDC-AGI-GRM-J71-076	AGGREGATE 10MM	10	694.70 	
80	2025-10-27	GRM	HVDC-AGI-GRM-J71-077	AGGREGATE 20MM	20	641.58 	
81	2025-10-31	GRM	HVDC-AGI-GRM-J71-078	AGGREGATE 5MM	5	738.70 	
82	2025-11-05	GRM	HVDC-AGI-GRM-J71-079	AGGREGATE 10MM	10	784.96 	
83	2025-11-11	GRM	HVDC-AGI-GRM-J71-080	DUNE SAND	Dunesand	758.56 	
84	2025-11-15	GRM	HVDC-AGI-GRM-J71-081	AGGREGATE 5MM	5	715.48 	
85	2025-11-20	GRM	HVDC-AGI-GRM-J71-082	DUNE SAND	Dunesand	735.62 	
86	2025-11-26	GRM	HVDC-AGI-GRM-J71-083	AGGREGATE 20MM	20	615.74 	
87	2025-12-05	GRM	HVDC-AGI-GRM-J71-084	AGGREGATE 5MM	5	662.04 	
88	2025-12-08	GRM	HVDC-AGI-GRM-J71-085	DUNE SAND	Dunesand	748.62 	
89	2025-11-13	GRM	HVDC-AGI-GRM-J71-086	AGGREGATE 20MM	20	734.12 	
90	2025-12-26	GRM	HVDC-AGI-GRM-J71-087	AGGREGATE 10MM	10	760.14 	
			HVDC-AGI-GRM-J71-088	DUNE SAND	Dunesand		"""

if __name__ == "__main__":
    # 두 파일 모두 업데이트
    files_to_update = [
        "Untitled-1.json",
        r"c:\Users\SAMSUNG\Desktop\OFCO_GRM.JSON"
    ]
    
    for json_file in files_to_update:
        if Path(json_file).exists():
            print(f"\n[FILE] 업데이트 중: {json_file}")
            add_new_records(json_file, new_data)
        else:
            print(f"[SKIP] 파일 없음: {json_file}")

