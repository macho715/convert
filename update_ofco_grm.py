import json

# OFCO_GRM.JSON 파일 읽기
json_file = r"c:\Users\SAMSUNG\Desktop\OFCO_GRM.JSON"

with open(json_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

# J71-088 레코드 찾기
j71_088 = [r for r in data if 'J71-088' in r.get('ID Number', '')]

print('J71-088 레코드 확인:')
if j71_088:
    print('이미 존재함')
    print(json.dumps(j71_088, ensure_ascii=False, indent=2))
else:
    print('없음 - 추가 중...')
    
    # 빈 Batch 레코드 추가
    new_record = {
        "Batch": "",
        "Loading Date": "",
        "Sub-Con": "GRM",
        "ID Number": "HVDC-AGI-GRM-J71-088",
        "Description of Material": "DUNE SAND",
        "Size(mm or bag)": "Dunesand",
        "Delivery Qty. in Ton": None,
        "OFCO INVOICE NO": ""
    }
    
    data.append(new_record)
    
    # 저장
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f'J71-088 레코드 추가 완료. 총 {len(data)}개 레코드')

