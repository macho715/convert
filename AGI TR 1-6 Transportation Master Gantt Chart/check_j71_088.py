import json

# JSON 파일 읽기
with open('Untitled-1.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# J71-088 레코드 찾기
j71_088 = [r for r in data if 'J71-088' in r.get('ID Number', '')]

print('J71-088 레코드:')
if j71_088:
    print(json.dumps(j71_088, ensure_ascii=False, indent=2))
else:
    print('없음 - 추가 필요')
    
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
    with open('Untitled-1.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print('J71-088 레코드 추가 완료')

