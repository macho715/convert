import json

# JSON 파일 읽기
with open('Untitled-1.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"총 레코드 수: {len(data)}")

# 배치 번호 추출
batches = [r['Batch'] for r in data if r.get('Batch')]
print(f"배치 범위: {min(batches)} ~ {max(batches)}")

# 새로 추가된 레코드 확인 (82-90)
print(f"\n새로 추가된 레코드 (82-90):")
new_recs = [r for r in data if r.get('Batch') in ['82','83','84','85','86','87','88','89','90']]
for r in sorted(new_recs, key=lambda x: int(x['Batch']) if x['Batch'].isdigit() else 999):
    print(f"  Batch {r['Batch']}: {r['ID Number']} - {r['Description of Material']} - {r.get('Delivery Qty. in Ton', 'N/A')} Ton")

# 마지막 레코드 확인
print(f"\n마지막 레코드:")
print(json.dumps(data[-1], ensure_ascii=False, indent=2))

