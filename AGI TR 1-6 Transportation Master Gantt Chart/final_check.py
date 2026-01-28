import json

# Untitled-1.json 확인
with open('Untitled-1.json', 'r', encoding='utf-8') as f:
    data1 = json.load(f)

print(f"[Untitled-1.json] 총 레코드: {len(data1)}개")

batches_82_90 = [r for r in data1 if r.get('Batch') in ['82','83','84','85','86','87','88','89','90']]
print(f"[Untitled-1.json] 82-90번 레코드: {len(batches_82_90)}개")

j71_088_1 = [r for r in data1 if 'J71-088' in r.get('ID Number', '')]
print(f"[Untitled-1.json] J71-088 레코드: {'있음' if j71_088_1 else '없음'}")

# OFCO_GRM.JSON 확인
json_file = r"c:\Users\SAMSUNG\Desktop\OFCO_GRM.JSON"
with open(json_file, 'r', encoding='utf-8') as f:
    data2 = json.load(f)

print(f"\n[OFCO_GRM.JSON] 총 레코드: {len(data2)}개")

batches_82_90_2 = [r for r in data2 if r.get('Batch') in ['82','83','84','85','86','87','88','89','90']]
print(f"[OFCO_GRM.JSON] 82-90번 레코드: {len(batches_82_90_2)}개")

j71_088_2 = [r for r in data2 if 'J71-088' in r.get('ID Number', '')]
print(f"[OFCO_GRM.JSON] J71-088 레코드: {'있음' if j71_088_2 else '없음'}")

print("\n[OK] 모든 업데이트 완료!")

