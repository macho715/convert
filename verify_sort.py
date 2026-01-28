import json

# JSON 파일 읽기
with open('Untitled-1.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print('날짜 순서 확인:')
print(f'첫 번째: {data[0]["Loading Date"]} (Batch {data[0]["Batch"]})')
print(f'50번째: {data[49]["Loading Date"]} (Batch {data[49]["Batch"]})')
print(f'100번째: {data[99]["Loading Date"]} (Batch {data[99]["Batch"]})')
print(f'마지막: {data[-1]["Loading Date"]} (Batch {data[-1]["Batch"]})')

# 날짜 순서 검증
print('\n날짜 순서 검증:')
prev_date = None
errors = []
for i, record in enumerate(data):
    date = record.get("Loading Date", "")
    if date:
        if prev_date and date < prev_date:
            errors.append(f"순서 오류: 인덱스 {i}, 날짜 {date} < 이전 날짜 {prev_date}")
        prev_date = date

if errors:
    print(f"오류 발견: {len(errors)}개")
    for error in errors[:5]:
        print(f"  {error}")
else:
    print("정렬 완료: 모든 레코드가 날짜 순서대로 정렬되었습니다.")

