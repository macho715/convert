import re
import sys
import io
from datetime import datetime
from collections import defaultdict

# UTF-8 출력 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

file_path = r'c:\Users\SAMSUNG\Downloads\CONVERT\DAS HISTOR.MD'
output_path = r'c:\Users\SAMSUNG\Downloads\CONVERT\DAS_HISTOR_Organized.md'

# Read the file
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 날짜 패턴: 24/MM/DD AM/PM H:MM
date_pattern = re.compile(r'^24/(\d{1,2})/(\d{1,2})\s+(AM|PM)\s+(\d{1,2}):(\d{2})')

# 날짜별로 메시지 그룹화
messages_by_date = defaultdict(list)
current_date = None
current_time = None
current_message = []
current_datetime = None

for line in lines:
    line = line.rstrip('\n\r')
    match = date_pattern.match(line)
    
    if match:
        # 이전 메시지 저장
        if current_date and current_message:
            messages_by_date[current_date].append({
                'datetime': current_datetime,
                'time': current_time,
                'content': '\n'.join(current_message)
            })
        
        # 새 날짜/시간 파싱
        month = match.group(1).zfill(2)
        day = match.group(2).zfill(2)
        am_pm = match.group(3)
        hour = int(match.group(4))
        minute = match.group(5)
        
        # 24시간 형식으로 변환
        if am_pm == 'PM' and hour != 12:
            hour += 12
        elif am_pm == 'AM' and hour == 12:
            hour = 0
        
        current_date = f'2024-{month}-{day}'
        current_time = f'{hour:02d}:{minute}'
        current_datetime = datetime.strptime(f'{current_date} {current_time}', '%Y-%m-%d %H:%M')
        
        # 날짜 형식 변경된 라인으로 저장
        new_date_str = f'{current_date} {current_time}'
        current_message = [line.replace(match.group(0), new_date_str)]
    else:
        # 메시지 내용 추가
        if current_date:
            if line.strip() or current_message:  # 빈 줄도 포함
                current_message.append(line)

# 마지막 메시지 저장
if current_date and current_message:
    messages_by_date[current_date].append({
        'datetime': current_datetime,
        'time': current_time,
        'content': '\n'.join(current_message)
    })

# 날짜순 정렬 및 마크다운 생성
md_content = []
md_content.append("# DAS Transformer Project - Communication History\n\n")
md_content.append(f"**기간**: {min(messages_by_date.keys())} ~ {max(messages_by_date.keys())}\n\n")
md_content.append(f"**총 일수**: {len(messages_by_date)}일\n\n")
md_content.append(f"**총 메시지 수**: {sum(len(msgs) for msgs in messages_by_date.values())}개\n\n")
md_content.append("---\n\n")

# 날짜순 정렬
sorted_dates = sorted(messages_by_date.keys())

for date in sorted_dates:
    # 날짜 헤더
    date_obj = datetime.strptime(date, '%Y-%m-%d')
    weekday = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'][date_obj.weekday()]
    date_formatted = date_obj.strftime(f'%Y년 %m월 %d일 ({weekday})')
    md_content.append(f"\n## {date_formatted}\n\n")
    md_content.append(f"**날짜**: `{date}`\n\n")
    
    # 시간순 정렬된 메시지
    messages = sorted(messages_by_date[date], key=lambda x: x['datetime'])
    md_content.append(f"**메시지 수**: {len(messages)}개\n\n")
    md_content.append("---\n\n")
    
    for msg in messages:
        md_content.append(f"### {msg['time']}\n\n")
        # 메시지 내용을 코드 블록으로 감싸기
        content_lines = msg['content'].split('\n')
        md_content.append("```\n")
        for content_line in content_lines:
            md_content.append(f"{content_line}\n")
        md_content.append("```\n\n")
        md_content.append("---\n\n")

# 파일 저장
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(''.join(md_content))

print("=" * 60)
print("마크다운 문서 변환 완료!")
print("=" * 60)
print(f"입력 파일: {file_path}")
print(f"출력 파일: {output_path}")
print(f"\n날짜 범위: {min(sorted_dates)} ~ {max(sorted_dates)}")
print(f"총 일수: {len(sorted_dates)}일")
print(f"총 메시지 수: {sum(len(msgs) for msgs in messages_by_date.values())}개")
print("\n날짜별 메시지 수 (처음 10일):")
for date in sorted_dates[:10]:
    print(f"  {date}: {len(messages_by_date[date])}개")
if len(sorted_dates) > 10:
    print(f"  ... 외 {len(sorted_dates)-10}일")
print("=" * 60)
