#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EMAI3333.MD를 날짜별로 정렬하여 AGI Transformers Transportation_email_20251229.md에 업데이트
"""

import re
from datetime import datetime
from pathlib import Path

def parse_date(date_str):
    """이메일 날짜 문자열을 datetime 객체로 변환"""
    months = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12
    }
    
    # 패턴 순서 중요: 더 구체적인 것부터
    # 1. "Friday, January 16, 2026" 형식
    pattern1 = r'(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})'
    match = re.search(pattern1, date_str, re.IGNORECASE)
    if match:
        try:
            month_name = match.group(1)
            day = int(match.group(2))
            year = int(match.group(3))
            month = months.get(month_name.title(), 1)
            return datetime(year, month, day)
        except:
            pass
    
    # 2. "Friday, 16 January 2026" 형식
    pattern2 = r'(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s*(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})'
    match = re.search(pattern2, date_str, re.IGNORECASE)
    if match:
        try:
            day = int(match.group(1))
            month_name = match.group(2)
            year = int(match.group(3))
            month = months.get(month_name.title(), 1)
            return datetime(year, month, day)
        except:
            pass
    
    # 3. "16 January 2026" 형식 (요일 없음)
    pattern3 = r'(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})'
    match = re.search(pattern3, date_str, re.IGNORECASE)
    if match:
        try:
            day = int(match.group(1))
            month_name = match.group(2)
            year = int(match.group(3))
            month = months.get(month_name.title(), 1)
            return datetime(year, month, day)
        except:
            pass
    
    return None

def parse_time(time_str):
    """시간 문자열을 시간과 분으로 파싱 (기본값: 12:00)"""
    # "10:32 PM", "1:17 PM" 등의 패턴
    time_match = re.search(r'(\d{1,2}):(\d{2})\s*(AM|PM)', time_str, re.IGNORECASE)
    if time_match:
        hour = int(time_match.group(1))
        minute = int(time_match.group(2))
        am_pm = time_match.group(3).upper()
        if am_pm == 'PM' and hour != 12:
            hour += 12
        elif am_pm == 'AM' and hour == 12:
            hour = 0
        return hour, minute
    return 12, 0

def extract_emails_from_email222(content):
    """EMAI3333.MD에서 개별 이메일 추출"""
    emails = []
    
    # "From:" 으로 시작하는 블록들을 찾기
    email_blocks = re.split(r'(?=^From:)', content, flags=re.MULTILINE)
    
    for block in email_blocks:
        block = block.strip()
        if block and block.startswith('From:'):
            emails.append(block)
    
    return emails

def extract_email_metadata(email_text):
    """이메일에서 메타데이터 추출"""
    metadata = {
        'from': '',
        'to': '',
        'cc': '',
        'subject': '',
        'sent_date': '',
        'date': None,
        'time': (12, 0),
        'body_start': 0
    }
    
    lines = email_text.split('\n')
    body_start_idx = 0
    found_subject = False
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        
        if line_stripped.startswith('From:'):
            metadata['from'] = line.replace('From:', '').strip()
        elif line_stripped.startswith('To:'):
            metadata['to'] = line.replace('To:', '').strip()
        elif line_stripped.startswith('Cc:'):
            metadata['cc'] = line.replace('Cc:', '').strip()
        elif line_stripped.startswith('Subject:'):
            metadata['subject'] = line.replace('Subject:', '').strip()
            found_subject = True
        elif line_stripped.startswith('Sent:'):
            sent_line = line.replace('Sent:', '').strip()
            metadata['sent_date'] = sent_line
            
            # 날짜 파싱
            date_obj = parse_date(sent_line)
            if date_obj:
                metadata['date'] = date_obj
            
            # 시간 파싱
            time_obj = parse_time(sent_line)
            if time_obj:
                metadata['time'] = time_obj
        
        # 본문 시작점 찾기 (Subject 이후 첫 빈 줄이 아닌 내용)
        if found_subject and i > 5:
            if line_stripped and not line_stripped.startswith(('From:', 'To:', 'Cc:', 'Subject:', 'Sent:')):
                if body_start_idx == 0 and not line_stripped.startswith('ALERT:'):
                    # 실제 본문 시작
                    body_start_idx = i
    
    # 본문 시작점이 없으면 Subject 다음 빈 줄 이후로 설정
    if body_start_idx == 0:
        for i in range(len(lines) - 1, -1, -1):
            if 'Subject:' in lines[i]:
                body_start_idx = i + 3  # Subject, 빈 줄, 다음 줄
                break
    
    metadata['body_start'] = body_start_idx
    return metadata

def format_email_for_agi(email_text, msg_num, metadata):
    """이메일을 AGI 파일 형식으로 변환"""
    lines = email_text.split('\n')
    
    # 본문 추출
    if metadata['body_start'] > 0 and metadata['body_start'] < len(lines):
        body_lines = lines[metadata['body_start']:]
    else:
        body_lines = lines[10:]  # 기본값
    
    # ALERT, CONFIDENTIALITY NOTICE 등 제거
    body_text = '\n'.join(body_lines)
    body_text = re.sub(r'ALERT:.*?attachment\.', '', body_text, flags=re.DOTALL | re.IGNORECASE)
    body_text = re.sub(r'CONFIDENTIALITY NOTICE.*?www\.mammoet\.com', '', body_text, flags=re.DOTALL | re.IGNORECASE)
    body_text = re.sub(r'CONFIDENTIALITY NOTICE.*?www\.ofco-int\.com', '', body_text, flags=re.DOTALL | re.IGNORECASE)
    body_text = re.sub(r'DISCLAIMER::.*?Thank you\.', '', body_text, flags=re.DOTALL | re.IGNORECASE)
    
    # 여러 개의 빈 줄을 하나로
    body_text = re.sub(r'\n{3,}', '\n\n', body_text)
    body_text = body_text.strip()
    
    # 날짜 포맷팅 (AGI 형식: 2025-12-30 14:00:00 +04:00)
    if metadata['date']:
        hour, minute = metadata['time']
        date_str = metadata['date'].strftime(f'%Y-%m-%d {hour:02d}:{minute:02d}:00 +04:00')
    else:
        date_str = '2026-01-16 12:00:00 +04:00'
    
    # From에서 이름 추출
    from_line = metadata['from']
    sender_name = from_line.split('<')[0].strip() if '<' in from_line else from_line.split()[0] if from_line else 'Unknown'
    
    # To 줄 길이 제한
    to_display = metadata['to'][:100] + '...' if len(metadata['to']) > 100 else metadata['to']
    
    formatted = f"""#### Msg {msg_num} — {sender_name} @ {date_str} {{#msg-{msg_num}}}
| Key | Value |
|---|---|
| From | {metadata['from']} |
| To | {to_display} |
| Cc | Multiple |
| Subject | {metadata['subject']} |

```text
{body_text}
```
"""
    
    return formatted

def main():
    # 현재 스크립트 위치 기준으로 경로 설정
    script_path = Path(__file__).parent
    base_path = script_path
    email3333_path = base_path / 'EMAI3333.MD'
    agi_path = base_path / 'AGI Transformers Transportation_email_20251229.md'
    
    print("=" * 60)
    print("EMAI3333.MD를 AGI 파일에 업데이트하는 중...")
    print("=" * 60)
    
    # EMAI3333.MD 읽기
    print("\n1. EMAI3333.MD 읽는 중...")
    try:
        with open(email3333_path, 'r', encoding='utf-8') as f:
            email3333_content = f.read()
        print(f"   [OK] 파일 읽기 완료 ({len(email3333_content)} 문자)")
    except Exception as e:
        print(f"   [ERROR] 오류: {e}")
        return
    
    # 이메일 추출
    print("\n2. 이메일 추출 중...")
    emails = extract_emails_from_email222(email3333_content)
    print(f"   [OK] 총 {len(emails)}개의 이메일 발견")
    
    # 이메일을 날짜별로 정렬
    print("\n3. 날짜별 정렬 중...")
    emails_with_dates = []
    for i, email in enumerate(emails):
        metadata = extract_email_metadata(email)
        if metadata['date']:
            emails_with_dates.append((metadata['date'], metadata['time'], email, metadata))
        else:
            # 날짜를 찾을 수 없는 경우 기본값 사용
            print(f"   [WARN] 경고: 이메일 {i+1}의 날짜를 찾을 수 없음")
            emails_with_dates.append((datetime(2026, 1, 16), (12, 0), email, metadata))
    
    # 날짜 및 시간순 정렬 (오래된 것부터)
    emails_with_dates.sort(key=lambda x: (x[0], x[1]))
    
    print(f"   [OK] {len(emails_with_dates)}개 이메일 정렬 완료")
    print(f"   - 첫 번째: {emails_with_dates[0][0].strftime('%Y-%m-%d')}")
    print(f"   - 마지막: {emails_with_dates[-1][0].strftime('%Y-%m-%d')}")
    
    # AGI 파일 읽기
    print("\n4. AGI 파일 읽는 중...")
    try:
        with open(agi_path, 'r', encoding='utf-8') as f:
            agi_content = f.read()
        print(f"   [OK] 파일 읽기 완료 ({len(agi_content)} 문자)")
    except Exception as e:
        print(f"   [ERROR] 오류: {e}")
        return
    
    # 기존 메시지 번호 찾기
    msg_num_matches = list(re.finditer(r'#### Msg (\d+)', agi_content))
    if msg_num_matches:
        last_msg_num = int(msg_num_matches[-1].group(1))
        next_msg_num = last_msg_num + 1
        print(f"   [OK] 마지막 메시지 번호: {last_msg_num}, 다음 번호: {next_msg_num}")
    else:
        next_msg_num = 78
        print(f"   [WARN] 메시지 번호를 찾을 수 없음, 시작 번호: {next_msg_num}")
    
    # 새 이메일들을 AGI 파일 형식으로 변환
    print("\n5. 이메일 형식 변환 중...")
    new_emails_text = []
    for i, (date, time, email, metadata) in enumerate(emails_with_dates):
        msg_num = next_msg_num + i
        formatted = format_email_for_agi(email, msg_num, metadata)
        new_emails_text.append(formatted)
    
    new_emails_combined = '\n\n'.join(new_emails_text)
    print(f"   [OK] {len(new_emails_text)}개 이메일 변환 완료")
    
    # AGI 파일의 끝 부분에 추가
    print("\n6. AGI 파일에 이메일 추가 중...")
    
    # 마지막 ``` 이후에 추가
    if agi_content.rstrip().endswith('```'):
        # 마지막 ``` 전에 추가
        last_marker = agi_content.rfind('```')
        if last_marker > 0:
            # 마지막 ``` 앞의 내용 확인
            before_marker = agi_content[:last_marker].rstrip()
            agi_content_updated = before_marker + '\n\n' + new_emails_combined + '\n\n```'
        else:
            agi_content_updated = agi_content + '\n\n' + new_emails_combined
    else:
        agi_content_updated = agi_content + '\n\n' + new_emails_combined
    
    # 파일 저장
    print("\n7. 업데이트된 파일 저장 중...")
    output_path = agi_path  # 원본 파일 업데이트
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(agi_content_updated)
        print(f"   [OK] 파일 저장 완료: {output_path}")
    except Exception as e:
        print(f"   [ERROR] 오류: {e}")
        return
    
    print("\n" + "=" * 60)
    print("업데이트 완료!")
    print("=" * 60)
    print(f"추가된 이메일: {len(emails_with_dates)}개")
    print(f"메시지 번호: {next_msg_num} ~ {next_msg_num + len(emails_with_dates) - 1}")
    print(f"날짜 범위: {emails_with_dates[0][0].strftime('%Y-%m-%d')} ~ {emails_with_dates[-1][0].strftime('%Y-%m-%d')}")
    print("=" * 60)

if __name__ == '__main__':
    main()
