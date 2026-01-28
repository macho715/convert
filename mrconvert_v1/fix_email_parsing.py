#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
이메일 파싱 수정 - 원본 본문 그대로 추출
"""

import json
import re
from pathlib import Path

def fix_email_thread():
    """이메일 스레드 수정 - 원본 본문 추가"""
    email_file = Path('mrconvert_v1/AGI Transformers Transportation_email_20251229.md')
    with open(email_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # aaaw.md 읽기
    aaaw_file = Path('mrconvert_v1/aaaw.md')
    with open(aaaw_file, 'r', encoding='utf-8') as f:
        aaaw_content = f.read()
    
    lines = aaaw_content.split('\n')
    
    # 이메일들을 역순으로 파싱 (최신부터)
    emails = []
    current_email = {}
    i = 0
    
    # 첫 번째 이메일 (Yulia Frolova) - Subject로 시작, From/To 없음
    if lines[0].startswith('Subject:'):
        subject = lines[0].replace('Subject:', '').strip()
        body_lines = []
        i = 2  # Subject 다음 빈 줄 건너뛰기
        while i < len(lines) and not lines[i].startswith('Public'):
            body_lines.append(lines[i])
            i += 1
        
        emails.append({
            'subject': subject,
            'from': 'Yulia Frolova <Yulia.Frolova@mammoet.com>',
            'from_name': 'Yulia Frolova',
            'from_email': 'Yulia.Frolova@mammoet.com',
            'to': ['Agency | OFCO', 'minkyu.cha@samsung.com'],
            'sent': 'December 30, 2025',
            'body': '\n'.join(body_lines).strip(),
            'order': 78  # 가장 최신
        })
        i += 1  # 'Public' 건너뛰기
    
    # 나머지 이메일들 파싱
    while i < len(lines):
        line = lines[i]
        
        if line.startswith('From:'):
            if current_email:
                emails.append(current_email)
            
            from_match = re.match(r'From: (.+?) <(.+?)>', line)
            if from_match:
                current_email = {
                    'from_name': from_match.group(1).strip(),
                    'from_email': from_match.group(2).strip(),
                    'from': f"{from_match.group(1).strip()} <{from_match.group(2).strip()}>",
                    'body': '',
                    'to': [],
                    'sent': ''
                }
        
        elif line.startswith('Sent:') and current_email:
            current_email['sent'] = line.replace('Sent:', '').strip()
        
        elif line.startswith('To:') and current_email:
            to_line = line.replace('To:', '').strip()
            # 다음 줄도 확인
            j = i + 1
            while j < len(lines) and not lines[j].startswith(('Cc:', 'Subject:', 'From:')):
                if lines[j].strip() and not lines[j].startswith('Subject:'):
                    to_line += ' ' + lines[j].strip()
                j += 1
            # 이메일 주소 추출
            to_emails = re.findall(r'<([^>]+)>', to_line)
            to_names = re.findall(r"'?([^<']+?)'?\s*<", to_line)
            current_email['to'] = to_emails if to_emails else [to_line.split(';')[0].strip()]
        
        elif line.startswith('Subject:') and current_email:
            current_email['subject'] = line.replace('Subject:', '').strip()
            # 본문 시작
            body_lines = []
            j = i + 1
            while j < len(lines) and not lines[j].startswith(('From:', 'Public', '---')):
                if lines[j].strip():
                    body_lines.append(lines[j])
                j += 1
            current_email['body'] = '\n'.join(body_lines).strip()
            i = j - 1
        
        i += 1
    
    if current_email:
        emails.append(current_email)
    
    # 날짜 매핑
    date_map = {
        'Wednesday, December 24, 2025 1:11 PM': '2025-12-24T13:11:00+04:00',
        'Wednesday, December 24, 2025 11:35 AM': '2025-12-24T11:35:00+04:00',
        'Wednesday, December 24, 2025 11:33 AM': '2025-12-24T11:33:00+04:00',
        'Tuesday, December 23, 2025 8:54 PM': '2025-12-23T20:54:00+04:00',
        '23 December 2025 17:56': '2025-12-23T17:56:00+04:00',
        'December 30, 2025': '2025-12-30T14:00:00+04:00'
    }
    
    # JSON 부분 찾기
    json_match = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
    if not json_match:
        print("JSON not found")
        return
    
    json_str = json_match.group(1)
    data = json.loads(json_str)
    
    # 기존 msg-73~77 제거 (잘못된 것들)
    data['messages'] = [m for m in data['messages'] if int(m['id'].split('-')[1]) < 73]
    
    # 새 메시지 추가 (올바른 순서로)
    # 1. Nanda Kumar (2025-12-23 20:54)
    # 2. OFCO urgent (2025-12-24 11:33)
    # 3. OFCO Teams link (2025-12-24 11:35)
    # 4. OFCO Yoonus MZ GC Ops (2025-12-24 13:11)
    # 5. 차민규 SSOT (2025-12-23 17:56) - 이미 msg-67에 있음
    # 6. Yulia Frolova AD Ports docs (2025-12-30)
    
    new_messages = []
    msg_order = 73
    
    # 이메일들을 날짜순으로 정렬
    sorted_emails = sorted(emails, key=lambda x: (
        '2025-12-23' if 'December 23' in x.get('sent', '') else
        '2025-12-24' if 'December 24' in x.get('sent', '') else
        '2025-12-30' if 'December 30' in x.get('sent', '') else '2025-12-31'
    ))
    
    for email in sorted_emails:
        # 날짜 파싱
        iso_date = None
        for sent_key, date_val in date_map.items():
            if sent_key in email.get('sent', ''):
                iso_date = date_val
                break
        
        if not iso_date:
            if 'Yulia Frolova' in email.get('from', ''):
                iso_date = '2025-12-30T14:00:00+04:00'
            elif 'December 23' in email.get('sent', ''):
                iso_date = '2025-12-23T20:54:00+04:00'
            elif 'December 24' in email.get('sent', ''):
                iso_date = '2025-12-24T13:11:00+04:00'
            else:
                iso_date = '2025-12-24T13:11:00+04:00'
        
        # Summary 생성
        body = email.get('body', '')
        if 'Pre-Arrival Cargo Declaration' in body:
            summary = "OFCO requests Pre-Arrival Cargo Declaration completion and shares MZ Supervisor Teams meeting link for pre-operational discussion."
        elif 'MZ GC Ops' in body or 'Ahmed Qasem' in body or 'احمد الخضر' in body:
            summary = "OFCO forwards MZ GC Ops meeting summary: ETA 06th Jan AM, 4 cargo (2 x 217t transformers), SPMT load from Yard 5, max GBP 5t, ramp certificate required."
        elif 'teams.microsoft.com' in body:
            if 'Please attend' in body or 'ASAP' in body:
                summary = "OFCO requests urgent attendance at Teams meeting."
            else:
                url_match = re.search(r'https?://[^\s]+', body)
                url = url_match.group(0) if url_match else 'Teams meeting link'
                summary = f"OFCO shares Teams meeting link: {url}."
        elif 'AD Ports Method Statement' in body or 'Welding Machine Calibration' in body:
            summary = "Mammoet submits AD Ports documents: Method Statement Form, Permit to work Overwater form, Hot works Permit form, RA for AGI TR, Welding Machine Calibration certificates, Mammoet Welder Performance qualification reports. Confirms marine crew has 15+ years experience with valid certifications."
        else:
            summary = body[:200].replace('\n', ' ') + "..." if len(body) > 200 else body
        
        json_msg = {
            "id": f"msg-{msg_order}",
            "order": msg_order,
            "isoDate": iso_date,
            "from": email.get('from', 'Unknown'),
            "to": email.get('to', []),
            "cc": ["Multiple"],
            "subject": email.get('subject', ''),
            "summary": summary,
            "body": body,  # 원본 본문 추가
            "snippetRef": f"#msg-{msg_order}"
        }
        
        if msg_order > 73:
            json_msg["inReplyTo"] = f"#msg-{msg_order-1}"
        
        if 'URGENT' in body.upper() or 'ASAP' in body.upper() or 'CRITICAL' in body.upper():
            json_msg["importance"] = "High"
        
        new_messages.append(json_msg)
        msg_order += 1
    
    # 메시지 추가
    data['messages'].extend(new_messages)
    
    # JSON 업데이트
    updated_json = json.dumps(data, indent=2, ensure_ascii=False)
    updated_content = content.replace(json_match.group(0), f'```json\n{updated_json}\n```')
    
    # Thread 섹션 업데이트
    thread_marker = "#### Msg 72 — Sayeed Ahmed @ 2025-12-30 10:07 +04:00 {#msg-72}"
    if thread_marker in updated_content:
        # 기존 msg-73 이후 부분 제거
        msg73_pos = updated_content.find("#### Msg 73")
        if msg73_pos > 0:
            # 파일 끝까지 또는 다음 섹션까지
            end_pos = updated_content.find("\n\n---\n", msg73_pos)
            if end_pos == -1:
                end_pos = len(updated_content)
            
            # 새 Thread 섹션 생성
            thread_sections = []
            for msg in new_messages:
                from_name = msg['from'].split('<')[0].strip()
                from_email = msg['from'].split('<')[1].replace('>', '').strip() if '<' in msg['from'] else ''
                date_display = msg['isoDate'].replace('T', ' ').replace('+04:00', ' +04:00')
                
                thread_section = f"\n#### Msg {msg['order']} — {from_name} @ {date_display} {{#msg-{msg['order']}}}\n"
                thread_section += "| Key | Value |\n|---|---|\n"
                thread_section += f"| From | {from_name} <{from_email}> |\n"
                thread_section += f"| To | {', '.join(msg.get('to', [])[:3])}{'...' if len(msg.get('to', [])) > 3 else ''} |\n"
                thread_section += f"| Cc | Multiple |\n"
                thread_section += f"| Subject | {msg['subject']} |\n\n"
                thread_section += "```text\n"
                thread_section += msg.get('body', '')
                thread_section += "\n```\n"
                
                thread_sections.append(thread_section)
            
            updated_content = (
                updated_content[:msg73_pos] +
                '\n'.join(thread_sections) +
                updated_content[end_pos:]
            )
    
    # 저장
    with open(email_file, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    
    print(f"Email thread fixed successfully!")
    print(f"   - Total messages: {len(data['messages'])}")

if __name__ == "__main__":
    fix_email_thread()

