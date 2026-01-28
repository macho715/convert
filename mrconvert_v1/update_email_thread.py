#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
이메일 스레드 업데이트 스크립트
aaaw.md의 원본 이메일을 AGI Transformers Transportation_email_20251229.md에 추가
"""

import json
import re
from pathlib import Path
from datetime import datetime

def parse_date_time(date_str, time_str):
    """날짜/시간 문자열을 ISO 형식으로 변환"""
    # "Wednesday, December 24, 2025 1:11 PM" 형식
    try:
        # 간단한 파싱
        if "December 24, 2025" in date_str:
            if "1:11 PM" in time_str or "13:11" in time_str:
                return "2025-12-24T13:11:00+04:00"
            elif "11:35 AM" in time_str or "11:35" in time_str:
                return "2025-12-24T11:35:00+04:00"
            elif "11:33 AM" in time_str or "11:33" in time_str:
                return "2025-12-24T11:33:00+04:00"
        elif "December 23, 2025" in date_str:
            if "8:54 PM" in time_str or "20:54" in time_str:
                return "2025-12-23T20:54:00+04:00"
            elif "17:56" in time_str or "5:56 PM" in time_str:
                return "2025-12-23T17:56:00+04:00"
        elif "December 30" in date_str:
            # Yulia의 최신 이메일 (시간 미상, 추정 오후)
            return "2025-12-30T14:00:00+04:00"
    except:
        pass
    return None

def extract_email_body(lines, start_idx):
    """이메일 본문 추출"""
    body_lines = []
    i = start_idx
    while i < len(lines):
        line = lines[i].strip()
        # 다음 이메일 시작 (From: 또는 Public)이면 중단
        if line.startswith('From:') or line == 'Public':
            break
        if line:
            body_lines.append(lines[i])
        i += 1
    return '\n'.join(body_lines).strip(), i

def update_email_thread():
    """이메일 스레드 업데이트"""
    # 기존 파일 읽기
    email_file = Path('mrconvert_v1/AGI Transformers Transportation_email_20251229.md')
    with open(email_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # JSON 부분 추출
    json_match = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
    if not json_match:
        print("❌ JSON 부분을 찾을 수 없습니다.")
        return
    
    json_str = json_match.group(1)
    data = json.loads(json_str)
    
    # aaaw.md 읽기
    aaaw_file = Path('mrconvert_v1/aaaw.md')
    with open(aaaw_file, 'r', encoding='utf-8') as f:
        aaaw_content = f.read()
    
    lines = aaaw_content.split('\n')
    
    # 새 메시지들
    new_messages = []
    current_msg = None
    i = 0
    
    # 이메일 파싱
    while i < len(lines):
        line = lines[i]
        
        # Subject로 새 이메일 시작
        if line.startswith('Subject:'):
            if current_msg:
                new_messages.append(current_msg)
            
            subject = line.replace('Subject:', '').strip()
            current_msg = {
                'subject': subject,
                'body': '',
                'from': '',
                'to': [],
                'cc': [],
                'sent': '',
                'order': len(data['messages']) + len(new_messages) + 1
            }
        
        # From 라인
        elif line.startswith('From:') and current_msg:
            from_match = re.match(r'From: (.+?) <(.+?)>', line)
            if from_match:
                current_msg['from_name'] = from_match.group(1).strip()
                current_msg['from_email'] = from_match.group(2).strip()
                current_msg['from'] = f"{current_msg['from_name']} <{current_msg['from_email']}>"
        
        # Sent 라인
        elif line.startswith('Sent:') and current_msg:
            sent_line = line
            # 다음 줄도 확인
            if i + 1 < len(lines):
                sent_line += ' ' + lines[i + 1]
            current_msg['sent'] = sent_line
        
        # To 라인
        elif line.startswith('To:') and current_msg:
            to_line = line.replace('To:', '').strip()
            # 여러 줄에 걸칠 수 있음
            j = i + 1
            while j < len(lines) and not lines[j].startswith(('Cc:', 'Subject:', 'From:')):
                if lines[j].strip():
                    to_line += ' ' + lines[j].strip()
                j += 1
            # 이메일 주소 추출
            to_emails = re.findall(r'<([^>]+)>', to_line)
            to_names = re.findall(r'([^<]+?)\s*<', to_line)
            current_msg['to'] = to_emails if to_emails else [to_line]
        
        # 본문 시작 (Subject 다음 빈 줄 후)
        elif current_msg and line.strip() and not line.startswith(('From:', 'Sent:', 'To:', 'Cc:', 'Subject:', 'Public', '---')):
            if not current_msg['body'] or current_msg['body'].startswith('Subject:'):
                # 본문 추출
                body, next_idx = extract_email_body(lines, i)
                current_msg['body'] = body
                i = next_idx - 1
        
        i += 1
    
    if current_msg:
        new_messages.append(current_msg)
    
    # 메시지 정리 및 JSON 형식으로 변환
    json_messages = []
    thread_sections = []
    
    for idx, msg in enumerate(new_messages, start=73):
        # 날짜/시간 파싱
        iso_date = None
        if msg.get('sent'):
            if 'December 24, 2025' in msg['sent']:
                if '1:11 PM' in msg['sent'] or '13:11' in msg['sent']:
                    iso_date = "2025-12-24T13:11:00+04:00"
                elif '11:35 AM' in msg['sent']:
                    iso_date = "2025-12-24T11:35:00+04:00"
                elif '11:33 AM' in msg['sent']:
                    iso_date = "2025-12-24T11:33:00+04:00"
            elif 'December 23, 2025' in msg['sent']:
                if '8:54 PM' in msg['sent'] or '20:54' in msg['sent']:
                    iso_date = "2025-12-23T20:54:00+04:00"
                elif '17:56' in msg['sent'] or '5:56 PM' in msg['sent']:
                    iso_date = "2025-12-23T17:56:00+04:00"
        
        # 기본값 설정
        if not iso_date:
            # 첫 번째 이메일 (Yulia, 날짜 미상)
            if 'Yulia Frolova' in msg.get('from', ''):
                iso_date = "2025-12-30T14:00:00+04:00"
            else:
                iso_date = "2025-12-24T13:11:00+04:00"
        
        # Summary 생성
        body_preview = msg.get('body', '')[:200].replace('\n', ' ')
        if 'Pre-Arrival Cargo Declaration' in msg.get('body', ''):
            summary = "OFCO requests Pre-Arrival Cargo Declaration completion and shares MZ Supervisor Teams meeting link for pre-operational discussion."
        elif 'MZ GC Ops' in msg.get('body', '') or 'Ahmed Qasem' in msg.get('body', ''):
            summary = "OFCO forwards MZ GC Ops meeting summary: ETA 06th Jan AM, 4 cargo (2 x 217t transformers), SPMT load from Yard 5, max GBP 5t, ramp certificate required, berth booking pending, PTW/RA/Method statement required."
        elif 'teams.microsoft.com' in msg.get('body', '') or 'Meeting ID' in msg.get('body', ''):
            if 'Please attend' in msg.get('body', ''):
                summary = "OFCO requests urgent attendance at Teams meeting."
            else:
                url_match = re.search(r'https?://[^\s]+', msg.get('body', ''))
                url = url_match.group(0) if url_match else 'Teams meeting link shared'
                summary = f"OFCO shares Teams meeting link: {url}."
        elif 'AD Ports Method Statement' in msg.get('body', '') or 'Welding Machine Calibration' in msg.get('body', ''):
            summary = "Mammoet submits AD Ports documents: Method Statement Form, Permit to work Overwater form, Hot works Permit form, RA for AGI TR (Hot works & RoRo Ops), Welding Machine Calibration certificates, Mammoet Welder Performance qualification reports. Confirms marine crew has 15+ years experience with valid certifications."
        else:
            summary = body_preview[:150] + "..." if len(body_preview) > 150 else body_preview
        
        # JSON 메시지 생성
        json_msg = {
            "id": f"msg-{idx}",
            "order": idx,
            "isoDate": iso_date,
            "from": msg.get('from', 'Unknown'),
            "to": msg.get('to', []),
            "cc": ["Multiple"],
            "subject": msg.get('subject', ''),
            "summary": summary,
            "snippetRef": f"#msg-{idx}"
        }
        
        # inReplyTo 설정
        if idx > 73:
            json_msg["inReplyTo"] = f"#msg-{idx-1}"
        
        # importance 설정
        if 'URGENT' in msg.get('body', '').upper() or 'ASAP' in msg.get('body', '').upper():
            json_msg["importance"] = "High"
        
        json_messages.append(json_msg)
        
        # Thread 섹션 생성
        from_name = msg.get('from_name', 'Unknown')
        from_email = msg.get('from_email', '')
        date_display = iso_date.replace('T', ' ').replace('+04:00', ' +04:00')
        
        thread_section = f"\n#### Msg {idx} — {from_name} @ {date_display} {{#msg-{idx}}}\n"
        thread_section += "| Key | Value |\n|---|---|\n"
        thread_section += f"| From | {from_name} <{from_email}> |\n"
        thread_section += f"| To | {', '.join(msg.get('to', [])[:3])}{'...' if len(msg.get('to', [])) > 3 else ''} |\n"
        thread_section += f"| Cc | Multiple |\n"
        thread_section += f"| Subject | {msg.get('subject', '')} |\n\n"
        thread_section += "```text\n"
        thread_section += msg.get('body', '')
        thread_section += "\n```\n"
        
        thread_sections.append(thread_section)
    
    # JSON 업데이트
    data['messages'].extend(json_messages)
    
    # dateRange 업데이트
    data['dateRange']['end'] = "2025-12-30"
    
    # topics 추가
    new_topics = [
        "pre-arrival-cargo-declaration",
        "mz-gc-ops-meeting",
        "teams-meeting-link",
        "ad-ports-documents-submission",
        "crew-qualifications",
        "welding-certifications",
        "ptw-permit-submission",
        "ramp-certificate-requirement",
        "berth-booking"
    ]
    for topic in new_topics:
        if topic not in data['topics']:
            data['topics'].append(topic)
    
    # actions 추가
    new_actions = [
        {
            "owner": "Samsung (minkyu.cha)",
            "action": "Complete and return Pre-Arrival Cargo Declaration (signed and stamped)",
            "relatedMsg": "#msg-73",
            "status": "pending"
        },
        {
            "owner": "All Teams",
            "action": "Attend Teams meeting with MZ Supervisor for pre-operational discussion",
            "relatedMsg": "#msg-73",
            "status": "scheduled"
        },
        {
            "owner": "LCT Bushra",
            "action": "Submit valid ramp certificate (max GBP 5t on berth side)",
            "relatedMsg": "#msg-76",
            "status": "pending"
        },
        {
            "owner": "OFCO",
            "action": "Submit berth booking requirements 24 hours in advance",
            "relatedMsg": "#msg-76",
            "status": "pending"
        },
        {
            "owner": "Mammoet",
            "action": "Submit AD Ports documents (PTW forms, RA, Method Statement, Welding certificates) - COMPLETED",
            "relatedMsg": "#msg-77",
            "status": "completed"
        }
    ]
    data['actions'].extend(new_actions)
    
    # issues 추가
    new_issues = [
        {
            "type": "pre-arrival-documentation",
            "description": "Pre-Arrival Cargo Declaration required to be completed, signed, and stamped by Samsung",
            "relatedMsg": "#msg-73",
            "status": "pending"
        },
        {
            "type": "ramp-certificate-required",
            "description": "Valid ramp certificate required showing max GBP 5t on berth side",
            "relatedMsg": "#msg-76",
            "status": "pending"
        },
        {
            "type": "berth-booking-pending",
            "description": "Berth booking will be advised based on availability; agent to send requirements 24 hours in advance",
            "relatedMsg": "#msg-76",
            "status": "pending"
        }
    ]
    data['issues'].extend(new_issues)
    
    # participants 추가
    nanda_kumar = {
        "@type": "schema:Person",
        "name": "Nanda Kumar",
        "email": "agency@ofco-int.com",
        "org": "OFCO",
        "role": "Agency"
    }
    ahmed_qasem = {
        "@type": "schema:Person",
        "name": "Ahmed Qasem",
        "email": None,
        "org": "AD Ports",
        "role": "Superintendent - Zayed Port Operations"
    }
    
    if not any(p.get('name') == 'Nanda Kumar' for p in data['participants']):
        data['participants'].append(nanda_kumar)
    if not any(p.get('name') == 'Ahmed Qasem' for p in data['participants']):
        data['participants'].append(ahmed_qasem)
    
    # JSON 문자열 생성
    updated_json = json.dumps(data, indent=2, ensure_ascii=False)
    
    # 파일 내용 업데이트
    updated_content = content.replace(
        json_match.group(0),
        f'```json\n{updated_json}\n```'
    )
    
    # Thread 섹션에 새 메시지 추가
    thread_marker = "#### Msg 72 — Sayeed Ahmed @ 2025-12-30 10:07 +04:00 {#msg-72}"
    if thread_marker in updated_content:
        # msg-72 다음에 추가
        insert_pos = updated_content.find(thread_marker) + len(thread_marker)
        # 다음 메시지 섹션 찾기
        next_section = updated_content.find('\n\n', insert_pos)
        if next_section == -1:
            next_section = len(updated_content)
        
        # 새 Thread 섹션 삽입
        new_thread_content = '\n'.join(thread_sections)
        updated_content = (
            updated_content[:next_section] + 
            new_thread_content + 
            updated_content[next_section:]
        )
    else:
        # 파일 끝에 추가
        updated_content += '\n' + '\n'.join(thread_sections)
    
    # 저장
    with open(email_file, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    
    print(f"Email thread updated successfully!")
    print(f"   - Added messages: {len(json_messages)}")
    print(f"   - Total messages: {len(data['messages'])}")

if __name__ == "__main__":
    update_email_thread()

