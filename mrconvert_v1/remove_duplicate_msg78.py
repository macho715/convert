#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""중복된 msg-78 제거"""

import json
import re
from pathlib import Path

def remove_duplicate():
    email_file = Path('mrconvert_v1/AGI Transformers Transportation_email_20251229.md')
    with open(email_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # JSON 부분 찾기
    json_match = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
    if not json_match:
        print("JSON not found")
        return
    
    json_str = json_match.group(1)
    data = json.loads(json_str)
    
    # msg-78 제거 (중복 - 이미 msg-67에 있음)
    data['messages'] = [m for m in data['messages'] if m['id'] != 'msg-78']
    
    # JSON 업데이트
    updated_json = json.dumps(data, indent=2, ensure_ascii=False)
    updated_content = content.replace(json_match.group(0), f'```json\n{updated_json}\n```')
    
    # Thread 섹션에서 msg-78 제거
    msg78_thread_start = updated_content.find("#### Msg 78 —")
    if msg78_thread_start > 0:
        # 다음 섹션 또는 파일 끝까지 찾기
        msg78_thread_end = updated_content.find("\n\n#### Msg ", msg78_thread_start + 1)
        if msg78_thread_end == -1:
            msg78_thread_end = updated_content.find("\n\n---\n", msg78_thread_start)
        if msg78_thread_end == -1:
            msg78_thread_end = len(updated_content)
        
        updated_content = (
            updated_content[:msg78_thread_start] +
            updated_content[msg78_thread_end:]
        )
    
    # 저장
    with open(email_file, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    
    print(f"Removed duplicate msg-78")
    print(f"   - Total messages: {len(data['messages'])}")

if __name__ == "__main__":
    remove_duplicate()

