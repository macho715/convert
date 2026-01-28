#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MHT 파일에서 HTML 내용을 추출하여 마크다운으로 변환"""

import email
from email import policy
import html2text
from pathlib import Path

def extract_mht_content(mht_path):
    """MHT 파일에서 HTML 내용 추출"""
    with open(mht_path, 'rb') as f:
        msg = email.message_from_bytes(f.read(), policy=policy.default)
    
    # HTML 본문 찾기
    html_content = None
    for part in msg.walk():
        if part.get_content_type() == 'text/html':
            html_content = part.get_payload(decode=True).decode('utf-8')
            break
    
    # HTML을 마크다운으로 변환
    if html_content:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.body_width = 0  # 줄바꿈 없이
        return h.handle(html_content)
    return None

if __name__ == "__main__":
    mht_path = Path(__file__).parent / "Logistics Document Guardian - GPTS 개발 가이드 (4).mht"
    output_path = Path(__file__).parent / "Logistics Document Guardian - GPTS 개발 가이드 (4).md"
    
    print(f"Extracting content from: {mht_path}")
    content = extract_mht_content(mht_path)
    
    if content:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"[OK] Extracted to: {output_path}")
        print(f"Content length: {len(content)} characters")
    else:
        print("❌ Failed to extract content")

