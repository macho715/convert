"""
이메일 파일 리더 - .eml/.txt/.html 파일 읽기
"""
from __future__ import annotations
from pathlib import Path
from typing import Dict, Optional
from ..core.io import read_text
from ..core.config import ENCODING_FALLBACKS
from ..core.errors import IoError


def read_email_file(file_path: Path) -> Dict[str, str]:
    """
    이메일 파일을 읽어서 헤더와 본문 추출
    
    Args:
        file_path: 이메일 파일 경로
        
    Returns:
        Dict[str, str]: {'subject': str, 'body': str, 'sender': str, 'date': str}
        
    Raises:
        IoError: 파일 읽기 실패 시
    """
    try:
        content = read_text(file_path, ENCODING_FALLBACKS)
        
        # 간단한 파싱 (실제로는 email.parser 사용 권장)
        lines = content.split('\n')
        
        subject = ""
        sender = ""
        date = ""
        body_start = 0
        
        for i, line in enumerate(lines):
            line = line.strip()
            if line.startswith('Subject:'):
                subject = line[8:].strip()
            elif line.startswith('From:'):
                sender = line[5:].strip()
            elif line.startswith('Date:'):
                date = line[5:].strip()
            elif line == '' and subject:  # 헤더 끝
                body_start = i + 1
                break
        
        # 본문 추출
        body = '\n'.join(lines[body_start:]).strip()
        
        return {
            'subject': subject,
            'body': body,
            'sender': sender,
            'date': date
        }
        
    except Exception as e:
        raise IoError(f"이메일 파일 읽기 실패: {file_path}") from e


def read_text_file(file_path: Path) -> str:
    """
    텍스트 파일 읽기
    
    Args:
        file_path: 텍스트 파일 경로
        
    Returns:
        str: 파일 내용
        
    Raises:
        IoError: 파일 읽기 실패 시
    """
    try:
        return read_text(file_path, ENCODING_FALLBACKS)
    except Exception as e:
        raise IoError(f"텍스트 파일 읽기 실패: {file_path}") from e
