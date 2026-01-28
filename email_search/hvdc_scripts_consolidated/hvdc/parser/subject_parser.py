"""
이메일 제목 파서 - 순수 함수로 구현
"""
from __future__ import annotations
from typing import Dict, List
from ..core.typing import ParseResult
from ..extractors.case import extract_cases
from ..extractors.site import extract_sites
from ..extractors.lpo import extract_lpos
from ..extractors.phase import extract_phases


def parse_subject(subject: str) -> ParseResult:
    """
    이메일 제목에서 메타데이터 추출
    
    Args:
        subject: 이메일 제목
        
    Returns:
        ParseResult: 추출된 케이스, 사이트, LPO, 페이즈 정보
    """
    # 케이스 추출
    cases = extract_cases(subject)
    
    # 사이트 추출
    sites = extract_sites(subject)
    
    # LPO 추출
    lpos = extract_lpos(subject)
    
    # 페이즈 추출
    phases = extract_phases(subject)
    
    return {
        "cases": cases,
        "sites": sites,
        "lpos": lpos,
        "phases": phases
    }


def parse_folder_title(folder_title: str) -> ParseResult:
    """
    폴더 제목에서 메타데이터 추출
    
    Args:
        folder_title: 폴더 제목
        
    Returns:
        ParseResult: 추출된 케이스, 사이트, LPO, 페이즈 정보
    """
    return parse_subject(folder_title)  # 동일한 로직 사용
