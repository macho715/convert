"""
파일 시스템 스캐너 - 폴더 순회 및 파일 필터링
"""
from __future__ import annotations
from pathlib import Path
from typing import List, Iterator
from ..core.config import EMAIL_ROOT, ALLOWED_EXT, MAX_FILES
from ..core.errors import ScanError


def scan_folder(root: Path = None) -> List[Path]:
    """
    폴더를 스캔하여 허용된 확장자의 파일 목록 반환
    
    Args:
        root: 스캔할 루트 폴더 (기본값: EMAIL_ROOT)
        
    Returns:
        List[Path]: 스캔된 파일 경로 목록
        
    Raises:
        ScanError: 스캔 실패 시
    """
    if root is None:
        root = EMAIL_ROOT
    
    try:
        files = []
        count = 0
        
        for file_path in _walk_files(root):
            if _is_allowed_file(file_path):
                files.append(file_path)
                count += 1
                
                # 샘플 제한이 있는 경우
                if MAX_FILES and count >= MAX_FILES:
                    break
                    
        return files
        
    except OSError as e:
        raise ScanError(f"폴더 스캔 실패: {root}") from e


def _walk_files(root: Path) -> Iterator[Path]:
    """폴더를 재귀적으로 순회하여 파일 반환"""
    try:
        for item in root.iterdir():
            if item.is_file():
                yield item
            elif item.is_dir():
                # Outlook 파일 제외
                if not _is_outlook_file(item):
                    yield from _walk_files(item)
    except PermissionError:
        # 권한이 없는 폴더는 스킵
        pass


def _is_allowed_file(file_path: Path) -> bool:
    """파일이 허용된 확장자인지 확인"""
    return file_path.suffix.lower() in ALLOWED_EXT


def _is_outlook_file(file_path: Path) -> bool:
    """Outlook 파일인지 확인 (.pst, .ost, .msg 제외)"""
    outlook_extensions = {'.pst', '.ost', '.msg'}
    return file_path.suffix.lower() in outlook_extensions
