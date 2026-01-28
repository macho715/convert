#!/usr/bin/env python3
"""
HVDC SATUS.JSON 파일 수정 스크립트
- 이중 인용부호(`""`)를 단일 인용부호(`"`)로 변환
- 올바른 JSON 배열 구조로 변환
- 원본 파일 직접 수정
"""

import json
import re
from pathlib import Path


def fix_line_breaks(content: str) -> str:
    """
    줄바꿈으로 나뉜 값 수정 (이중 따옴표 변환 전)
    ""key"": ""\n"value", -> ""key"": ""value"",
    """
    # 정규식으로 직접 패턴 매칭 및 치환
    # 패턴: ""key"": ""\n"value", -> ""key"": ""value"",
    pattern = r'(""[^"]+"":\s*"")\s*\n\s*"([^"]+)",'
    replacement = r'\1\2",'
    content = re.sub(pattern, replacement, content, flags=re.MULTILINE)
    
    # 패턴: ""key"": ""\n"value"" -> ""key"": ""value""
    pattern2 = r'(""[^"]+"":\s*"")\s*\n\s*"([^"]+)""'
    replacement2 = r'\1\2""'
    content = re.sub(pattern2, replacement2, content, flags=re.MULTILINE)
    
    return content


def fix_line_breaks_after_normalization(content: str) -> str:
    """
    줄바꿈으로 나뉜 값 수정 (이중 따옴표 변환 후)
    "key": "\n"value", -> "key": "value",
    """
    # 더 강력한 패턴 매칭 - 여러 패턴 시도
    # 패턴 1: "key": "\n"value",
    content = re.sub(
        r'("[^"]+":\s*")\s*\n\s*"([^"]+)",',
        r'\1\2",',
        content,
        flags=re.MULTILINE
    )
    
    # 패턴 2: "key": "\n"value" (쉼표 없음)
    content = re.sub(
        r'("[^"]+":\s*")\s*\n\s*"([^"]+)"',
        r'\1\2"',
        content,
        flags=re.MULTILINE
    )
    
    # 패턴 3: 들여쓰기가 있는 경우
    content = re.sub(
        r'(\s+"[^"]+":\s*")\s*\n\s*"([^"]+)",',
        r'\1\2",',
        content,
        flags=re.MULTILINE
    )
    
    return content


def normalize_double_quotes(content: str) -> str:
    """
    이중 인용부호를 단일 인용부호로 변환
    ""key"": ""value"" -> "key": "value"
    """
    # 일반적인 이중 인용부호를 단일 인용부호로 변환
    content = content.replace('""', '"')
    return content


def ensure_array_structure(content: str) -> str:
    """
    JSON 배열 구조 확인 및 수정
    - 첫 줄이 `[`로 시작하는지 확인
    - 마지막 줄이 `]`로 끝나는지 확인
    """
    content = content.strip()
    
    # 문자열 언래핑 (앞뒤 따옴표 제거)
    if content.startswith('"['):
        content = content[2:]  # "[" 제거
    elif content.startswith('"'):
        content = content[1:]  # 첫 " 제거
    
    if content.endswith(']"'):
        content = content[:-2] + ']'  # ]" 제거하고 ] 추가
    elif content.endswith('"'):
        content = content[:-1]  # 마지막 " 제거
    
    # 배열 시작 확인
    if not content.startswith('['):
        content = '[' + content
    
    # 배열 끝 확인
    content = content.rstrip()
    if not content.endswith(']'):
        # 마지막 쉼표 제거 후 배열 닫기
        content = content.rstrip(',').rstrip() + '\n]'
    
    return content


def fix_hvdc_status_file(filepath: Path) -> bool:
    """
    HVDC SATUS.JSON 파일 수정
    """
    print(f"\n[처리 중] {filepath.name}")
    
    try:
        # 파일 읽기
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"[OK] 파일 읽기 완료 ({len(content):,} 문자)")
        
        # 1. 이중 인용부호 정규화 (먼저 변환)
        content = normalize_double_quotes(content)
        print("[OK] 이중 인용부호 정규화 완료")
        
        # 2. 줄바꿈으로 나뉜 값 수정 (이중 따옴표 변환 후 처리)
        content = fix_line_breaks_after_normalization(content)
        print("[OK] 줄바꿈 문제 수정 완료")
        
        # 3. 배열 구조 확인 및 수정
        content = ensure_array_structure(content)
        print("[OK] 배열 구조 정규화 완료")
        
        # 4. JSON 파싱 및 검증
        try:
            data = json.loads(content)
            print(f"[OK] JSON 유효성 검증 통과 ({len(data):,}개 레코드)")
        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON 파싱 오류: {e}")
            print(f"[ERROR] 오류 위치: line {e.lineno}, column {e.colno}")
            # 오류 위치 주변 출력
            lines = content.split('\n')
            start = max(0, e.lineno - 3)
            end = min(len(lines), e.lineno + 2)
            print(f"[ERROR] 오류 위치 주변:")
            for i in range(start, end):
                marker = ">>> " if i == e.lineno - 1 else "    "
                print(f"{marker}{i+1}: {lines[i][:100]}")
            return False
        
        # 5. 원본 파일 백업
        backup_path = filepath.with_suffix('.json.backup')
        with open(backup_path, 'w', encoding='utf-8') as f:
            with open(filepath, 'r', encoding='utf-8') as orig:
                f.write(orig.read())
        print(f"[OK] 백업 파일 생성: {backup_path.name}")
        
        # 6. 정규화된 JSON 저장 (들여쓰기 2칸)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print("[OK] 파일 수정 완료")
        
        # 7. 파일 크기 비교
        backup_size = backup_path.stat().st_size
        new_size = filepath.stat().st_size
        print(f"[INFO] 파일 크기: {backup_size:,} bytes -> {new_size:,} bytes")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """
    메인 함수
    """
    base_dir = Path(__file__).parent
    file_path = base_dir / "HVDC SATUS.JSON"
    
    print("=" * 60)
    print("HVDC SATUS.JSON 파일 수정 시작")
    print("=" * 60)
    
    if not file_path.exists():
        print(f"[ERROR] 파일 없음: {file_path}")
        return
    
    if fix_hvdc_status_file(file_path):
        print("\n" + "=" * 60)
        print("수정 완료!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("수정 실패 - 오류 로그를 확인하세요")
        print("=" * 60)


if __name__ == "__main__":
    main()
