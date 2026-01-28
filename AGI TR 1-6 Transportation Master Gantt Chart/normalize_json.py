#!/usr/bin/env python3
"""
JSON 정규화 스크립트
- 이중 인용부호(`""`)를 단일 인용부호(`"`)로 변환
- 올바른 JSON 배열 구조로 변환
- JSON 유효성 검증
"""

import json
import re
import sys
from pathlib import Path


def normalize_double_quotes(content: str) -> str:
    """
    이중 인용부호를 단일 인용부호로 변환
    ""key"": ""value"" -> "key": "value"
    
    특수 케이스 처리:
    - 값이 여러 줄에 걸쳐 있는 경우: ""value1"\n"" -> "value1"
    """
    # 먼저 줄바꿈이 값 안에 포함된 케이스 처리
    # 패턴: ""value"\n"" -> "value"
    import re
    # 값이 줄바꿈으로 끝나는 경우: ""value"\n"" -> "value"
    content = re.sub(r'""([^"]*)"\n""', r'"\1"', content)
    
    # 일반적인 이중 인용부호를 단일 인용부호로 변환
    content = content.replace('""', '"')
    
    return content


def ensure_array_structure(content: str) -> str:
    """
    JSON 배열 구조 확인 및 수정
    - 첫 줄이 `[`로 시작하는지 확인
    - 마지막 줄이 `]`로 끝나는지 확인
    """
    lines = content.split('\n')
    
    # 첫 줄 확인 및 수정
    first_line = lines[0].strip()
    if not first_line.startswith('['):
        # 첫 줄이 `{`로 시작하면 `[`를 앞에 추가
        if first_line.startswith('{'):
            lines[0] = '[' + lines[0]
        else:
            # 문자열로 래핑된 경우 제거
            if first_line.startswith('"['):
                lines[0] = first_line[2:]  # "[" 제거
            elif first_line.startswith('"'):
                lines[0] = '[' + lines[0].lstrip('"')
    
    # 마지막 줄 확인 및 수정
    last_line = lines[-1].strip()
    if last_line.endswith(','):
        # 마지막 쉼표 제거 후 `]` 추가
        lines[-1] = last_line.rstrip(',') + '\n]'
    elif not last_line.endswith(']'):
        # `]`가 없으면 추가
        if last_line.endswith('}'):
            lines[-1] = last_line + '\n]'
        else:
            lines[-1] = last_line + '\n]'
    
    return '\n'.join(lines)


def normalize_json_file(input_path: Path, output_path: Path = None) -> bool:
    """
    JSON 파일 정규화
    """
    if output_path is None:
        output_path = input_path.with_suffix('.normalized.json')
    
    print(f"[입력] {input_path}")
    print(f"[출력] {output_path}")
    
    try:
        # 파일 읽기
        with open(input_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"[OK] 파일 읽기 완료 ({len(content)} 문자)")
        
        # 1. 이중 인용부호 정규화
        content = normalize_double_quotes(content)
        print("[OK] 이중 인용부호 정규화 완료")
        
        # 2. 배열 구조 확인 및 수정
        content = ensure_array_structure(content)
        print("[OK] 배열 구조 정규화 완료")
        
        # 3. JSON 파싱 및 검증
        try:
            data = json.loads(content)
            print(f"[OK] JSON 유효성 검증 통과 ({len(data)}개 레코드)")
        except json.JSONDecodeError as e:
            print(f"[WARN] JSON 파싱 오류: {e}")
            print("[WARN] 원본 파일을 백업하고 수동 수정이 필요할 수 있습니다.")
            # 오류 위치 정보 출력
            if hasattr(e, 'pos'):
                start = max(0, e.pos - 100)
                end = min(len(content), e.pos + 100)
                print(f"[WARN] 오류 위치 주변:\n{content[start:end]}")
            return False
        
        # 4. 정규화된 JSON 저장 (들여쓰기 2칸)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print("[OK] 정규화된 JSON 저장 완료")
        
        # 5. 파일 크기 비교
        input_size = input_path.stat().st_size
        output_size = output_path.stat().st_size
        print(f"[INFO] 파일 크기: {input_size:,} bytes -> {output_size:,} bytes")
        
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
    # 작업 디렉토리
    base_dir = Path(__file__).parent
    
    # 처리할 파일 목록
    files_to_process = [
        base_dir / "hvdc logistics status_3.json",
        base_dir / "hvdc logistics status_4.json",
    ]
    
    print("=" * 60)
    print("JSON 정규화 시작")
    print("=" * 60)
    
    success_count = 0
    for file_path in files_to_process:
        if not file_path.exists():
            print(f"[WARN] 파일 없음: {file_path}")
            continue
        
        print(f"\n{'=' * 60}")
        if normalize_json_file(file_path):
            success_count += 1
        print(f"{'=' * 60}\n")
    
    print("=" * 60)
    print(f"정규화 완료: {success_count}/{len(files_to_process)} 파일")
    print("=" * 60)


if __name__ == "__main__":
    main()
