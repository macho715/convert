#!/usr/bin/env python3
"""
JSON 정규화 스크립트 (최종 버전)
- 이중 인용부호(`""`)를 단일 인용부호(`"`)로 변환
- 줄바꿈으로 분리된 문자열 값 합치기
- 제어 문자 제거
- 올바른 JSON 배열 구조로 변환
"""

import json
import re
from pathlib import Path


def normalize_double_quotes(content: str) -> str:
    """
    이중 인용부호를 단일 인용부호로 변환
    더 정교한 방법: 모든 ""를 "로 변환하되, 연속된 ""는 하나의 "로
    """
    # 모든 이중 인용부호를 단일 인용부호로 변환
    # "" -> "
    content = content.replace('""', '"')
    
    return content


def fix_split_string_values(content: str) -> str:
    """
    줄바꿈으로 분리된 문자열 값을 합치기
    예: "HVDC-ADOP"\n"T-SCT-0041" -> "HVDC-ADOPT-SCT-0041"
    """
    # 패턴: "value1"\n"value2" (키 다음 값이 분리된 경우)
    # 정규식으로 처리
    # 패턴: ": "value1"\n"value2"
    pattern = r'(":\s*")([^"]*)"\s*\n\s*"([^"]*)"'
    
    def replace_func(match):
        key_part = match.group(1)  # ": "
        value1 = match.group(2)     # 첫 번째 값 부분
        value2 = match.group(3)     # 두 번째 값 부분
        return f'{key_part}{value1}{value2}"'
    
    content = re.sub(pattern, replace_func, content)
    
    return content


def fix_control_characters(content: str) -> str:
    """
    JSON에서 허용되지 않는 제어 문자 제거 또는 이스케이프
    """
    # 문자열 값 안의 제어 문자 처리
    # JSON 표준에 따라 \n, \r, \t는 허용되지만 다른 제어 문자는 제거
    result = []
    in_string = False
    escape_next = False
    
    for i, char in enumerate(content):
        if escape_next:
            result.append(char)
            escape_next = False
            continue
        
        if char == '\\':
            escape_next = True
            result.append(char)
            continue
        
        if char == '"':
            in_string = not in_string
            result.append(char)
            continue
        
        if in_string:
            # 문자열 안에서는 특수 제어 문자만 제거
            # \n, \r, \t는 유지
            if ord(char) < 32 and char not in ['\t', '\n', '\r']:
                # 허용되지 않는 제어 문자는 제거
                continue
            result.append(char)
        else:
            # 문자열 밖에서는 모든 문자 허용
            result.append(char)
    
    return ''.join(result)


def normalize_json_content(content: str) -> str:
    """
    JSON 내용 정규화
    """
    # 1. 줄바꿈으로 분리된 문자열 값 합치기 (이중 인용부호 변환 전에 처리)
    content = fix_split_string_values(content)
    
    # 2. 이중 인용부호 정규화
    content = normalize_double_quotes(content)
    
    # 3. 줄바꿈이 값 안에 포함된 케이스 처리 (이미 이중 인용부호가 변환된 후)
    # 패턴: "value"\n" -> "value" (불필요한 줄바꿈과 인용부호 제거)
    content = re.sub(r'"([^"]*)"\s*\n\s*"', r'"\1"', content)
    
    # 4. 제어 문자 처리
    content = fix_control_characters(content)
    
    # 5. 남은 이중 인용부호 제거 (혹시 모를 경우)
    content = content.replace('""', '"')
    
    return content


def ensure_array_structure(content: str) -> str:
    """
    JSON 배열 구조 확인 및 수정
    """
    lines = content.split('\n')
    
    # 첫 줄 확인 및 수정
    first_line = lines[0].strip()
    
    # 문자열로 래핑된 경우 제거
    if first_line.startswith('"['):
        lines[0] = first_line[2:]  # "[" 제거
    elif first_line.startswith('"') and first_line[1:].strip().startswith('['):
        lines[0] = first_line[1:]  # 첫 " 제거
    
    # 배열 시작 확인
    if not lines[0].strip().startswith('['):
        # 첫 줄이 `{`로 시작하면 `[`를 앞에 추가
        if lines[0].strip().startswith('{'):
            lines[0] = '[' + lines[0]
        else:
            # 공백만 있으면 제거하고 `[` 추가
            lines[0] = '[' + lines[0].lstrip()
    
    # 마지막 줄 확인 및 수정
    last_non_empty = -1
    for i in range(len(lines) - 1, -1, -1):
        if lines[i].strip():
            last_non_empty = i
            break
    
    if last_non_empty >= 0:
        last_line = lines[last_non_empty].strip()
        
        # 마지막 쉼표 제거
        if last_line.endswith(','):
            lines[last_non_empty] = last_line.rstrip(',')
            last_line = lines[last_non_empty].strip()
        
        # 배열 닫기 추가
        if not last_line.endswith(']'):
            if last_line.endswith('}'):
                lines[last_non_empty] = last_line + '\n]'
            else:
                # 빈 줄들 제거하고 `]` 추가
                for i in range(last_non_empty + 1, len(lines)):
                    lines[i] = ''
                lines[last_non_empty] = last_line + '\n]'
    
    return '\n'.join(lines)


def normalize_json_file(input_path: Path, output_path: Path = None) -> bool:
    """
    JSON 파일 정규화
    """
    if output_path is None:
        # 원본 파일을 백업하고 정규화된 파일로 교체
        backup_path = input_path.with_suffix('.backup.json')
        output_path = input_path
    
    print(f"[입력] {input_path}")
    print(f"[출력] {output_path}")
    
    try:
        # 원본 백업
        if output_path == input_path:
            import shutil
            if not backup_path.exists():
                shutil.copy2(input_path, backup_path)
                print(f"[백업] {backup_path}")
        
        # 파일 읽기
        with open(input_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"[OK] 파일 읽기 완료 ({len(content):,} 문자)")
        
        # 1. 이중 인용부호 정규화
        content = normalize_json_content(content)
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
            print(f"[WARN] 오류 위치: line {e.lineno}, column {e.colno}")
            
            # 오류 위치 주변 출력
            lines = content.split('\n')
            error_line_idx = e.lineno - 1
            start_line = max(0, error_line_idx - 3)
            end_line = min(len(lines), error_line_idx + 4)
            
            print(f"[WARN] 오류 위치 주변 (line {start_line+1}-{end_line}):")
            for i in range(start_line, end_line):
                marker = ">>> " if i == error_line_idx else "    "
                line_content = lines[i] if i < len(lines) else ""
                print(f"{marker}{i+1:5d}: {line_content[:100]}")
            
            # 수동 수정을 위한 힌트
            print("\n[INFO] 수동 수정이 필요할 수 있습니다.")
            # 중간 결과 저장 (디버깅용)
            debug_path = input_path.with_suffix('.debug.json')
            with open(debug_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"[DEBUG] 중간 결과 저장: {debug_path}")
            return False
        
        # 4. 정규화된 JSON 저장 (들여쓰기 2칸)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print("[OK] 정규화된 JSON 저장 완료")
        
        # 5. 파일 크기 비교
        input_size = input_path.stat().st_size
        output_size = output_path.stat().st_size
        print(f"[INFO] 파일 크기: {input_size:,} bytes -> {output_size:,} bytes")
        
        # 6. 검증: 다시 읽어서 파싱 확인
        with open(output_path, 'r', encoding='utf-8') as f:
            verify_data = json.load(f)
        print(f"[OK] 최종 검증 통과 ({len(verify_data)}개 레코드)")
        
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
    print("JSON 정규화 시작 (최종 버전)")
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
