#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC SATUS.JSON 파일 수정 스크립트 (개선 버전)
- 이중 따옴표(Double Quote) 정규화
- 줄바꿈으로 분할된 값(Line Break Value Split) 병합
- 배열 문자열 래핑 제거
- JSON 유효성 검증 및 저장
"""

import json
import re
from pathlib import Path


def remove_array_wrapping(lines):
    """배열 문자열 래핑 제거: "[" -> [ , ]" -> ]"""
    if not lines:
        return lines
    
    # 첫 줄 처리
    first_line = lines[0].strip()
    if first_line.startswith('"['):
        lines[0] = lines[0].replace('"[', '[', 1)
    elif first_line.startswith('"') and first_line[1:].strip().startswith('['):
        lines[0] = '[' + lines[0].lstrip('"').lstrip()
    
    # 마지막 줄 처리
    if lines:
        last_line = lines[-1].strip()
        if last_line.endswith(']"'):
            lines[-1] = lines[-1].replace(']"', ']', 1)
        elif last_line.endswith('"') and last_line.rstrip('"').strip().endswith(']'):
            lines[-1] = lines[-1].rstrip('"').rstrip() + ']'
    
    return lines


def merge_split_values(lines):
    """
    줄바꿈으로 분할된 값 병합 (이중 따옴표 상태에서 처리)
    패턴: ""key"": ""\n"value"", -> ""key"": ""value"",
    """
    fixed_lines = []
    i = 0
    
    while i < len(lines):
        current_line = lines[i]
        
        # 현재 줄이 ""key"": ""로 끝나는지 확인 (공백 포함)
        # 패턴: ""CBM"": ""
        if re.search(r'""[^"]+"":\s*""\s*$', current_line):
            # 다음 줄 확인
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                
                # 다음 줄이 "value"", 또는 "value", 형태인지 확인
                # 패턴: "391.3"", 또는 "391.3",
                # 이중 따옴표로 끝나는 경우와 단일 따옴표로 시작하는 경우 모두 처리
                match1 = re.match(r'^\s*"([^"]+)",?\s*$', next_line.strip())
                match2 = re.match(r'^\s*"([^"]+)""', next_line.strip())
                
                if match1 or match2:
                    value = match1.group(1) if match1 else match2.group(1)
                    # 현재 줄의 키와 들여쓰기 추출
                    key_match = re.match(r'^(\s*""[^"]+"":\s*)""\s*$', current_line)
                    if key_match:
                        # 병합: ""key"": ""value"",
                        merged_line = key_match.group(1) + '""' + value + '"",'
                        fixed_lines.append(merged_line)
                        i += 2
                        continue
        
        fixed_lines.append(current_line)
        i += 1
    
    return fixed_lines


def normalize_double_quotes(content):
    """
    이중 따옴표를 단일 따옴표로 변환
    ""key"": ""value"" -> "key": "value"
    """
    return content.replace('""', '"')


def fix_hvdc_status_file(input_path, output_path=None):
    """
    HVDC SATUS.JSON 파일 수정 메인 함수
    """
    input_path = Path(input_path)
    
    if output_path is None:
        # 백업 생성
        backup_path = input_path.with_suffix('.json.backup')
        if not backup_path.exists():
            import shutil
            shutil.copy2(input_path, backup_path)
            print(f"[백업] {backup_path.name} 생성 완료")
        
        output_path = input_path  # 원본 파일 직접 수정
    
    print(f"[처리 시작] {input_path.name}")
    print(f"[파일 크기] {input_path.stat().st_size:,} bytes")
    
    try:
        # 1. 파일 읽기 (라인별)
        with open(input_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
        
        print(f"[읽기 완료] {len(lines):,} 줄")
        
        # 2. 배열 래핑 제거
        lines = remove_array_wrapping(lines)
        print("[처리 완료] 배열 문자열 래핑 제거")
        
        # 3. 줄바꿈으로 분할된 값 병합 (이중 따옴표 상태에서)
        lines = merge_split_values(lines)
        print("[처리 완료] 줄바꿈 분할 값 병합")
        
        # 4. 전체 내용을 문자열로 변환
        content = ''.join(lines)
        
        # 5. 이중 따옴표 정규화
        content = normalize_double_quotes(content)
        print("[처리 완료] 이중 따옴표 정규화")
        
        # 6. 배열 구조 최종 확인
        content = content.strip()
        if not content.startswith('['):
            content = '[' + content
        if not content.rstrip().endswith(']'):
            content = content.rstrip().rstrip(',') + '\n]'
        
        # 7. JSON 파싱 및 유효성 검증
        try:
            data = json.loads(content)
            print(f"[검증 성공] JSON 유효성 통과 ({len(data):,}개 레코드)")
        except json.JSONDecodeError as e:
            print(f"[오류] JSON 파싱 실패: {e}")
            print(f"[오류 위치] line {e.lineno}, column {e.colno}")
            
            # 오류 위치 주변 출력
            error_lines = content.split('\n')
            start = max(0, e.lineno - 3)
            end = min(len(error_lines), e.lineno + 2)
            print(f"[오류 주변 컨텍스트]:")
            for i in range(start, end):
                marker = ">>> " if i == e.lineno - 1 else "    "
                print(f"{marker}{i+1}: {error_lines[i][:100]}")
            
            return False
        
        # 8. 정규화된 JSON 저장
        output_path = Path(output_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        print(f"[저장 완료] {output_path.name}")
        
        # 9. 파일 크기 비교
        new_size = output_path.stat().st_size
        print(f"[파일 크기] {new_size:,} bytes")
        
        return True
        
    except Exception as e:
        print(f"[오류] 처리 중 예외 발생: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 실행 함수"""
    base_dir = Path(__file__).parent
    input_file = base_dir / "HVDC SATUS.JSON"
    
    print("=" * 60)
    print("HVDC SATUS.JSON 파일 수정 스크립트 (개선 버전)")
    print("=" * 60)
    
    if not input_file.exists():
        print(f"[오류] 파일을 찾을 수 없습니다: {input_file}")
        return
    
    success = fix_hvdc_status_file(input_file)
    
    print("=" * 60)
    if success:
        print("[완료] 파일 수정이 성공적으로 완료되었습니다!")
    else:
        print("[실패] 파일 수정 중 오류가 발생했습니다.")
    print("=" * 60)


if __name__ == "__main__":
    main()
