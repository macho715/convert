#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC SATUS.JSON 파일 최종 수정 스크립트 (JSON Spec 준수)
- Top-level 배열 문자열 래핑 제거
- 줄바꿈 분할된 값 병합
- 이중 따옴표 정규화
- JSON 유효성 검증
"""

import json
import re
from pathlib import Path


def strip_array_wrapping(lines):
    """Top-level 배열을 문자열 래핑 제거."""
    if lines and lines[0].strip().startswith('"['):
        lines[0] = lines[0].replace('"[', '[', 1)
    if lines and lines[-1].strip().endswith(']"'):
        lines[-1] = lines[-1].replace(']"', ']', 1)
    return lines


def merge_broken_values(lines):
    """
    줄바꿈 분할된 값 병합:
    ""Key"": ""
    "Value"",
    → ""Key"": "Value",
    """
    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # "Key": "" <-- pattern
        if re.search(r'""[^"]+"":\s*""\s*$', line):
            if i + 1 < len(lines):
                next_line = lines[i + 1].rstrip()
                # next_line = "Value", or "Value",
                m = re.match(r'^\s*"([^"]+)",?\s*$', next_line)
                if m:
                    key_header = line.rstrip()
                    value_text = m.group(1)
                    merged = re.sub(r'""\s*$', f'"{value_text}",', key_header)
                    out.append(merged + "\n")
                    i += 2
                    continue
        out.append(line)
        i += 1
    return out


def normalize_quotes(text):
    """
    불필요한 연속 이중 따옴표 제거:
    ""Key"" → "Key"
    """
    text = re.sub(r'""([^"]+)""', r'"\1"', text)
    return text


def check_control_characters(content):
    """Control character 검사 및 보고"""
    issues = []
    lines = content.split('\n')
    for i, line in enumerate(lines, 1):
        # 이스케이프되지 않은 제어 문자 검사
        if re.search(r'[^\x20-\x7E\t\n\r]', line):
            # JSON 문자열 값 내부의 제어 문자 확인
            if '"' in line:
                issues.append(f"Line {i}: Possible unescaped control character")
    return issues


def fix_file(input_file: str, output_file: str = None):
    """메인 수정 함수"""
    p = Path(input_file)
    
    if output_file is None:
        output_file = input_file
        # 백업 생성
        backup_path = p.with_suffix('.json.backup_final')
        if not backup_path.exists():
            import shutil
            shutil.copy2(p, backup_path)
            print(f"[백업] {backup_path.name} 생성 완료")
    
    print(f"[처리 시작] {p.name}")
    print(f"[파일 크기] {p.stat().st_size:,} bytes")
    
    try:
        # 1. 파일 읽기
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        
        print(f"[읽기 완료] {len(lines):,} 줄")
        
        # 2. 배열 래핑 제거
        lines = strip_array_wrapping(lines)
        print("[처리 완료] 배열 문자열 래핑 제거")
        
        # 3. 줄바꿈 분할 값 병합
        lines = merge_broken_values(lines)
        print("[처리 완료] 줄바꿈 분할 값 병합")
        
        # 4. 전체 내용을 문자열로 변환
        content = "".join(lines)
        
        # 5. 이중 따옴표 정규화
        content = normalize_quotes(content)
        print("[처리 완료] 이중 따옴표 정규화")
        
        # 6. Control character 검사
        control_issues = check_control_characters(content)
        if control_issues:
            print(f"[경고] Control character 이슈 {len(control_issues)}개 발견")
            for issue in control_issues[:5]:  # 처음 5개만 표시
                print(f"  - {issue}")
        else:
            print("[검사 완료] Control character 문제 없음")
        
        # 7. 배열 구조 최종 확인
        content = content.strip()
        if not content.startswith('['):
            content = '[' + content
        if not content.rstrip().endswith(']'):
            content = content.rstrip().rstrip(',') + '\n]'
        
        # 8. JSON 파싱 및 유효성 검증
        try:
            parsed = json.loads(content)
            print(f"[검증 성공] JSON 유효성 통과 ({len(parsed):,}개 레코드)")
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
        
        # 9. 정규화된 JSON 저장
        output_path = Path(output_file)
        with open(output_path, "w", encoding="utf-8") as out:
            json.dump(parsed, out, indent=2, ensure_ascii=False)
        
        print(f"[저장 완료] {output_path.name}")
        
        # 10. 파일 크기 비교
        new_size = output_path.stat().st_size
        print(f"[파일 크기] {new_size:,} bytes")
        
        # 11. 추가 검증: 레코드 구조 확인
        if parsed:
            sample_keys = list(parsed[0].keys()) if isinstance(parsed[0], dict) else []
            print(f"[레코드 구조] 샘플 키 수: {len(sample_keys)}")
            if sample_keys:
                print(f"[레코드 구조] 샘플 키 (처음 5개): {', '.join(sample_keys[:5])}")
        
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
    print("HVDC SATUS.JSON 파일 최종 수정 스크립트 (JSON Spec 준수)")
    print("=" * 60)
    
    if not input_file.exists():
        print(f"[오류] 파일을 찾을 수 없습니다: {input_file}")
        return
    
    success = fix_file(str(input_file))
    
    print("=" * 60)
    if success:
        print("[완료] 파일 수정이 성공적으로 완료되었습니다!")
        print("[다음 단계] JSON 파싱 검증 완료 - 파일 사용 가능")
    else:
        print("[실패] 파일 수정 중 오류가 발생했습니다.")
    print("=" * 60)


if __name__ == "__main__":
    main()
