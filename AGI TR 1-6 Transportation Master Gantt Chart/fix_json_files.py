import json
import os
import re

def fix_json_file(filepath):
    """JSON 파일의 형식 문제를 수정"""
    print(f"\nProcessing: {filepath}")
    
    # 파일 읽기
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 이중 따옴표를 단일 따옴표로 변환
    content = ''.join(lines)
    content = content.strip('"')
    content = content.replace('""', '"')
    
    # 제어 문자 제거
    content = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', content)
    
    # 라인별로 처리하여 줄바꿈 문제 수정
    fixed_lines = []
    i = 0
    while i < len(content.split('\n')):
        lines_list = content.split('\n')
        if i >= len(lines_list):
            break
            
        line = lines_list[i]
        
        # 다음 라인 확인
        if i + 1 < len(lines_list):
            next_line = lines_list[i + 1].strip()
            
            # 패턴 1: "value", 다음 줄에 "로 시작하는 경우 (잘못된 형식)
            if line.rstrip().endswith('","') and next_line.startswith('"'):
                # 잘못된 따옴표 제거
                line = line.rstrip().rstrip('","') + '",'
                fixed_lines.append(line)
                i += 1
                continue
            
            # 패턴 2: "value" 다음 줄에 "continuation", 형태
            if line.rstrip().endswith('"') and not line.rstrip().endswith('",') and next_line.strip().startswith('"') and next_line.strip().endswith('",'):
                # 두 줄을 합치기
                continuation = next_line.strip().strip('"').rstrip(',')
                line = line.rstrip().rstrip('"') + continuation + '",'
                fixed_lines.append(line)
                i += 2
                continue
            
            # 패턴 3: "key" 다음 줄에 "continuation": 형태
            if line.rstrip().endswith('"') and next_line.strip().startswith('"') and '":' in next_line:
                continuation = next_line.strip().strip('"').split('":')[0]
                rest = next_line.split('":', 1)[1] if '":' in next_line else ''
                line = line.rstrip().rstrip('"') + continuation + '":' + rest
                fixed_lines.append(line)
                i += 2
                continue
            
            # 패턴 4: 빈 줄이나 공백만 있는 줄 다음에 "로 시작하는 경우
            if not line.strip() and next_line.strip().startswith('"') and i + 2 < len(lines_list):
                # 다음 다음 줄 확인
                next_next = lines_list[i + 2].strip()
                if next_line.strip().startswith('"') and not '":' in next_line and next_next.strip().startswith('"'):
                    # 중간 빈 줄 제거하고 합치기
                    continuation = next_line.strip().strip('"')
                    next_continuation = next_next.strip().strip('"').split('":')[0] if '":' in next_next else next_next.strip().strip('"')
                    rest = next_next.split('":', 1)[1] if '":' in next_next else ''
                    line = '"' + continuation + next_continuation + '":' + rest
                    fixed_lines.append(line)
                    i += 3
                    continue
        
        fixed_lines.append(line)
        i += 1
    
    content = '\n'.join(fixed_lines)
    
    # 파일 끝 정리
    content = content.rstrip()
    
    # 배열 시작 확인 및 추가
    if not content.startswith('['):
        content = '[\n' + content
    
    # 배열 끝 확인 및 추가
    if not content.rstrip().endswith(']'):
        content = content.rstrip().rstrip(',') + '\n]'
    
    # JSON 파싱하여 유효성 검증
    try:
        data = json.loads(content)
        print(f"  JSON parsing success: {len(data)} items")
        
        # 유효한 JSON으로 저장
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        print(f"  File fixed successfully")
        return True
        
    except json.JSONDecodeError as e:
        print(f"  JSON parsing error: {e}")
        print(f"  Error at line {e.lineno}, column {e.colno}")
        # 오류 위치 주변 출력
        lines = content.split('\n')
        start = max(0, e.lineno - 3)
        end = min(len(lines), e.lineno + 2)
        print(f"  Context around error:")
        for i in range(start, end):
            marker = ">>> " if i == e.lineno - 1 else "    "
            line_content = lines[i] if i < len(lines) else ""
            print(f"  {marker}{i+1}: {line_content[:100]}")
        return False

# 두 파일 수정
files = [
    'hvdc logistics status_1.json',
    'hvdc logistics status_2.json'
]

for filename in files:
    if os.path.exists(filename):
        fix_json_file(filename)
    else:
        print(f"\nFile not found: {filename}")

print("\nAll files processed.")
