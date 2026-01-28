import re
from datetime import datetime
import os

def convert_date(date_str):
    """DD-Mon-YY 형식을 YYYY-MM-DD로 변환"""
    try:
        dt = datetime.strptime(date_str, '%d-%b-%y')
        return dt.strftime('%Y-%m-%d')
    except Exception as e:
        print(f"날짜 변환 오류: {date_str} - {e}")
        return date_str

def parse_line(line):
    """MMT 라인을 파싱하여 컴포넌트 추출"""
    line = line.strip()
    if not line:
        return None
    
    # Activity ID로 시작하는 경우 (A로 시작하는 코드)
    id_match = re.match(r'^([A]\d+)\s+(.+?)\s+(\d+\.?\d*)\s+(\d{2}-[A-Za-z]{3}-\d{2})\s+(\d{2}-[A-Za-z]{3}-\d{2})$', line)
    if id_match:
        activity_id, name, duration, start, finish = id_match.groups()
        return {
            'type': 'activity',
            'id': activity_id,
            'name': name,
            'duration': duration,
            'start': start,
            'finish': finish
        }
    
    # 그룹/요약 항목 (Activity ID 없음)
    # 마지막 4개 필드가: Duration, Start, Finish 형태
    group_match = re.match(r'^(.+?)\s+(\d+\.?\d*)\s+(\d{2}-[A-Za-z]{3}-\d{2})\s+(\d{2}-[A-Za-z]{3}-\d{2})$', line)
    if group_match:
        name, duration, start, finish = group_match.groups()
        return {
            'type': 'group',
            'name': name.strip(),
            'duration': duration,
            'start': start,
            'finish': finish
        }
    
    return None

def convert_mmt_to_optionb(input_file, output_file):
    """MMT TSV를 Option B TSV 포맷으로 변환"""
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    output_lines = []
    output_lines.append("Activity ID\tActivity ID\tActivity ID\tActivity Name\tOriginal Duration\tPlanned Start\tPlanned Finish\n")
    
    # 계층 구조 추적
    level1 = None  # MOBILIZATION, DEMOBILIZATION, OPERATIONAL
    level2 = None  # SPMT, MARINE, JACKING EQUIPMENT, Beam Replacement, Deck Preparations, AGI TR Unit X
    
    for line in lines:
        if 'Samsung C&T' in line or 'HVDC Transformers' in line:
            continue
        
        parsed = parse_line(line)
        if not parsed:
            continue
        
        # Level 1: 메인 카테고리
        if parsed['name'] == 'MOBILIZATION':
            level1 = 'MOBILIZATION'
            level2 = None
            output_lines.append(f"{level1}\t\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        elif parsed['name'] == 'DEMOBILIZATION':
            level1 = 'DEMOBILIZATION'
            level2 = None
            output_lines.append(f"{level1}\t\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        elif parsed['name'] == 'OPERATIONAL':
            level1 = 'OPERATIONAL'
            level2 = None
            output_lines.append(f"{level1}\t\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        # Level 2: 서브 카테고리
        elif parsed['name'] in ['SPMT', 'MARINE', 'JACKING EQUIPMENT, STEEL BRIDGE']:
            if level1:
                level2 = parsed['name']
                output_lines.append(f"{level1}\t{level2}\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        elif parsed['name'] == 'Beam Replacement':
            if level1 == 'OPERATIONAL':
                level2 = 'Beam Replacement'
                output_lines.append(f"{level1}\t{level2}\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        elif parsed['name'] == 'Deck Preparations':
            if level1 == 'OPERATIONAL':
                level2 = 'Deck Preparations'
                output_lines.append(f"{level1}\t{level2}\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        elif parsed['name'].startswith('AGI TR Unit'):
            if level1 == 'OPERATIONAL':
                level2 = parsed['name']
                output_lines.append(f"{level1}\t{level2}\t\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
        
        # Level 3: 실제 활동 (Activity ID 포함)
        elif parsed['type'] == 'activity':
            if level1 and level2:
                output_lines.append(f"{level1}\t{level2}\t{parsed['id']}\t{parsed['name']}\t{parsed['duration']}\t{convert_date(parsed['start'])}\t{convert_date(parsed['finish'])}\n")
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.writelines(output_lines)
    
    print(f"변환 완료: {output_file}")
    print(f"   총 {len(output_lines)-1}개 행 변환됨")

# 실행
if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(base_dir, "option_b_mmt.tsv")
    output_file = os.path.join(base_dir, "option_b_mmt_converted.tsv")
    
    convert_mmt_to_optionb(input_file, output_file)
