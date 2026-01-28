import csv
import json
import re
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime

def parse_value(value: str) -> Any:
    """값 파싱 헬퍼 (숫자, > 기호 등 처리)"""
    if not value:
        return None
    value = value.strip()
    if value.startswith('>'):
        try:
            num_val = value[1:].strip()
            if num_val.replace('.', '').replace('-', '').isdigit():
                return {"operator": ">", "value": float(num_val)}
            return value
        except:
            return value
    try:
        return float(value)
    except:
        return value

def parse_table_6_1(lines: List[str], start_idx: int) -> Tuple[Dict, int]:
    """Table 6.1: Initial Loading Condition - 2줄 헤더 처리"""
    table = {
        "table_number": "6.1",
        "title": "Initial Loading Condition",
        "headers": ["S.No.", "Description", "Weight (MT)", "LCG(1) (M)", "TCG(1) (M)", "VCG(1) (M)"],
        "data": []
    }
    
    i = start_idx
    # 빈 줄 건너뛰기
    while i < len(lines) and not lines[i].strip():
        i += 1
    
    # 헤더 2줄 건너뛰기 (라인 6-7)
    header_count = 0
    while i < len(lines) and header_count < 2:
        line = lines[i].strip()
        if 'S.No.' in line or 'Description' in line:
            header_count += 1
        i += 1
    
    # 데이터 행 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith('6.3') or line.startswith('Table 6.2'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 6:
            try:
                # S.No.가 숫자이거나 "Total"인 경우
                if parts[0].isdigit() or parts[0] == "Total":
                    row = {
                        "S.No.": parts[0] if parts[0] else None,
                        "Description": parts[1] if len(parts) > 1 else None,
                        "Weight (MT)": float(parts[2]) if len(parts) > 2 and parts[2] else None,
                        "LCG(1) (M)": float(parts[3]) if len(parts) > 3 and parts[3] else None,
                        "TCG(1) (M)": float(parts[4]) if len(parts) > 4 and parts[4] else None,
                        "VCG(1) (M)": float(parts[5]) if len(parts) > 5 and parts[5] else None
                    }
                    table["data"].append(row)
            except (ValueError, IndexError) as e:
                pass
        i += 1
    
    return table, i

def parse_table_6_2(lines: List[str], start_idx: int) -> Tuple[Dict, int]:
    """Table 6.2: Vessel Floating Condition - 키-값 쌍 형태"""
    table = {
        "table_number": "6.2",
        "title": "Vessel Floating Condition",
        "headers": ["Description", "Value"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기
    while i < len(lines) and ('Description' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 키-값 쌍 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('7.'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 2:
            key = parts[0]
            value = parts[1]
            # 숫자 변환 시도
            try:
                if '(' in value:
                    # "0.83 (+ve by stern)" 같은 경우
                    num_part = value.split('(')[0].strip()
                    value_obj = {"numeric": float(num_part), "note": value}
                    table["data"].append({
                        "Description": key,
                        "Value": value_obj
                    })
                else:
                    value_num = float(value)
                    table["data"].append({
                        "Description": key,
                        "Value": value_num
                    })
            except:
                table["data"].append({
                    "Description": key,
                    "Value": value
                })
        i += 1
    
    return table, i

def parse_table_7_1_7_3(lines: List[str], start_idx: int, table_num: str) -> Tuple[Dict, int]:
    """Table 7.1, 7.3: Ballast Arrangement - 서브헤더 처리"""
    table = {
        "table_number": table_num,
        "title": "Ballast Arrangement at each stage",
        "headers": ["Stages", "Weight (MT)", "FW1.P (% Filled)", "FW1.S (% Filled)", 
                   "FWB2.P (% Filled)", "FWB2.S (% Filled)", "FW2.P (% Filled)", "FW2.S (% Filled)"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기
    while i < len(lines) and ('Stages' in lines[i] or '% Filled' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 데이터 행 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('Table') or line.startswith('Notes:'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 8 and parts[0].isdigit():
            try:
                row = {
                    "Stages": int(parts[0]),
                    "Weight (MT)": int(parts[1]) if parts[1] else None,
                    "FW1.P (% Filled)": int(parts[2]) if len(parts) > 2 and parts[2] else None,
                    "FW1.S (% Filled)": int(parts[3]) if len(parts) > 3 and parts[3] else None,
                    "FWB2.P (% Filled)": int(parts[4]) if len(parts) > 4 and parts[4] else None,
                    "FWB2.S (% Filled)": int(parts[5]) if len(parts) > 5 and parts[5] else None,
                    "FW2.P (% Filled)": int(parts[6]) if len(parts) > 6 and parts[6] else None,
                    "FW2.S (% Filled)": int(parts[7]) if len(parts) > 7 and parts[7] else None
                }
                table["data"].append(row)
            except (ValueError, IndexError):
                pass
        i += 1
    
    return table, i

def parse_table_7_2_7_4(lines: List[str], start_idx: int, table_num: str) -> Tuple[Dict, int]:
    """Table 7.2, 7.4: Vessel Floating Condition at each stage"""
    table = {
        "table_number": table_num,
        "title": "Vessel Floating Condition at each stage",
        "headers": ["Stages", "Weight (MT)", "Draft Aft (m)", "Draft Fwd (m)", 
                   "Draft Mid (m)", "Trim (m)", "Heel (deg)"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기
    while i < len(lines) and ('Stages' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 데이터 행 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('Table') or line.startswith('Notes:'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 7 and parts[0].isdigit():
            try:
                row = {
                    "Stages": int(parts[0]),
                    "Weight (MT)": int(parts[1]) if parts[1] else None,
                    "Draft Aft (m)": float(parts[2]) if len(parts) > 2 and parts[2] else None,
                    "Draft Fwd (m)": float(parts[3]) if len(parts) > 3 and parts[3] else None,
                    "Draft Mid (m)": float(parts[4]) if len(parts) > 4 and parts[4] else None,
                    "Trim (m)": float(parts[5]) if len(parts) > 5 and parts[5] else None,
                    "Heel (deg)": float(parts[6]) if len(parts) > 6 and parts[6] else None
                }
                table["data"].append(row)
            except (ValueError, IndexError):
                pass
        i += 1
    
    return table, i

def parse_table_8_1_8_2(lines: List[str], start_idx: int, table_num: str) -> Tuple[Dict, int]:
    """Table 8.1, 8.2: Intact Stability Results - 여러 줄 헤더 처리"""
    transformer_num = "1" if table_num == "8.1" else "2"
    table = {
        "table_number": table_num,
        "title": f"Intact Stability Results for Loadout of transformer-{transformer_num}",
        "headers": ["SL.No.", "Particulars", "Required", "Stage 1", "Stage 2", "Stage 3", 
                   "Stage 4", "Stage 5", "Stage 6", "Remark"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기 ("SL.\nNo." 포함)
    while i < len(lines) and ('SL.' in lines[i] or 'Particulars' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 데이터 행 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('Table'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 3 and parts[0].isdigit():
            try:
                # 여러 줄에 걸친 Particulars 처리
                particulars = parts[1] if len(parts) > 1 else ""
                if i + 1 < len(lines):
                    next_line = lines[i+1].strip()
                    next_parts = [p.strip() for p in next_line.split('\t') if p.strip()]
                    # 다음 줄이 데이터 행이 아니면 현재 행의 연속
                    if next_parts and not next_parts[0].isdigit() and len(next_parts) < 3:
                        particulars = particulars + " " + next_parts[0] if next_parts else particulars
                        i += 1
                
                row = {
                    "SL.No.": int(parts[0]),
                    "Particulars": particulars,
                    "Required": parts[2] if len(parts) > 2 else None,
                    "Stage 1": parse_value(parts[3]) if len(parts) > 3 else None,
                    "Stage 2": parse_value(parts[4]) if len(parts) > 4 else None,
                    "Stage 3": parse_value(parts[5]) if len(parts) > 5 else None,
                    "Stage 4": parse_value(parts[6]) if len(parts) > 6 else None,
                    "Stage 5": parse_value(parts[7]) if len(parts) > 7 else None,
                    "Stage 6": parse_value(parts[8]) if len(parts) > 8 else None,
                    "Remark": parts[9] if len(parts) > 9 else None
                }
                table["data"].append(row)
            except (ValueError, IndexError) as e:
                pass
        i += 1
    
    return table, i

def parse_table_9_1(lines: List[str], start_idx: int) -> Tuple[Dict, int]:
    """Table 9.1: Allowable Bending moment and shear force - 2열 구조"""
    table = {
        "table_number": "9.1",
        "title": "Allowable Bending moment and shear force for \"LCT Bushra\"",
        "headers": ["Category", "Parameter", "Value", "Unit"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기
    while i < len(lines) and ('Shear Force' in lines[i] or 'Bending Moment' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 2열 구조 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('Table'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 2:
            # 첫 번째 행: 헤더 행
            if 'Shear Force' in parts[0] or 'Allowable' in parts[0]:
                if len(parts) >= 4:
                    # Shear Force 열
                    table["data"].append({
                        "Category": "Shear Force",
                        "Parameter": parts[0],
                        "Value": parse_value(parts[1]),
                        "Unit": parts[2] if len(parts) > 2 else None
                    })
                    # Bending Moment 열
                    if len(parts) >= 4:
                        table["data"].append({
                            "Category": "Bending Moment",
                            "Parameter": parts[2] if len(parts) > 2 else None,
                            "Value": parse_value(parts[3]) if len(parts) > 3 else None,
                            "Unit": parts[4] if len(parts) > 4 else None
                        })
            else:
                # 데이터 행
                if len(parts) >= 4:
                    table["data"].extend([
                        {
                            "Category": "Shear Force",
                            "Parameter": parts[0],
                            "Value": parse_value(parts[1]),
                            "Unit": None
                        },
                        {
                            "Category": "Bending Moment",
                            "Parameter": parts[2] if len(parts) > 2 else None,
                            "Value": parse_value(parts[3]) if len(parts) > 3 else None,
                            "Unit": None
                        }
                    ])
        i += 1
    
    return table, i

def parse_table_9_2(lines: List[str], start_idx: int) -> Tuple[Dict, int]:
    """Table 9.2: Bending Moment and Shear force - 서브헤더 처리"""
    table = {
        "table_number": "9.2",
        "title": "Bending Moment and Shear force for loadout of Transformer-1",
        "headers": ["Stages", "Weight (MT)", "Bending Moment Maximum (MT.m)", 
                   "Bending Moment Allowable (MT.m)", "Shear Force Maximum (MT)", 
                   "Shear Force Allowable (MT)"],
        "data": []
    }
    
    i = start_idx
    # 헤더 건너뛰기
    while i < len(lines) and ('Stages' in lines[i] or 'Maximum' in lines[i] or 'Allowable' in lines[i] or not lines[i].strip()):
        i += 1
    
    # 데이터 행 파싱
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('10.'):
            break
        
        parts = [p.strip() for p in line.split('\t') if p.strip()]
        if len(parts) >= 2 and parts[0].isdigit():
            try:
                row = {
                    "Stages": int(parts[0]),
                    "Weight (MT)": int(parts[1]) if parts[1] else None,
                    "Bending Moment Maximum (MT.m)": int(parts[2]) if len(parts) > 2 and parts[2] else None,
                    "Bending Moment Allowable (MT.m)": int(parts[3]) if len(parts) > 3 and parts[3] else None,
                    "Shear Force Maximum (MT)": int(parts[4]) if len(parts) > 4 and parts[4] else None,
                    "Shear Force Allowable (MT)": int(parts[5]) if len(parts) > 5 and parts[5] else None
                }
                table["data"].append(row)
            except (ValueError, IndexError):
                pass
        i += 1
    
    return table, i

def parse_csv_to_json_enhanced(csv_file_path: str) -> Dict[str, Any]:
    """개선된 CSV to JSON 변환기 - 모든 테이블 구조 처리"""
    
    with open(csv_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
        lines = content.split('\n')
    
    result = {
        "meta": {
            "source": csv_file_path,
            "type": "stability_calculation_report",
            "parsed_at": datetime.now().isoformat(),
            "total_tables": 10,
            "sections": []
        },
        "sections": []
    }
    
    current_section = None
    i = 0
    
    # 첫 번째 셀 처리 (여러 줄로 구성된 따옴표 셀)
    if i < len(lines) and lines[i].strip().startswith('"'):
        first_cell_lines = []
        j = i
        while j < len(lines):
            line = lines[j]
            first_cell_lines.append(line)
            if line.rstrip().endswith('"') and j > i:
                break
            j += 1
        
        # 첫 번째 셀 내용 파싱
        first_cell_content = '\n'.join(first_cell_lines).strip('"').strip()
        
        # 섹션 6.2 감지
        section_match = re.search(r'^(\d+\.\d*)\s+(.+)$', first_cell_content, re.MULTILINE)
        if section_match:
            section_num = section_match.group(1)
            section_title = section_match.group(2).split('\n')[0].strip()
            current_section = {
                "section_number": section_num,
                "title": section_title,
                "description": first_cell_content.split('\n')[1] if '\n' in first_cell_content else "",
                "tables": []
            }
        
        # Table 6.1 감지
        table_match = re.search(r'Table\s+(\d+\.\d+\.?):\s*(.+)$', first_cell_content, re.MULTILINE)
        if table_match:
            table_num = table_match.group(1).rstrip('.')
            table_title = table_match.group(2).strip().strip('"')
            # Table 6.1 파싱
            table, i = parse_table_6_1(lines, j + 1)
            if current_section:
                current_section["tables"].append(table)
        
        i = j + 1
    
    while i < len(lines):
        line = lines[i].strip()
        
        # 따옴표 제거
        if line.startswith('"') and line.endswith('"'):
            line = line[1:-1].strip()
        elif line.startswith('"'):
            # 여러 줄에 걸친 따옴표 처리
            line = line[1:]
            while i + 1 < len(lines) and not lines[i+1].strip().endswith('"'):
                i += 1
                line += " " + lines[i].strip()
            if i + 1 < len(lines):
                i += 1
                line += " " + lines[i].strip().rstrip('"')
        
        # 테이블 제목 감지 및 파싱 (섹션보다 우선)
        table_match = re.match(r'^Table\s+(\d+\.\d+\.?):\s*(.+)$', line)
        if table_match:
            table_num = table_match.group(1).rstrip('.')
            table_title = table_match.group(2).strip().strip('"')
            
            # 섹션 6.2가 아직 없으면 생성
            if table_num == "6.1" and not current_section:
                current_section = {
                    "section_number": "6.2",
                    "title": "Initial Loading Condition",
                    "description": "",
                    "tables": []
                }
            
            i += 1
            
            # 테이블별 파서 호출
            try:
                if table_num == "6.1":
                    table, i = parse_table_6_1(lines, i)
                elif table_num == "6.2":
                    table, i = parse_table_6_2(lines, i)
                elif table_num == "7.1":
                    table, i = parse_table_7_1_7_3(lines, i, table_num)
                elif table_num == "7.2":
                    table, i = parse_table_7_2_7_4(lines, i, table_num)
                elif table_num == "7.3":
                    table, i = parse_table_7_1_7_3(lines, i, table_num)
                elif table_num == "7.4":
                    table, i = parse_table_7_2_7_4(lines, i, table_num)
                elif table_num == "8.1":
                    table, i = parse_table_8_1_8_2(lines, i, table_num)
                elif table_num == "8.2":
                    table, i = parse_table_8_1_8_2(lines, i, table_num)
                elif table_num == "9.1":
                    table, i = parse_table_9_1(lines, i)
                elif table_num == "9.2":
                    table, i = parse_table_9_2(lines, i)
                else:
                    # 기본 파서
                    table = {"table_number": table_num, "title": table_title, "headers": [], "data": []}
                
                if current_section:
                    current_section["tables"].append(table)
            except Exception as e:
                print(f"[WARNING] 테이블 {table_num} 파싱 오류: {e}")
            
            continue
        
        # 섹션 헤더 감지
        section_match = re.match(r'^(\d+\.\d*)\s+(.+)$', line)
        if section_match:
            if current_section:
                result["sections"].append(current_section)
            
            section_num = section_match.group(1)
            section_title = section_match.group(2)
            current_section = {
                "section_number": section_num,
                "title": section_title,
                "description": "",
                "tables": []
            }
            i += 1
            continue
        if table_match:
            table_num = table_match.group(1).rstrip('.')
            table_title = table_match.group(2).strip().strip('"')
            
            i += 1
            
            # 테이블별 파서 호출
            try:
                if table_num == "6.1":
                    table, i = parse_table_6_1(lines, i)
                elif table_num == "6.2":
                    table, i = parse_table_6_2(lines, i)
                elif table_num == "7.1":
                    table, i = parse_table_7_1_7_3(lines, i, table_num)
                elif table_num == "7.2":
                    table, i = parse_table_7_2_7_4(lines, i, table_num)
                elif table_num == "7.3":
                    table, i = parse_table_7_1_7_3(lines, i, table_num)
                elif table_num == "7.4":
                    table, i = parse_table_7_2_7_4(lines, i, table_num)
                elif table_num == "8.1":
                    table, i = parse_table_8_1_8_2(lines, i, table_num)
                elif table_num == "8.2":
                    table, i = parse_table_8_1_8_2(lines, i, table_num)
                elif table_num == "9.1":
                    table, i = parse_table_9_1(lines, i)
                elif table_num == "9.2":
                    table, i = parse_table_9_2(lines, i)
                else:
                    # 기본 파서
                    table = {"table_number": table_num, "title": table_title, "headers": [], "data": []}
                
                if current_section:
                    current_section["tables"].append(table)
            except Exception as e:
                print(f"[WARNING] 테이블 {table_num} 파싱 오류: {e}")
            
            continue
        
        # 섹션 설명 수집
        if current_section and line and not line.startswith('Table'):
            if current_section["description"]:
                current_section["description"] += " " + line
            else:
                current_section["description"] = line
        
        i += 1
    
    # 마지막 섹션 저장
    if current_section:
        result["sections"].append(current_section)
    
    # 검증 리포트 생성
    result["validation"] = validate_tables(result)
    
    return result

def validate_tables(data: Dict) -> Dict:
    """테이블 데이터 검증"""
    validation = {
        "total_tables_found": 0,
        "tables_with_data": 0,
        "tables_missing_data": [],
        "expected_tables": [
            "6.1", "6.2", "7.1", "7.2", "7.3", "7.4", 
            "8.1", "8.2", "9.1", "9.2"
        ],
        "found_tables": [],
        "row_counts": {}
    }
    
    for section in data["sections"]:
        for table in section["tables"]:
            table_num = table["table_number"]
            validation["total_tables_found"] += 1
            validation["found_tables"].append(table_num)
            
            row_count = len(table.get("data", []))
            validation["row_counts"][table_num] = row_count
            
            if row_count > 0:
                validation["tables_with_data"] += 1
            else:
                validation["tables_missing_data"].append(table_num)
    
    validation["all_tables_present"] = (
        set(validation["found_tables"]) == set(validation["expected_tables"])
    )
    
    return validation

def convert_csv_to_json_enhanced(input_file: str, output_file: str = None):
    """개선된 CSV to JSON 변환"""
    if output_file is None:
        output_file = input_file.replace('.csv', '_enhanced.json')
    
    json_data = parse_csv_to_json_enhanced(input_file)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=2, ensure_ascii=False)
    
    # 검증 리포트 출력
    validation = json_data["validation"]
    print(f"[OK] 변환 완료: {input_file} -> {output_file}")
    print(f"\n[검증 결과]")
    print(f"  - 발견된 테이블 수: {validation['total_tables_found']}/10")
    print(f"  - 데이터가 있는 테이블: {validation['tables_with_data']}")
    print(f"  - 모든 테이블 존재: {'YES' if validation['all_tables_present'] else 'NO'}")
    print(f"\n[테이블별 행 수]")
    for table_num in sorted(validation['row_counts'].keys()):
        count = validation['row_counts'][table_num]
        print(f"  - Table {table_num}: {count} rows")
    
    if validation['tables_missing_data']:
        print(f"\n[WARNING] 데이터 누락 테이블: {validation['tables_missing_data']}")
    else:
        print(f"\n[OK] 모든 테이블에 데이터가 있습니다!")
    
    return output_file

if __name__ == "__main__":
    convert_csv_to_json_enhanced("sddddd.csv", "sddddd_enhanced.json")
