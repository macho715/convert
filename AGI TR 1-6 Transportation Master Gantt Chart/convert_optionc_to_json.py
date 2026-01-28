import json
import csv
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional

def parse_tsv_to_hierarchical_json(tsv_file_path: str, json_file_path: Optional[str] = None) -> Dict:
    """
    Option C TSV 파일을 계층 구조 JSON으로 변환
    
    Args:
        tsv_file_path: 입력 TSV 파일 경로
        json_file_path: 출력 JSON 파일 경로 (None이면 자동 생성)
    
    Returns:
        변환된 JSON 데이터 딕셔너리
    """
    activities = []
    current_level1 = None
    current_level2 = None
    
    with open(tsv_file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='\t')
        
        # 헤더 스킵
        header = next(reader)
        
        for row in reader:
            # 빈 행 스킵
            if not any(row):
                continue
            
            # TSV 구조: [Activity ID, Activity ID, Activity ID, Activity Name, Original Duration, Planned Start, Planned Finish]
            # 인덱스로 직접 접근
            level1_raw = row[0].strip() if len(row) > 0 else ''
            level2_raw = row[1].strip() if len(row) > 1 else ''
            level3_raw = row[2].strip() if len(row) > 2 else ''
            activity_name = row[3].strip() if len(row) > 3 else ''
            duration = row[4].strip() if len(row) > 4 else ''
            start = row[5].strip() if len(row) > 5 else ''
            finish = row[6].strip() if len(row) > 6 else ''
            
            # Level 1 결정: 값이 있으면 사용, 없으면 이전 값 유지
            if level1_raw in ['MOBILIZATION', 'DEMOBILIZATION', 'OPERATIONAL']:
                # Level 1 카테고리가 변경되면 업데이트하고 Level 2 리셋
                current_level1 = level1_raw
                current_level2 = None
            elif level1_raw:
                # 다른 값이면 업데이트 (일반적으로는 발생하지 않음)
                current_level1 = level1_raw
            # level1_raw가 비어있으면 current_level1 유지 (이전 값)
            
            # Level 2 결정: 값이 있으면 사용, 없으면 이전 값 유지
            if level2_raw:
                if level2_raw not in ['MOBILIZATION', 'DEMOBILIZATION', 'OPERATIONAL']:
                    current_level2 = level2_raw
            # level2_raw가 비어있으면 current_level2 유지 (이전 값)
            
            # Activity 객체 생성 (현재 상태 사용)
            activity = {
                'level1': current_level1,
                'level2': current_level2,
                'activity_id': level3_raw if level3_raw else None,
                'activity_name': activity_name,
                'duration': float(duration) if duration else None,
                'planned_start': start if start else None,
                'planned_finish': finish if finish else None
            }
            
            # 빈 활동 제외
            if activity_name:
                activities.append(activity)
    
    # 날짜 범위 계산
    start_dates = [a['planned_start'] for a in activities if a['planned_start']]
    finish_dates = [a['planned_finish'] for a in activities if a['planned_finish']]
    
    # JSON 구조 생성
    result = {
        'document_metadata': {
            'title': 'AGI TR 1-6 Transportation Master Gantt Chart - Option C',
            'source_file': Path(tsv_file_path).name,
            'generated_at': datetime.now().isoformat() + 'Z',
            'total_activities': len(activities)
        },
        'activities': activities,
        'summary': {
            'mobilization_count': sum(1 for a in activities if a['level1'] == 'MOBILIZATION'),
            'demobilization_count': sum(1 for a in activities if a['level1'] == 'DEMOBILIZATION'),
            'operational_count': sum(1 for a in activities if a['level1'] == 'OPERATIONAL'),
            'total_activities': len(activities),
            'date_range': {
                'start': min(start_dates) if start_dates else None,
                'finish': max(finish_dates) if finish_dates else None
            }
        }
    }
    
    # JSON 파일 저장
    if json_file_path is None:
        json_file_path = str(Path(tsv_file_path).with_suffix('.json'))
    
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 변환 완료: {len(activities)}개 활동")
    print(f"[FILE] 출력 파일: {json_file_path}")
    
    return result

def parse_tsv_to_flat_json(tsv_file_path: str, json_file_path: Optional[str] = None) -> List[Dict]:
    """
    Option C TSV 파일을 평면 구조 JSON 배열로 변환
    
    Args:
        tsv_file_path: 입력 TSV 파일 경로
        json_file_path: 출력 JSON 파일 경로 (None이면 자동 생성)
    
    Returns:
        변환된 JSON 데이터 리스트
    """
    activities = []
    
    with open(tsv_file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='\t')
        
        # 헤더 스킵
        header = next(reader)
        
        for row in reader:
            # 빈 행 스킵
            if not any(row):
                continue
            
            # TSV 구조: [Activity ID, Activity ID, Activity ID, Activity Name, Original Duration, Planned Start, Planned Finish]
            activity = {
                'activity_id_level1': row[0].strip() if len(row) > 0 and row[0].strip() else None,
                'activity_id_level2': row[1].strip() if len(row) > 1 and row[1].strip() else None,
                'activity_id_level3': row[2].strip() if len(row) > 2 and row[2].strip() else None,
                'activity_name': row[3].strip() if len(row) > 3 and row[3].strip() else None,
                'original_duration': float(row[4]) if len(row) > 4 and row[4].strip() else None,
                'planned_start': row[5].strip() if len(row) > 5 and row[5].strip() else None,
                'planned_finish': row[6].strip() if len(row) > 6 and row[6].strip() else None
            }
            
            # 빈 활동 제외
            if activity['activity_name']:
                activities.append(activity)
    
    # JSON 파일 저장
    if json_file_path is None:
        json_file_path = str(Path(tsv_file_path).with_suffix('_flat.json'))
    
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(activities, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 변환 완료: {len(activities)}개 활동")
    print(f"[FILE] 출력 파일: {json_file_path}")
    
    return activities

# 실행
if __name__ == "__main__":
    base_dir = Path(__file__).parent
    input_file = base_dir / "option_c.tsv"
    
    # 계층 구조 JSON 변환
    print("=" * 60)
    print("계층 구조 JSON 변환")
    print("=" * 60)
    hierarchical_data = parse_tsv_to_hierarchical_json(
        str(input_file),
        str(base_dir / "option_c.json")
    )
    
    # 평면 구조 JSON 변환
    print("\n" + "=" * 60)
    print("평면 구조 JSON 변환")
    print("=" * 60)
    flat_data = parse_tsv_to_flat_json(
        str(input_file),
        str(base_dir / "option_c_flat.json")
    )
    
    # 통계 출력
    print("\n" + "=" * 60)
    print("[STATS] 변환 통계")
    print("=" * 60)
    print(f"   - 총 활동 수: {hierarchical_data['document_metadata']['total_activities']}")
    print(f"   - Mobilization: {hierarchical_data['summary']['mobilization_count']}")
    print(f"   - Demobilization: {hierarchical_data['summary']['demobilization_count']}")
    print(f"   - Operational: {hierarchical_data['summary']['operational_count']}")
    
    if hierarchical_data['summary']['date_range']['start']:
        print(f"   - 날짜 범위: {hierarchical_data['summary']['date_range']['start']} ~ {hierarchical_data['summary']['date_range']['finish']}")
    
    # 샘플 출력
    print("\n" + "=" * 60)
    print("[SAMPLE] 샘플 데이터 (처음 3개)")
    print("=" * 60)
    print(json.dumps(hierarchical_data['activities'][:3], ensure_ascii=False, indent=2))
