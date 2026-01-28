import json
from datetime import datetime

def sort_json_by_date(json_file_path):
    """
    JSON 파일을 Loading Date 기준으로 정렬
    
    Args:
        json_file_path: JSON 파일 경로
    """
    # JSON 파일 읽기
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    print(f"[INFO] 총 레코드 수: {len(data)}개")
    
    # 날짜 파싱 및 정렬 함수
    def parse_date(date_str):
        """날짜 문자열을 datetime 객체로 변환"""
        if not date_str or date_str.strip() == "":
            # 빈 날짜는 가장 뒤로
            return datetime(9999, 12, 31)
        try:
            return datetime.strptime(date_str.strip(), "%Y-%m-%d")
        except (ValueError, AttributeError):
            # 파싱 실패 시 가장 뒤로
            return datetime(9999, 12, 31)
    
    # 날짜 기준으로 정렬 (같은 날짜면 Batch 번호로 정렬)
    def sort_key(record):
        date = parse_date(record.get("Loading Date", ""))
        batch = record.get("Batch", "")
        
        # Batch 번호를 숫자로 변환 시도
        try:
            if batch and batch.strip():
                # 숫자로 시작하는 경우
                if batch[0].isdigit():
                    batch_num = int(batch) if batch.isdigit() else int(batch.rstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
                else:
                    # 문자로 시작하는 경우 (A1, D1, E1 등)
                    batch_num = ord(batch[0]) * 1000 + (int(batch[1:]) if len(batch) > 1 and batch[1:].isdigit() else 0)
            else:
                batch_num = 999999
        except:
            batch_num = 999999
        
        return (date, batch_num)
    
    # 정렬 실행
    sorted_data = sorted(data, key=sort_key)
    
    # 정렬 결과 확인
    print(f"\n[INFO] 정렬 완료")
    print(f"[INFO] 첫 번째 레코드 날짜: {sorted_data[0].get('Loading Date', 'N/A')}")
    print(f"[INFO] 마지막 레코드 날짜: {sorted_data[-1].get('Loading Date', 'N/A')}")
    
    # JSON 파일 저장
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(sorted_data, f, ensure_ascii=False, indent=2)
    
    print(f"[OK] 파일 저장 완료: {json_file_path}")
    
    # 날짜별 통계
    date_counts = {}
    for record in sorted_data:
        date = record.get("Loading Date", "N/A")
        date_counts[date] = date_counts.get(date, 0) + 1
    
    print(f"\n[STATS] 날짜별 레코드 수 (처음 10개):")
    for i, (date, count) in enumerate(sorted(date_counts.items())[:10]):
        print(f"  {date}: {count}개")
    
    return sorted_data

if __name__ == "__main__":
    # 두 파일 모두 정렬
    files_to_sort = [
        "Untitled-1.json",
        r"c:\Users\SAMSUNG\Desktop\OFCO_GRM.JSON"
    ]
    
    for json_file in files_to_sort:
        try:
            print(f"\n{'='*60}")
            print(f"[FILE] 처리 중: {json_file}")
            print(f"{'='*60}")
            sort_json_by_date(json_file)
        except FileNotFoundError:
            print(f"[SKIP] 파일 없음: {json_file}")
        except Exception as e:
            print(f"[ERROR] 오류 발생: {e}")

