#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Content Calendar Python Application

Excel 파일(content-calendar.xlsx)을 완전한 Python 애플리케이션으로 변환
"""

import sys
import io
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Tuple
from datetime import date, datetime, timedelta
from enum import Enum
import json

# UTF-8 출력 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass


@dataclass
class ContentItem:
    """콘텐츠 항목"""
    id: Optional[str] = None
    date: Optional[date] = None
    title: Optional[str] = None
    description: Optional[str] = None
    status: Optional[str] = None
    hashtags: List[str] = field(default_factory=list)
    platform: Optional[str] = None
    url: Optional[str] = None
    
    def to_dict(self) -> Dict:
        """딕셔너리로 변환"""
        return {
            'id': self.id,
            'date': self.date.isoformat() if self.date else None,
            'title': self.title,
            'description': self.description,
            'status': self.status,
            'hashtags': self.hashtags,
            'platform': self.platform,
            'url': self.url
        }


@dataclass
class CalendarSettings:
    """캘린더 설정"""
    year: int = 2025
    month: int = 12
    start_day_of_week: int = 1  # 1=일요일, 2=월요일 등
    
    def to_dict(self) -> Dict:
        """딕셔너리로 변환"""
        return {
            'year': self.year,
            'month': self.month,
            'start_day_of_week': self.start_day_of_week
        }


class CalendarCalculator:
    """캘린더 계산 로직 (Excel 함수를 Python으로 변환)"""
    
    @staticmethod
    def get_first_day_of_month(year: int, month: int) -> date:
        """월의 첫 번째 날짜 (Excel DATE 함수)"""
        return date(year, month, 1)
    
    @staticmethod
    def get_weekday(date_val: date, return_type: int = 1) -> int:
        """
        요일 반환 (Excel WEEKDAY 함수)
        return_type=1: 1(일)~7(토)
        return_type=2: 1(월)~7(일)
        """
        weekday = date_val.weekday()  # 0=월요일, 6=일요일
        
        if return_type == 1:
            # 1(일요일) ~ 7(토요일)
            return weekday + 2 if weekday < 6 else weekday - 5
        elif return_type == 2:
            # 1(월요일) ~ 7(일요일)
            return weekday + 1
        return weekday + 1
    
    @staticmethod
    def get_calendar_start_date(base_date: date, start_day: int) -> date:
        """
        캘린더 시작 날짜 계산
        Excel: =DATE(P6,Q8,1)-(WEEKDAY(DATE(P6,Q8,1),1)-(P10-1))-IF((WEEKDAY(DATE(P6,Q8,1),1)-(P10-1))<=0,7,0)+1
        """
        first_day = CalendarCalculator.get_first_day_of_month(
            base_date.year, base_date.month
        )
        weekday = CalendarCalculator.get_weekday(first_day, 1)
        
        offset = weekday - (start_day - 1)
        if offset <= 0:
            offset += 7
        
        return first_day - timedelta(days=offset - 1)
    
    @staticmethod
    def generate_week_dates(start_date: date, week_num: int = 0) -> List[date]:
        """
        주별 날짜 목록 생성 (Excel의 M3, N3, O3... 로직)
        M3 = J2-WEEKDAY(J2,1)+2+7*(J3-1)
        """
        # 주 시작 날짜 계산
        week_start = start_date + timedelta(days=7 * week_num)
        
        # 일주일 날짜 생성
        dates = []
        for day in range(7):
            dates.append(week_start + timedelta(days=day))
        
        return dates
    
    @staticmethod
    def generate_calendar_dates(start_date: date, weeks: int = 6) -> List[date]:
        """캘린더 날짜 목록 생성"""
        dates = []
        current_date = start_date
        
        for week in range(weeks):
            for day in range(7):
                dates.append(current_date)
                current_date += timedelta(days=1)
        
        return dates


class ContentRepository:
    """콘텐츠 데이터 저장소"""
    
    def __init__(self):
        self.items: Dict[date, List[ContentItem]] = {}
        self.all_items: List[ContentItem] = []
    
    def add_item(self, item: ContentItem):
        """콘텐츠 항목 추가"""
        if item.date:
            if item.date not in self.items:
                self.items[item.date] = []
            self.items[item.date].append(item)
        self.all_items.append(item)
    
    def get_items_for_date(self, target_date: date) -> List[ContentItem]:
        """특정 날짜의 콘텐츠 조회 (Excel VLOOKUP 로직)"""
        return self.items.get(target_date, [])
    
    def get_items_for_range(self, start_date: date, end_date: date) -> List[ContentItem]:
        """날짜 범위의 콘텐츠 조회"""
        result = []
        current = start_date
        while current <= end_date:
            result.extend(self.get_items_for_date(current))
            current += timedelta(days=1)
        return result
    
    def get_all_items(self) -> List[ContentItem]:
        """모든 콘텐츠 항목 반환"""
        return self.all_items


class CalendarView:
    """캘린더 뷰 생성 (Calendar 시트 로직)"""
    
    def __init__(self, calculator: CalendarCalculator, repository: ContentRepository):
        self.calculator = calculator
        self.repository = repository
    
    def generate_month_view(self, year: int, month: int, start_day: int = 1) -> Dict:
        """월별 캘린더 뷰 생성"""
        base_date = date(year, month, 1)
        start_date = self.calculator.get_calendar_start_date(base_date, start_day)
        dates = self.calculator.generate_calendar_dates(start_date, weeks=6)
        
        # 주별로 그룹화
        weeks = []
        for week_start in range(0, len(dates), 7):
            week_dates = dates[week_start:week_start + 7]
            week_data = []
            
            for day_date in week_dates:
                items = self.repository.get_items_for_date(day_date)
                week_data.append({
                    'date': day_date.isoformat(),
                    'day': day_date.day,
                    'is_current_month': day_date.month == month,
                    'is_today': day_date == date.today(),
                    'weekday': day_date.strftime('%A'),
                    'items': [self._item_to_dict(item) for item in items],
                    'item_count': len(items)
                })
            
            weeks.append(week_data)
        
        return {
            'year': year,
            'month': month,
            'month_name': base_date.strftime('%B %Y').upper(),
            'start_date': start_date.isoformat(),
            'weeks': weeks
        }
    
    def _item_to_dict(self, item: ContentItem) -> Dict:
        """ContentItem을 딕셔너리로 변환"""
        return item.to_dict()


class ContentCalendarApplication:
    """완전한 Content Calendar 애플리케이션"""
    
    def __init__(self):
        self.calculator = CalendarCalculator()
        self.repository = ContentRepository()
        self.view = CalendarView(self.calculator, self.repository)
        self.settings = CalendarSettings(year=2025, month=12, start_day_of_week=1)
    
    def load_from_excel(self, excel_path: str):
        """Excel 파일에서 데이터 로드"""
        try:
            from excel_python_engine import ExcelWorkbook
            
            print(f"📂 Excel 파일 로드: {excel_path}")
            workbook = ExcelWorkbook.load_from_excel(excel_path)
            workbook.calculate_all()
            
            # 데이터 추출 및 변환
            self._import_data(workbook)
            print("✅ 데이터 로드 완료!")
            
        except Exception as e:
            print(f"❌ Excel 로드 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def _import_data(self, workbook):
        """Excel 데이터를 Python 객체로 변환"""
        # Settings 로드
        settings_sheet = workbook.get_sheet("Settings")
        if settings_sheet:
            year_cell = settings_sheet.get_cell("P6")
            month_cell = settings_sheet.get_cell("Q8")
            start_day_cell = settings_sheet.get_cell("P10")
            
            if year_cell:
                try:
                    year_val = year_cell.get_value()
                    if year_val:
                        self.settings.year = int(year_val)
                except:
                    pass
            
            if month_cell:
                try:
                    month_val = month_cell.get_value()
                    if month_val:
                        self.settings.month = int(month_val)
                except:
                    pass
            
            if start_day_cell:
                try:
                    start_day_val = start_day_cell.get_value()
                    if start_day_val:
                        self.settings.start_day_of_week = int(start_day_val)
                except:
                    pass
        
        # Content 시트 로드
        content_sheet = workbook.get_sheet("Content")
        if content_sheet:
            self._load_content_items(content_sheet)
        
        # Settings 시트에서 날짜별 콘텐츠 매핑 로드
        if settings_sheet:
            self._load_content_mappings(settings_sheet)
    
    def _load_content_items(self, content_sheet):
        """Content 시트에서 콘텐츠 항목 로드"""
        # Content 시트의 실제 구조에 맞게 파싱
        # 행 4부터 데이터가 시작되는 것으로 가정
        for row in range(4, min(27, content_sheet.rows + 1)):
            # 실제 Excel 구조에 맞게 컬럼 매핑 필요
            # 예시: C=날짜, D=제목, F=설명 등
            try:
                date_cell = content_sheet.get_cell(f"C{row}")
                title_cell = content_sheet.get_cell(f"D{row}")
                desc_cell = content_sheet.get_cell(f"F{row}")
                
                if date_cell or title_cell:
                    item = ContentItem()
                    item.id = str(row)
                    
                    # 날짜 파싱
                    if date_cell:
                        date_val = date_cell.get_value()
                        if date_val:
                            if isinstance(date_val, date):
                                item.date = date_val
                            elif isinstance(date_val, (int, float)):
                                # Excel 날짜 시리얼 번호 변환
                                base_date = date(1900, 1, 1)
                                item.date = base_date + timedelta(days=int(date_val) - 2)
                    
                    # 제목
                    if title_cell:
                        item.title = str(title_cell.get_value() or "")
                    
                    # 설명
                    if desc_cell:
                        item.description = str(desc_cell.get_value() or "")
                    
                    if item.date or item.title:
                        self.repository.add_item(item)
            except Exception as e:
                # 개별 행 오류는 무시하고 계속 진행
                pass
    
    def _load_content_mappings(self, settings_sheet):
        """Settings 시트에서 날짜별 콘텐츠 매핑 로드"""
        # Settings 시트의 A45:C94 범위에서 VLOOKUP 데이터 로드
        # Excel: =IFERROR(VLOOKUP(A4,Settings!$A$45:$C$94,3,FALSE),"")
        try:
            for row in range(45, min(95, settings_sheet.rows + 1)):
                date_cell = settings_sheet.get_cell(f"A{row}")
                content_cell = settings_sheet.get_cell(f"C{row}")
                
                if date_cell and content_cell:
                    date_val = date_cell.get_value()
                    content_val = content_cell.get_value()
                    
                    if date_val and content_val:
                        if isinstance(date_val, date):
                            target_date = date_val
                        elif isinstance(date_val, (int, float)):
                            base_date = date(1900, 1, 1)
                            target_date = base_date + timedelta(days=int(date_val) - 2)
                        else:
                            continue
                        
                        # 기존 항목에 추가하거나 새로 생성
                        items = self.repository.get_items_for_date(target_date)
                        if not items:
                            item = ContentItem(
                                id=f"settings_{row}",
                                date=target_date,
                                title=str(content_val),
                                description=""
                            )
                            self.repository.add_item(item)
        except Exception as e:
            pass
    
    def get_calendar_view(self) -> Dict:
        """현재 설정으로 캘린더 뷰 생성"""
        return self.view.generate_month_view(
            self.settings.year,
            self.settings.month,
            self.settings.start_day_of_week
        )
    
    def add_content_item(self, item: ContentItem):
        """콘텐츠 항목 추가"""
        self.repository.add_item(item)
    
    def export_to_json(self, output_path: str):
        """JSON으로 내보내기"""
        try:
            calendar_data = self.get_calendar_view()
            calendar_data['settings'] = self.settings.to_dict()
            calendar_data['all_content_items'] = [
                item.to_dict() for item in self.repository.get_all_items()
            ]
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(calendar_data, f, indent=2, default=str, ensure_ascii=False)
            print(f"✅ JSON 저장 완료: {output_path}")
        except OSError as e:
            print(f"⚠️ JSON 저장 실패 (디스크 공간 부족): {e}")
        except Exception as e:
            print(f"⚠️ JSON 저장 오류: {e}")
    
    def export_to_html(self, output_path: str):
        """HTML 캘린더 생성"""
        calendar_data = self.get_calendar_view()
        
        html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Content Calendar - {calendar_data['month_name']}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            text-align: center;
            color: #333;
            margin-bottom: 30px;
            font-size: 2em;
        }}
        .calendar {{
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 1px;
            background-color: #ddd;
            border: 1px solid #ddd;
        }}
        .day-header {{
            background-color: #4a90e2;
            color: white;
            padding: 15px;
            text-align: center;
            font-weight: bold;
            font-size: 0.9em;
        }}
        .day {{
            background-color: white;
            padding: 10px;
            min-height: 120px;
            border: 1px solid #ddd;
            position: relative;
        }}
        .day.other-month {{
            background-color: #f9f9f9;
            color: #999;
        }}
        .day.today {{
            background-color: #fff9e6;
            border: 2px solid #ffd700;
        }}
        .day-number {{
            font-weight: bold;
            font-size: 1.1em;
            margin-bottom: 5px;
            color: #333;
        }}
        .day.other-month .day-number {{
            color: #999;
        }}
        .day.today .day-number {{
            color: #4a90e2;
            font-size: 1.2em;
        }}
        .items {{
            margin-top: 5px;
        }}
        .item {{
            font-size: 0.85em;
            padding: 3px 5px;
            margin: 2px 0;
            background-color: #e8f4f8;
            border-left: 3px solid #4a90e2;
            border-radius: 3px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .item-count {{
            font-size: 0.75em;
            color: #666;
            margin-top: 5px;
        }}
        @media (max-width: 768px) {{
            .calendar {{
                grid-template-columns: 1fr;
            }}
            .day {{
                min-height: 80px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>{calendar_data['month_name']}</h1>
        <div class="calendar">
            <div class="day-header">일</div>
            <div class="day-header">월</div>
            <div class="day-header">화</div>
            <div class="day-header">수</div>
            <div class="day-header">목</div>
            <div class="day-header">금</div>
            <div class="day-header">토</div>
"""
        
        for week in calendar_data['weeks']:
            for day in week:
                day_date = datetime.fromisoformat(day['date']).date()
                css_class = ""
                if not day['is_current_month']:
                    css_class = "other-month"
                if day['is_today']:
                    css_class += " today"
                
                html += f"""
            <div class="day {css_class}">
                <div class="day-number">{day['day']}</div>
                <div class="items">
"""
                for item in day['items'][:3]:  # 최대 3개만 표시
                    title = item.get('title', '')[:30]
                    html += f'                    <div class="item" title="{item.get("description", "")}">{title}</div>\n'
                
                if day['item_count'] > 3:
                    html += f'                    <div class="item-count">+{day["item_count"] - 3} more</div>\n'
                
                html += """
                </div>
            </div>
"""
        
        html += """
        </div>
    </div>
</body>
</html>
"""
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html)
            print(f"✅ HTML 저장 완료: {output_path}")
        except OSError as e:
            print(f"⚠️ HTML 저장 실패 (디스크 공간 부족): {e}")
        except Exception as e:
            print(f"⚠️ HTML 저장 오류: {e}")
    
    def print_summary(self):
        """요약 정보 출력"""
        calendar_view = self.get_calendar_view()
        total_items = len(self.repository.get_all_items())
        
        print("\n" + "=" * 70)
        print("Content Calendar 요약")
        print("=" * 70)
        print(f"설정: {calendar_view['month_name']}")
        print(f"시작 날짜: {calendar_view['start_date']}")
        print(f"총 콘텐츠 항목: {total_items}개")
        print(f"주 수: {len(calendar_view['weeks'])}주")
        
        # 날짜별 항목 수
        items_by_date = {}
        for item in self.repository.get_all_items():
            if item.date:
                if item.date not in items_by_date:
                    items_by_date[item.date] = 0
                items_by_date[item.date] += 1
        
        if items_by_date:
            print(f"\n날짜별 콘텐츠:")
            for day_date, count in sorted(items_by_date.items())[:10]:
                print(f"  {day_date}: {count}개")
            if len(items_by_date) > 10:
                print(f"  ... 외 {len(items_by_date) - 10}개 날짜")
        
        print("=" * 70)


def main():
    """메인 실행 함수"""
    import sys
    from pathlib import Path
    
    print("=" * 70)
    print("Content Calendar Python Application")
    print("=" * 70)
    
    # Excel 파일 경로
    excel_path = "content-calendar_calculated.xlsx"
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    
    if not Path(excel_path).exists():
        print(f"❌ 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    # 애플리케이션 생성
    app = ContentCalendarApplication()
    
    # Excel 파일에서 로드
    app.load_from_excel(excel_path)
    
    # 요약 출력
    app.print_summary()
    
    # JSON으로 내보내기 (선택적)
    try:
        app.export_to_json("calendar_output.json")
    except:
        print("⚠️ JSON 저장 건너뜀")
    
    # HTML로 내보내기 (선택적)
    try:
        app.export_to_html("calendar_output.html")
    except:
        print("⚠️ HTML 저장 건너뜀")
    
    # 콘솔에 캘린더 뷰 미리보기 출력
    print("\n" + "=" * 70)
    print("캘린더 뷰 미리보기")
    print("=" * 70)
    calendar_view = app.get_calendar_view()
    print(f"월: {calendar_view['month_name']}")
    print(f"시작 날짜: {calendar_view['start_date']}")
    print(f"총 주 수: {len(calendar_view['weeks'])}주")
    
    # 첫 주 미리보기
    if calendar_view['weeks']:
        print("\n첫 주 미리보기:")
        first_week = calendar_view['weeks'][0]
        for day in first_week:
            status = "✓" if day['is_current_month'] else " "
            today_mark = " [오늘]" if day['is_today'] else ""
            items_mark = f" ({day['item_count']}개)" if day['item_count'] > 0 else ""
            print(f"  {status} {day['date']} ({day['weekday'][:3]}) - {day['day']}일{today_mark}{items_mark}")
    
    print("\n✅ 실행 완료!")


if __name__ == "__main__":
    main()

