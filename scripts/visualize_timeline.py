import json
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import numpy as np

# 주요 이벤트 데이터
events = [
    {"date": "2025-09-19", "id": "msg-13", "summary": "프로젝트 시작: 217t 변압기, 12축, 12m 링크스판", "type": "init"},
    {"date": "2025-09-22", "id": "msg-14/15", "summary": "Aries 응답 및 범위 제안", "type": "response"},
    {"date": "2025-09-30", "id": "msg-20", "summary": "Aries 필요 도면 목록 요청", "type": "request"},
    {"date": "2025-10-10", "id": "msg-6", "summary": "Aries 하중 케이스 제출 (201.6t)", "type": "submission"},
    {"date": "2025-10-14", "id": "msg-8/9", "summary": "수정된 하중 계산 및 bow 구조 계산 진행", "type": "calculation"},
    {"date": "2025-10-21", "id": "msg-24", "summary": "책임 명확화: Mammoet 방법론, Aries 안정성 체크", "type": "clarification"},
    {"date": "2025-10-24", "id": "msg-25", "summary": "Aries 상세 범위 설명 (7-10일 소요)", "type": "scope"},
    {"date": "2025-10-25", "id": "msg-30", "summary": "✓ Bow deck 201t 적합성 확인", "type": "approval"},
    {"date": "2025-10-26", "id": "msg-26/28", "summary": "✓ Samsung ELC 요청, LCT Bushra 데이터 제공 완료", "type": "completion"},
    {"date": "2025-10-27", "id": "msg-31", "summary": "Mammoet 12m linkspan 동원 승인 요청", "type": "request"},
    {"date": "2025-10-27", "id": "msg-32", "summary": "범위 오해 명확화: bow deck 확인, link span 분석 진행 중", "type": "clarification"},
    {"date": "2025-10-28", "id": "msg-33", "summary": "✓✓ Samsung 최종 승인: 12m linkspan import", "type": "final_approval"}
]

# 날짜 파싱
dates = [datetime.strptime(e["date"], "%Y-%m-%d") for e in events]
y_positions = np.arange(len(events))

# 타입별 색상
type_colors = {
    "init": "#9C27B0",  # 보라 (시작)
    "response": "#2196F3",  # 파랑 (응답)
    "request": "#FF9800",  # 주황 (요청)
    "submission": "#00BCD4",  # 청록 (제출)
    "calculation": "#03A9F4",  # 하늘색 (계산)
    "clarification": "#FFC107",  # 황금색 (명확화)
    "scope": "#FF5722",  # 빨강-주황 (범위)
    "approval": "#4CAF50",  # 초록 (승인)
    "completion": "#8BC34A",  # 연두 (완료)
    "final_approval": "#1B5E20"  # 진한 초록 (최종 승인)
}

colors = [type_colors[e["type"]] for e in events]

# 그래프 생성
plt.figure(figsize=(16, 10))
plt.rcParams['font.family'] = 'Malgun Gothic'  # 한글 폰트
plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

# 타임라인 플롯
plt.scatter(dates, y_positions, c=colors, s=300, alpha=0.8, zorder=3, edgecolors='black', linewidths=1.5)

# 연결선
for i in range(len(dates)-1):
    plt.plot([dates[i], dates[i+1]], [y_positions[i], y_positions[i+1]], 
             'k--', alpha=0.3, linewidth=1)

# 각 이벤트에 라벨 추가
for i, event in enumerate(events):
    plt.text(dates[i], y_positions[i], f'  [{event["id"]}] {event["summary"]}', 
            va='center', ha='left', fontsize=10, 
            bbox=dict(boxstyle='round,pad=0.6', facecolor='white', alpha=0.9, edgecolor=colors[i], linewidth=2))

# 주요 마일스톤 하이라이트
approval_date = datetime.strptime("2025-10-25", "%Y-%m-%d")
final_approval_date = datetime.strptime("2025-10-28", "%Y-%m-%d")

plt.axvline(approval_date, color='green', linestyle=':', linewidth=2, alpha=0.4, label='Bow Deck 승인')
plt.axvline(final_approval_date, color='darkgreen', linestyle='-', linewidth=3, alpha=0.5, label='최종 승인')

# 레이아웃 설정
ax = plt.gca()
ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
plt.xticks(rotation=45, ha='right', fontsize=10)

ax.set_yticks([])
ax.set_xlabel('일자 (2025년)', fontsize=13, fontweight='bold')
ax.set_title('HVDC 프로젝트 타임라인: AGI 변압기 운송\n선박 안정성 및 데크 강도 분석 (9/19 - 10/28, 2025)', 
             fontsize=15, fontweight='bold', pad=20)

# 범례
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor='#9C27B0', edgecolor='black', label='프로젝트 시작'),
    Patch(facecolor='#2196F3', edgecolor='black', label='응답/제출'),
    Patch(facecolor='#FF9800', edgecolor='black', label='요청'),
    Patch(facecolor='#FFC107', edgecolor='black', label='명확화'),
    Patch(facecolor='#4CAF50', edgecolor='black', label='승인'),
    Patch(facecolor='#1B5E20', edgecolor='black', label='최종 승인')
]
ax.legend(handles=legend_elements, loc='upper left', fontsize=11, framealpha=0.9)

# 그리드
ax.grid(True, axis='x', alpha=0.3, linestyle='--')
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_visible(False)

# 기간 표시
plt.text(0.02, 0.98, f'프로젝트 기간: 40일 (9/19 - 10/28)', 
         transform=ax.transAxes, fontsize=11, verticalalignment='top',
         bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))

plt.tight_layout()
plt.savefig('project_timeline.png', dpi=300, bbox_inches='tight', facecolor='white')
print('✓ 타임라인 시각화 완료: project_timeline.png')
print(f'✓ 총 이벤트: {len(events)}개')
print(f'✓ 프로젝트 기간: 40일')
print(f'✓ 주요 마일스톤: 2개 (Bow Deck 승인, 최종 승인)')



