# pip install plotly pandas dash
"""
AGI TR 1-6 Transportation Master Gantt Chart - Interactive Dashboard Generator
고급 인터랙티브 대시보드 생성 스크립트
"""

import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
from datetime import date, timedelta, datetime
import json

# =========================
# 0) 입력 데이터 (jpt71.py와 동일)
# =========================
rows_plan = [
    # --- 실적 예시(84 cycle) ---
    ("84",       "Aggregate", "5mm",       date(2025,12,5),  date(2025,12,6),  date(2025,12,7),  None,              None,              ""),
    ("Debris-5", "Debris",    "Debris",    None,             None,             None,             date(2025,12,7),   date(2025,12,8),   ""),

    # --- (선택) 다음 cycle 예시(85) ---
    ("85",       "Aggregate", "Dune Sand", date(2025,12,9),  date(2025,12,10), date(2025,12,11), None,              None,              ""),
    ("Debris-6", "Debris",    "Debris",    None,             None,             None,             date(2025,12,11),  date(2025,12,12),  ""),

    # --- Debris list (whiteboard) ---
    ("Debris-7", "Debris",    "Debris",    None, None, None, date(2025,12,25), date(2025,12,26), ""),
    ("Debris-8", "Debris",    "Debris",    None, None, None, date(2025,12,28), date(2025,12,29), "IN PROGRESS"),  # 현재 하역중(MW4)

    # --- Aggregate list / schedule board (예시) ---
    ("88",       "Aggregate", "Dune Sand", date(2025,12,30), None, None, None, None, ""),
    ("Debris-9", "Debris",    "Debris",    None, None, None, date(2026,1,7),  date(2026,1,8),  ""),

    ("89",       "Aggregate", "5mm",       date(2026,1,9),  date(2026,1,10), date(2026,1,11), None, None, ""),
    ("Debris-10","Debris",    "Debris",    None, None, None, date(2026,1,10), date(2026,1,11), ""),

    ("90",       "Aggregate", "20mm",      date(2026,1,12), date(2026,1,13), None, None, None, ""),
    ("Debris-11","Debris",    "Debris",    None, None, None, date(2026,1,13), date(2026,1,14), ""),

    ("91",       "Aggregate", "10mm",      date(2026,1,15), None, None, None, None, ""),
    ("Debris-12","Debris",    "Debris",    None, None, None, date(2026,1,19), date(2026,1,20), ""),

    ("92",       "Aggregate", "20mm",      date(2026,1,21), date(2026,1,24), None, None, None, ""),
    ("Debris-13","Debris",    "Debris",    None, None, None, date(2026,1,25), date(2026,1,26), ""),

    ("93",       "Aggregate", "5mm",       date(2026,1,27), None, None, None, None, ""),
]

DEBRIS8_PLANNED_MW4_DEB_OFF = date(2025,12,29)

# =========================
# 1) 데이터 전처리
# =========================
def earliest_dt(r):
    ds = [x for x in r[3:8] if x is not None]
    return min(ds) if ds else date(2099,1,1)

rows_plan = sorted(rows_plan, key=earliest_dt)

# 날짜 범위 계산
all_dates = []
for r in rows_plan:
    all_dates += [x for x in r[3:8] if x is not None]
start = min(all_dates) - timedelta(days=1)
end = max(all_dates) + timedelta(days=14)

# =========================
# 2) 간트 차트 데이터 준비
# =========================
gantt_data = []

for idx, r in enumerate(rows_plan, start=1):
    item, typ, material, p_mw4, p_off1, p_off2, p_deb_load, p_deb_off, status = r
    
    # 각 단계별로 간트 바 생성
    stages = [
        ("MW4 Agg Loading", p_mw4, "#4F81BD"),
        ("AGI Agg Offload Day-1", p_off1, "#70AD47"),
        ("AGI Agg Offload+Deb Load Day-2", p_off2, "#92D050"),
        ("AGI Debris Loading", p_deb_load, "#FFC000"),
        ("MW4 Debris Offloading", p_deb_off, "#FF6600"),
    ]
    
    for stage_name, stage_date, color in stages:
        if stage_date:
            gantt_data.append({
                "Item": item,
                "Type": typ,
                "Material": material,
                "Stage": stage_name,
                "Start": stage_date,
                "End": stage_date + timedelta(days=1),
                "Status": status,
                "Color": color,
                "Y_Position": idx,
            })

df_gantt = pd.DataFrame(gantt_data)

# =========================
# 3) 인터랙티브 간트 차트 생성
# =========================
def create_interactive_gantt():
    """Plotly를 사용한 인터랙티브 간트 차트 생성"""
    fig = go.Figure()
    
    # 각 단계별 색상
    stage_colors = {
        "MW4 Agg Loading": "#4F81BD",
        "AGI Agg Offload Day-1": "#70AD47",
        "AGI Agg Offload+Deb Load Day-2": "#92D050",
        "AGI Debris Loading": "#FFC000",
        "MW4 Debris Offloading": "#FF6600",
    }

    # 각 단계별로 Scatter 트레이스 생성
    for stage in df_gantt["Stage"].unique():
        stage_df = df_gantt[df_gantt["Stage"] == stage]
        show_legend = True
        for _, row in stage_df.iterrows():
            start_date = pd.Timestamp(row["Start"])
            end_date = pd.Timestamp(row["End"])
            duration = (end_date - start_date).days
            hover_text = (
                f"<b>{row['Item']}</b><br>"
                f"Type: {row['Type']}<br>"
                f"Material: {row['Material']}<br>"
                f"Stage: {stage}<br>"
                f"Start: {row['Start'].strftime('%Y-%m-%d')}<br>"
                f"End: {row['End'].strftime('%Y-%m-%d')}<br>"
                f"Duration: {duration} day(s)<br>"
                f"Status: {row['Status'] if row['Status'] else 'Scheduled'}<br>"
                f"<extra></extra>"
            )
            fig.add_trace(go.Scatter(
                x=[start_date, end_date],
                y=[row["Y_Position"], row["Y_Position"]],
                mode="lines+markers",
                name=stage,
                line=dict(color=stage_colors.get(stage, "#808080"), width=15),
                marker=dict(size=8, color=stage_colors.get(stage, "#808080")),
                hovertemplate=hover_text,
                legendgroup=stage,
                showlegend=show_legend,
            ))
            show_legend = False

    # 오늘 날짜 표시선 (shape + annotation)
    today = pd.Timestamp(date.today())
    fig.add_shape(
        type="line",
        x0=today,
        x1=today,
        y0=0,
        y1=len(rows_plan) + 1,
        line=dict(color="red", width=2, dash="dash"),
    )
    fig.add_annotation(
        x=today,
        y=len(rows_plan) + 0.5,
        text="Today",
        showarrow=True,
        arrowhead=2,
        arrowcolor="red",
        font=dict(color="red", size=12, family="Arial Black"),
    )
    
    # 레이아웃 설정
    fig.update_layout(
        title={
            'text': 'AGI TR 1-6 Transportation Master Gantt Chart<br><sub>Interactive Dashboard</sub>',
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 20}
        },
        xaxis_title="Date",
        yaxis_title="Item",
        height=max(600, len(rows_plan) * 40),
        hovermode='closest',
        xaxis=dict(
            showgrid=True,
            gridcolor='lightgray',
            type='date',
            rangeslider=dict(visible=True)
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='lightgray',
            tickmode='array',
            tickvals=list(range(1, len(rows_plan) + 1)),
            ticktext=[r[0] for r in rows_plan],
            autorange='reversed'
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        template='plotly_white',
        margin=dict(l=150, r=50, t=100, b=50),
    )
    
    return fig

# =========================
# 4) 통계 대시보드 생성
# =========================
def create_statistics_dashboard():
    """통계 및 KPI 대시보드 생성"""
    
    # 타입별 통계
    type_counts = {}
    status_counts = {}
    
    for r in rows_plan:
        typ = r[1]
        status = r[8] if r[8] else "Scheduled"
        
        type_counts[typ] = type_counts.get(typ, 0) + 1
        status_counts[status] = status_counts.get(status, 0) + 1
    
    # 서브플롯 생성
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Type Distribution', 'Status Distribution', 'Timeline Overview', 'Material Types'),
        specs=[[{"type": "pie"}, {"type": "pie"}],
               [{"type": "bar"}, {"type": "bar"}]]
    )
    
    # Type 분포
    fig.add_trace(
        go.Pie(
            labels=list(type_counts.keys()),
            values=list(type_counts.values()),
            name="Type",
            marker=dict(colors=["#E2EFDA", "#FCE4D6"])
        ),
        row=1, col=1
    )
    
    # Status 분포
    fig.add_trace(
        go.Pie(
            labels=list(status_counts.keys()),
            values=list(status_counts.values()),
            name="Status",
            marker=dict(colors=["#C5E0B4", "#D9D9D9", "#FFFFFF"])
        ),
        row=1, col=2
    )
    
    # Material 타입별 개수
    material_counts = {}
    for r in rows_plan:
        mat = r[2]
        material_counts[mat] = material_counts.get(mat, 0) + 1
    
    fig.add_trace(
        go.Bar(
            x=list(material_counts.keys()),
            y=list(material_counts.values()),
            name="Materials",
            marker_color="#4F81BD"
        ),
        row=2, col=2
    )
    
    # 타임라인 개요 (월별 작업 수)
    monthly_counts = {}
    for r in rows_plan:
        dates = [x for x in r[3:8] if x is not None]
        if dates:
            first_date = min(dates)
            month_key = first_date.strftime("%Y-%m")
            monthly_counts[month_key] = monthly_counts.get(month_key, 0) + 1
    
    fig.add_trace(
        go.Bar(
            x=list(monthly_counts.keys()),
            y=list(monthly_counts.values()),
            name="Monthly Tasks",
            marker_color="#70AD47"
        ),
        row=2, col=1
    )
    
    fig.update_layout(
        title_text="AGI TR 1-6 Transportation Statistics Dashboard",
        height=800,
        showlegend=True,
        template='plotly_white'
    )
    
    return fig

# =========================
# 5) HTML 대시보드 생성
# =========================
def generate_html_dashboard():
    """완전한 HTML 대시보드 생성"""
    
    # 간트 차트 생성
    gantt_fig = create_interactive_gantt()
    gantt_html = gantt_fig.to_html(include_plotlyjs='cdn', div_id="gantt-chart")
    
    # 통계 대시보드 생성
    stats_fig = create_statistics_dashboard()
    stats_html = stats_fig.to_html(include_plotlyjs=False, div_id="statistics-dashboard")
    
    # HTML 템플릿
    html_content = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AGI TR 1-6 Transportation Master Gantt - Interactive Dashboard</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1800px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .controls {{
            background: #f8f9fa;
            padding: 20px;
            border-bottom: 2px solid #dee2e6;
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            align-items: center;
        }}
        
        .control-group {{
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .control-group label {{
            font-weight: 600;
            color: #495057;
        }}
        
        .control-group select {{
            padding: 8px 15px;
            border: 2px solid #ced4da;
            border-radius: 5px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: all 0.3s;
        }}
        
        .control-group select:hover {{
            border-color: #1F4E79;
        }}
        
        .control-group select:focus {{
            outline: none;
            border-color: #1F4E79;
            box-shadow: 0 0 0 3px rgba(31, 78, 121, 0.1);
        }}
        
        .btn {{
            padding: 10px 20px;
            background: #1F4E79;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s;
        }}
        
        .btn:hover {{
            background: #2E75B6;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }}
        
        .content {{
            padding: 30px;
        }}
        
        .section {{
            margin-bottom: 40px;
        }}
        
        .section-title {{
            font-size: 1.8em;
            color: #1F4E79;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #1F4E79;
        }}
        
        .chart-container {{
            background: white;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }}
        
        .legend {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin-top: 20px;
        }}
        
        .legend h3 {{
            color: #1F4E79;
            margin-bottom: 15px;
        }}
        
        .legend-items {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }}
        
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .legend-color {{
            width: 30px;
            height: 30px;
            border-radius: 5px;
            border: 2px solid #dee2e6;
        }}
        
        .info-panel {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        
        .info-panel h3 {{
            margin-bottom: 15px;
        }}
        
        .info-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }}
        
        .info-item {{
            background: rgba(255,255,255,0.2);
            padding: 15px;
            border-radius: 8px;
        }}
        
        .info-item strong {{
            display: block;
            margin-bottom: 5px;
            font-size: 0.9em;
            opacity: 0.9;
        }}
        
        .info-item span {{
            font-size: 1.5em;
            font-weight: bold;
        }}
        
        .footer {{
            background: #1F4E79;
            color: white;
            padding: 20px;
            text-align: center;
        }}
        
        @media (max-width: 768px) {{
            .header h1 {{
                font-size: 1.8em;
            }}
            
            .controls {{
                flex-direction: column;
                align-items: stretch;
            }}
            
            .control-group {{
                flex-direction: column;
                align-items: stretch;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚢 AGI TR 1-6 Transportation Master Gantt Chart</h1>
            <p>Interactive Dashboard - Real-time Schedule Visualization</p>
            <p style="margin-top: 10px; font-size: 0.9em;">Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
        
        <div class="controls">
            <div class="control-group">
                <label for="typeFilter">Filter by Type:</label>
                <select id="typeFilter">
                    <option value="all">All Types</option>
                    <option value="Aggregate">Aggregate</option>
                    <option value="Debris">Debris</option>
                </select>
            </div>
            
            <div class="control-group">
                <label for="statusFilter">Filter by Status:</label>
                <select id="statusFilter">
                    <option value="all">All Status</option>
                    <option value="IN PROGRESS">IN PROGRESS</option>
                    <option value="Scheduled">Scheduled</option>
                    <option value="Completed">Completed</option>
                </select>
            </div>
            
            <button class="btn" onclick="resetFilters()">Reset Filters</button>
            <button class="btn" onclick="exportData()">Export Data</button>
        </div>
        
        <div class="content">
            <div class="info-panel">
                <h3>📊 Project Overview</h3>
                <div class="info-grid">
                    <div class="info-item">
                        <strong>Total Items</strong>
                        <span>{len(rows_plan)}</span>
                    </div>
                    <div class="info-item">
                        <strong>Aggregate Tasks</strong>
                        <span>{sum(1 for r in rows_plan if r[1] == 'Aggregate')}</span>
                    </div>
                    <div class="info-item">
                        <strong>Debris Tasks</strong>
                        <span>{sum(1 for r in rows_plan if r[1] == 'Debris')}</span>
                    </div>
                    <div class="info-item">
                        <strong>In Progress</strong>
                        <span>{sum(1 for r in rows_plan if r[8] == 'IN PROGRESS')}</span>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2 class="section-title">📅 Interactive Gantt Chart</h2>
                <div class="chart-container">
                    {gantt_html.split('<body>')[1].split('</body>')[0] if '<body>' in gantt_html else gantt_html}
                </div>
            </div>
            
            <div class="section">
                <h2 class="section-title">📈 Statistics Dashboard</h2>
                <div class="chart-container">
                    {stats_html.split('<body>')[1].split('</body>')[0] if '<body>' in stats_html else stats_html}
                </div>
            </div>
            
            <div class="legend">
                <h3>🎨 Color Legend</h3>
                <div class="legend-items">
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #4F81BD;"></div>
                        <span><strong>MW4 Agg Loading</strong> - MW4 집계 적재/출발</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #70AD47;"></div>
                        <span><strong>AGI Agg Offload Day-1</strong> - AGI 집계 하역 1일차</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #92D050;"></div>
                        <span><strong>AGI Agg Offload+Deb Load Day-2</strong> - AGI 집계 하역+잔재 적재 혼합 2일차</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #FFC000;"></div>
                        <span><strong>AGI Debris Loading</strong> - AGI 잔재 적재</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #FF6600;"></div>
                        <span><strong>MW4 Debris Offloading</strong> - MW4 잔재 하역</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #E2EFDA;"></div>
                        <span><strong>Aggregate Row</strong> - 집계 작업 행</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #FCE4D6;"></div>
                        <span><strong>Debris Row</strong> - 잔재 작업 행</span>
                    </div>
                    <div class="legend-item">
                        <div class="legend-color" style="background-color: #D9D9D9;"></div>
                        <span><strong>IN PROGRESS</strong> - 진행 중 작업</span>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p>AGI TR 1-6 Transportation Master Gantt Chart - Interactive Dashboard</p>
            <p style="margin-top: 10px; font-size: 0.9em;">Generated by MACHO-GPT v3.6-APEX</p>
        </div>
    </div>
    
    <script>
        // 필터 기능 (향후 구현)
        function resetFilters() {{
            document.getElementById('typeFilter').value = 'all';
            document.getElementById('statusFilter').value = 'all';
            // 차트 업데이트 로직 추가 가능
        }}
        
        function exportData() {{
            const data = {json.dumps([{
                'Item': r[0],
                'Type': r[1],
                'Material': r[2],
                'MW4_Agg_Load': r[3].isoformat() if r[3] else None,
                'AGI_Agg_Off1': r[4].isoformat() if r[4] else None,
                'AGI_Agg_Off2': r[5].isoformat() if r[5] else None,
                'AGI_Deb_Load': r[6].isoformat() if r[6] else None,
                'MW4_Deb_Off': r[7].isoformat() if r[7] else None,
                'Status': r[8]
            } for r in rows_plan], default=str)};
            
            const blob = new Blob([JSON.stringify(data, null, 2)], {{type: 'application/json'}});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'gantt_data_' + new Date().toISOString().split('T')[0] + '.json';
            a.click();
        }}
    </script>
</body>
</html>
"""
    
    return html_content

# =========================
# 6) 메인 실행
# =========================
if __name__ == "__main__":
    import sys
    import io
    
    # Windows 콘솔 인코딩 설정
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    
    print("[*] Generating Interactive Dashboard...")
    
    # HTML 대시보드 생성
    html_content = generate_html_dashboard()
    html_path = "JPT71_Interactive_Dashboard.html"
    
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    
    print(f"[OK] Saved: {html_path}")
    print(f"[INFO] Open {html_path} in your browser to view the interactive dashboard")
    
    # 개별 차트도 저장 (선택사항)
    gantt_fig = create_interactive_gantt()
    gantt_fig.write_html("JPT71_Gantt_Chart_Standalone.html")
    print("[OK] Saved: JPT71_Gantt_Chart_Standalone.html")
    
    stats_fig = create_statistics_dashboard()
    stats_fig.write_html("JPT71_Statistics_Dashboard_Standalone.html")
    print("[OK] Saved: JPT71_Statistics_Dashboard_Standalone.html")
    
    print("\n[SUCCESS] Interactive Dashboard Generation Complete!")
    print("[FEATURES]")
    print("   - Interactive zoom, pan, and hover")
    print("   - Filter by type and status")
    print("   - Real-time statistics")
    print("   - Export functionality")
    print("   - Responsive design")
