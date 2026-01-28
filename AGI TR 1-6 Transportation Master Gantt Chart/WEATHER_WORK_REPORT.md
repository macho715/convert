# AGI TR Weather 파이프라인 및 스케줄 HTML 적용 작업 보고

**작업 기간**: 2026-01-28  
**범위**: PDF 파서 → WEATHER.PY/WEATHER_DASHBOARD.py → 스케줄 HTML 히트맵 삽입, Base64 임베딩

---

## 1. PDF 파서 → WEATHER.PY 파이프라인 구축

### 1.1 배경
- PDF 파서(`weather_parse.py`) 출력(txt)이 WEATHER.PY 입력으로 연결되어 있지 않았음.
- ADNOC PDF(28·29·30·31 Jan, 파고 6 ft 등)를 파싱해 4일 히트맵에 반영할 필요.

### 1.2 구현 내용

| 단계 | 스크립트/경로 | 입·출력 |
|------|----------------|--------|
| **1** | `scripts/weather_parse.py` | 입력: `weather/YYYYMMDD/` (PDF·JPG) → 출력: `out/weather_parsed/YYYYMMDD/*.txt` |
| **2** | `scripts/parsed_to_weather_json.py` | 입력: `out/weather_parsed/YYYYMMDD/*.txt` → 출력: `out/weather_parsed/YYYYMMDD/weather_for_weather_py.json` |
| **3** | `WEATHER.PY` (4일 모드) | 입력: 위 JSON(존재 시 자동 사용) → 출력: `out/weather_4day_heatmap.png` |

- **parsed_to_weather_json.py**: ADNOC 텍스트에서 `VALID FROM 28/01`, `WAVE H. 2 - 3 / 4 FT`, `OUTLOOK 29/01`, `THU/FRI WAVE H. 2 - 4 / 6 FT` 등 파싱 → 날짜·파고(ft)·풍속(kt) 추출, 30·31일 6 ft(운행 어려움) 반영, `risk_level` HIGH, `wave_max_m` 1.83.
- **WEATHER.PY**: `SCHEDULE_4DAY_MODE = True` 시 `date.today()` 기준 4일(D~D+3), 파이프라인 JSON 경로 `out/weather_parsed/<YYYYMMDD>/weather_for_weather_py.json` 존재 시 해당 파일을 `WEATHER_JSON_PATH`로 사용.

### 1.3 실행 순서(표준)
```text
1) python scripts/weather_parse.py "AGI TR 1-6.../weather/YYYYMMDD" --out out/weather_parsed/YYYYMMDD
2) python scripts/parsed_to_weather_json.py out/weather_parsed/YYYYMMDD
3) (Gantt 폴더에서) python WEATHER.PY
```

---

## 2. WEATHER.PY: 투명 PNG 및 4일 출력

- **Figure/axes**: `fig.patch.set_facecolor("none")`, `fig.patch.set_alpha(0)`, 각 `ax.set_facecolor("none")`, `ax.patch.set_alpha(0)`.
- **저장**: `plt.savefig(..., facecolor="none", transparent=True)` → PNG 배경 투명.
- **4일 모드**: `START_DATE = date.today()`, `END_DATE = today + 3일`, `OUTPUT_PATH = <SCRIPT_DIR>/out/weather_4day_heatmap.png`.

---

## 3. 스케줄 HTML에 히트맵 삽입

### 3.1 대상 파일 및 이미지
- **AGI TR Unit 1 Schedule_20260126.html**, **AGI TR Unit 1 Schedule_20260128.html**  
  - 이미지: `out/weather_4day_heatmap.png` (이후 Base64로 대체).
- **files/AGI20Unit20Schedule_20260126_redesigned_v2.html**  
  - 이미지: `weather_4day_heatmap_dashboard.png` (이후 Base64로 대체).

### 3.2 삽입 방식
- Weather & Marine Risk Update 블록 내, 날짜별 문단 아래에 히트맵 삽입.
- **래퍼 div**: `background: rgba(17, 24, 39, 0.5);` `border-radius: 8px;` `padding: 8px;` `border: 1px solid var(--border-subtle);`
- **img**: `max-width:100%; height:auto; border-radius:8px; display:block;`

---

## 4. WEATHER_DASHBOARD.py (다크 테마 4일 히트맵)

### 4.1 경로·데이터
- **위치**: `AGI TR 1-6 Transportation Master Gantt Chart/files/WEATHER_DASHBOARD.py`
- **출력**: `files/out/weather_4day_heatmap.png` (사용 시 `files/weather_4day_heatmap_dashboard.png`로 복사해 HTML에서 참조).

### 4.2 적용 사항
- **투명도 80%**: `fig.patch.set_alpha(0.2)`, 각 `ax.patch.set_alpha(0.2)`.
- **Weather Analysis Summary 박스**: `stats_box.patch.set_alpha(0.1)` (90% 투명).
- **Data Coverage 박스**: `cov_box.patch.set_alpha(0.1)` (90% 투명).
- **상단 제목 삭제**: `ax1.set_title("AGI TR Transportation - Weather Risk Heatmap (Jan-Feb 2026, Multi-Source)", ...)` 제거.
- **저장**: `plt.savefig(..., facecolor="none", transparent=True)`.
- **데이터 경로**: 스크립트가 `files/`에 있으므로 `_convert_root = os.path.dirname(os.path.dirname(SCRIPT_DIR))`로 CONVERT 루트 지정, 파이프라인 JSON 사용. 기본 JSON은 상위 폴더(Gantt Chart)의 `weather_data_20260106.json` 사용.
- **출력**: Windows cp949 대비 `✅` → `[OK]` 로 변경.

---

## 5. Base64 이미지 임베딩 (모바일/경로 독립)

### 5.1 목적
- 상대 경로(`out/weather_4day_heatmap.png`, `weather_4day_heatmap_dashboard.png`)가 모바일·다른 환경에서 표시되지 않는 문제 해결.
- 파일 경로 의존 없이 HTML 단일 파일로 이미지 표시.

### 5.2 구현
- **스크립트**: `files/embed_heatmap_base64.py`
  - `out/weather_4day_heatmap.png` → Base64 인코딩 → Schedule_20260126.html, Schedule_20260128.html의 `src="out/weather_4day_heatmap.png"` 를 `src="data:image/png;base64,..."` 로 치환.
  - `files/weather_4day_heatmap_dashboard.png` → Base64 인코딩 → `files/AGI20Unit20Schedule_20260126_redesigned_v2.html`의 `src="weather_4day_heatmap_dashboard.png"` 를 `src="data:image/png;base64,..."` 로 치환.

### 5.3 재적용
- PNG 갱신 후 다시 임베딩할 때: `files/` 폴더에서 `python embed_heatmap_base64.py` 실행.

---

## 6. 스킬·문서 반영

- **`.cursor/skills/agi-schedule-daily-update/SKILL.md`**  
  - **2c) PDF 파서 → WEATHER.PY 파이프라인** 절차(1→2→3단계, 4일 히트맵 PNG 생성 및 스케줄 HTML 삽입) 명시.

---

## 7. 결과물·경로 요약

| 구분 | 경로/파일 |
|------|-----------|
| 파서 txt | `out/weather_parsed/YYYYMMDD/*.txt` |
| 파이프라인 JSON | `out/weather_parsed/YYYYMMDD/weather_for_weather_py.json` |
| 4일 히트맵 PNG (WEATHER.PY) | `AGI TR 1-6.../out/weather_4day_heatmap.png` |
| 4일 대시보드 PNG (WEATHER_DASHBOARD) | `files/out/weather_4day_heatmap.png`, 사용 시 `files/weather_4day_heatmap_dashboard.png` |
| Base64 임베딩 스크립트 | `files/embed_heatmap_base64.py` |
| 스케줄 HTML (히트맵·Base64 적용) | `AGI TR Unit 1 Schedule_20260126.html`, `AGI TR Unit 1 Schedule_20260128.html`, `files/AGI20Unit20Schedule_20260126_redesigned_v2.html` |

---

## 8. 참고: Go/No-Go 로직

- **weathergonnologic.md**: 해상 운행(SEA TRANSIT) 전용 Go/No-Go 로직(파고·풍속·Squall 버퍼·연속 window 등) 정리. 파이프라인에서 추출한 파고(ft)→m 변환 및 risk_level 부여와 연계 가능.

---

**보고 종료.**
