---
name: agi-schedule-daily-update
description: 통합 파이프라인 2단계. AGI TR Schedule HTML(files 폴더)의 공지란·Weather & Marine Risk 블록 매일 갱신. 기본 파일 files/AGI TR SCHEDULE_20260128.html, 업데이트 시 가장 최근 날짜 파일 사용.
---

# agi-schedule-daily-update

## 파이프라인 위치

- **통합 파이프라인(에이전트)** 에서 **항상 2번째**로 적용되는 스킬이다.
- 순서: 1) agi-schedule-shift → **2) agi-schedule-daily-update** → 3) agi-schedule-pipeline-check → 4) weather-go-nogo.
- 모든 요청에 대해 본 스킬(공지·날씨 갱신)이 **항상** 호출된다.

## 작업 범위

- **모든 작업은 `AGI TR 1-6 Transportation Master Gantt Chart/files/` 폴더 안에서만 수행한다.**

## 대상 파일

- **기본 파일**: `AGI TR 1-6 Transportation Master Gantt Chart/files/AGI TR SCHEDULE_20260128.html`
- **대상 패턴**: `AGI TR 1-6 Transportation Master Gantt Chart/files/AGI TR SCHEDULE_*.html`
- **업데이트 시 소스**: `files/` 안에서 **파일명에 포함된 날짜(YYYYMMDD)가 가장 최근인** `AGI TR SCHEDULE_*.html` 을 읽어서 갱신한다.
  예: `AGI TR SCHEDULE_20260128.html`, `AGI TR SCHEDULE_20260129.html` 이 있으면 → `AGI TR SCHEDULE_20260129.html` 를 소스로 사용.

## 언제 사용

- AGI TR Schedule HTML의 **공지란(Operational Notice)**을 사용자 제공 내용으로 갱신할 때
- **Weather & Marine Risk Update** 블록 갱신: (1) Mina Zayed 인근 = 인터넷 검색, (2) 해상 = 사용자 수동 다운로드한 PDF/JPG 파싱 후 삽입
- "AGI schedule 공지", "날씨 블록 업데이트", "Mina Zayed weather 반영", "weather 폴더 PDF 파싱" 요청 시

## 입력

| 항목 | 제공자 | 형식 | 예시 |
|------|--------|------|------|
| 공지 날짜 | 사용자 | YYYY-MM-DD | 2026-01-28 |
| 공지 텍스트 | 사용자 | 한 줄 또는 여러 줄 | Action : Request Material entry pass |
| 날씨/해상 정보 | 에이전트(웹 검색) 또는 사용자 다운로드(해상) | 아래 "날씨 입력 2원화" 참조 | - |

## 날씨 입력 (2원화)

| 구분 | 입력 경로 | 동작 |
|------|-----------|------|
| **1. Mina Zayed Port 인근 날씨** | 에이전트(웹 검색) | 인터넷 검색 후 요약 문단을 날씨란에 삽입(기존 방식). |
| **2. 해상 날씨** | 사용자 수동 다운로드 | PDF·JPG를 `files/weather/YYYYMMDD/` 에 두면, 에이전트가 **PDF/JPG 파싱** 후 해상 날씨 문단을 날씨란에 삽입. |

- **해상용 파일 경로**: `AGI TR 1-6 Transportation Master Gantt Chart/files/weather/YYYYMMDD/` (YYYYMMDD = 해당일).
- **해상용 파일**: PDF(예: ADNOC_DAILY_FORECAST_*.pdf, ADNOC-TR *.pdf, ENGLISH *.pdf), JPG(예: MR.JPG, SF.jpg, ST.jpg).
- **해상 파싱**: PDF는 텍스트 추출(mrconvert 또는 PyMuPDF 등), JPG는 OCR 또는 이미지 분석 후 해상 예보·해상 상태 요약문 생성 → Weather & Marine Risk 블록에 **해상(Marine)** 문단으로 추가.

## 절차

### 1) 공지란 갱신

- HTML에서 `<!-- Operational Notice -->` 다음의 `div.weather-alert` 블록을 찾는다.
- **사용자가 해당일 공지 내용을 제공한 경우**: 기존 날짜·본문을 **제공한 날짜 + 텍스트**로 교체. 마크업 유지: `<strong style="color: var(--accent-amber); font-weight: 600;">YYYY-MM-DD</strong><br>` + 공지 텍스트.
- **해당일에 내용이 없으면**: **날짜만** 해당일(갱신일)로 변경하고, **기준(기존)에 있던 내용은 전부 삭제**한다. 공지란에는 날짜만 남기고 본문은 비운다.

### 2) Weather & Marine Risk Update 갱신

- **포맷**: 기존 HTML과 동일하게 유지.
  - 제목: `Weather & Marine Risk Update (Mina Zayed Port)`
  - 서브: `Last Updated: DD Mon YYYY | Update Frequency: Weekly` (오늘 날짜로 갱신)
  - 본문: (1) Mina Zayed 인근 요약 문단, (2) 해상(Marine) 요약 문단. 날짜별 `<p><strong>DD Mon:</strong> …</p>` 유지.
- **예보 일수**: 조회(갱신)하는 날짜 기준 **당일 포함 4일치** 예보를 항상 삽입한다.
  - 예: 1월 28일 갱신 시 → 28 Jan, 29 Jan, 30 Jan, 31 Jan (4일).
  - 일반: 갱신일 D, D+1, D+2, D+3 (총 4일). 날짜별 문단 `<p><strong>DD Mon:</strong> …</p>` 유지.

**2a) Mina Zayed Port 인근 날씨**
- **웹 검색**: "Mina Zayed Port Abu Dhabi weather", "Arabian Gulf sea state wind 7 days" 등으로 검색.
- 검색 결과를 요약해 **위 4일치**에 맞춰 날짜별 문단으로 삽입.

**2b) 해상 날씨 (수동 다운로드)**
- **경로**: `AGI TR 1-6 Transportation Master Gantt Chart/files/weather/YYYYMMDD/` (해당일 폴더). 사용자가 PDF·JPG를 수동 다운로드해 둠.
- **파싱**: 해당일 폴더 내 PDF는 텍스트 추출(mrconvert 또는 동일 도구), JPG는 OCR 후 해상 예보/해상 상태 요약 추출.
- **PDF 파서 불가 시 fallback** (보안/DRM 등으로 PDF 직접 파싱이 실패할 때):
  1. **PDF 파일을 실행(열기)** 한다. (뷰어로 화면에 표시.)
  2. **화면에 표시된 PDF를 스크린 캡처**한다. (전체 창 또는 해당 페이지만.)
  3. **스크린 캡처한 이미지 파일**을 **파서**한다. (이미지 OCR/이미지 분석으로 텍스트·표 추출 후 해상 예보 요약.)
  - fallback 후에도 동일하게 추출 결과를 날씨란 "Marine / Offshore" 문단으로 삽입.
- **삽입**: 추출한 해상 예보를 날씨란에 "Marine / Offshore" 또는 **동일 4일치** 날짜별 해상 문단으로 삽입. 해당일 폴더가 없거나 파일이 없으면 2b는 생략.
- **JPG OCR 설정**: JPG 파싱은 Tesseract OCR 사용. `eng.traineddata`가 없으면 프로젝트에 `out/tessdata/eng.traineddata`를 두고, `scripts/weather_parse.py`가 자동으로 `TESSDATA_PREFIX=CONVERT/out/tessdata`를 사용. **날씨 파서 확인**: `python scripts/weather_parse.py "AGI TR 1-6.../files/weather/YYYYMMDD"` 로 PDF·JPG 파싱 가능 여부 확인.

**2c) PDF 파서 → WEATHER_DASHBOARD → 이미지 임베드 파이프라인 (4일 히트맵용)**
- **순서**: (1) PDF/JPG 파싱 → (2) 파서 txt → WEATHER용 JSON 변환 → (3) WEATHER_DASHBOARD.py 실행 → (4) embed_heatmap_base64.py 또는 파일 참조.
- **1단계**: `python scripts/weather_parse.py "AGI TR 1-6.../files/weather/YYYYMMDD" --out files/out/weather_parsed/YYYYMMDD`
- **2단계**: `python scripts/parsed_to_weather_json.py files/out/weather_parsed/YYYYMMDD` → `files/out/weather_parsed/YYYYMMDD/weather_for_weather_py.json` 생성 (ADNOC 텍스트에서 날짜·파고(ft)·풍속(kt) 추출).
- **3단계**: `files/` 폴더에서 `python WEATHER_DASHBOARD.py` 실행. 4일 모드일 때 위 JSON이 있으면 자동 사용 → `files/out/weather_4day_heatmap.png` 생성.
- **4단계(선택)**: `files/out/weather_4day_heatmap.png` 를 `files/weather_4day_heatmap_dashboard.png` 로 복사 후, `files/` 에서 `python embed_heatmap_base64.py` 실행 → `files/AGI TR SCHEDULE_*.html` 내 이미지를 Base64 인라인으로 교체. (되돌리려면 `python replace_img_ref.py`.)
- 스케줄 HTML Weather 블록에 위 파이프라인 실행 후 생성된 PNG(파일 참조 또는 Base64)를 사용.

- 위 (2a+2b) 결과를 합쳐 `<!-- Weather Alert -->` 다음 블록 전체를 교체한다.
- `class="weather-alert"`, 아이콘(⚠️), 인라인 스타일은 유지.

### 3) 저장

- **수정한 날짜(YYYYMMDD)를 파일명 뒤에 넣어 신규 파일로 `files/` 안에 저장**한다. 원본은 덮어쓰지 않는다.
- 예: 소스가 `files/AGI TR SCHEDULE_20260128.html` 이고 갱신일이 2026-01-29 이면 → `files/AGI TR SCHEDULE_20260129.html` 로 신규 생성.
- 파일명 형식: `files/AGI TR SCHEDULE_YYYYMMDD.html` (YYYYMMDD = 갱신을 수행한 날짜).

## 안전 규칙

- **`files/` 폴더 밖**의 HTML을 읽거나 쓰지 않는다.
- 공지란·날씨 블록 **밖**의 HTML(간트 데이터, 스크립트, KPI, Voyage Cards 등)은 수정하지 않는다. **KPI Grid(Total Days 재계산, SPMT Set=1)** 는 본 스킬 범위가 아니며, **3단계 agi-schedule-pipeline-check** 점검 F에서만 수행한다. 갱신 시 KPI 규칙 적용 책임은 pipeline-check 스킬에 있다.
- 공지 미제공 시: 해당일로 날짜만 변경하고 기존 내용은 삭제(날짜만 남김). 날씨만 갱신할 때도 공지란에 이 규칙 적용.

## 통합

- Subagent `/agi-schedule-updater`: 이 스킬을 사용해 공지·날씨 블록 갱신 수행. 모든 작업은 `files/` 안에서만 수행.
