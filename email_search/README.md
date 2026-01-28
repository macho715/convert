# Outlook Email Search & Thread Tracking System

## 프로젝트 개요

Outlook에서 내보낸 Excel 파일을 기반으로 이메일 검색 및 스레드 추적 시스템을 구축했습니다.

**주요 목표:**
- Excel 기반 이메일 데이터 검색 (AQS-lite 쿼리 지원)
- 관련 이메일 자동 그룹핑 및 스레드 추적
- 맥락 기반 검색 (케이스/사이트/LPO 자동 연결)
- 향후 Microsoft Graph 연동 준비 (SSOT 구조)

**데이터 규모:**
- 총 이메일: 21,086개
- 컬럼 수: 20개
- 시트: 전체_데이터

---

## 구현된 기능

### 1. AQS-lite 검색 CLI (`scripts/outlook_aqs_searcher.py`)

**지원 기능:**
- 필드 검색: `subject:`, `from:`, `to:`, `body:`, `case:`, `site:`, `lpo:`
- 날짜 범위: `received:2025-10-01..2025-10-31`
- 복합 쿼리: `(subject:tr OR subject:agi) AND NOT from:log`
- 괄호 그룹핑 및 NOT 연산자 지원
- AST 기반 쿼리 파싱

**주요 클래스:**
- `SchemaValidator`: 컬럼 별칭 검증 및 매핑
- `EmailNormalizer`: 데이터 정규화
- `AQSParser`: 쿼리 파싱 (AST 생성)
- `OutlookAqsSearcher`: 검색 실행

**CLI 옵션:**
```bash
python scripts/outlook_aqs_searcher.py <excel_file> -q "<query>" \
  --config config/aqs_column_aliases.json \
  --max-results 50 \
  --schema-report outputs/reports/schema.json \
  --auto-schema \
  --export outputs/searches/results.xlsx
```

### 2. 컬럼 별칭 시스템 (`config/aqs_column_aliases.json`)

**구조:**
- 파일/시트별 별칭 정의
- 기본(built-in) 별칭과 사용자 정의(custom) 별칭 병합
- 별칭 우선순위: custom > built-in

**예시:**
```json
{
  "OUTLOOK_HVDC_ALL_rev.xlsx": {
    "전체_데이터": {
      "subject": ["제목", "Subject", "subject"],
      "from": ["발신자", "From", "SenderName", "SenderEmail"],
      "case": ["케이스번호", "CaseNumber", "case_numbers", "hvdc_cases"]
    }
  }
}
```

### 3. 스키마 리포트 시스템

**생성 정보:**
- 해결된 별칭 (resolved_aliases)
- 별칭 출처 (config/built-in/merged)
- 시도한/매칭된/미매칭 별칭 목록
- 행/컬럼 수, 사용 가능한 컬럼 목록

**출력 파일:**
- `data/OUTLOOK_HVDC_ALL_rev.schema.json` (자동 생성)
- `outputs/reports/schema_report.json` (명시적 생성)

### 4. 스레드 추적 시스템 (Option A - 구현 예정)

**설계:**
- Union-Find 기반 connected component 구성
- 역색인(inverted index) 활용 (O(n log n) 성능)
- 휴리스틱 스레딩 (Subject + 엔티티 + 시간창)

**파일:**
- `scripts/outlook_thread_tracker_v3.py` (구현 완료)
- `scripts/export_email_threads_cli.py` (CLI, 구현 완료)

---

## 파일 구조

```
email_search/
├── README.md                        # 프로젝트 문서 (이 파일)
├── docs/
│   ├── ARCHITECTURE.md             # 아키텍처 문서 (예정)
│   ├── API_REFERENCE.md            # API 참조 (예정)
│   └── MIGRATION_GUIDE.md           # Graph 전환 가이드 (예정)
├── scripts/
│   ├── outlook_aqs_searcher.py     # AQS 검색 CLI
│   ├── outlook_thread_tracker_v3.py # 스레드 추적
│   └── export_email_threads_cli.py  # 스레드 CLI
├── config/
│   └── aqs_column_aliases.json     # 별칭 설정
├── data/
│   ├── OUTLOOK_HVDC_ALL_rev.xlsx   # 원본 데이터
│   ├── OUTLOOK_HVDC_ALL_rev_analysis.json
│   └── OUTLOOK_HVDC_ALL_rev.schema.json
├── outputs/
│   ├── reports/                    # 실행 리포트
│   │   ├── _run_report_aqs.json
│   │   └── schema_report.json
│   ├── threads/                    # 스레드 분석 결과
│   │   ├── threads.json (예정)
│   │   └── edges.csv (예정)
│   └── searches/                   # 검색 결과
│       └── search_*.csv (예정)
└── sql/                            # Supabase DDL (Option B)
    └── email_ssot.sql (예정)
```

---

## 사용 예시

### 기본 검색

```bash
# 제목으로 검색
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "subject:tr" \
  --config config/aqs_column_aliases.json

# 발신자로 검색
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "from:karthik" \
  --config config/aqs_column_aliases.json

# 케이스 번호로 검색
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "case:123" \
  --config config/aqs_column_aliases.json
```

### 복합 쿼리

```bash
# OR + AND + NOT
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "(subject:tr OR subject:agi) AND NOT from:log" \
  --config config/aqs_column_aliases.json \
  --max-results 20
```

### 스키마 리포트 생성

```bash
# 자동 생성 (data/OUTLOOK_HVDC_ALL_rev.schema.json)
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "subject:test" \
  --config config/aqs_column_aliases.json \
  --auto-schema

# 명시적 생성
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx \
  -q "subject:test" \
  --config config/aqs_column_aliases.json \
  --schema-report outputs/reports/schema_report.json
```

---

## 데이터 스키마

### 현재 Excel 컬럼 (20개)

**기본 메타데이터:**
- `no`: 번호
- `Month`: 월

**이메일 헤더:**
- `Subject`: 제목
- `SenderName`: 발신자 이름
- `SenderEmail`: 발신자 이메일
- `RecipientTo`: 수신자

**시간:**
- `DeliveryTime`: 수신일시
- `CreationTime`: 생성일시

**프로젝트 메타데이터:**
- `case_numbers`, `hvdc_cases`, `primary_case`: 케이스 번호
- `sites`, `primary_site`, `site`: 사이트
- `lpo`, `lpo_numbers`: LPO 번호
- `phase`, `stage`, `stage_hits`: 단계/스테이지

**본문:**
- `PlainTextBody`: 본문 (평문)

### 누락된 필드 (향후 Graph 연동 시 보강)

- RFC 헤더: `Message-ID`, `In-Reply-To`, `References`
- Graph 메타: `conversationId`, `conversationIndex`, `internetMessageId`
- 수신자: `RecipientCc`, `RecipientBcc`
- 플래그: `HasAttachment`, `IsFlagged`, `Category`

---

## 향후 계획

### Option A: Excel-only 개선 (현재 진행)

- [x] AQS 검색 CLI 구현
- [x] 별칭 시스템 구현
- [x] 스키마 리포트 생성
- [x] Union-Find 기반 스레드 추적 (v3)
- [ ] 역색인 성능 최적화
- [x] 대시보드 UI (Streamlit)

**KPI 목표:**
- 스레드 정답률: ≥85%
- 검색 응답시간: ≤2.00s
- 인덱싱 시간: ≤60.00s (21,086 rows)

### Option B: Graph 연동 (SSOT)

- [ ] Supabase DDL 적용
- [ ] Microsoft Graph API 연동
- [ ] 헤더/대화ID 수집
- [ ] Excel → Supabase 마이그레이션

**KPI 목표:**
- 스레드 정답률: ≥90%
- 동기화 지연: ≤300.00s

### Option C: 운영형 (히스토리 추적)

- [ ] Delta query 기반 변경 동기화
- [ ] 이벤트 로그 (append-only)
- [ ] 실시간 대시보드
- [ ] 트리/타임라인 시각화

---

## 기술 스택

- **Python 3.8+**
- **pandas**: 데이터 처리
- **openpyxl**: Excel 읽기/쓰기
- **typer/argparse**: CLI
- **json**: 설정/리포트

**향후 추가:**
- **streamlit**: 대시보드
- **plotly**: 시각화
- **supabase**: SSOT 저장소
- **msal**: Microsoft Graph 인증

---

## 참고 문서

- [Microsoft Graph - message resource](https://learn.microsoft.com/en-us/graph/api/resources/message)
- [RFC 5322 - Internet Message Format](https://datatracker.ietf.org/doc/html/rfc5322)
- [Microsoft Graph - Delta query](https://learn.microsoft.com/en-us/graph/delta-query-overview)

---

## 변경 이력

### 2026-01-23
- AQS 검색 CLI 구현 완료
- 별칭 시스템 구현 완료
- 스키마 리포트 생성 기능 추가
- `--auto-schema` 옵션 추가
- 별칭 우선순위 조정 (SenderName 우선)
- 파일 구조 정리 및 문서화

### 2026-01-25
- 스레드 추적 v3 구현 완료 (`scripts/outlook_thread_tracker_v3.py`)
- 스레드 CLI 구현 완료 (`scripts/export_email_threads_cli.py`)
- 대시보드 UI 구현 완료 (`dashboard/`)
- README 문서 현행화

### 향후
- Graph 연동 준비
