# 폴더별 주요 작업 내역 보고서

**작성일**: 2026-01-28  
**작성자**: AI Assistant  
**범위**: CONVERT 폴더 전체

---

## 📋 Executive Summary

CONVERT 폴더는 HVDC 프로젝트 관련 물류 자동화 및 문서 변환 시스템을 포함하고 있습니다. 주요 작업 영역은 다음과 같습니다:

1. **문서 변환 시스템** (mrconvert_v1)
2. **이메일 검색 및 스레드 추적** (email_search)
3. **AGI TR 운송 일정 관리** (AGI TR 1-6, AGI DOCS)
4. **상업 송장 및 포장 목록 생성** (CIPL)
5. **선박 안정성 계산** (vessel_stability_python)
6. **콘텐츠 캘린더** (JPT71)
7. **유틸리티 스크립트** (scripts)

---

## 📁 폴더별 상세 작업 내역

### 1. mrconvert_v1/ - 문서 변환 시스템

**목적**: PDF/DOCX/XLSX → TXT/MD/JSON 변환 (OCR 지원)

**주요 작업**:
- ✅ PDF/DOCX/XLSX/XLSB → TXT/MD/JSON 변환 엔진 구현
- ✅ OCR 지원 (ocrmypdf, pytesseract)
- ✅ 테이블 추출 및 CSV 변환
- ✅ 온톨로지 문서 변환 (ontology_machine_readable/)
- ✅ 플러그인 아키텍처 설계 (AGENTS.md 기준)
- ✅ Cursor 통합 패키지 (cursor_only_pack_v1/)

**주요 파일**:
- `src/`: 핵심 변환 로직
- `ontology_machine_readable/`: 기계 가독형 온톨로지 문서 (JSON/MD/TXT/CSV)
- `cursor_only_pack_v1/`: Cursor IDE 통합 패키지
- `AGENTS.md`: 프로젝트 가이드라인

**상태**: ✅ 운영 중

---

### 2. email_search/ - 이메일 검색 및 스레드 추적 시스템

**목적**: Outlook Excel 데이터 기반 이메일 검색 및 스레드 추적

**주요 작업**:
- ✅ AQS-lite 검색 CLI 구현 (`outlook_aqs_searcher.py`)
- ✅ 컬럼 별칭 시스템 (`config/aqs_column_aliases.json`)
- ✅ 스키마 리포트 자동 생성
- ✅ Union-Find 기반 스레드 추적 (v3)
- ✅ Streamlit 대시보드 UI 구현
- ✅ 스레드 CLI (`export_email_threads_cli.py`)

**데이터 규모**:
- 총 이메일: 21,086개
- 컬럼 수: 20개

**주요 기능**:
- 필드 검색: `subject:`, `from:`, `to:`, `body:`, `case:`, `site:`, `lpo:`
- 날짜 범위 검색: `received:2025-10-01..2025-10-31`
- 복합 쿼리: `(subject:tr OR subject:agi) AND NOT from:log`

**KPI 목표**:
- 스레드 정답률: ≥85%
- 검색 응답시간: ≤2.00s
- 인덱싱 시간: ≤60.00s

**상태**: ✅ 운영 중 (Option A: Excel-only)

---

### 3. AGI DOCS/ - AGI TR 문서 추적 시스템

**목적**: HVDC AGI TR Transportation 프로젝트용 문서 추적 시스템

**주요 작업**:
- ✅ Python 빌더 시스템 구현 (5가지 빌더)
- ✅ VBA 모듈 통합 (6개 모듈)
- ✅ 정규화 모델 및 기존 모델 지원
- ✅ DocGap v3.1 통합
- ✅ 통합 빌더 (`통합빌더.py`)
- ✅ 문서화 완료 (한국어 가이드 포함)

**빌더 종류**:
1. 정규화 모델 (권장) - `통합빌더.py`
2. 기존 모델 - `create_tr_document_tracker_v2.py`
3. 기존 모델 + DocGap 통합
4. DocGap v2 → v3 Full Options
5. DocGap v3.1 Operational 패치

**주요 기능**:
- 시트 구조: S_Voyages, M_DocCatalog, M_Parties, R_DeadlineRules, T_Tracker, D_Dashboard
- 룰테이블 기반 DueDate 자동 계산
- VBA 단축키 지원 (Ctrl+Shift+R, Ctrl+Shift+P, Ctrl+Shift+E)

**상태**: ✅ 운영 중

---

### 4. AGI TR 1-6 Transportation Master Gantt Chart/ - 운송 일정 관리

**목적**: AGI 변압기 운송 일정 관리 및 간트 차트 생성

**주요 작업**:
- ✅ 다중 시나리오 간트 차트 생성
- ✅ VBA 기반 간트 차트 자동화
- ✅ 날씨 리스크 히트맵 생성
- ✅ LCT 운송 현황 추적
- ✅ SPMT 운영 현황 추적
- ✅ 항차별 운송 일정 관리
- ✅ JSON/CSV/HTML 변환

**주요 파일**:
- `AGI_TR_MultiScenario_Master_Gantt_*.xlsm`: 다중 시나리오 간트 차트
- `AGI_TR6_MasterSuite_READY_v3_1/`: MasterSuite 패키지
- `AGI_TR7_Dynamic_Gantt/`: 동적 간트 차트
- `new/`: 최신 버전 개발 파일

**상태**: ✅ 운영 중 (지속적 업데이트)

---

### 5. CIPL/ - 상업 송장 및 포장 목록 생성 시스템

**목적**: Commercial Invoice & Packing List 자동 생성

**주요 작업**:
- ✅ 성능 최적화 적용 (2026-01-14)
  - Border 처리 최적화 (~90% 향상)
  - Alignment 캐싱 (~30% 향상)
  - Merged Cell 최적화 (O(n*m) → O(1))
- ✅ excel_helpers 모듈 통합
- ✅ 폴더 구조 정리 (CIPL_PATCH_PACKAGE/ 활성, CIPL_LEGACY/ 레거시)
- ✅ 시각적 출력 검증 완료

**주요 파일**:
- `CIPL_PATCH_PACKAGE/`: 최신 최적화 버전 ✅
  - `COMMERCIAL INVOICE.PY`
  - `PACKING LIST.PY`
  - `CI RIDER.PY`
  - `PACKING LIST ATTACHED RIDER.PY`
  - `CIPL.py` (데이터 매퍼)
  - `make_cipl_set.py` (통합 빌더)
- `CIPL_LEGACY/`: 이전 버전 (참고용) ⚠️

**성능 개선**:
- 전체 처리 시간: ~40-50% 단축

**상태**: ✅ 운영 중 (v2.0)

---

### 6. vessel_stability_python/ - 선박 안정성 계산

**목적**: Excel 기반 선박 안정성 계산 함수를 Python으로 변환

**주요 작업**:
- ✅ Excel 함수 분석 및 Python 변환
- ✅ Hydrostatic 계산 구현
- ✅ Stability 계산 구현
- ✅ 검증 리포트 생성
- ✅ 문서화 완료

**주요 파일**:
- `src/vessel_stability_functions.py`: 핵심 계산 함수
- `src/excel_to_python_stability.py`: Excel 변환 로직
- `docs/`: 상세 문서 (function_reference.md, implementation_guide.md 등)
- `tests/`: 테스트 코드

**상태**: ✅ 운영 중

---

### 7. JPT71/ - 콘텐츠 캘린더 애플리케이션

**목적**: Excel 기반 콘텐츠 캘린더를 Python 애플리케이션으로 변환

**주요 작업**:
- ✅ Excel 엔진 구현 (`excel_python_engine.py`)
- ✅ Content Calendar 애플리케이션 구현
- ✅ JSON/HTML 내보내기 기능
- ✅ 캘린더 계산 로직 구현

**주요 기능**:
- Excel 파일 로드
- 캘린더 계산 로직 (DATE, WEEKDAY 함수 매핑)
- 콘텐츠 데이터 관리
- JSON/HTML 내보내기

**상태**: ✅ 운영 중

---

### 8. scripts/ - 유틸리티 스크립트

**목적**: 다양한 유틸리티 스크립트 모음

**주요 스크립트**:
- `build_email_threads.py`: 이메일 스레드 빌드
- `convert_pdf_to_md_improved.py`: PDF → Markdown 변환
- `email_derived_fields.py`: 이메일 파생 필드 생성
- `email_thread_tracker_v2_enhanced.py`: 이메일 스레드 추적 (v2)
- `repair_hvdc_json.py`: HVDC JSON 수리
- `update_weather_from_csv.py`: CSV에서 날씨 데이터 업데이트
- `validate_mammoet_data.py`: Mammoet 데이터 검증

**상태**: ✅ 운영 중

---

### 9. mammoet/ - Mammoet 관련 파일

**목적**: Mammoet 프로젝트 관련 문서 및 데이터

**주요 내용**:
- PDF 문서 (52개)
- 이미지 파일 (PNG 20개, JPG 18개)
- Gate Pass 상세 추출 스크립트
- 데이터 검증 스크립트

**상태**: ✅ 자료 보관

---

### 10. out/ 및 output/ - 출력 폴더

**목적**: 변환 및 생성된 파일 저장

**주요 내용**:
- `out/`: JSON 리포트, 변환된 문서
- `output/`: PDF 출력 파일

**상태**: ✅ 운영 중

---

## 📊 전체 통계

### 파일 수 (대략)
- Python 파일: ~728개
- Markdown 문서: ~245개
- JSON 파일: ~92개
- Excel 파일: 다수 (정확한 수 미집계)

### 주요 프로젝트
1. **mrconvert_v1**: 문서 변환 시스템
2. **email_search**: 이메일 검색/스레드 추적 (21,086개 이메일)
3. **AGI DOCS**: 문서 추적 시스템 (5가지 빌더)
4. **AGI TR 1-6**: 운송 일정 관리 (다중 시나리오)
5. **CIPL**: 상업 송장/포장 목록 생성 (성능 최적화 완료)
6. **vessel_stability_python**: 선박 안정성 계산
7. **JPT71**: 콘텐츠 캘린더

---

## 🔄 최근 업데이트 (2026-01)

### 2026-01-28
- 폴더별 작업 내역 보고서 작성

### 2026-01-25
- email_search: 스레드 추적 v3 구현 완료
- email_search: 대시보드 UI 구현 완료

### 2026-01-23
- email_search: AQS 검색 CLI 구현 완료
- email_search: 별칭 시스템 구현 완료

### 2026-01-19
- AGI DOCS: TR_DocHub_AGI_2026 통합 패키지 완성

### 2026-01-14
- CIPL: 성능 최적화 적용 (v2.0)
- CIPL: 폴더 구조 정리 완료

---

## 🎯 주요 성과

1. **자동화 시스템 구축**
   - 문서 변환 자동화 (mrconvert)
   - 이메일 검색/분석 자동화 (email_search)
   - 문서 추적 자동화 (AGI DOCS)
   - 상업 송장 생성 자동화 (CIPL)

2. **성능 최적화**
   - CIPL: ~40-50% 처리 시간 단축
   - email_search: ≤2.00s 검색 응답시간

3. **통합 및 표준화**
   - 플러그인 아키텍처 (mrconvert)
   - 정규화 모델 (AGI DOCS)
   - 컬럼 별칭 시스템 (email_search)

4. **문서화**
   - 모든 주요 프로젝트에 README 및 가이드 문서 완비
   - 한국어 가이드 포함 (AGI DOCS)

---

## 📝 향후 계획

### email_search
- [ ] 역색인 성능 최적화
- [ ] Microsoft Graph API 연동 (Option B)
- [ ] Delta query 기반 변경 동기화 (Option C)

### mrconvert
- [ ] 추가 플러그인 개발
- [ ] OCR 품질 개선

### AGI TR
- [ ] 실시간 업데이트 기능
- [ ] 웹 대시보드 연동

---

## 🔗 관련 문서

- `AGENTS.md`: mrconvert 프로젝트 가이드라인
- `email_search/README.md`: 이메일 검색 시스템 문서
- `CIPL/UPDATE_LOG.md`: CIPL 업데이트 로그
- `AGI DOCS/TR_DocHub_AGI_2026_Complete/README.md`: 문서 추적 시스템 문서

---

**보고서 작성 완료**: 2026-01-28
