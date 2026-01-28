문서 간 관계 분석 중. 통합 가이드 작성 중.


[4 tools called]


## CIPL 시스템 통합 가이드

### 문서 계층 구조

```
┌─────────────────────────────────────────────────────────────┐
│  Level 1: 전략/정책 문서                                      │
│  - LDG_v8.2_Full_Guide (5).md                               │
│  - CHAT GPT GPTS지침.txt                                     │
│  - Logistics Document Guardian - GPTS 개발 가이드 (4).md   │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Level 2: 실행 가이드                                         │
│  - GUIDE.MD (LDG → Commons → Excel 통합 실행)               │
│  - CIPL Excel AutoBuild – Full Implementation Guide v2.5.MD │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Level 3: 기술 파이프라인                                     │
│  - PIPELINE.MD (make_cipl_set.py 실행 흐름)                 │
└─────────────────────────────────────────────────────────────┘
```

### 문서별 역할 및 관계

#### 1. LDG_v8.2_Full_Guide (5).md
- 역할: LDG v8.2 전체 표준
- 범위: OCR → 정규화 → 검증 → LDG 연동
- 핵심 내용:
  - 6단계 OCR 파이프라인
  - 5단계 자동 검증
  - LDG_PAYLOAD v2.4 스키마
  - KPI 게이트 (MeanConf ≥ 0.92, TableAcc ≥ 98%)
  - ZERO Fail-Safe 정책

#### 2. CHAT GPT GPTS지침.txt
- 역할: GPT 실행 지침
- 범위: LDG CIPL Builder GPT 동작 규칙
- 핵심 내용:
  - ZERO-QUESTION 기본 동작
  - LDG_PAYLOAD v2.4 생성
  - Commons Mode 우선
  - make_cipl_set.py 경로 지정
  - Soft Warning 모드 (Excel 생성 차단 없음)

#### 3. Logistics Document Guardian - GPTS 개발 가이드 (4).md
- 역할: GPTs 구축 가이드
- 범위: Custom GPT 설계 및 구현
- 핵심 내용:
  - GPT Builder 설정
  - Knowledge 사용 주의사항
  - Data Analysis vs Actions 옵션
  - 필드 매핑 표준
  - 보안/프라이버시 고려사항

#### 4. GUIDE.MD
- 역할: 실행 통합 가이드
- 범위: LDG_PAYLOAD → Commons → Excel 변환
- 핵심 내용:
  - LDG_PAYLOAD → Commons Mode 변환 로직
  - 필드 매핑 테이블
  - 파이프라인 실행 예시
  - ZERO Fail-Safe 처리

#### 5. CIPL Excel AutoBuild – Full Implementation Guide v2.5.MD
- 역할: Excel 생성 품질 가이드
- 범위: Style-first, 서식 동일성 우선
- 핵심 내용:
  - SSoT 파일 참조 규칙
  - 4-Page 구조 고정
  - Style-first 체크리스트
  - KPI Soft Warning 정책

#### 6. PIPELINE.MD
- 역할: 기술 파이프라인 설계
- 범위: make_cipl_set.py 실행 흐름
- 핵심 내용:
  - 입력 모드 감지 (Commons vs Expanded)
  - 4개 시트 생성 순서
  - 모듈 동적 로드
  - excel_helpers 최적화

### 통합 실행 워크플로우

```
┌─────────────────────────────────────────────────────────────┐
│  Phase 1: 입력 처리 (LDG_v8.2_Full_Guide 기준)              │
│  - Pre-Prep: 이미지 전처리                                   │
│  - Vision OCR: 레이아웃/텍스트 인식                          │
│  - Smart Table Parser: 테이블 구조화                         │
│  - NLP Refine: 단위 정규화                                   │
│  - Field Tagger: 키-값 태깅                                  │
│  - Payload Builder: LDG_PAYLOAD v2.4 생성                   │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Phase 2: 검증 (LDG_v8.2_Full_Guide 기준)                    │
│  - KPI 체크: MeanConf, TableAcc, NumericIntegrity           │
│  - 교차 검증: CIPL↔BL↔PL Δ 계산                              │
│  - LDG_AUDIT v2.4 생성                                       │
│  - Soft Warning 모드 (Excel 생성 차단 없음)                 │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Phase 3: 데이터 변환 (GUIDE.MD 기준)                        │
│  - LDG_PAYLOAD → Commons Mode JSON 변환                      │
│  - 필드 매핑 (필드 매핑 테이블 참조)                          │
│  - Rider Items 자동 생성 (없을 경우)                         │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Phase 4: Excel 생성 (PIPELINE.MD + AutoBuild Guide 기준)   │
│  - make_cipl_set.py 실행                                     │
│  - CIPL.py.make_4page_data_dicts() 호출                      │
│  - 4개 시트 생성 (Style-first 준수)                          │
│  - excel_helpers 최적화 적용                                  │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Phase 5: 출력 및 검증                                        │
│  - CIPL_{Invoice_No}_{YYYY-MM-DD}.xlsx 생성                  │
│  - LDG_AUDIT 최종 보고                                        │
│  - Auto Guard Summary 출력                                   │
└─────────────────────────────────────────────────────────────┘
```

### 문서 간 데이터 흐름

```
LDG_PAYLOAD v2.4 (LDG_v8.2_Full_Guide)
    │
    ├─ meta: {mean_conf, table_acc, ...}
    ├─ data: {Invoice_No, Shipper, ...}
    ├─ evidence: [...]
    └─ warnings: [...]
    │
    ▼ (GUIDE.MD 변환 로직)
    │
Commons Mode JSON
    │
    ├─ commons: {pol, pod, invoice_no, ...}
    ├─ static_parties: {shipper_name, consignee_name1, ...}
    ├─ ci_rider_items: [] (선택적)
    └─ pl_rider_items: [] (선택적)
    │
    ▼ (PIPELINE.MD 실행)
    │
CIPL.py.make_4page_data_dicts()
    │
    ├─ ci_p1: {...}
    ├─ ci_rider_p2: {...}
    ├─ pl_p1: {...}
    └─ pl_rider_p2: {...}
    │
    ▼ (AutoBuild Guide Style-first)
    │
Excel Workbook (4 sheets)
    │
    ├─ Commercial_Invoice_P1 (COMMERCIAL INVOICE.PY)
    ├─ CI_Rider_P2 (CI RIDER.PY)
    ├─ Packing_List_P1 (PACKING LIST.PY)
    └─ PL_Rider_P2 (PACKING LIST ATTACHED RIDER.PY)
```

### 주요 정책 통합

#### KPI 및 품질 게이트

| 문서 | KPI 기준 | 모드 | 차단 여부 |
|------|---------|------|----------|
| LDG_v8.2_Full_Guide | MeanConf ≥ 0.92, TableAcc ≥ 98% | OCR 파이프라인 | Hard Block (ZERO) |
| CHAT GPT GPTS지침 | 동일 | Excel 생성 | Soft Warning (차단 없음) |
| AutoBuild Guide | 참고용 | Excel 생성 | Soft Warning (차단 없음) |

#### ZERO Fail-Safe 정책

| 문서 | ZERO 처리 | Excel 생성 |
|------|----------|-----------|
| LDG_v8.2_Full_Guide | Hard Block (중단 로그) | - |
| CHAT GPT GPTS지침 | Soft Warning (기록만) | 계속 진행 |
| GUIDE.MD | Soft Warning + Fallback | 최선 노력 |

### 실행 우선순위

1. CHAT GPT GPTS지침.txt (GPT 실행 규칙)
2. GUIDE.MD (데이터 변환 로직)
3. PIPELINE.MD (기술 실행 흐름)
4. AutoBuild Guide (서식 품질)
5. LDG_v8.2_Full_Guide (전체 표준 참조)

### GPT Instructions 통합 예시

```markdown
# LDG CIPL Builder (Samsung Format) - 통합 Instructions

## Core Mission
Extract shipment/trade info from text/PDF/screenshots and generate 
company-format CIPL Excel (4 sheets: CI P1, CI Rider P2, PL P1, PL Rider P2).

## Workflow (5 Phases)
1. Extract → LDG_PAYLOAD v2.4 (mean_conf ≥ 0.92, table_acc ≥ 98%)
2. Validate → LDG_AUDIT v2.4 (Soft Warning mode, no blocking)
3. Transform → Commons Mode JSON (GUIDE.MD mapping table)
4. Generate → Excel via make_cipl_set.py (Style-first, SSoT files)
5. Report → Auto Guard Summary + LDG_AUDIT

## Key Policies
- ZERO-QUESTION: Best-effort extraction, NULL + "가정:" in warnings
- Soft Warning: KPI 미달 시도 Excel 생성 (차단 없음)
- Style-first: Template fidelity > Data completeness
- HallucinationBan: NULL + assumption + Human Gate

## File Paths (SSoT)
- Builder: CIPL_PATCH_PACKAGE/make_cipl_set.py
- Mapper: CIPL_PATCH_PACKAGE/CIPL.py
- Templates: CIPL_PATCH_PACKAGE/*.PY
- Logo: CIPL_PATCH_PACKAGE/SAMSUNGLOGO.PNG

## Output Naming
CIPL_{Invoice_No}_{YYYY-MM-DD}.xlsx
(If missing: CIPL_UNK_{timestamp}.xlsx)
```

### 체크리스트

#### GPT 구축 시
- [ ] CHAT GPT GPTS지침.txt 기준 Instructions 설정
- [ ] GUIDE.MD 필드 매핑 테이블 참조
- [ ] PIPELINE.MD 실행 흐름 이해
- [ ] AutoBuild Guide Style-first 준수
- [ ] LDG_v8.2_Full_Guide 전체 표준 참조

#### 실행 시
- [ ] LDG_PAYLOAD v2.4 생성 (KPI 기록)
- [ ] Commons Mode 변환 (필드 매핑 정확)
- [ ] make_cipl_set.py 실행 (경로 확인)
- [ ] Style-first 검증 (서식 동일성)
- [ ] LDG_AUDIT 최종 보고

이 통합 가이드로 전체 시스템을 일관되게 운영할 수 있습니다.