# TR_DocHub_AGI_2026 통합 패키지

HVDC AGI TR Transportation 프로젝트용 문서 추적 시스템 통합 패키지

## 📦 빠른 시작

### 1. Python 환경 설정

```bash
pip install -r 06_Requirements/requirements_tr_doc_tracker.txt
```

### 2. 빌더 실행

```bash
cd 01_Python_Builders
python run_all_builders.py
```

빌더 선택:
- **1**: 정규화 모델 (권장) - `통합빌더.py`
- **2**: 기존 모델 - `create_tr_document_tracker_v2.py`
- **3**: 기존 모델 + DocGap 통합
- **4**: DocGap v2 → v3 Full Options
- **5**: DocGap v3.1 Operational 패치

### 3. 빌더 선택 가이드

자세한 내용은 `04_Documentation/Builder_Selection_Guide.md` 참조

### 4. VBA 모듈 임포트

`04_Documentation/VBA_Import_Guide.md` 참조

---

## 📁 폴더 구조

```
TR_DocHub_AGI_2026_Complete/
├── 01_Python_Builders/          # Python 빌더 스크립트
│   ├── 통합빌더.py               # 정규화 모델 빌더
│   ├── create_tr_document_tracker_v2.py  # 기존 모델 빌더
│   ├── build_docgap_v3_1_operational.py   # DocGap 운영 패치
│   ├── build_docgap_v3_fulloptions.py     # DocGap 전체옵션
│   ├── run_all_builders.py      # 통합 실행 스크립트
│   └── run_builder.py            # 개별 빌더 실행 헬퍼
│
├── 02_VBA_Modules/              # VBA 모듈 파일
│   ├── modControlTower.bas      # 통합 엔트리포인트
│   ├── modOperations.bas        # 정규화 모델 운영 함수
│   ├── TR_DocTracker_VBA_Module.bas  # TR 기능
│   ├── modTRDocTracker.bas      # Python 연동
│   ├── DocGapMacros_v3_1.bas   # DocGap 기능
│   └── ThisWorkbook_Shortcuts.bas  # 단축키
│
├── 03_Sheet_Codes/              # 시트 이벤트 코드
│   ├── Document_Tracker_Sheet_Code.txt  # 기존 모델용
│   └── T_Tracker_Sheet_Code.txt         # 정규화 모델용
│
├── 04_Documentation/            # 문서
│   ├── Builder_Selection_Guide.md      # 빌더 선택 가이드
│   ├── Sheet_Mapping_Guide.md           # 시트명 매핑
│   ├── Build_Checklist.md               # 빌드 체크리스트
│   ├── VBA_Import_Guide.md              # VBA 임포트 가이드
│   ├── TR_Document_Tracker_VBA_Guide_KR.md  # VBA 사용 가이드
│   ├── Phase 1, 2, 3 전체 구현 코드입니다.MD  # 구현 단계별 코드
│   ├── 통합.MD                           # 통합 설계 문서
│   └── 통합 12.MD                        # 정규화 모델 설계
│
├── 05_Templates/                # 빌더로 생성된 최신 템플릿 파일
│   └── (빌더 실행 시 자동 생성, 타임스탬프 포함)
│
├── 06_Requirements/             # Python 패키지 요구사항
│   └── requirements_tr_doc_tracker.txt
│
├── 07_Reference/                # 참고 문서
│   └── gate_pass_customs_checklist_EN.html  # Gate Pass 체크리스트
│
├── 08_Source_Templates/         # 원본/중간 버전 템플릿 보관소
│   └── README.md                # 원본 템플릿 설명
│
└── image/                       # 문서용 이미지 파일
```

---

## 🔧 주요 기능

### 정규화 모델 (통합빌더.py)

- **시트 구조**: S_Voyages, M_DocCatalog, M_Parties, R_DeadlineRules, T_Tracker, D_Dashboard
- **특징**: 룰테이블 기반 DueDate 자동 계산, 정규화된 데이터 모델
- **VBA**: modOperations.bas 필요 (InitializeWorkbook, GenerateTrackerRows, RecalcDeadlines)

### 기존 모델 (create_tr_document_tracker_v2.py)

- **시트 구조**: Voyage_Schedule, Doc_Matrix, Document_Tracker, Dashboard
- **특징**: 시나리오 지원, Python REFRESH 모드
- **VBA**: TR_DocTracker_VBA_Module.bas, modTRDocTracker.bas

### DocGap 통합

- **패치**: build_docgap_v3_1_operational.py
- **기능**: Inputs 시나리오 선택, Lead Time 매핑, OFCO_Req/NOC_Req 확장
- **VBA**: DocGapMacros_v3_1.bas

---

## 📋 실행 체크리스트

빌드 후 다음 단계를 확인하세요:

1. **빌드**
   - [ ] 빌더 실행 완료
   - [ ] 출력 파일이 `05_Templates/`에 생성됨

2. **Excel 패키징**
   - [ ] `.xlsx` → `.xlsm` 변환
   - [ ] VBA 모듈 임포트 (6개)
   - [ ] 시트 코드 추가
   - [ ] ThisWorkbook 단축키 추가

3. **검증**
   - [ ] `RefreshAll_ControlTower()` 실행
   - [ ] Dashboard KPI 업데이트 확인
   - [ ] Inputs → Voyage 1 연동 확인 (기존 모델)

자세한 내용: `04_Documentation/Build_Checklist.md`

---

## ⌨️ 단축키

- **Ctrl+Shift+R**: `RefreshAll_ControlTower()` - 전체 갱신
- **Ctrl+Shift+P**: `EXP_ExportToPDF()` - PDF 내보내기
- **Ctrl+Shift+E**: `TR_Draft_Reminder_Emails()` - 리마인더 이메일 초안

---

## 📚 주요 문서

| 문서 | 설명 |
|------|------|
| `Builder_Selection_Guide.md` | 빌더 선택 가이드 및 시나리오별 권장사항 |
| `Sheet_Mapping_Guide.md` | 정규화 모델 ↔ 기존 모델 시트명 매핑 |
| `Build_Checklist.md` | 빌드 후 검증 체크리스트 |
| `VBA_Import_Guide.md` | VBA 모듈 임포트 단계별 가이드 |
| `TR_Document_Tracker_VBA_Guide_KR.md` | VBA 사용 가이드 (한국어) |
| `Phase 1, 2, 3 전체 구현 코드입니다.MD` | 구현 단계별 코드 변경사항 |
| `통합.MD` | 통합 설계 문서 (Dashboard, Calendar, VBA_Pasteboard) |
| `통합 12.MD` | 정규화 모델 설계 문서 (룰테이블, 정규화 스키마) |

---

## 🔄 시나리오별 실행 순서

### 시나리오 1: 정규화 모델 (권장)

```bash
cd 01_Python_Builders
python run_all_builders.py
# 선택: 1
```

생성된 파일: `05_Templates/TR_DocHub_AGI_2026_Normalized_YYYYMMDD_HHMMSS.xlsx`

### 시나리오 2: 기존 모델

```bash
cd 01_Python_Builders
python run_all_builders.py
# 선택: 2
```

생성된 파일: `05_Templates/TR_Document_Tracker_v2_YYYYMMDD_HHMMSS.xlsx`

### 시나리오 3: 기존 모델 + DocGap 통합

```bash
cd 01_Python_Builders
python run_all_builders.py
# 선택: 3
```

생성된 파일: `05_Templates/TR_DocHub_AGI_2026_Integrated_YYYYMMDD_HHMMSS.xlsx`

### 시나리오 4: DocGap v2 → v3 Full Options

```bash
cd 01_Python_Builders
python run_all_builders.py
# 선택: 4
# DocGap v2 소스 파일 경로 입력
```

생성된 파일: 
- `05_Templates/OFCO_AGI_TR1_DocGap_Tracker_v3_FULLOPTIONS_YYYYMMDD_HHMMSS.xlsx`
- `05_Templates/OFCO_AGI_TR1_DocGap_Tracker_v3_FULLOPTIONS_YYYYMMDD_HHMMSS.xlsm`

### 시나리오 5: DocGap v3.1 Operational 패치 (기존 파일)

```bash
cd 01_Python_Builders
python run_all_builders.py
# 선택: 5
# 패치할 파일 경로 입력
```

생성된 파일: `05_Templates/TR_DocHub_AGI_2026_Patched_YYYYMMDD_HHMMSS.xlsx`

---

## 🚀 다음 단계

1. **빌더 실행**: `run_all_builders.py`로 템플릿 생성
2. **Excel 변환**: `.xlsx` → `.xlsm` 변환
3. **VBA 임포트**: `02_VBA_Modules/`의 모든 `.bas` 파일 임포트
4. **시트 코드 추가**: `03_Sheet_Codes/`의 코드를 해당 시트에 추가
5. **검증**: `RefreshAll_ControlTower()` 실행 및 KPI 확인

---

## 📝 버전 정보

- **Version**: 1.0
- **Date**: 2026-01-19
- **Project**: HVDC AGI TR Transportation
- **Python**: 3.11+
- **Excel**: 2021 LTSC / Microsoft 365

---

## ⚠️ 주의사항

1. **Excel 파일 잠금**: `.xlsm` 파일이 열려있으면 이동/삭제 불가
2. **Python 경로**: `modTRDocTracker.bas`는 상대 경로 사용 (통합 폴더 구조 기준)
3. **VBA 보안**: 매크로 보안 설정 확인 필요
4. **템플릿 정리**: 
   - 빌더로 생성된 템플릿은 `05_Templates/`에 자동 저장
   - 원본 템플릿은 `08_Source_Templates/`에 보관
   - Excel 파일이 열려있으면 이동/삭제 불가 (파일 닫은 후 정리)
   - 중복 파일은 최신 버전 기준으로 하나만 유지
   - `AGI DOCS/` 루트의 임시 파일(`~$*.xlsm`)은 Excel 종료 시 자동 삭제

---

## 📂 템플릿 정리 기준

### 폴더 역할

- **`05_Templates/`**: 
  - 빌더(`run_all_builders.py`)로 **새로 생성**된 템플릿만 보관
  - 파일명 형식: `TR_DocHub_AGI_2026_[Type]_YYYYMMDD_HHMMSS.xlsx`
  - VBA 임포트 및 검증 완료된 최종 버전

- **`08_Source_Templates/`**: 
  - 통합 패키지 생성 **이전**의 원본/중간 버전 보관
  - 참고용으로만 사용 (새 템플릿 생성 시 사용하지 않음)
  - `README.md`에 파일 목록 및 설명 포함

### 중복 파일 처리

- `05_Templates/`와 `08_Source_Templates/`에 동일 파일이 있으면:
  1. `05_Templates/`의 파일이 최신 버전인지 확인
  2. 최신 버전은 `05_Templates/`에 유지
  3. 원본/중간 버전만 `08_Source_Templates/`에 보관
  4. 중복 파일은 하나만 유지 (최신 기준)

### 정리 체크리스트

- [ ] `05_Templates/`에는 빌더 생성 파일만 존재
- [ ] `08_Source_Templates/`에는 원본/중간 버전만 존재
- [ ] 중복 파일 제거 완료
- [ ] `AGI DOCS/` 루트의 임시 파일(`~$*.xlsm`) 정리

---

## 📞 지원

문제 발생 시:
1. `04_Documentation/`의 가이드 문서 확인
2. `Build_Checklist.md`의 체크리스트 확인
3. VBA 오류는 `VBA_Import_Guide.md`의 문제 해결 섹션 참조
