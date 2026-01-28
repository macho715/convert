## Executive Summary (3~5줄)

* “ONE FILE(.xlsm)”로 **항차(Voyage 1~N) × 문서유형 × Responsible Party**를 SSOT(단일 진실원천) 테이블로 관리하고, **룰테이블 기반 Due Date 자동 산출**로 누락/지연을 구조적으로 차단합니다.
* **DAS Island / AGI Site를 분리 적용**(필수 문서/리드타임/승인흐름 차등)하되, 워크북은 하나로 유지합니다.
* Dashboard에서 Responsible Party별 **다가오는 데드라인/Overdue/증빙 누락**을 즉시 가시화하고, 버튼 1회로 **항차별 PDF/CSV Export Pack**을 생성합니다.
* 과기능을 배제하고, **테이블·드롭다운·조건부서식·간단 Gantt + 최소 VBA 모듈**만 통합합니다.

---

## EN Sources (≤3: title+date)

1. Microsoft Docs — *Workbook.ExportAsFixedFormat method (Excel)* (Jan 17, 2023). ([Microsoft Learn][1])
2. Microsoft Docs — *ListObject object (Excel)* (Sep 12, 2021). ([Microsoft Learn][2])
3. Microsoft Docs — *XlFileFormat enumeration (Excel)* (Jul 18, 2024). ([Microsoft Learn][3])

---

## Visual 표

| No | Item            | Value                                            | Risk              | Evidence/가정                                    |
| -: | --------------- | ------------------------------------------------ | ----------------- | ---------------------------------------------- |
|  1 | ONE FILE 원칙     | 단일 .xlsm + SSOT 테이블(Excel Table/ListObject)      | 데이터 중복/버전 분기      | ListObject 기반 테이블 운영 권장 ([Microsoft Learn][2]) |
|  2 | Site 분리 적용      | DAS/PTW≥48h, AGI/PTW≥24h 등 룰 차등                  | Site 혼용으로 리드타임 실패 | (가정) AD Ports/현장 운영 기준은 Site별로 상이              |
|  3 | Deadline 룰 고정   | DueDate = AnchorDate + OffsetDays                | 룰 변경 시 영향범위 불명확   | 룰테이블 버전/적용일(EffectiveFrom) 필수                  |
|  4 | 증빙/포털 추적        | Portal, RefNo, FileLink, SubmittedAt, VerifiedAt | 증빙 파일 분실/링크 깨짐    | Maqta 업로드·파일명 규칙 운영 가정                         |
|  5 | PDF Export Pack | ExportAsFixedFormat로 항차별 PDF 생성                  | 프린트영역/서식 깨짐       | ExportAsFixedFormat 지원 ([Microsoft Learn][1])  |
|  6 | CSV Export Pack | XlFileFormat로 CSV 저장                             | 인코딩/로케일 이슈        | CSV 포맷 열거형 관리 필요 ([Microsoft Learn][3])        |
|  7 | 운영 체크리스트 연계     | “필수 문서+담당+마감+검증자” 구조                             | 책임 분산/누락          | Mandatory docs 표 구조 존재                         |
|  8 | Maqta/ATLP 운영성  | 제출/확인 상태를 RefNo로 추적                              | 포털 상태 미동기         | Maqta 확인/업로드 운영 존재                             |

---

## 통합 시트 설계(필수/옵션)

### 공통 설계 원칙 (DAS/AGI 공통)

* **모든 핵심 데이터는 Excel Table(ListObject)**로만 보관(필터/피벗/구조참조 안정). ([Microsoft Learn][2])
* Key는 **정수 ID + 합성키(UniqueKey)**를 병행: 사람이 읽기 쉬우면서도 VBA에서 충돌 방지.
* Site별 차등은 “코드/룰/적용여부”로만 처리(시트 분기 최소화).

---

### 1) 필수 시트 (Minimum Viable Workbook)

#### (1) `00_HOME`

* 목적: 파일 사용법, 버튼, 상태 요약(오늘 Overdue, D+7 Due, Evidence Missing)
* 주요 영역: “오늘 기준(AsOfDate)”, “Site 선택(기본값)”, “빠른 링크”
* 버튼: `Refresh_All`, `ExportPack_Voyage`, `Add_Voyage`

#### (2) `10_MASTER_DOCS`  (문서 카탈로그: SSOT)

* **Table: `tDocType`**
* Key:

  * `DocTypeID`(PK, 숫자)
  * `DocCode`(예: PTW_HOTWORK, GATEPASS_QR, CUSTOMS_DECL)
  * `DocName`
* 필드(권장):

  * `Category`(PTW/Customs/HM/MWS/Marine/Operations 등)
  * `SiteApplicability` 또는 `Req_DAS`(Y/N), `Req_AGI`(Y/N)
  * `DefaultOwnerParty`(Responsible Party)
  * `DefaultVerifierParty`
  * `Portal`(ATLP/MTG/Maqta/Email/Other)
  * `MandatoryFlag`(Mandatory/Important/Standard)
  * `NamingPattern`(예: `[DocCode]_[VoyageID]_[YYYYMMDD].pdf`) — Maqta 파일명 규칙 운영 가정 

> DAS/AGI 분리 적용: `Req_DAS`, `Req_AGI`로 “생성 대상 문서”가 달라지며, Dashboard에서 Site별 필수 누락을 즉시 표시.

#### (3) `20_VOYAGES` (항차 일정 입력: 요구사항 4)

* **Table: `tVoyage`**
* Key:

  * `VoyageID`(PK, 예: V001)
  * `SiteCode`(DAS / AGI)
* 입력 필드(요구사항 고정):

  * `TR_Units` (예: “TR Units 1-2”)
  * `MZP_Arrival`
  * `LoadOut`
  * `MZP_Departure`
  * `AGI_Arrival`
  * `Doc_Deadline`
  * `Land_Permit_By` (담당/요청 주체 or 승인기관/담당자명)
* 파생(자동):

  * `Anchor_LoadOut = [@LoadOut]`
  * `Anchor_DocDeadline = [@Doc_Deadline]`
  * `VoyageStatus`(Planned/Active/Closed)

> 참고: 실제 운영에서 “문서 제출 절차/포맷(PDF) 및 업로드 플랫폼(Maqta)”가 정의되어 있다고 가정 

#### (4) `30_RULES` (Doc Deadline 산출 규칙: 요구사항 3)

* **Table A: `tAnchorType`**

  * `AnchorCode` = `MZP_ARRIVAL`, `LOADOUT`, `MZP_DEP`, `AGI_ARRIVAL`, `DOC_DEADLINE`

* **Table B: `tDueRule`** (룰테이블 고정/확장 가능)

  * Key: `RuleID`(PK), `DocTypeID`(FK), `SiteCode`
  * 필드:

    * `AnchorCode`
    * `OffsetDays`(정수, 예: -1, -2)
    * `LeadTimeNote`(예: “DAS PTW ≥48h”)
    * `EffectiveFrom`, `EffectiveTo`(룰 버전관리)

* **계산식(고정)**

  * `DueDate = AnchorDate + OffsetDays`
  * AnchorDate는 `tVoyage`의 해당 컬럼을 `XLOOKUP`(또는 `INDEX/MATCH`)로 매핑
  * 예시 룰:

    * Gate Pass = Load-out - 1
    * Customs = Doc Deadline - 2
  * (운영예) Doc Deadline이 “도착 3일 전”으로 관리되는 케이스 존재 

> DAS/AGI 분리 적용: 동일 DocType이라도 `SiteCode`에 따라 `OffsetDays`/`AnchorCode`가 달라질 수 있음(예: DAS PTW -2일, AGI PTW -1일).

#### (5) `40_SUBMISSIONS` (핵심 운영 테이블: 항차별 제출 상태)

* **Table: `tSubmission`** (항차×문서유형의 “업무 단위”)

* Key:

  * `SubmissionID`(PK, GUID 또는 연번)
  * `UniqueKey` = `VoyageID & "|" & DocTypeID` (Unique)

* 필드:

  * `VoyageID`(FK), `SiteCode`(복제 보관)
  * `DocTypeID`(FK), `DocCode`, `DocName`(조회값 캐시)
  * `ResponsibleParty`(드롭다운)
  * `VerifierParty`
  * `AnchorCode`, `AnchorDate`, `OffsetDays`, `DueDate`(자동)
  * `Status`(Not Started / In Progress / Submitted / Approved / Rejected / N/A)
  * `SubmittedAt`, `VerifiedAt`
  * `Portal`, `RefNo`(포털 접수번호)
  * `EvidenceLink`(파일/폴더 하이퍼링크)
  * `Remarks`

* 생성 로직(초기화):

  * `tVoyage` 입력 후, 버튼 1회로 `tDocType` × Site 필터를 펼쳐 `tSubmission`을 자동 생성(필수 문서만 또는 전체).

> “문서명/제출자/마감/검증자” 구조는 기존 필수 제출 문서 리스트 형태로 존재 

#### (6) `50_DASH_RACI` (대시보드/책임)

* **DAS Island Dashboard**

  * Responsible Party별: `Overdue`, `Due in 7d`, `Missing Evidence`, `Submitted not Verified`
  * Voyage 필터(드롭다운), Category 필터
* **AGI Site Dashboard**

  * 동일 구조 + AGI 전용 항목(예: Stability Template/MWS 사전검증 등)을 `Req_AGI` 기반으로 표시
* 최소 Gantt(권장):

  * Voyage 타임라인: `MZP_Arrival → LoadOut → MZP_Departure → (DAS/AGI Arrival)`
  * 문서 DueDate는 마커(조건부서식)로만 표시

---

### 2) 옵션 시트 (필요 시만 추가)

#### (7) `60_EVIDENCE_LOG` (증빙 상세 로그)

* Table: `tEvidence`
* Key: `EvidenceID`(PK), `SubmissionID`(FK)
* 필드: `FileName`, `FilePath`, `UploadedBy`, `UploadedAt`, `Checksum(선택)`, `Notes`

#### (8) `70_AUDIT` (변경 이력/감사)

* Table: `tAudit`
* 필드: `When`, `User`, `Action`, `Entity`, `Key`, `OldValue`, `NewValue`
* 구현: 최소로 “상태 변경/기한 변경/증빙 링크 변경”만 로깅

#### (9) `90_ADMIN`

* 드롭다운 마스터:

  * `tParty`(Responsible/Verifier 목록)
  * `tStatus`(상태 코드)
  * `tSiteConfig`(DAS/AGI 리드타임, 알림 D-일 기준 등)

---

## VBA 기능 목록(최소 기능 세트)

### 공통 (DAS/AGI 공통)

1. **Initialize / Rebuild Submissions**

   * 입력: `tVoyage`, `tDocType`, `tDueRule`
   * 출력: `tSubmission` 생성/갱신(Upsert), Site 적용 필터링
2. **Refresh_All**

   * DueDate 재계산(룰 변경 반영), 피벗/차트 Refresh, 조건부서식 재적용
3. **Validation Gate (Preflight)**

   * 필수 체크:

     * Voyage 날짜 누락(AnchorDate 없음) → 해당 Submission `Status=HOLD`
     * DueDate < Today & Status≠Approved → Overdue flag
     * Status=Submitted인데 EvidenceLink/RefNo 공란 → “Submitted Incomplete”
4. **Export Pack 생성(요구사항 5)**

   * Voyage 선택 → 폴더 생성 → 아래 파일 생성

     * `Dashboard_Voyage.pdf` (ExportAsFixedFormat) ([Microsoft Learn][1])
     * `Submissions_Voyage.csv`, `Overdue.csv`, `MissingEvidence.csv`
     * (선택) `EvidenceIndex.csv`
5. **One-Click Navigation**

   * Dashboard에서 클릭 시 해당 Submission 행으로 점프(필터 동기)

### DAS Island 전용(차등)

* `tSiteConfig(DAS)` 기준:

  * PTW 계열은 기본 `OffsetDays=-2`(≥48h)로 룰 제공
  * Hot Work 제한 등은 `LeadTimeNote`로 경고 표시(룰 기반)

### AGI Site 전용(차등)

* `tSiteConfig(AGI)` 기준:

  * PTW 계열 `OffsetDays=-1`(≥24h)
  * Stability/MWS/Trim Sheet 등 AGI 요구 문서는 `Req_AGI=Y`로만 생성

> 구현 디테일: **Excel Table(ListObject) 이름 고정**으로 VBA가 Range를 찾지 않고 테이블을 직접 참조(유지보수성 향상). ([Microsoft Learn][2])

---

## “VBA_Pasteboard 시트” 설계 (복사-붙여넣기 목적)

### A) 시트 레이아웃

1. **Header 블록**

   * Workbook 버전 / Last Build / 설치 체크리스트(체크박스)
2. **Module Registry 표 (`tVbaModules`)**

   * `Seq`, `ModuleName`, `Type`(Std/Class/ThisWorkbook), `Purpose`, `Dependencies`, `ButtonMap`
3. **Reference Libraries 표 (`tVbaRefs`)**

   * `RefName`, `Required(Y/N)`, `Reason`, `LateBindPossible(Y/N)`
4. **Button Map 표 (`tButtons`)**

   * `ButtonCaption`, `MacroName`, `TargetSheet`, `Tooltip`, `RoleLimit(옵션)`
5. **Install Order & Tests**

   * Step 1~8 설치 순서 + “Self Test” 실행 결과 기록 칸

### B) 모듈 최소 세트(권장)

1. `modConst` : 테이블명/시트명/Status 코드 상수
2. `modTables` : ListObject get/create, 컬럼 인덱싱 유틸
3. `modRules` : AnchorDate 매핑, DueDate 계산(룰 적용/버전 필터)
4. `modSubmissions` : Upsert 생성/갱신, UniqueKey 충돌 처리
5. `modDashboard` : 피벗/슬라이서/조건부서식 Refresh
6. `modExportPack` : 폴더 생성, PDF/CSV Export
7. `clsLogger`(옵션) : `tAudit` 기록

### C) 설치 순서(권장)

1. 시트/테이블 이름을 설계서대로 생성(또는 Python 빌더로 생성)
2. `modConst` → `modTables` → `modRules` → `modSubmissions` → `modDashboard` → `modExportPack`
3. `ThisWorkbook`에 이벤트 최소 삽입(예: Open 시 Refresh 여부만)
4. 버튼(Form Control) 생성 및 `tButtons` 매핑대로 Macro 연결
5. References 설정 → Self Test 실행 → 샘플 Voyage 1건으로 Export Pack 검증

### D) 참조 라이브러리 체크(권장)

* **필수(가능하면 Late Binding 회피)**

  * Microsoft Excel Object Library(기본)
  * Microsoft Office Object Library(기본)
* **권장(Optional)**

  * Microsoft Scripting Runtime (Dictionary/FSO) — *미참조 운영을 위해 Late Binding 권장*
* CSV/PDF:

  * PDF는 `ExportAsFixedFormat`로 충분 ([Microsoft Learn][1])
  * CSV 저장 시 `XlFileFormat` 기반 선택(UTF-8 여부 포함) ([Microsoft Learn][3])

---

## Python 빌더 아이디어 (openpyxl/pandas/xlwings 등)

목표: 운영용 .xlsm은 “ONE FILE”로 유지하되, **초기 생성/표준화/배포는 Python으로 자동화**.

1. **Workbook Skeleton Builder (openpyxl)**

* 시트/테이블/Named Range/드롭다운/조건부서식 “표준 골격” 자동 생성
* 이미 유사한 방식으로 다중 시트를 고정 순서로 생성하는 템플릿 생성기가 존재(16시트, NamedRanges, TableStyle) 
* 산출물: `HVDC_DOC_CONTROL_TEMPLATE.xlsx` + (별도) `VBA.bas` 패키지

2. **Rules & Doc Catalog Compiler (pandas)**

* `tDocType`, `tDueRule`를 CSV/YAML로 관리 → 빌더가 엑셀 테이블로 주입
* Site별 룰/필수문서 차이를 Git diff로 추적(감사 대응)

3. **Deployment Automation (xlwings/pywin32)**

* 배포 시점에:

  * 템플릿 복제 → 프로젝트명/Run ID 반영
  * (선택) VBA 모듈을 자동 삽입(보안정책 허용 시)
  * 버튼/인쇄영역/Export 폴더 기본값 세팅

---

## Options A/B/C (Cost/Risk/Time)

| Option | Scope                                   | Cost        | Risk                      | Time        |
| ------ | --------------------------------------- | ----------- | ------------------------- | ----------- |
| A      | **수식+테이블 중심**, VBA는 Export Pack만        | Low         | Medium (수작업 Upsert/오류 가능) | Fast        |
| B      | A + **Upsert 생성/검증/대시보드 자동 Refresh** 포함 | Medium      | Low~Med (VBA 품질 의존)       | Medium      |
| C      | B + **Python 빌더(표준화/배포/룰관리)** + 경량 거버넌스 | Medium~High | Low (표준 유지), 단 운영환경/권한 이슈 | Medium~Long |

권장: **B를 기준선(Baseline)으로 구축 후, 운영 안정화 뒤 C로 확장**.

---

## Roadmap: Prepare→Pilot→Build→Operate→Scale + KPI

1. **Prepare (1~2주)**

* 문서 카탈로그(Req_DAS/Req_AGI), Responsible Party 리스트 확정
* 룰테이블(Anchor/Offset) 초안 입력, 샘플 Voyage 1건으로 검증
  **KPI:** 필수 DocType 커버리지 100%, 룰 공란 0건

2. **Pilot (1~2주)**

* Voyage 1~2 대상으로 운영 리허설
* Export Pack(PDF/CSV) 포털 제출 흐름과 맞춤(파일명/RefNo) 
  **KPI:** Overdue 0, Evidence 누락률 < 5%

3. **Build (2~4주)**

* VBA Upsert/검증/대시보드 고도화(최소 기능만)
* Site별 차등 룰 완성(DAS/AGI)
  **KPI:** 수작업 행 추가 90% 감소

4. **Operate (상시)**

* 주간 운영(Responsible Party별 리뷰), 룰 변경은 `EffectiveFrom`으로만 반영
  **KPI:** On-time submission ≥ 95%, 평균 지연일수 ≤ 0.3일

5. **Scale (N 항차/다중 프로젝트)**

* Python 빌더로 프로젝트별 템플릿 자동 배포, 룰/문서 카탈로그 중앙관리
  **KPI:** 신규 프로젝트 셋업 리드타임 ≤ 1일

---

## QA 체크리스트

* 데이터(SSOT)

  * [ ] `tVoyage`의 7개 일정 필드 누락 없이 입력됨
  * [ ] `tDocType`의 Req_DAS/Req_AGI가 문서 요구사항과 일치
  * [ ] `tDueRule`에 DocType×Site 조합의 룰 누락 0
* 계산/룰

  * [ ] DueDate가 `AnchorDate + OffsetDays`로만 산출됨(하드코딩 금지)
  * [ ] 룰 버전(EffectiveFrom/To) 변경 시 과거 항차 재계산 정책이 정의됨
* 상태/증빙

  * [ ] Status=Submitted이면 EvidenceLink 또는 RefNo가 필수
  * [ ] Status=Approved이면 VerifiedAt 필수
* 대시보드

  * [ ] Responsible Party별 Overdue/DueSoon 수치가 원장과 일치
  * [ ] DAS/AGI 필수문서 누락이 즉시 식별됨
* Export Pack

  * [ ] PDF가 지정 템플릿/인쇄영역으로 정상 생성됨 ([Microsoft Learn][1])
  * [ ] CSV가 지정 FileFormat으로 저장됨(인코딩 정책 포함) ([Microsoft Learn][3])

---

## 가정

* (공통) 문서 제출은 ATLP/MTG 및(또는) Maqta 등 지정 포털을 통해 수행되며, 내부적으로 **Ref No.**로 추적 가능하다고 가정합니다. 
* (DAS Island) PTW 선행 기한이 **≥48h**, Hot Work 제한이 더 엄격하다는 운영 전제를 룰테이블에 반영합니다(OffsetDays 차등).
* (AGI Site) PTW 선행 기한이 **≥24h**, Pilotage Exemption 등 절차 차등이 존재한다는 전제를 룰테이블/DocType 적용여부로 반영합니다.
* 증빙은 워크북에 파일을 “첨부”하기보다, **표준 폴더 경로/하이퍼링크로 관리**(ONE FILE 용량/성능 리스크 최소화)합니다.

[1]: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat?utm_source=chatgpt.com "Workbook.ExportAsFixedFormat method (Excel)"
[2]: https://learn.microsoft.com/en-us/office/vba/api/excel.listobject?utm_source=chatgpt.com "ListObject object (Excel)"
[3]: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat?utm_source=chatgpt.com "XlFileFormat enumeration (Excel)"
