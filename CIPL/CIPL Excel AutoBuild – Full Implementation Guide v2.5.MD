**CIPL Excel AutoBuild – Full Implementation Guide v2.4 (Style-first / Project-SSoT)**

*(2026-01-06 · Asia/Dubai · Template-Fidelity 우선 · KPI Soft Warning)*

## 1) 목적(가장 중요)

* 사용자가 CIPL/BL/유사 종료 서류를 제공하면, **프로젝트 파일(템플릿/로고/빌더 스크립트)** 을 최대한 참조해 **“원본과 거의 동일한 형태”** 의 CIPL Excel(.xlsx)을 생성한다.
* 우선순위:

  1. **Style/Format Fidelity(서식 동일성)**
  2. Data Fill(채움률)
  3. Audit-grade KPI는 참고(경고)
* EN-KR 1L: *Pixel/format fidelity first; data second; KPI is advisory.*

---

## 2) SSoT(단일 진실원) — 프로젝트 파일 참조 규칙

### 2.1 SSoT 파일(필수)

* 로고: `SAMSUNGLOGO.PNG`
* 빌더: `make_cipl_set.py`
* 매퍼: `CIPL.py`
* 페이지 템플릿:

  * `COMMERCIAL INVOICE.PY`
  * `CI RIDER.PY`
  * `PACKING LIST.PY`
  * `PACKING LIST ATTACHED RIDER.PY`

### 2.2 SSoT 준수 원칙

* **서식은 템플릿 스크립트가 “정답”** 이다.
* OCR 결과가 불완전하더라도:

  * 셀 병합/테두리/열폭/회색영역/프린트영역/로고 위치는 템플릿 그대로 유지
  * 값은 NULL로 남기고 warnings로만 표시(임의 추정 금지)

---

## 3) 입력 문서 감지(Trigger & Routing)

### 3.1 문서 타입 판별(휴리스틱)

* CIPL류: “Commercial Invoice”, “Invoice No”, “Consignee”, “Total Amount”, “Incoterms” 등 키워드
* PL류: “Packing List”, “Packages”, “Gross Weight”, “Net Weight”, “Dimensions/CBM” 등
* BL류: “Bill of Lading”, “Shipper/Consignee/Notify”, “Port of Loading/Discharge”, “Vessel/Voyage”, “Container No” 등

### 3.2 Routing 규칙

* 입력이 CIPL/PL/BL/유사 종료 서류로 판단되면:

  * **항상 Excel AutoBuild 실행**(KPI와 무관)
* 단, 아래 “치명 실패”면 ZERO(중단)

---

## 4) Excel 산출물 구조(4-Page 고정)

* Sheet/Page 구성(고정):

  1. CI Page(Commercial Invoice)
  2. CI Rider(추가 조건/서명/리마크)
  3. PL Page(Packing List)
  4. PL Rider(Attached Rider/추가 페이지)

### 4.1 필드 매핑(개념)

* 공통 헤더 영역(회사/문서 제목/페이지 정보/로고)
* Parties(Shipper/Consignee/Notify)
* Refs(Invoice No, Date, BL No, PO/Project)
* Cargo Summary(Packages, Gross/Net, CBM, Origin/Destination)
* Line Items(품명/규격/수량/단가/금액, HS 후보 등)
* Remarks/Sign

### 4.2 데이터 소스 우선순위(충돌 시)

* 금액/통화/인보이스 정보: CI 우선
* 운송 정보(선박/항차/폴/포드): BL 우선
* 포장/중량/CBM 상세: PL 우선
* 충돌 발생 시:

  * Excel에는 **기본 우선순위대로 채우고**
  * `warnings[]`에 “상충”을 기록(결론 단정 금지)

---

## 5) Style-first 규격(서식 동일성 체크리스트)

> 아래 항목은 “데이터가 없어도 반드시 유지”되는 서식 요소다.

### 5.1 Page Setup(인쇄/레이아웃)

* 용지/여백/방향/스케일/프린트영역: 템플릿 스크립트 값 그대로
* 헤더/푸터, 반복 인쇄 행, 페이지 번호 표기 방식 고정

### 5.2 Geometry(행/열)

* 열폭(Column Width), 행높이(Row Height) 고정
* 병합셀(Merge) 패턴 고정(템플릿과 1:1)
* “틀(프레임) 테두리”는 최우선 유지

### 5.3 Borders(테두리)

* Thin/Thick, 색상(블랙/삼성블루) 규칙 템플릿 그대로
* “특정 셀 구간에 선 삽입” 같은 예외 규칙도 템플릿 우선
  (예: D5–E5, D8–E8 사이 라인 등은 템플릿 로직으로 고정)

### 5.4 Fonts/Alignment/Fill

* 폰트, 크기, Bold, 색상(빨강 경고 텍스트 포함)
* Alignment(가운데/좌/우, wrap, shrink, indent)
* 회색 영역/제목 바(Fill 색상) 고정

### 5.5 Logo

* `SAMSUNGLOGO.PNG` 사용
* 삽입 위치/크기/앵커(셀 기준) 고정
* 로고 비율 변형 금지

---

## 6) KPI 정책(Excel 전용) — Soft Warning

### 6.1 원칙

* KPI는 **Excel 생성 차단 조건이 아니다.**
* KPI는 **경고/등급/검토 필요성** 표시용이다.

### 6.2 권장 등급(표시만)

| Grade | MeanConf | TableAcc | Excel 생성 | Human Gate |
| ----- | -------: | -------: | -------- | ---------- |
| A     |   ≥ 0.92 | ≥ 98.00% | YES      | 선택         |
| B     |   ≥ 0.85 | ≥ 90.00% | YES      | 권장         |
| C     |   < 0.85 | < 90.00% | YES      | 필수         |

### 6.3 NumericIntegrity 처리

* NumericIntegrity가 불안정해도:

  * **서식 유지 + 값은 비우고 경고**
  * “합계 자동 계산” 같은 추정 로직은 금지(원문 근거 없으면)

---

## 7) ZERO(중단) — Excel 전용 “치명 실패”만

**아래 3가지에만 ZERO를 허용한다.**

1. **SSoT 로드 실패**: 템플릿/빌더/매퍼 파일 접근 불가, import 실패
2. **입력 파일 파손/열기 불가**: PDF/이미지 unreadable
3. **저장 실패**: xlsx write error(권한/경로/파일 잠김)

**중단 템플릿(Excel 전용)**

| 단계            | 이유            | 위험    | 요청데이터      | 다음조치        |
| ------------- | ------------- | ----- | ---------- | ----------- |
| Template Load | 프로젝트 파일 로드 실패 | 생성 불가 | 파일 재업로드/경로 | 로드 복구 후 재실행 |
| Input Open    | 원본 파손         | 생성 불가 | 원본 재전송     | 재시도         |
| Save          | 저장 실패         | 결과 유실 | 저장 경로/권한   | 경로 변경 후 재시도 |

---

## 8) 실행 인터페이스(표준)

### 8.1 기본 실행(권장)

```bash
python make_cipl_set.py --in <voyage_input.json> --out <CIPL.xlsx>
```

### 8.2 다항차(항차가 바뀌어도 자동 적용)

* JSON 입력은 “항차별 반복 구조”를 허용해야 하며,
* `CIPL.py`에서 commons → 4page dict 생성 후,
* `make_cipl_set.py`가 4개 페이지 템플릿에 값을 주입한다.

### 8.3 출력 파일명 규칙

* 우선: `CIPL_{Invoice_No or BL_No}_{YYYY-MM-DD}.xlsx`
* 누락 시: `CIPL_UNK_{YYYY-MM-DD_HHMM}.xlsx`
* 절대: 임의 추정으로 파일명을 “그럴듯하게” 만들지 않는다.

---

## 9) 검증(테스트) — “서식 동일성” 중심

### 9.1 필수 테스트(Regression)

* (1) Merge 패턴 비교: 템플릿과 완전 동일
* (2) Column Width/Row Height 비교
* (3) Border map 비교(셀별 테두리 스타일)
* (4) Print Area / Page Setup 비교
* (5) Logo 위치/크기 비교

### 9.2 데이터 테스트(부차)

* Invoice/BL/HS/Weights/Packages 주요 필드:

  * 문서 원문 ↔ Payload ↔ Excel 값 일치 확인
* 불일치 시:

  * Excel 값 강제 수정 금지
  * warnings에 기록 + Human Gate 안내

---

## 10) 운영 가이드(현업 Runbook)

### 10.1 현업 사용 흐름(무질의 기본)

1. CIPL/BL/PDF/이미지 업로드
2. AutoBuild 실행(자동)
3. 결과로:

   * CIPL Excel(.xlsx)
   * (선택) LDG_PAYLOAD/LDG_AUDIT 요약
4. KPI가 낮으면 “검토 필요”만 표시(생성은 유지)

### 10.2 현업 체크(3PL/WMS 관점)

* Consignee/Shipper, Incoterms, HS 후보, Packages/GW/NW/CBM
* berth/gate pass 일정과 연결되는 ETA/Port 정보(있을 경우)
* DEM/DET 리스크 판단은 별도 모듈에서 수행(Excel 생성 자체와 분리)

---

## 11) 운영 로그(감사 최소세트)

* `doc_hash`
* 사용 템플릿 버전(파일명/해시)
* 생성 결과 상태: PASS/WARN/FAIL
* warnings 요약(Top 5)
* (필요 시) key fields snapshot(Invoice/BL/HS/Weights/Packages)

---

## 12) 변경 관리(Changelog 규칙)

* 템플릿 변경(셀 병합/테두리/인쇄영역)은 “서식 Breaking Change”로 취급
* 변경 시:

  * 버전 증가(v2.4.x)
  * 회귀 테스트 결과 첨부(서식 동일성 5종)

---

원하시면 다음 단계로, 위 **Full Guide #2 기준에 맞춰** “프로젝트 스크립트(make_cipl_set.py/CIPL.py)에서 어떤 함수가 어떤 필드를 채우는지”를 **함수·필드 매핑 표**로 더 쪼개서(필드 100% 리스트업) 붙여드릴 수 있습니다.
