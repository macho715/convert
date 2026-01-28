# VBA 모듈 임포트 가이드

## TR_DocHub_AGI_2026.xlsx → .xlsm 변환 및 VBA 모듈 임포트

### Step 1: 파일 변환

1. Excel에서 `TR_DocHub_AGI_2026.xlsx` 열기
2. **File → Save As**
3. **File Format**: `Excel Macro-Enabled Workbook (*.xlsm)` 선택
4. 파일명: `TR_DocHub_AGI_2026.xlsm`로 저장

---

### Step 2: VBA Editor 열기

1. **Alt + F11** (VBA Editor 열기)
2. **Project Explorer**에서 `VBAProject (TR_DocHub_AGI_2026.xlsm)` 확인

---

### Step 3: VBA 모듈 임포트 (순서 중요)

**중요**: 다음 순서대로 임포트하세요.

#### 3.1 Control Tower (최우선 - 엔트리포인트)

1. **File → Import File...**
2. 파일 선택: `modControlTower.bas`
3. 확인: `modControlTower` 모듈이 Project Explorer에 표시되는지 확인

#### 3.2 TR Document Tracker (TR 기능)

1. **File → Import File...**
2. 파일 선택: `TR_DocTracker_VBA_Module.bas`
3. 확인: `TR_DocTracker_VBA_Module` 모듈 확인

#### 3.3 TR Python 연동

1. **File → Import File...**
2. 파일 선택: `modTRDocTracker.bas`
3. 확인: `modTRDocTracker` 모듈 확인

#### 3.4 DocGap Macros

1. **File → Import File...**
2. 파일 선택: `DocGapMacros_v3_1.bas`
3. 확인: `DocGapMacros_v3_1` 모듈 확인

---

### Step 4: ThisWorkbook 단축키 추가

1. **Project Explorer**에서 `ThisWorkbook` 더블클릭
2. 기존 코드가 있다면 확인 후, 아래 코드 추가:

```vba
' ThisWorkbook_Shortcuts.bas의 내용을 복사하여 붙여넣기
```

또는:

1. **File → Import File...**
2. 파일 선택: `ThisWorkbook_Shortcuts.bas` (단, 이미 ThisWorkbook 코드가 있다면 수동으로 병합 필요)

---

### Step 5: Document_Tracker 시트 이벤트 코드 추가

1. **Project Explorer**에서 `Sheets("Document_Tracker")` 더블클릭
2. `Document_Tracker_Sheet_Code.txt`의 내용을 복사하여 붙여넣기
3. 코드는 `Worksheet_Change` 이벤트로 자동 인식됨

---

### Step 6: 매크로 보안 설정

1. **File → Options → Trust Center → Trust Center Settings**
2. **Macro Settings**: 
   - "Disable all macros with notification" (권장)
   - 또는 "Enable all macros" (개발 환경에서만)
3. **Trusted Locations**: 이 파일이 있는 폴더를 신뢰 위치로 추가 (선택사항)

---

### Step 7: 검증 테스트

#### 7.1 단축키 테스트

1. **Ctrl + Shift + R**: `RefreshAll_ControlTower` 실행 확인
2. 메시지 박스에 "Control Tower Refresh completed" 표시 확인

#### 7.2 Dashboard 버튼 매핑 확인

1. **Dashboard** 시트로 이동
2. 각 버튼 우클릭 → **Assign Macro** 확인:
   - **Refresh All** → `RefreshAll_ControlTower`
   - **Export PDF** → `EXP_ExportToPDF`
   - **Send Reminder** → `TR_Draft_Reminder_Emails`
   - **Python Refresh** → `RefreshAll_WithPython`

#### 7.3 VBA_Pasteboard 확인

1. **VBA_Pasteboard** 시트 확인
2. 모든 모듈 코드가 텍스트로 저장되어 있는지 확인 (참고용)

---

## 문제 해결

### 문제: "Sub or Function not defined" 오류

**원인**: 모듈 임포트 순서 문제 또는 모듈 누락

**해결**:
1. 모든 모듈이 임포트되었는지 확인
2. `modControlTower`가 최우선으로 임포트되었는지 확인
3. 프로시저명이 정확한지 확인 (TR_, DG_, EXP_ 접두어)

### 문제: 단축키가 작동하지 않음

**원인**: ThisWorkbook 코드 누락 또는 Workbook_Open 이벤트 미실행

**해결**:
1. ThisWorkbook 코드 확인
2. 파일을 닫고 다시 열어 Workbook_Open 이벤트 트리거
3. 수동으로 `ThisWorkbook.Workbook_Open` 실행 (Immediate Window에서)

### 문제: "Worksheet does not exist" 오류

**원인**: 시트명 불일치

**해결**:
1. 시트명 확인:
   - `Dashboard` (대소문자 구분)
   - `Document_Tracker`
   - `Inputs`
   - `Executive_Summary`
   - `OFCO_Req_1_15`
   - `NOC_Req_1_6`
2. 시트명이 정확한지 확인

---

## 완료 체크리스트

- [ ] `.xlsx` → `.xlsm` 변환 완료
- [ ] `modControlTower.bas` 임포트
- [ ] `TR_DocTracker_VBA_Module.bas` 임포트
- [ ] `modTRDocTracker.bas` 임포트
- [ ] `DocGapMacros_v3_1.bas` 임포트
- [ ] `ThisWorkbook` 단축키 코드 추가
- [ ] `Document_Tracker` 시트 이벤트 코드 추가
- [ ] 단축키 테스트 (Ctrl+Shift+R/P/E)
- [ ] Dashboard 버튼 매핑 확인
- [ ] `RefreshAll_ControlTower` 실행 테스트

---

## 다음 단계

1. **Inputs 시트 설정**: Voyage 1 날짜 입력
2. **Doc_Matrix 설정**: 문서 요구사항 및 리드타임 설정
3. **Party_Contacts 설정**: 담당자 연락처 입력
4. **첫 번째 Refresh 실행**: `RefreshAll_ControlTower` 또는 Ctrl+Shift+R

---

**작성일**: 2026-01-19  
**버전**: 1.0
