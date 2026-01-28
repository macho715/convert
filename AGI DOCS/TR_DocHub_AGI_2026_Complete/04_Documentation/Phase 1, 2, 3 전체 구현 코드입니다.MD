Phase 1, 2, 3 전체 구현 코드입니다. 파일별 변경사항을 정리했습니다.

## Phase 1: VBA 모듈 리네임

### 1.1 modTRDocTracker.bas 변경사항

```vba
Attribute VB_Name = "modTRDocTracker"
'===============================================================================
' HVDC TR TRANSPORTATION - DOCUMENT TRACKER VBA MODULE
' Version: 2.0 (Integrated - TR_ prefix)
' Date: 2026-01-XX
'===============================================================================

Option Explicit

'===============================================================================
' 1. STATUS CONDITIONAL FORMATTING - 상태별 색상 자동 적용
'===============================================================================
Public Sub TR_ApplyStatusFormatting()
    ' [기존 ApplyStatusFormatting() 코드 그대로]
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Document_Tracker 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim statusCols As Variant
    statusCols = Array(8, 11, 14, 17)
    
    Dim col As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For Each col In statusCols
        For Each cell In ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
            Select Case cell.Value
                Case "Complete"
                    cell.Interior.Color = RGB(144, 238, 144)
                    cell.Font.Color = RGB(0, 100, 0)
                    cell.Font.Bold = True
                Case "In Progress"
                    cell.Interior.Color = RGB(255, 255, 153)
                    cell.Font.Color = RGB(128, 128, 0)
                    cell.Font.Bold = False
                Case "Not Started"
                    cell.Interior.Color = RGB(255, 204, 204)
                    cell.Font.Color = RGB(139, 0, 0)
                    cell.Font.Bold = False
                Case "Pending Review"
                    cell.Interior.Color = RGB(173, 216, 230)
                    cell.Font.Color = RGB(0, 0, 139)
                    cell.Font.Bold = False
                Case "N/A"
                    cell.Interior.Color = RGB(211, 211, 211)
                    cell.Font.Color = RGB(128, 128, 128)
                    cell.Font.Bold = False
                Case Else
                    cell.Interior.ColorIndex = xlNone
                    cell.Font.ColorIndex = xlAutomatic
                    cell.Font.Bold = False
            End Select
        Next cell
    Next col
    
    Application.ScreenUpdating = True
    MsgBox "상태별 색상 서식이 적용되었습니다.", vbInformation
End Sub

'===============================================================================
' 2. PROGRESS CALCULATOR - 진행률 계산
'===============================================================================
Public Sub TR_CalculateProgress()
    ' [기존 CalculateProgress() 코드 그대로, 이름만 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' 3. PARTY-WISE PROGRESS CALCULATOR
'===============================================================================
Public Sub TR_CalculatePartyProgress()
    ' [기존 CalculatePartyProgress() 코드 그대로, 이름만 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' 4. HIGHLIGHT OVERDUE DOCUMENTS
'===============================================================================
Public Sub TR_HighlightOverdue()
    ' [기존 HighlightOverdue() 코드 그대로, 이름만 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' 5. EXPORT TO PDF
'===============================================================================
Public Sub EXP_ExportToPDF()
    ' [기존 ExportToPDF() 코드 그대로, 이름만 EXP_로 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' 6. SEND EMAIL REMINDER
'===============================================================================
Public Sub EXP_SendEmailReminder()
    ' [기존 SendEmailReminder() 코드 그대로, 이름만 EXP_로 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' 7. AUTO-REFRESH ALL (Deprecated - Use RefreshAll_ControlTower instead)
'===============================================================================
Public Sub TR_AutoRefreshAll()
    ' [기존 AutoRefreshAll() 코드 그대로, 이름만 변경]
    ' 참고: 이제는 RefreshAll_ControlTower()를 사용하세요
    Call TR_ApplyStatusFormatting
    Call TR_CalculateProgress
    Call TR_CalculatePartyProgress
    Call TR_HighlightOverdue
    
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    wsDash.Cells(3, 1).Value = "Last Updated: " & Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    MsgBox "모든 데이터가 업데이트되었습니다!", vbInformation
End Sub

'===============================================================================
' 8. REFRESH DOCUMENT TRACKER (Python External)
'===============================================================================
Public Sub TR_Refresh_Document_Tracker()
    On Error GoTo EH

    Dim py As String: py = GetPythonExe()
    Dim script As String: script = GetScriptPath()
    Dim wbFile As String: wbFile = ThisWorkbook.FullName

    Dim cmd As String
    cmd = py & " " & Chr(34) & script & Chr(34) & " --refresh " & Chr(34) & wbFile & Chr(34)

    Shell cmd, vbHide
    MsgBox "Refresh started. Re-open workbook after script completes.", vbInformation
    Exit Sub

EH:
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub

'===============================================================================
' 9. DRAFT REMINDER EMAILS
'===============================================================================
Public Sub TR_Draft_Reminder_Emails()
    ' [기존 Draft_Reminder_Emails() 코드 그대로, 이름만 변경]
    ' ... (전체 코드 동일)
End Sub

'===============================================================================
' Helper Functions (unchanged)
'===============================================================================
Private Function GetPythonExe() As String
    GetPythonExe = "py"
End Function

Private Function GetScriptPath() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    GetScriptPath = wbPath & Application.PathSeparator & "create_tr_document_tracker_v2.py"
End Function

'===============================================================================
' Helper: Get Config Value (shared with modControlTower)
'===============================================================================
Public Function GetCfgValue(key As String, defaultValue As Variant) As Variant
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    If wsConfig Is Nothing Then
        GetCfgValue = defaultValue
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = wsConfig.Range("A:A").Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetCfgValue = defaultValue
    Else
        GetCfgValue = rng.Offset(0, 1).Value
        If IsEmpty(GetCfgValue) Then GetCfgValue = defaultValue
    End If
End Function
```

### 1.2 DocGapMacros_v3_1.bas 변경사항

```vba
Attribute VB_Name = "modDocGapMacros"
Option Explicit

'=====================
' Doc Gap Tracker Macros (v3.1) - Integrated with TR Tracker
' All macros prefixed with DG_ to avoid conflicts
'=====================

Public Sub DG_RefreshAll()
    'Recalculate
    Application.CalculateFull
End Sub

Public Sub DG_FilterMissing()
    'Filter current sheet by Status = Missing (assumes headers on row 2)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode = False Then Exit Sub
    ws.Range("A2").AutoFilter Field:=4, Criteria1:="Missing"
End Sub

Public Sub DG_ClearFilters()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If
End Sub

Public Sub DG_StampUpdated()
    'Stamp last updated time in Executive_Summary!C1
    On Error Resume Next
    With ThisWorkbook.Sheets("Executive_Summary")
        .Range("C1").Value = "Last updated: " & Format(Now, "dd-mmm-yy hh:nn")
    End With
    On Error GoTo 0
End Sub

Public Sub DG_ApplyScenarioToManual()
    'Copy auto schedule dates (Column B) to manual (Column C) and set scenario to CUSTOM
    On Error Resume Next
    With ThisWorkbook.Sheets("Inputs")
        .Range("C3").Value = .Range("B3").Value
        .Range("C4").Value = .Range("B4").Value
        .Range("C5").Value = .Range("B5").Value
        .Range("C6").Value = .Range("B6").Value
        .Range("B9").Value = "CUSTOM"
    End With
    On Error GoTo 0
    Application.CalculateFull
End Sub
```

### 1.3 modControlTower.bas (신규 생성)

```vba
Attribute VB_Name = "modControlTower"
Option Explicit

'===============================================================================
' CONTROL TOWER - Single Entry Point for All Refresh Operations
' Version: 1.0
' Date: 2026-01-XX
'===============================================================================

'===============================================================================
' Application State Guard (Performance & Safety)
'===============================================================================
Private Sub AppStateGuard_Begin()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub

Private Sub AppStateGuard_End()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

'===============================================================================
' MAIN ENTRY POINT: RefreshAll_ControlTower()
' Runs all in-workbook refresh operations (no Python external calls)
'===============================================================================
Public Sub RefreshAll_ControlTower()
    On Error GoTo EH
    
    Dim startTime As Double
    startTime = Timer
    
    AppStateGuard_Begin
    
    ' 1. DocGap operations
    Call DG_RefreshAll
    Call DG_StampUpdated
    
    ' 2. TR Tracker operations
    Call TR_ApplyStatusFormatting
    Call TR_CalculateProgress
    Call TR_CalculatePartyProgress
    Call TR_HighlightOverdue
    
    ' 3. Dashboard KPI update (D-7/D-3/D-1)
    Call UpdateDashboardKPIs
    
    ' 4. Update timestamp
    Call UpdateDashboardTimestamp
    
    AppStateGuard_End
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox "Control Tower Refresh 완료" & vbCrLf & _
           "실행 시간: " & Format(elapsed, "0.00") & "초", vbInformation
    
    Exit Sub
    
EH:
    AppStateGuard_End
    MsgBox "오류 발생: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

'===============================================================================
' Dashboard KPI 업데이트 (D-7/D-3/D-1)
'===============================================================================
Private Sub UpdateDashboardKPIs()
    Dim wsDash As Worksheet
    Dim wsConfig As Worksheet
    
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    
    If wsDash Is Nothing Or wsConfig Is Nothing Then
        Exit Sub
    End If
    
    ' Get threshold values from Config (with fallback defaults)
    Dim amberDays As Long: amberDays = CLng(GetCfgValue("Amber_Threshold_Days", 7))
    Dim redDays As Long: redDays = CLng(GetCfgValue("Red_Threshold_Days", 3))
    Dim criticalDays As Long: criticalDays = CLng(GetCfgValue("Critical_Threshold_Days", 1))
    
    ' D-7 Count (Amber) - Dashboard B11
    On Error Resume Next
    wsDash.Range("B11").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & amberDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""))"
    
    ' D-3 Count (Red) - Dashboard B12
    wsDash.Range("B12").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & redDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""))"
    
    ' D-1 Count (Critical) - Dashboard B13
    wsDash.Range("B13").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & criticalDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""))"
    On Error GoTo 0
End Sub

'===============================================================================
' Update Dashboard Timestamp
'===============================================================================
Private Sub UpdateDashboardTimestamp()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    If Not wsDash Is Nothing Then
        wsDash.Cells(3, 1).Value = "Last Updated: " & Format(Now, "YYYY-MM-DD HH:MM:SS")
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' Helper: Get Config Value (uses modTRDocTracker function if available)
'===============================================================================
Private Function GetCfgValue(key As String, defaultValue As Variant) As Variant
    ' Try to use modTRDocTracker's function first
    On Error Resume Next
    GetCfgValue = modTRDocTracker.GetCfgValue(key, defaultValue)
    If Err.Number = 0 Then
        Exit Function
    End If
    
    ' Fallback: direct Config sheet lookup
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("Config")
    If wsConfig Is Nothing Then
        GetCfgValue = defaultValue
        Exit Function
    End If
    
    Dim rng As Range
    Set rng = wsConfig.Range("A:A").Find(key, LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then
        GetCfgValue = defaultValue
    Else
        GetCfgValue = rng.Offset(0, 1).Value
        If IsEmpty(GetCfgValue) Then GetCfgValue = defaultValue
    End If
End Function

'===============================================================================
' External Refresh (Python) - Separate entry point with user prompt
'===============================================================================
Public Sub RefreshAll_WithPython()
    Dim response As VbMsgBoxResult
    response = MsgBox("Python 스크립트를 실행하여 Document_Tracker를 재생성합니다." & vbCrLf & _
                     "완료 후 파일을 다시 열어야 합니다." & vbCrLf & vbCrLf & _
                     "계속하시겠습니까?", vbYesNo + vbQuestion, "External Refresh")
    
    If response = vbYes Then
        Call TR_Refresh_Document_Tracker
    End If
End Sub
```

---

## Phase 2: Python 빌더 수정

### 2.1 create_tr_document_tracker_v2.py 변경사항

Config 키 추가 (line 174 근처):

```python
cfg_rows = [
    ("DocCutoff_Days", 3, "Doc Deadline = MZP Departure - DocCutoff_Days"),
    ("LandPermitLead_Days", 7, "Land Permit By = AGI Arrival - LandPermitLead_Days"),
    ("DueSoon_Threshold_Days", 7, "Dashboard/CF: upcoming items within N days"),
    ("Amber_Threshold_Days", 7, "Dashboard KPI: D-7 Amber warning threshold"),
    ("Red_Threshold_Days", 3, "Dashboard KPI: D-3 Red warning threshold"),
    ("Critical_Threshold_Days", 1, "Dashboard KPI: D-1 Critical alert threshold"),
    ("Default_Year", dt.date.today().year, "Year used for sample schedule MM-DD parsing"),
]
```

Dashboard KPI 행 추가 (line 600 근처, B10 다음):

```python
    ws_dash["B10"] = "=IFERROR(B7/B6,0)"
    ws_dash["B10"].number_format = "0%"

    # Add D-7/D-3/D-1 KPI rows
    ws_dash["A11"] = "D-7 Count (Amber)"
    ws_dash["A12"] = "D-3 Count (Red)"
    ws_dash["A13"] = "D-1 Count (Critical)"
    
    ws_dash["B11"] = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*(tblTracker[SUBMISSION DATE]<=TODAY()+Config!$B$7)*(tblTracker[Status]<>\"Approved\")*(tblTracker[Status]<>\"Submitted\")*(tblTracker[Status]<>\"Not Required\")*(tblTracker[SUBMISSION DATE]<>\"\"))"
    ws_dash["B12"] = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*(tblTracker[SUBMISSION DATE]<=TODAY()+Config!$B$8)*(tblTracker[Status]<>\"Approved\")*(tblTracker[Status]<>\"Submitted\")*(tblTracker[Status]<>\"Not Required\")*(tblTracker[SUBMISSION DATE]<>\"\"))"
    ws_dash["B13"] = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*(tblTracker[SUBMISSION DATE]<=TODAY()+Config!$B$9)*(tblTracker[Status]<>\"Approved\")*(tblTracker[Status]<>\"Submitted\")*(tblTracker[Status]<>\"Not Required\")*(tblTracker[SUBMISSION DATE]<>\"\"))"

    for rr in range(6, 14):  # Changed from range(6, 11)
        ws_dash[f"A{rr}"].border = border_thin
        ws_dash[f"B{rr}"].border = border_thin
        ws_dash[f"B{rr}"].font = Font(bold=True, size=11)
```

VBA_Code → VBA_Pasteboard 통합 (line 797 근처):

```python
    # =========================
    # VBA_Pasteboard (unified code repository)
    # =========================
    if "VBA_Pasteboard" not in wb.sheetnames:
        ws_vba = wb.create_sheet("VBA_Pasteboard")
        row_vba = 1
        ws_vba.cell(row_vba, 1).value = "VBA Code Repository (Copy & Paste into VBA Editor)"
        ws_vba.cell(row_vba, 1).font = font_title
        row_vba += 2
    else:
        ws_vba = wb["VBA_Pasteboard"]
        row_vba = ws_vba.max_row + 3
        ws_vba.cell(row_vba, 1).value = "=" * 50
        row_vba += 1
        ws_vba.cell(row_vba, 1).value = "TR Document Tracker Macros"
        row_vba += 1
        ws_vba.cell(row_vba, 1).value = "=" * 50
        row_vba += 1

    vba_text = '''Attribute VB_Name = "modTRDocTracker"
'===============================================================================
' HVDC TR TRANSPORTATION - DOCUMENT TRACKER VBA MODULE
'===============================================================================
Option Explicit

Private Function GetPythonExe() As String
    GetPythonExe = "py"
End Function

Private Function GetScriptPath() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    GetScriptPath = wbPath & Application.PathSeparator & "create_tr_document_tracker_v2.py"
End Function

Public Sub TR_Refresh_Document_Tracker()
    On Error GoTo EH
    Dim py As String: py = GetPythonExe()
    Dim script As String: script = GetScriptPath()
    Dim wbFile As String: wbFile = ThisWorkbook.FullName
    Dim cmd As String
    cmd = py & " " & Chr(34) & script & Chr(34) & " --refresh " & Chr(34) & wbFile & Chr(34)
    Shell cmd, vbHide
    MsgBox "Refresh started. Re-open workbook after script completes.", vbInformation
    Exit Sub
EH:
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub
'''

    # Write VBA code line by line
    vba_lines = vba_text.split('\n')
    for line in vba_lines:
        ws_vba.cell(row_vba, 1).value = line
        row_vba += 1
    
    # Format: Make column A wider and wrap text
    ws_vba.column_dimensions['A'].width = 120
    for row_num in range(1, ws_vba.max_row + 1):
        cell = ws_vba.cell(row_num, 1)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        if row_num == 1:
            cell.font = font_title
```

### 2.2 build_docgap_v3_1_operational.py 변경사항

VBA_Module → VBA_Pasteboard 통합 (line 219 근처):

```python
    # ---------------- VBA_Pasteboard sheet (unified) ----------------
    # Create or append to VBA_Pasteboard
    if "VBA_Pasteboard" in wb.sheetnames:
        ws_vba = wb["VBA_Pasteboard"]
        # Append DocGap section
        row = ws_vba.max_row + 3
        ws_vba.cell(row, 1).value = "=" * 50
        row += 1
        ws_vba.cell(row, 1).value = "DocGap Tracker Macros (v3.1)"
        row += 1
        ws_vba.cell(row, 1).value = "=" * 50
        row += 1
    else:
        ws_vba = wb.create_sheet("VBA_Pasteboard")
        row = 1
        ws_vba.cell(row, 1).value = "VBA Code Repository (Copy & Paste into VBA Editor)"
        ws_vba.cell(row, 1).font = Font(bold=True, size=16, color="1E3A5F")
        row += 2
    
    # Write change log header
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_vba.cell(row, 1).value = f"v3.1 Operational - {timestamp}"
    row += 1
    ws_vba.cell(row, 1).value = "=" * 50
    row += 1
    ws_vba.cell(row, 1).value = ""
    row += 1
    
    # Write change log items (same as before)
    change_items = [
        "주요 변경사항:",
        "",
        "1. Schedule Scenario 선택 기능 추가 (Inputs!B9)",
        "   - SCN-01 ~ SCN-05, CUSTOM 옵션",
        "   - 시나리오 테이블 (A12:E17)",
        "",
        "2. Manual Date 입력 컬럼 추가 (Inputs Column C)",
        "   - Auto 날짜 공식 (Column B) + Manual 입력 (Column C)",
        "   - CUSTOM 선택 시 Manual 값 사용",
        "",
        "3. Lead Time & Anchor 매핑 테이블 추가 (Inputs A21:C40)",
        "   - Submit To → Default Lead Time → Anchor Key 매핑",
        "   - HSE, HM, Maqta Gateway, Port Ops, Customs, AD Maritime",
        "",
        "4. OFCO_Req_1_15 / NOC_Req_1_6 시트 확장",
        "   - Lead Time Default (WD) - Column I",
        "   - Lead Time Override (WD) - Column J",
        "   - Lead Time Effective (WD) - Column K",
        "   - Anchor (Auto) - Column L",
        "   - Target Submit Date - Column M",
        "",
        "5. VBA 매크로 추가 (DG_ 접두어)",
        "   - DG_RefreshAll: 전체 재계산",
        "   - DG_FilterMissing: Missing 상태 필터",
        "   - DG_ClearFilters: 필터 제거",
        "   - DG_StampUpdated: 업데이트 시간 기록",
        "   - DG_ApplyScenarioToManual: 시나리오 → 수동 복사",
        "",
        "사용법:",
        "- VBA Editor에서 이 시트의 매크로 코드를 복사하여 모듈에 붙여넣기",
        "- 또는 별도 .bas 파일로 저장 후 Import",
    ]
    
    for item in change_items:
        ws_vba.cell(row, 1).value = item
        row += 1
    
    # Write VBA macros section
    row += 1
    ws_vba.cell(row, 1).value = "=" * 50
    row += 1
    ws_vba.cell(row, 1).value = "VBA MACRO CODE"
    row += 1
    ws_vba.cell(row, 1).value = "=" * 50
    row += 1
    ws_vba.cell(row, 1).value = "Copy the code below and paste into VBA Editor Module:"
    row += 1
    ws_vba.cell(row, 1).value = ""
    row += 1
    
    # Write VBA code line by line
    vba_lines = VBA_MACROS.split('\n')
    for line in vba_lines:
        ws_vba.cell(row, 1).value = line
        row += 1
    
    # Format: Make column A wider and wrap text
    ws_vba.column_dimensions['A'].width = 100
    for row_num in range(1, ws_vba.max_row + 1):
        cell = ws_vba.cell(row_num, 1)
        cell.alignment = Alignment(wrap_text=True, vertical='top')
```

VBA_MACROS 문자열도 업데이트 (line 25 근처):

```python
# VBA 매크로 코드 (DG_ 접두어 적용)
VBA_MACROS = """Option Explicit

'=====================
' Doc Gap Tracker Macros (v3.1) - Integrated
'=====================

Public Sub DG_RefreshAll()
    'Recalculate
    Application.CalculateFull
End Sub

Public Sub DG_FilterMissing()
    'Filter current sheet by Status = Missing (assumes headers on row 2)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode = False Then Exit Sub
    ws.Range("A2").AutoFilter Field:=4, Criteria1:="Missing"
End Sub

Public Sub DG_ClearFilters()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If
End Sub

Public Sub DG_StampUpdated()
    'Stamp last updated time in Executive_Summary!C1
    With ThisWorkbook.Sheets("Executive_Summary")
        .Range("C1").Value = "Last updated: " & Format(Now, "dd-mmm-yy hh:nn")
    End With
End Sub

Public Sub DG_ApplyScenarioToManual()
    'Copy auto schedule dates (Column B) to manual (Column C) and set scenario to CUSTOM
    With ThisWorkbook.Sheets("Inputs")
        .Range("C3").Value = .Range("B3").Value
        .Range("C4").Value = .Range("B4").Value
        .Range("C5").Value = .Range("B5").Value
        .Range("C6").Value = .Range("B6").Value
        .Range("B9").Value = "CUSTOM"
    End With
    Application.CalculateFull
End Sub
"""
```

---

## Phase 3: Inputs ↔ Voyage_Schedule 연동

### 3.1 create_tr_document_tracker_v2.py 변경사항

Inputs 시트 읽기 로직 추가 (Voyage_Schedule 생성 전, line 202 근처):

```python
    # =========================
    # Sheet: Voyage_Schedule (Input) - Linked to Inputs if available
    # =========================
    ws_voy = wb.create_sheet("Voyage_Schedule")
    ws_voy["A1"] = "Voyage Schedule (Input)"
    ws_voy["A1"].font = font_title
    ws_voy["A2"] = f"Last Updated: {now_str()}"
    ws_voy["A2"].font = Font(italic=True, size=10)
    
    # Check if Inputs sheet exists (from DocGap integration)
    has_inputs = "Inputs" in wb.sheetnames
    if has_inputs:
        ws_inputs = wb["Inputs"]
        ws_voy["A2"].Value = ws_voy["A2"].Value + " | Linked to Inputs sheet (scenario-based)"
```

Voyage_Schedule 수식에서 Inputs 참조 (line 241 근처):

```python
        # Date formulas - link to Inputs if available, else use direct dates
        if has_inputs:
            # Use Inputs!B3:B6 (Auto dates from scenario)
            ws_voy.cell(row=r, column=4, value=f"=IFERROR(Inputs!$B$3,\"\"")  # MZP Arrival
            ws_voy.cell(row=r, column=5, value=f"=IFERROR(Inputs!$B$4,\"\"")  # Load-out Start
            ws_voy.cell(row=r, column=6, value=f"=IFERROR(Inputs!$B$6,\"\"")  # MZP Departure
            ws_voy.cell(row=r, column=7, value=f"=IFERROR(Inputs!$B$6+1,\"\"")  # AGI Arrival (Departure + 1 day)
        else:
            # Original direct date assignment
            ws_voy.cell(row=r, column=4, value=mmdd_to_date(extract_start_mmdd(mzp_arr) or "", year))
            ws_voy.cell(row=r, column=5, value=mmdd_to_date(extract_start_mmdd(loadout) or "", year))
            ws_voy.cell(row=r, column=6, value=mmdd_to_date(extract_start_mmdd(mzp_dep) or "", year))
            ws_voy.cell(row=r, column=7, value=mmdd_to_date(extract_start_mmdd(agi_arr) or "", year))
```

Inputs 시트 생성 옵션 추가 (build_template 함수 시작 부분):

```python
def build_template(output_path: Path) -> Path:
    wb = Workbook()
    
    # ... (기존 스타일 정의)
    
    # =========================
    # Sheet: Inputs (DocGap-style, optional integration)
    # =========================
    # Note: This sheet is created if DocGap integration is desired
    # For standalone TR Tracker, this can be skipped
    create_inputs_sheet = True  # Set to False to skip Inputs sheet
    
    if create_inputs_sheet:
        ws_inputs = wb.create_sheet("Inputs")
        ws_inputs["A1"] = "Schedule Inputs (DocGap Integration)"
        ws_inputs["A1"].font = font_title
        
        # Scenario selection
        ws_inputs.cell(9, 1).value = "Schedule Scenario (Select)"
        ws_inputs.cell(9, 2).value = "SCN-01"
        ws_inputs.cell(9, 3).value = "(Select SCN code; choose CUSTOM to use manual values)"
        
        # Scenario table headers
        ws_inputs.cell(11, 1).value = "Scenario"
        ws_inputs.cell(11, 2).value = "Arrival"
        ws_inputs.cell(11, 3).value = "RoRo Start"
        ws_inputs.cell(11, 4).value = "RoRo End"
        ws_inputs.cell(11, 5).value = "Departure"
        
        # Default scenario (will be populated by build_docgap_v3_1_operational.py)
        ws_inputs.cell(12, 1).value = "SCN-01"
        # Dates will be set by DocGap builder
        
        # Manual/Auto columns
        ws_inputs.cell(2, 2).value = "Value (Auto)"
        ws_inputs.cell(2, 3).value = "Value (Manual)"
        
        # Auto date formulas (will reference scenario table)
        ws_inputs.cell(3, 2).value = '=IF($B$9="CUSTOM",$C3,VLOOKUP($B$9,$A$12:$E$17,2,FALSE))'
        ws_inputs.cell(4, 2).value = '=IF($B$9="CUSTOM",$C4,VLOOKUP($B$9,$A$12:$E$17,3,FALSE))'
        ws_inputs.cell(5, 2).value = '=IF($B$9="CUSTOM",$C5,VLOOKUP($B$9,$A$12:$E$17,4,FALSE))'
        ws_inputs.cell(6, 2).value = '=IF($B$9="CUSTOM",$C6,VLOOKUP($B$9,$A$12:$E$17,5,FALSE))'
        
        # Labels
        ws_inputs.cell(3, 1).value = "Vessel Arrival to Mina Zayed (UTC+4)"
        ws_inputs.cell(4, 1).value = "RoRo / Load-out Start (UTC+4)"
        ws_inputs.cell(5, 1).value = "RoRo / Load-out End (UTC+4)"
        ws_inputs.cell(6, 1).value = "Vessel Departure from Mina Zayed (UTC+4)"
```

---

## ThisWorkbook 코드 (단축키 설정)

```vba
Private Sub Workbook_Open()
    ' Ctrl+Shift+R = RefreshAll_ControlTower (in-workbook only)
    Application.OnKey "^+R", "RefreshAll_ControlTower"
    
    ' Ctrl+Shift+P = Python Refresh (external, requires reopen)
    Application.OnKey "^+P", "RefreshAll_WithPython"
    
    ' Ctrl+Shift+E = Draft Email Reminders
    Application.OnKey "^+E", "TR_Draft_Reminder_Emails"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Remove shortcuts
    Application.OnKey "^+R"
    Application.OnKey "^+P"
    Application.OnKey "^+E"
End Sub
```

---

## 실행 순서

1. Python 빌더 실행:
   ```bash
   python create_tr_document_tracker_v2.py --output TR_Tracker_Template.xlsx
   python build_docgap_v3_1_operational.py --in TR_Tracker_Template.xlsx --out HVDC_Integrated.xlsx
   ```

2. Excel에서 .xlsm 변환 및 VBA Import:
   - 파일 → 다른 이름으로 저장 → `.xlsm`
   - Alt+F11 → 파일 → 가져오기:
     - `modControlTower.bas`
     - `modTRDocTracker.bas` (리네임된 버전)
     - `modDocGapMacros.bas` (리네임된 버전)
   - Document_Tracker 시트 코드에 `Document_Tracker_Sheet_Code.txt` 붙여넣기
   - ThisWorkbook에 단축키 코드 추가

3. 테스트:
   - Ctrl+Shift+R 실행
   - Dashboard KPI 확인 (D-7/D-3/D-1)
   - Inputs 시나리오 변경 시 Voyage_Schedule 자동 업데이트 확인

모든 변경사항을 적용할 준비가 되었습니다. Agent 모드로 전환하시면 파일 편집을 진행할 수 있습니다.