Attribute VB_Name = "mod_JPT71_Auto"
Option Explicit

'==========================================================
' JPT71 AutoSuite - COMPLETE MODULE
'
' 핵심 목표:
' 1) Plan 시트의 INPUT(Q:R) + tblPlan(A:O)만 수정하면 모든 시트 자동 갱신
'    - Shipping_List / Mail_Draft : 수식 링크로 즉시 반영 (Workbook 계산)
'    - Cross_Gantt / Gantt_Chart / Calendar_View : VBA가 자동 재생성/렌더링
' 2) LOG: 셀 변경 이력(시간/사용자/시트/주소/이전값/신규값)
' 3) LOG_SNAPSHOT: Plan INPUT(Q:R) 블록 버전 스냅샷 저장(감사 대응)
' 4) Gantt Chart: Stacked Bar(Offset+Duration) 자동 생성/갱신
' 5) Calendar_View: Vertex42 스타일 월간 캘린더 뷰(일자별 n개 이벤트 + 기간막대)
' 6) Settings: 사이트/포트/휴일/옵션(월/년/일자별 최대 이벤트 수)
' 7) Export_FINAL_Values(): 외부 발송용 "값만" 파일 생성
'
' ---------------------------------------------------------
' ★ 반드시 함께 넣어야 하는 이벤트 코드
'   (1) Plan 시트 코드(워크시트 모듈):
'
'   Option Explicit
'   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'       If Target.Cells.CountLarge = 1 Then
'           mod_JPT71_Auto.gPrevAddr = Target.Address(False, False)
'           mod_JPT71_Auto.gPrevValue = Target.Value
'       End If
'   End Sub
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       On Error GoTo EH
'       Dim rngWatch As Range
'       Set rngWatch = Union(Me.Range("Q2:R200"), Me.ListObjects("tblPlan").Range)
'       If Intersect(Target, rngWatch) Is Nothing Then Exit Sub
'
'       Application.EnableEvents = False
'       mod_JPT71_Auto.Init_LOG
'       mod_JPT71_Auto.Init_Snapshot
'
'       If Target.Cells.CountLarge = 1 Then
'           mod_JPT71_Auto.Log_Change Me.Name, Target.Address(False, False), mod_JPT71_Auto.gPrevValue, Target.Value
'       Else
'           mod_JPT71_Auto.Log_Event Me.Name, "MULTI", "", "", "Multi-cell change (" & Target.Cells.CountLarge & " cells)"
'       End If
'
'       mod_JPT71_Auto.Schedule_Refresh 2   '2초 디바운스
' EH:
'       Application.EnableEvents = True
'   End Sub
'
'   (2) ThisWorkbook 코드:
'
'   Option Explicit
'   Private Sub Workbook_Open()
'       mod_JPT71_Auto.Ensure_Core_Sheets
'       mod_JPT71_Auto.Schedule_Refresh 1
'   End Sub
'==========================================================

'======================
' Globals (for debounce/log)
'======================
Public gPrevAddr As String
Public gPrevValue As Variant
Public gNextRun As Date
Public gRefreshScheduled As Boolean

'======================
' Sheet constants
'======================
Private Const PLAN_SHEET As String = "Plan"
Private Const LOG_SHEET As String = "LOG"
Private Const LOG_SNAP As String = "LOG_SNAPSHOT"
Private Const CROSS_SHEET As String = "Cross_Gantt"
Private Const SETTINGS_SHEET As String = "Settings"
Private Const CAL_DATA As String = "Calendar_Data"
Private Const CAL_VIEW As String = "Calendar_View"
Private Const GANTT_SHEET As String = "Gantt_Chart"
Private Const GANTT_DATA As String = "Gantt_Data"
Private Const DASH As String = "Dashboard"

'Plan INPUT(Q:R) snapshot range
Private Const INPUT_FIRST_ROW As Long = 2
Private Const INPUT_LAST_ROW As Long = 60
Private Const INPUT_COL_Q As String = "Q"
Private Const INPUT_COL_R As String = "R"

'In-progress trip identifier (optional)
Private Const INPROG_TRIP As String = "Debris-8"

'======================
' Public: Buttons / Entry points
'======================
Public Sub Refresh_All()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Ensure_Core_Sheets

    Build_CrossGantt_FromPlan
    Apply_Borders ThisWorkbook.Worksheets(CROSS_SHEET)

    Build_Or_Refresh_GanttChart

    Build_Calendar_Data_FromCrossGantt
    Render_Calendar_View
    Update_Dashboard

    'Audit snapshot after refresh
    Snapshot_InputBlock "Auto snapshot after Refresh_All"

    Log_Event "SYSTEM", "Refresh_All", "", "", "OK"
    GoTo FIN

EH:
    Log_Event "SYSTEM", "Refresh_All", "", "", "ERR: " & Err.Description
    MsgBox "Refresh_All Error: " & Err.Description, vbExclamation

FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub Export_FINAL_Values()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wbSrc As Workbook: Set wbSrc = ThisWorkbook
    Dim wbNew As Workbook: Set wbNew = Workbooks.Add

    'remove default sheets
    Do While wbNew.Worksheets.Count > 1
        wbNew.Worksheets(1).Delete
    Loop

    CopySheetAsValues wbSrc, "Shipping_List", wbNew, 1
    CopySheetAsValues wbSrc, "Mail_Draft", wbNew, 2
    CopySheetAsValues wbSrc, CROSS_SHEET, wbNew, 3
    If SheetExists(GANTT_SHEET, wbSrc) Then CopySheetAsValues wbSrc, GANTT_SHEET, wbNew, 4
    If SheetExists(CAL_VIEW, wbSrc) Then CopySheetAsValues wbSrc, CAL_VIEW, wbNew, wbNew.Worksheets.Count + 1

    Dim outPath As String
    outPath = Left(wbSrc.FullName, InStrRev(wbSrc.FullName, ".") - 1) & "_FINAL.xlsx"
    wbNew.SaveAs outPath, xlOpenXMLWorkbook
    wbNew.Close False

    Log_Event "SYSTEM", "Export_FINAL_Values", "", "", outPath
    MsgBox "FINAL created:" & vbCrLf & outPath, vbInformation
    GoTo FIN

EH:
    MsgBox "Export_FINAL_Values Error: " & Err.Description, vbExclamation
FIN:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'======================
' Debounce scheduler
'======================
Public Sub Schedule_Refresh(Optional ByVal SecondsDelay As Double = 2)
    On Error Resume Next

    If gRefreshScheduled Then
        Application.OnTime EarliestTime:=gNextRun, Procedure:="mod_JPT71_Auto.Run_Scheduled_Refresh", Schedule:=False
        gRefreshScheduled = False
    End If

    gNextRun = Now + TimeSerial(0, 0, SecondsDelay)
    Application.OnTime EarliestTime:=gNextRun, Procedure:="mod_JPT71_Auto.Run_Scheduled_Refresh", Schedule:=True
    gRefreshScheduled = True
End Sub

Public Sub Run_Scheduled_Refresh()
    gRefreshScheduled = False
    Refresh_All
End Sub

'======================
' Ensure sheets & minimal defaults
'======================
Public Sub Ensure_Core_Sheets()
    Init_LOG
    Init_Snapshot

    If Not SheetExists(SETTINGS_SHEET, ThisWorkbook) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = SETTINGS_SHEET
    End If
    If Not SheetExists(CAL_DATA, ThisWorkbook) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = CAL_DATA
        ThisWorkbook.Worksheets(CAL_DATA).Visible = xlSheetVeryHidden
    End If
    If Not SheetExists(CAL_VIEW, ThisWorkbook) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = CAL_VIEW
    End If
    If Not SheetExists(DASH, ThisWorkbook) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = DASH
    End If

    Init_SettingsDefaults
End Sub

Private Sub Init_SettingsDefaults()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    'Default option cells (H4/H5/H6)
    If ws.Range("H4").Value = "" Then ws.Range("H4").Value = 6 'max events/day
    If ws.Range("H5").Value = "" Then ws.Range("H5").Value = Month(Date) 'month
    If ws.Range("H6").Value = "" Then ws.Range("H6").Value = Year(Date)  'year
End Sub

'======================
' LOG
'======================
Public Sub Init_LOG()
    If Not SheetExists(LOG_SHEET, ThisWorkbook) Then
        With ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            .Name = LOG_SHEET
            .Range("A1:F1").Value = Array("Timestamp", "User", "Sheet", "Address", "OldValue", "NewValue/Msg")
            .Rows(1).Font.Bold = True
            .Columns("A:F").ColumnWidth = 22
        End With
    End If
End Sub

Public Sub Log_Change(ByVal wsName As String, ByVal addr As String, ByVal oldV As Variant, ByVal newV As Variant)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim n As Long: n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(n, 1).Value = Now
    ws.Cells(n, 2).Value = Environ$("USERNAME")
    ws.Cells(n, 3).Value = wsName
    ws.Cells(n, 4).Value = addr
    ws.Cells(n, 5).Value = oldV
    ws.Cells(n, 6).Value = newV
End Sub

Public Sub Log_Event(ByVal wsName As String, ByVal addr As String, ByVal oldV As Variant, ByVal newV As Variant, ByVal msg As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim n As Long: n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(n, 1).Value = Now
    ws.Cells(n, 2).Value = Environ$("USERNAME")
    ws.Cells(n, 3).Value = wsName
    ws.Cells(n, 4).Value = addr
    ws.Cells(n, 5).Value = oldV
    ws.Cells(n, 6).Value = msg
End Sub

'======================
' LOG_SNAPSHOT (INPUT block versioning)
'======================
Public Sub Init_Snapshot()
    If Not SheetExists(LOG_SNAP, ThisWorkbook) Then
        With ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            .Name = LOG_SNAP
            .Range("A1:F1").Value = Array("Version", "Timestamp", "User", "Range", "Hash", "Note")
            .Rows(1).Font.Bold = True
            .Columns("A:F").ColumnWidth = 22
            .Columns("F").ColumnWidth = 50
        End With
    End If
End Sub

Public Sub Snapshot_InputBlock(Optional ByVal note As String = "")
    On Error GoTo EH

    Dim wsPlan As Worksheet: Set wsPlan = ThisWorkbook.Worksheets(PLAN_SHEET)
    Dim wsSnap As Worksheet: Set wsSnap = ThisWorkbook.Worksheets(LOG_SNAP)

    Dim rng As Range
    Set rng = wsPlan.Range(INPUT_COL_Q & INPUT_FIRST_ROW & ":" & INPUT_COL_R & INPUT_LAST_ROW)

    Dim ver As Long: ver = NextSnapshotVersion(wsSnap)
    Dim hashV As String: hashV = SimpleHashRange(rng)

    Dim startRow As Long: startRow = wsSnap.Cells(wsSnap.Rows.Count, 1).End(xlUp).Row + 1

    wsSnap.Cells(startRow, 1).Value = ver
    wsSnap.Cells(startRow, 2).Value = Now
    wsSnap.Cells(startRow, 3).Value = Environ$("USERNAME")
    wsSnap.Cells(startRow, 4).Value = rng.Address(False, False)
    wsSnap.Cells(startRow, 5).Value = hashV
    wsSnap.Cells(startRow, 6).Value = note

    rng.Copy
    wsSnap.Cells(startRow + 1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    Exit Sub
EH:
    Log_Event LOG_SNAP, "SNAPSHOT", "", "", "ERR: " & Err.Description
End Sub

Private Function NextSnapshotVersion(ws As Worksheet) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then NextSnapshotVersion = 1 Else NextSnapshotVersion = CLng(ws.Cells(lastRow, 1).Value) + 1
End Function

Private Function SimpleHashRange(rng As Range) As String
    Dim c As Range, s As Long: s = 0
    For Each c In rng.Cells
        s = (s + Len(CStr(c.Value)) * 7 + AscW(Left(CStr(c.Value) & " ", 1))) Mod 1000003
    Next c
    SimpleHashRange = CStr(s)
End Function

'======================
' Cross_Gantt from Plan tblPlan
'======================
Private Sub Build_CrossGantt_FromPlan()
    Dim wsPlan As Worksheet: Set wsPlan = ThisWorkbook.Worksheets(PLAN_SHEET)
    Dim lo As ListObject: Set lo = wsPlan.ListObjects("tblPlan")

    Dim ws As Worksheet
    If SheetExists(CROSS_SHEET, ThisWorkbook) Then
        Set ws = ThisWorkbook.Worksheets(CROSS_SHEET)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=wsPlan)
        ws.Name = CROSS_SHEET
    End If

    Dim delayRef As String: delayRef = wsPlan.Name & "!$R$4"

    Dim n As Long: n = lo.DataBodyRange.Rows.Count
    If n = 0 Then Exit Sub

    Dim cTrip As Long, cType As Long, cMat As Long
    Dim cE As Long, cF As Long, cG As Long, cH As Long
    cTrip = lo.ListColumns("Trip").Index
    cType = lo.ListColumns("Type").Index
    cMat = lo.ListColumns("Material").Index
    cE = lo.ListColumns("Plan_MW4_Depart_Agg").Index
    cF = lo.ListColumns("Plan_AGI_Offload_Agg").Index
    cG = lo.ListColumns("Plan_AGI_Debris_Load").Index
    cH = lo.ListColumns("Plan_MW4_Debris_Offload").Index

    'min/max dates for timeline
    Dim planStart As Date, planEnd As Date, hasDate As Boolean
    Dim i As Long
    For i = 1 To n
        UpdateMinMaxDate lo.DataBodyRange.Cells(i, cE).Value, hasDate, planStart, planEnd
        UpdateMinMaxDate lo.DataBodyRange.Cells(i, cF).Value, hasDate, planStart, planEnd
        UpdateMinMaxDate lo.DataBodyRange.Cells(i, cG).Value, hasDate, planStart, planEnd
        UpdateMinMaxDate lo.DataBodyRange.Cells(i, cH).Value, hasDate, planStart, planEnd
    Next i
    If Not hasDate Then Exit Sub
    Dim viewEnd As Date: viewEnd = planEnd + 14

    'in-progress seq
    Dim inprogSeq As Long: inprogSeq = 999999
    For i = 1 To n
        If Trim$(CStr(lo.DataBodyRange.Cells(i, cTrip).Value)) = INPROG_TRIP Then inprogSeq = i: Exit For
    Next i

    Dim headers As Variant, hidden As Variant
    headers = Array("Seq","Trip","Type","Material","MW4 Depart (Agg)","AGI Offload (Agg)","AGI Debris Loading (Deb)","MW4 Debris Offloading (Deb)","Status")
    hidden = Array("ShiftFlag","Plan_MW4_Depart","Plan_AGI_Offload","Plan_AGI_Deb_Load","Plan_MW4_Deb_Off")

    Dim col As Long: col = 1
    For i = LBound(headers) To UBound(headers): ws.Cells(1, col).Value = headers(i): col = col + 1: Next i
    For i = LBound(hidden) To UBound(hidden): ws.Cells(1, col).Value = hidden(i): col = col + 1: Next i

    Dim dateStartCol As Long: dateStartCol = col
    Dim d As Date: d = planStart
    Do While d <= viewEnd
        ws.Cells(1, col).Value = d
        ws.Cells(1, col).NumberFormat = "mm-dd"
        ws.Columns(col).ColumnWidth = 5
        col = col + 1
        d = d + 1
    Loop
    Dim lastDateCol As Long: lastDateCol = col - 1

    'columns
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 14

    'hide internal cols J:N
    Dim shiftCol As Long: shiftCol = 10 'J
    ws.Columns(shiftCol).Hidden = True
    ws.Columns(shiftCol + 1).Resize(, 4).Hidden = True

    ws.Range("J2").Select
    ws.Application.ActiveWindow.FreezePanes = True

    Dim r As Long, planStartCol As Long: planStartCol = shiftCol + 1 'K
    For i = 1 To n
        r = 1 + i
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = lo.DataBodyRange.Cells(i, cTrip).Value
        ws.Cells(r, 3).Value = lo.DataBodyRange.Cells(i, cType).Value
        ws.Cells(r, 4).Value = lo.DataBodyRange.Cells(i, cMat).Value
        ws.Cells(r, 9).Value = IIf(Trim$(CStr(ws.Cells(r, 2).Value)) = INPROG_TRIP, "IN PROGRESS", "")

        ws.Cells(r, shiftCol).Value = IIf(i > inprogSeq, 1, 0)

        'hidden plan dates
        ws.Cells(r, planStartCol + 0).Value = lo.DataBodyRange.Cells(i, cE).Value
        ws.Cells(r, planStartCol + 1).Value = lo.DataBodyRange.Cells(i, cF).Value
        ws.Cells(r, planStartCol + 2).Value = lo.DataBodyRange.Cells(i, cG).Value
        ws.Cells(r, planStartCol + 3).Value = lo.DataBodyRange.Cells(i, cH).Value
        ws.Range(ws.Cells(r, planStartCol), ws.Cells(r, planStartCol + 3)).NumberFormat = "yyyy-mm-dd"

        'adjusted visible dates E:H
        ws.Cells(r, 5).Formula = "=IF(" & ws.Cells(r, planStartCol).Address(False, False) & "="""","""",IF(" & ws.Cells(r, shiftCol).Address(False, False) & "=1," & ws.Cells(r, planStartCol).Address(False, False) & "+" & delayRef & "," & ws.Cells(r, planStartCol).Address(False, False) & "))"
        ws.Cells(r, 6).Formula = "=IF(" & ws.Cells(r, planStartCol + 1).Address(False, False) & "="""","""",IF(" & ws.Cells(r, shiftCol).Address(False, False) & "=1," & ws.Cells(r, planStartCol + 1).Address(False, False) & "+" & delayRef & "," & ws.Cells(r, planStartCol + 1).Address(False, False) & "))"
        ws.Cells(r, 7).Formula = "=IF(" & ws.Cells(r, planStartCol + 2).Address(False, False) & "="""","""",IF(" & ws.Cells(r, shiftCol).Address(False, False) & "=1," & ws.Cells(r, planStartCol + 2).Address(False, False) & "+" & delayRef & "," & ws.Cells(r, planStartCol + 2).Address(False, False) & "))"
        ws.Cells(r, 8).Formula = "=IF(" & ws.Cells(r, planStartCol + 3).Address(False, False) & "="""","""",IF(" & ws.Cells(r, shiftCol).Address(False, False) & "=1," & ws.Cells(r, planStartCol + 3).Address(False, False) & "+" & delayRef & "," & ws.Cells(r, planStartCol + 3).Address(False, False) & "))"
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 8)).NumberFormat = "yyyy-mm-dd"
    Next i

    'header style
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastDateCol))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    'timeline conditional formatting
    Apply_Timeline_CF ws, 2, n + 1, dateStartCol, lastDateCol, (1 + inprogSeq)

End Sub

Private Sub UpdateMinMaxDate(ByVal v As Variant, ByRef hasDate As Boolean, ByRef minD As Date, ByRef maxD As Date)
    If IsDate(v) Then
        If Not hasDate Then
            minD = CDate(v): maxD = CDate(v): hasDate = True
        Else
            If CDate(v) < minD Then minD = CDate(v)
            If CDate(v) > maxD Then maxD = CDate(v)
        End If
    End If
End Sub

Private Sub Apply_Timeline_CF(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstDateCol As Long, ByVal lastDateCol As Long, ByVal inprogRow As Long)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0

    Dim rngTop As Range, rngBot As Range
    Dim topLeftCol As String: topLeftCol = Split(ws.Cells(1, firstDateCol).Address(True, False), "$")(1)

    'Colors
    Dim cMW4 As Long: cMW4 = RGB(79, 129, 189)
    Dim cAGIOff As Long: cAGIOff = RGB(198, 239, 206)
    Dim cDebLoad As Long: cDebLoad = RGB(244, 176, 132)
    Dim cDebOff As Long: cDebOff = RGB(248, 203, 173)

    If inprogRow - 1 >= firstRow Then
        Set rngTop = ws.Range(ws.Cells(firstRow, firstDateCol), ws.Cells(inprogRow - 1, lastDateCol))
        AddCF rngTop, "=" & topLeftCol & "$1=$E" & firstRow, cMW4
        AddCF rngTop, "=" & topLeftCol & "$1=$F" & firstRow, cAGIOff
        AddCF rngTop, "=" & topLeftCol & "$1=$G" & firstRow, cDebLoad
        AddCF rngTop, "=" & topLeftCol & "$1=$H" & firstRow, cDebOff
    End If

    If inprogRow + 1 <= lastRow Then
        Set rngBot = ws.Range(ws.Cells(inprogRow + 1, firstDateCol), ws.Cells(lastRow, lastDateCol))
        AddCF rngBot, "=" & topLeftCol & "$1=$E" & (inprogRow + 1), cMW4
        AddCF rngBot, "=" & topLeftCol & "$1=$F" & (inprogRow + 1), cAGIOff
        AddCF rngBot, "=" & topLeftCol & "$1=$G" & (inprogRow + 1), cDebLoad
        AddCF rngBot, "=" & topLeftCol & "$1=$H" & (inprogRow + 1), cDebOff
    End If

    'In-progress row grey
    Dim c As Long
    For c = firstDateCol To lastDateCol
        ws.Cells(inprogRow, c).Interior.Color = RGB(217, 217, 217)
    Next c
End Sub

Private Sub AddCF(ByVal rng As Range, ByVal formula As String, ByVal colorRGB As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    fc.Interior.Color = colorRGB
End Sub

Private Sub Apply_Borders(ws As Worksheet)
    Dim ur As Range: Set ur = ws.UsedRange
    With ur.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
End Sub

'======================
' Gantt Chart (COMPLETE)
'======================
Private Sub Build_Or_Refresh_GanttChart()
    'Build Gantt_Data (Task / Offset / DurationDays) from Cross_Gantt, then build stacked bar chart.
    Dim wsCG As Worksheet: Set wsCG = ThisWorkbook.Worksheets(CROSS_SHEET)

    Dim wsD As Worksheet
    If SheetExists(GANTT_DATA, ThisWorkbook) Then
        Set wsD = ThisWorkbook.Worksheets(GANTT_DATA)
        wsD.Cells.Clear
    Else
        Set wsD = ThisWorkbook.Worksheets.Add(After:=wsCG)
        wsD.Name = GANTT_DATA
    End If
    wsD.Visible = xlSheetVeryHidden

    wsD.Range("A1:C1").Value = Array("Task", "Offset", "DurationDays")
    wsD.Rows(1).Font.Bold = True

    'Base date: first date in header row (timeline start)
    Dim baseCol As Long: baseCol = FindTimelineStartCol(wsCG)
    Dim baseDate As Date
    If IsDate(wsCG.Cells(1, baseCol).Value) Then
        baseDate = CDate(wsCG.Cells(1, baseCol).Value)
    Else
        baseDate = Date
    End If

    Dim lastRow As Long: lastRow = wsCG.Cells(wsCG.Rows.Count, 2).End(xlUp).Row
    Dim outR As Long: outR = 2

    Dim r As Long
    For r = 2 To lastRow
        Dim task As String: task = CStr(wsCG.Cells(r, 2).Value)
        If Len(task) = 0 Then GoTo CONT

        Dim sDate As Variant, eDate As Variant
        sDate = MinDate4(wsCG.Cells(r, 5).Value, wsCG.Cells(r, 6).Value, wsCG.Cells(r, 7).Value, wsCG.Cells(r, 8).Value)
        eDate = MaxDate4(wsCG.Cells(r, 5).Value, wsCG.Cells(r, 6).Value, wsCG.Cells(r, 7).Value, wsCG.Cells(r, 8).Value)

        If IsDate(sDate) And IsDate(eDate) Then
            Dim off As Long: off = CLng(CDate(sDate) - baseDate)
            Dim dur As Long: dur = CLng(CDate(eDate) - CDate(sDate) + 1)
            If dur < 1 Then dur = 1

            wsD.Cells(outR, 1).Value = task
            wsD.Cells(outR, 2).Value = off
            wsD.Cells(outR, 3).Value = dur
            outR = outR + 1
        End If
CONT:
    Next r

    Dim lastDataRow As Long: lastDataRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 3 Then Exit Sub

    'Create/clear Gantt_Chart sheet
    Dim wsC As Worksheet
    If SheetExists(GANTT_SHEET, ThisWorkbook) Then
        Set wsC = ThisWorkbook.Worksheets(GANTT_SHEET)
        wsC.Cells.Clear
        On Error Resume Next
        wsC.ChartObjects.Delete
        On Error GoTo 0
    Else
        Set wsC = ThisWorkbook.Worksheets.Add(After:=wsCG)
        wsC.Name = GANTT_SHEET
    End If

    wsC.Range("A1").Value = "Gantt Chart (Auto)"
    wsC.Range("A1").Font.Bold = True

    'Copy table for reference (optional visible)
    wsC.Range("A3:C" & lastDataRow).Value = wsD.Range("A1:C" & lastDataRow).Value
    wsC.Columns("A:C").AutoFit

    'Build chart
    Dim co As ChartObject
    Set co = wsC.ChartObjects.Add(Left:=20, Top:=80, Width:=980, Height:=420)

    With co.Chart
        .ChartType = xlBarStacked
        .SetSourceData Source:=wsC.Range("A3:C" & lastDataRow)
        .HasTitle = True
        .ChartTitle.Text = "Project Gantt (Base: " & Format$(baseDate, "yyyy-mm-dd") & ")"

        'Series1 Offset = invisible
        .SeriesCollection(1).Format.Fill.Visible = msoFalse
        .SeriesCollection(1).Format.Line.Visible = msoFalse

        'Reverse category order (top = first task)
        .Axes(xlCategory).ReversePlotOrder = True

        'Value axis formatting (days from base)
        With .Axes(xlValue)
            .MinimumScale = 0
            .MajorUnit = 1
            .TickLabels.NumberFormat = "0"
            .HasTitle = True
            .AxisTitle.Text = "Days from base (" & Format$(baseDate, "yyyy-mm-dd") & ")"
        End With

        'Legend off (optional)
        .HasLegend = False
    End With

    'Optional: Color Duration series by Type (Agg/Deb) - left for future enhancement.
End Sub

Private Function FindTimelineStartCol(ws As Worksheet) As Long
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count
        If IsDate(ws.Cells(1, c).Value) Then
            FindTimelineStartCol = c
            Exit Function
        End If
    Next c
    FindTimelineStartCol = 15
End Function

'======================
' Calendar: Data + View
'======================
Private Sub Build_Calendar_Data_FromCrossGantt()
    Dim wsCG As Worksheet: Set wsCG = ThisWorkbook.Worksheets(CROSS_SHEET)

    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(CAL_DATA)
    wsD.Cells.Clear
    wsD.Visible = xlSheetVeryHidden

    wsD.Range("A1:H1").Value = Array("EventDate","Trip","Type","Action","Location","StartDate","EndDate","ConflictFlag")
    wsD.Rows(1).Font.Bold = True

    Dim lastRow As Long: lastRow = wsCG.Cells(wsCG.Rows.Count, 2).End(xlUp).Row
    Dim outR As Long: outR = 2

    Dim r As Long
    For r = 2 To lastRow
        Dim trip As String: trip = CStr(wsCG.Cells(r, 2).Value)
        Dim typ As String: typ = CStr(wsCG.Cells(r, 3).Value)

        AddEventRow wsD, outR, wsCG.Cells(r, 5).Value, trip, typ, "MW4 Depart (Agg)", "MW4"
        AddEventRow wsD, outR, wsCG.Cells(r, 6).Value, trip, typ, "AGI Offload (Agg)", "AGI"
        AddEventRow wsD, outR, wsCG.Cells(r, 7).Value, trip, typ, "AGI Debris Loading", "AGI"
        AddEventRow wsD, outR, wsCG.Cells(r, 8).Value, trip, typ, "MW4 Debris Offloading", "MW4"

        'period bar (start~end)
        Dim sDate As Variant, eDate As Variant
        sDate = MinDate4(wsCG.Cells(r, 5).Value, wsCG.Cells(r, 6).Value, wsCG.Cells(r, 7).Value, wsCG.Cells(r, 8).Value)
        eDate = MaxDate4(wsCG.Cells(r, 5).Value, wsCG.Cells(r, 6).Value, wsCG.Cells(r, 7).Value, wsCG.Cells(r, 8).Value)

        If IsDate(sDate) And IsDate(eDate) Then
            wsD.Cells(outR, 1).Value = CDate(sDate)
            wsD.Cells(outR, 2).Value = trip
            wsD.Cells(outR, 3).Value = typ
            wsD.Cells(outR, 4).Value = "PERIOD"
            wsD.Cells(outR, 5).Value = ""
            wsD.Cells(outR, 6).Value = CDate(sDate)
            wsD.Cells(outR, 7).Value = CDate(eDate)
            outR = outR + 1
        End If
    Next r

    MarkConflicts wsD
End Sub

Private Sub AddEventRow(ws As Worksheet, ByRef outR As Long, ByVal dt As Variant, ByVal trip As String, ByVal typ As String, ByVal action As String, ByVal loc As String)
    If IsDate(dt) Then
        ws.Cells(outR, 1).Value = CDate(dt)
        ws.Cells(outR, 2).Value = trip
        ws.Cells(outR, 3).Value = typ
        ws.Cells(outR, 4).Value = action
        ws.Cells(outR, 5).Value = loc
        ws.Cells(outR, 6).Value = CDate(dt)
        ws.Cells(outR, 7).Value = CDate(dt)
        outR = outR + 1
    End If
End Sub

Private Sub MarkConflicts(ws As Worksheet)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To lastRow
        Dim dt As Variant: dt = ws.Cells(r, 1).Value
        Dim loc As String: loc = CStr(ws.Cells(r, 5).Value)
        Dim act As String: act = CStr(ws.Cells(r, 4).Value)

        If IsDate(dt) And loc <> "" And act <> "PERIOD" Then
            Dim k As String: k = Format$(CDate(dt), "yyyymmdd") & "|" & loc
            dict(k) = dict(k) + 1
        End If
    Next r

    For r = 2 To lastRow
        Dim dt2 As Variant: dt2 = ws.Cells(r, 1).Value
        Dim loc2 As String: loc2 = CStr(ws.Cells(r, 5).Value)
        Dim act2 As String: act2 = CStr(ws.Cells(r, 4).Value)
        If IsDate(dt2) And loc2 <> "" And act2 <> "PERIOD" Then
            Dim k2 As String: k2 = Format$(CDate(dt2), "yyyymmdd") & "|" & loc2
            If dict.Exists(k2) Then ws.Cells(r, 8).Value = IIf(dict(k2) >= 2, 1, 0)
        End If
    Next r
End Sub

Public Sub Render_Calendar_View()
    Dim wsV As Worksheet: Set wsV = ThisWorkbook.Worksheets(CAL_VIEW)
    Dim wsS As Worksheet: Set wsS = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(CAL_DATA)

    Dim yy As Long: yy = CLng(wsS.Range("H6").Value)
    Dim mm As Long: mm = CLng(wsS.Range("H5").Value)
    Dim maxEv As Long: maxEv = CLng(wsS.Range("H4").Value)

    Dim firstDay As Date: firstDay = DateSerial(yy, mm, 1)
    Dim startDate As Date: startDate = firstDay - (Weekday(firstDay, vbMonday) - 1)

    Dim r As Long, c As Long, d As Date
    d = startDate

    For r = 5 To 10
        For c = 1 To 7
            Dim cell As Range: Set cell = wsV.Cells(r, c)
            cell.Value = BuildDayCellText(wsD, d, maxEv)
            cell.WrapText = True
            cell.HorizontalAlignment = xlLeft
            cell.VerticalAlignment = xlTop

            'weekend gray
            If Weekday(d, vbMonday) >= 6 Then
                cell.Interior.Color = RGB(242, 242, 242)
            Else
                cell.Interior.Color = vbWhite
            End If

            'holiday light red
            If IsHoliday(wsS, d) Then cell.Interior.Color = RGB(255, 235, 235)

            'outside month gray text
            If Month(d) <> mm Then cell.Font.Color = RGB(150, 150, 150) Else cell.Font.Color = vbBlack

            d = d + 1
        Next c
    Next r
End Sub

Private Function BuildDayCellText(wsD As Worksheet, ByVal dt As Date, ByVal maxEv As Long) As String
    Dim s As String: s = Format$(dt, "dd") & vbCrLf

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    Dim cnt As Long: cnt = 0
    Dim more As Long: more = 0
    Dim r As Long

    'period line (first match)
    For r = 2 To lastRow
        If CStr(wsD.Cells(r, 4).Value) = "PERIOD" Then
            Dim sD As Variant: sD = wsD.Cells(r, 6).Value
            Dim eD As Variant: eD = wsD.Cells(r, 7).Value
            If IsDate(sD) And IsDate(eD) Then
                If dt >= CDate(sD) And dt <= CDate(eD) Then
                    s = s & "▮ " & wsD.Cells(r, 2).Value & vbCrLf
                    Exit For
                End If
            End If
        End If
    Next r

    'daily events
    For r = 2 To lastRow
        If IsDate(wsD.Cells(r, 1).Value) Then
            If CDate(wsD.Cells(r, 1).Value) = dt Then
                Dim act As String: act = CStr(wsD.Cells(r, 4).Value)
                If act <> "PERIOD" Then
                    If cnt < maxEv Then
                        s = s & "• " & act & " (" & wsD.Cells(r, 5).Value & ")" & vbCrLf
                        cnt = cnt + 1
                    Else
                        more = more + 1
                    End If
                End If
            End If
        End If
    Next r

    If more > 0 Then s = s & "… +" & more & " more" & vbCrLf
    BuildDayCellText = Trim$(s)
End Function

Private Function IsHoliday(wsS As Worksheet, ByVal dt As Date) As Boolean
    Dim lastRow As Long: lastRow = wsS.Cells(wsS.Rows.Count, "E").End(xlUp).Row
    Dim r As Long
    For r = 4 To lastRow
        If IsDate(wsS.Cells(r, "E").Value) Then
            If CDate(wsS.Cells(r, "E").Value) = dt Then
                IsHoliday = True
                Exit Function
            End If
        End If
    Next r
    IsHoliday = False
End Function

Private Sub Update_Dashboard()
    If Not SheetExists(DASH, ThisWorkbook) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DASH)
    If Not SheetExists(CAL_DATA, ThisWorkbook) Then Exit Sub
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(CAL_DATA)

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    Dim nextDt As Variant: nextDt = ""
    Dim conflicts As Long: conflicts = 0

    Dim r As Long
    For r = 2 To lastRow
        If IsDate(wsD.Cells(r, 1).Value) Then
            Dim d As Date: d = CDate(wsD.Cells(r, 1).Value)
            If d >= Date Then
                If nextDt = "" Then nextDt = d Else If d < nextDt Then nextDt = d
            End If
        End If
        If wsD.Cells(r, 8).Value = 1 Then conflicts = conflicts + 1
    Next r

    ws.Range("B4").Value = nextDt
    ws.Range("B5").Value = conflicts
End Sub

'======================
' Utilities
'======================
Private Function SheetExists(ByVal nm As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub CopySheetAsValues(wbSrc As Workbook, ByVal srcName As String, wbTgt As Workbook, ByVal idx As Long)
    Dim wsS As Worksheet: Set wsS = wbSrc.Worksheets(srcName)
    Dim wsT As Worksheet

    If wbTgt.Worksheets.Count < idx Then wbTgt.Worksheets.Add After:=wbTgt.Worksheets(wbTgt.Worksheets.Count)
    Set wsT = wbTgt.Worksheets(idx)
    wsT.Name = srcName

    wsS.UsedRange.Copy
    wsT.Range("A1").PasteSpecial xlPasteColumnWidths
    wsS.UsedRange.Copy
    wsT.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    wsS.UsedRange.Copy
    wsT.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Function MinDate4(a, b, c, d) As Variant
    Dim v As Variant, m As Date, has As Boolean
    has = False
    For Each v In Array(a, b, c, d)
        If IsDate(v) Then
            If Not has Then m = CDate(v): has = True Else If CDate(v) < m Then m = CDate(v)
        End If
    Next v
    If has Then MinDate4 = m Else MinDate4 = ""
End Function

Private Function MaxDate4(a, b, c, d) As Variant
    Dim v As Variant, m As Date, has As Boolean
    has = False
    For Each v In Array(a, b, c, d)
        If IsDate(v) Then
            If Not has Then m = CDate(v): has = True Else If CDate(v) > m Then m = CDate(v)
        End If
    Next v
    If has Then MaxDate4 = m Else MaxDate4 = ""
End Function
