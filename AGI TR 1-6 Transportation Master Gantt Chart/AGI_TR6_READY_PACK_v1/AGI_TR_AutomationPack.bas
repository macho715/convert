
' ═══════════════════════════════════════════════════════════════════════════
' MACHO-GPT AGI TR Schedule VBA Automation Module
' ═══════════════════════════════════════════════════════════════════════════
' Project: Samsung C&T HVDC Logistics - AGI TR Transportation
' Version: 1.0.0
' Created: 2026-01-07
' Features:
'   1. D0-driven auto-update system
'   2. Weather/Tide integration
'   3. Gantt chart auto-refresh
'   4. Daily briefing generator
'   5. Progress tracker
'   6. Export functions
' ═══════════════════════════════════════════════════════════════════════════

Option Explicit

' ===== CONSTANTS =====
Private Const CTRL_SHEET As String = "Control_Panel"
Private Const DATA_SHEET As String = "Schedule_Data"
Private Const GANTT_SHEET As String = "Gantt_Chart"
Private Const TIDE_SHEET As String = "Tide_Data"
Private Const D0_CELL As String = "C5"
Private Const DATA_START_ROW As Long = 6
Private Const GANTT_DATE_START_COL As Long = 9

' ===== COLOR CONSTANTS =====
Private Const CLR_MILESTONE As Long = 12566463   ' C00000
Private Const CLR_LOADOUT As Long = 13281587     ' 4472C4
Private Const CLR_SAIL As Long = 4766583         ' 70AD47
Private Const CLR_AGI As Long = 3171389          ' ED7D31
Private Const CLR_JACKDOWN As Long = 7346080     ' 7030A0
Private Const CLR_BUFFER As Long = 10921638      ' A6A6A6
Private Const CLR_TURNING As Long = 49407        ' FFC000
Private Const CLR_SHAMAL As Long = 13421772      ' FFCCCC

' ═══════════════════════════════════════════════════════════════════════════
' MAIN FUNCTIONS
' ═══════════════════════════════════════════════════════════════════════════

Public Sub UpdateAllDates()
    '-----------------------------------------------------------
    ' Master function: Updates all dates based on D0 input
    ' Keyboard shortcut: Ctrl+Shift+U
    '-----------------------------------------------------------
    Dim d0Date As Date
    Dim wsCtrl As Worksheet, wsData As Worksheet, wsGantt As Worksheet
    Dim lastRow As Long, i As Long
    Dim offset As Double, duration As Double
    Dim startDate As Date, endDate As Date
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsGantt = ThisWorkbook.Sheets(GANTT_SHEET)
    
    ' Get D0 date from Control Panel
    d0Date = wsCtrl.Range(D0_CELL).Value
    
    If d0Date < Date Then
        If MsgBox("D0 date is in the past. Continue?", vbYesNo + vbQuestion) = vbNo Then
            GoTo Cleanup
        End If
    End If
    
    ' Update Schedule_Data dates
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    For i = DATA_START_ROW To lastRow
        If wsData.Cells(i, 1).Value <> "" Then
            offset = Val(wsData.Cells(i, 6).Value)   ' Column F = Offset
            duration = Val(wsData.Cells(i, 9).Value) ' Column I = Duration
            
            startDate = d0Date + offset
            If duration > 0 Then
                endDate = startDate + duration - 1
            Else
                endDate = startDate
            End If
            
            wsData.Cells(i, 7).Value = startDate    ' Column G = Start
            wsData.Cells(i, 8).Value = endDate      ' Column H = End
        End If
    Next i
    
    ' Update Control Panel summary
    Call UpdateControlPanelSummary(wsCtrl, d0Date)
    
    ' Refresh Gantt Chart
    Call RefreshGanttChart(wsGantt, d0Date)
    
    ' Check tide/weather conflicts
    Call CheckScheduleConflicts(wsData, wsCtrl)
    
    MsgBox "Schedule updated successfully!" & vbCrLf & _
           "D0 (V1 Load-out): " & Format(d0Date, "yyyy-mm-dd") & vbCrLf & _
           "Project End: " & Format(d0Date + 46, "yyyy-mm-dd"), vbInformation
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error updating dates: " & Err.Description, vbCritical
    Resume Cleanup
End Sub


Public Sub RefreshGanttChart(wsGantt As Worksheet, d0Date As Date)
    '-----------------------------------------------------------
    ' Refreshes the Gantt chart visual based on current dates
    '-----------------------------------------------------------
    Dim wsData As Worksheet
    Dim lastRow As Long, i As Long, col As Long
    Dim startDate As Date, endDate As Date, chartDate As Date
    Dim ganttStartDate As Date, ganttEndDate As Date
    Dim phase As String
    Dim startCol As Long, endCol As Long
    
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Gantt timeline: D0-5 to D0+50
    ganttStartDate = d0Date - 5
    ganttEndDate = d0Date + 50
    
    ' Clear existing Gantt bars (columns I onwards)
    wsGantt.Range(wsGantt.Cells(5, GANTT_DATE_START_COL), _
                  wsGantt.Cells(lastRow, GANTT_DATE_START_COL + 55)).Interior.ColorIndex = xlNone
    
    ' Update date headers
    For col = GANTT_DATE_START_COL To GANTT_DATE_START_COL + 55
        chartDate = ganttStartDate + (col - GANTT_DATE_START_COL)
        wsGantt.Cells(4, col).Value = chartDate
        wsGantt.Cells(4, col).NumberFormat = "d"
        
        ' Highlight weekends
        If Weekday(chartDate, vbMonday) > 5 Then
            wsGantt.Cells(4, col).Interior.Color = RGB(242, 242, 242)
        Else
            wsGantt.Cells(4, col).Interior.ColorIndex = xlNone
        End If
        
        ' Highlight Shamal risk period (Jan 14-18)
        If chartDate >= DateSerial(2026, 1, 14) And chartDate <= DateSerial(2026, 1, 18) Then
            wsGantt.Cells(4, col).Interior.Color = RGB(255, 204, 204)
        End If
    Next col
    
    ' Draw Gantt bars for each task
    For i = DATA_START_ROW To lastRow
        If wsData.Cells(i, 1).Value <> "" Then
            startDate = wsData.Cells(i, 7).Value
            endDate = wsData.Cells(i, 8).Value
            phase = wsData.Cells(i, 5).Value
            
            ' Calculate column positions
            startCol = GANTT_DATE_START_COL + (startDate - ganttStartDate)
            endCol = GANTT_DATE_START_COL + (endDate - ganttStartDate)
            
            ' Ensure within visible range
            If startCol >= GANTT_DATE_START_COL And startCol <= GANTT_DATE_START_COL + 55 Then
                If endCol > GANTT_DATE_START_COL + 55 Then endCol = GANTT_DATE_START_COL + 55
                If endCol < startCol Then endCol = startCol
                
                ' Apply phase color
                wsGantt.Range(wsGantt.Cells(i, startCol), wsGantt.Cells(i, endCol)).Interior.Color = _
                    GetPhaseColor(phase)
                
                ' Add milestone marker
                If phase = "MILESTONE" Or phase = "JACKDOWN" Then
                    wsGantt.Cells(i, startCol).Value = IIf(phase = "JACKDOWN", "★", "▶")
                    wsGantt.Cells(i, startCol).Font.Bold = True
                End If
            End If
        End If
    Next i
End Sub


Private Function GetPhaseColor(phase As String) As Long
    '-----------------------------------------------------------
    ' Returns color for each phase type
    '-----------------------------------------------------------
    Select Case phase
        Case "MILESTONE": GetPhaseColor = RGB(192, 0, 0)
        Case "LOADOUT": GetPhaseColor = RGB(68, 114, 196)
        Case "SAIL": GetPhaseColor = RGB(112, 173, 71)
        Case "AGI_UNLOAD": GetPhaseColor = RGB(237, 125, 49)
        Case "JACKDOWN": GetPhaseColor = RGB(112, 48, 160)
        Case "TURNING": GetPhaseColor = RGB(255, 192, 0)
        Case "SEAFAST": GetPhaseColor = RGB(0, 176, 240)
        Case "RETURN": GetPhaseColor = RGB(146, 208, 80)
        Case "BUFFER": GetPhaseColor = RGB(166, 166, 166)
        Case "MOBILIZATION": GetPhaseColor = RGB(68, 114, 196)
        Case "DECK_PREP": GetPhaseColor = RGB(0, 176, 240)
        Case Else: GetPhaseColor = RGB(200, 200, 200)
    End Select
End Function


Public Sub CheckScheduleConflicts(wsData As Worksheet, wsCtrl As Worksheet)
    '-----------------------------------------------------------
    ' Checks for tide and weather conflicts
    '-----------------------------------------------------------
    Dim lastRow As Long, i As Long
    Dim taskDate As Date, phase As String, taskID As String
    Dim conflicts As String, tideRisk As String, weatherRisk As String
    Dim conflictCount As Long
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    conflicts = ""
    conflictCount = 0
    
    For i = DATA_START_ROW To lastRow
        taskID = wsData.Cells(i, 1).Value
        phase = wsData.Cells(i, 5).Value
        taskDate = wsData.Cells(i, 7).Value
        
        ' Check LOADOUT tasks for tide
        If phase = "LOADOUT" Or phase = "AGI_UNLOAD" Then
            tideRisk = GetTideRisk(taskDate)
            If tideRisk = "HIGH" Then
                conflicts = conflicts & "⚠️ " & taskID & " (" & Format(taskDate, "mm/dd") & "): Low tide risk" & vbCrLf
                conflictCount = conflictCount + 1
                wsData.Cells(i, 1).Interior.Color = RGB(255, 199, 206)
            ElseIf tideRisk = "MEDIUM" Then
                wsData.Cells(i, 1).Interior.Color = RGB(255, 235, 156)
            Else
                wsData.Cells(i, 1).Interior.ColorIndex = xlNone
            End If
        End If
        
        ' Check SAIL tasks for weather
        If phase = "SAIL" Then
            weatherRisk = GetWeatherRisk(taskDate)
            If weatherRisk = "SHAMAL_HIGH" Then
                conflicts = conflicts & "🌪️ " & taskID & " (" & Format(taskDate, "mm/dd") & "): Shamal HIGH risk" & vbCrLf
                conflictCount = conflictCount + 1
                wsData.Cells(i, 1).Interior.Color = RGB(255, 199, 206)
            ElseIf weatherRisk = "SHAMAL_MEDIUM" Then
                wsData.Cells(i, 1).Interior.Color = RGB(255, 235, 156)
            Else
                wsData.Cells(i, 1).Interior.ColorIndex = xlNone
            End If
        End If
    Next i
    
    ' Update Control Panel conflicts section
    wsCtrl.Range("C15").Value = conflictCount
    If conflictCount > 0 Then
        wsCtrl.Range("C15").Interior.Color = RGB(255, 199, 206)
        wsCtrl.Range("C16").Value = conflicts
    Else
        wsCtrl.Range("C15").Interior.Color = RGB(198, 239, 206)
        wsCtrl.Range("C16").Value = "✅ No conflicts detected"
    End If
End Sub


Private Function GetTideRisk(checkDate As Date) As String
    '-----------------------------------------------------------
    ' Returns tide risk level for a given date
    ' Based on Mina Zayed tide data (Jan-Feb 2026)
    '-----------------------------------------------------------
    Dim wsTide As Worksheet
    Dim rng As Range
    
    On Error Resume Next
    Set wsTide = ThisWorkbook.Sheets(TIDE_SHEET)
    If wsTide Is Nothing Then
        GetTideRisk = "UNKNOWN"
        Exit Function
    End If
    
    Set rng = wsTide.Range("A:A").Find(checkDate, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rng Is Nothing Then
        GetTideRisk = wsTide.Cells(rng.Row, 4).Value ' Column D = Risk
    Else
        GetTideRisk = "UNKNOWN"
    End If
    On Error GoTo 0
End Function


Private Function GetWeatherRisk(checkDate As Date) As String
    '-----------------------------------------------------------
    ' Returns weather risk level for a given date
    ' Shamal period: Jan 14-18 HIGH, Jan 19-20 MEDIUM
    '-----------------------------------------------------------
    If checkDate >= DateSerial(2026, 1, 14) And checkDate <= DateSerial(2026, 1, 18) Then
        GetWeatherRisk = "SHAMAL_HIGH"
    ElseIf checkDate >= DateSerial(2026, 1, 19) And checkDate <= DateSerial(2026, 1, 20) Then
        GetWeatherRisk = "SHAMAL_MEDIUM"
    Else
        GetWeatherRisk = "OK"
    End If
End Function


Private Sub UpdateControlPanelSummary(wsCtrl As Worksheet, d0Date As Date)
    '-----------------------------------------------------------
    ' Updates the Control Panel summary section
    '-----------------------------------------------------------
    Dim endDate As Date
    Dim daysToMarch As Long
    
    endDate = d0Date + 46 ' Project complete offset
    daysToMarch = DateSerial(2026, 3, 1) - endDate
    
    wsCtrl.Range("C7").Value = d0Date + 4   ' V2 Start
    wsCtrl.Range("C8").Value = d0Date + 18  ' V4 Start
    wsCtrl.Range("C9").Value = d0Date + 32  ' V6 Start
    wsCtrl.Range("C10").Value = endDate     ' Project End
    
    wsCtrl.Range("C12").Value = daysToMarch ' Days before March
    If daysToMarch >= 7 Then
        wsCtrl.Range("C12").Interior.Color = RGB(198, 239, 206) ' Green
    ElseIf daysToMarch >= 0 Then
        wsCtrl.Range("C12").Interior.Color = RGB(255, 235, 156) ' Yellow
    Else
        wsCtrl.Range("C12").Interior.Color = RGB(255, 199, 206) ' Red
    End If
End Sub


' ═══════════════════════════════════════════════════════════════════════════
' UTILITY FUNCTIONS
' ═══════════════════════════════════════════════════════════════════════════

Public Sub GenerateDailyBriefing()
    '-----------------------------------------------------------
    ' Generates a daily briefing for today's and upcoming tasks
    ' Keyboard shortcut: Ctrl+Shift+B
    '-----------------------------------------------------------
    Dim wsData As Worksheet
    Dim lastRow As Long, i As Long
    Dim taskDate As Date, taskID As String, taskName As String, phase As String
    Dim todayTasks As String, upcomingTasks As String
    Dim today As Date, tomorrow As Date, nextWeek As Date
    
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    today = Date
    tomorrow = today + 1
    nextWeek = today + 7
    
    todayTasks = "📅 TODAY (" & Format(today, "yyyy-mm-dd ddd") & "):" & vbCrLf & String(40, "-") & vbCrLf
    upcomingTasks = vbCrLf & "📆 NEXT 7 DAYS:" & vbCrLf & String(40, "-") & vbCrLf
    
    For i = DATA_START_ROW To lastRow
        taskID = wsData.Cells(i, 1).Value
        taskName = wsData.Cells(i, 3).Value
        phase = wsData.Cells(i, 5).Value
        taskDate = wsData.Cells(i, 7).Value
        
        If taskDate = today Then
            todayTasks = todayTasks & "• " & taskID & ": " & Left(taskName, 40) & vbCrLf
        ElseIf taskDate > today And taskDate <= nextWeek Then
            upcomingTasks = upcomingTasks & Format(taskDate, "mm/dd") & " - " & taskID & ": " & Left(taskName, 35) & vbCrLf
        End If
    Next i
    
    MsgBox todayTasks & upcomingTasks, vbInformation, "AGI TR Schedule - Daily Briefing"
End Sub


Public Sub ExportToCSV()
    '-----------------------------------------------------------
    ' Exports schedule data to CSV file
    ' Keyboard shortcut: Ctrl+Shift+E
    '-----------------------------------------------------------
    Dim wsData As Worksheet
    Dim filePath As String
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim lineText As String
    Dim fNum As Integer
    
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = 10 ' Columns A-J
    
    filePath = ThisWorkbook.Path & "\AGI_TR_Schedule_Export_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    
    ' Write header
    Print #fNum, "ID,WBS,Task,Phase,Owner,Offset,Start,End,Duration,Notes"
    
    ' Write data
    For i = DATA_START_ROW To lastRow
        lineText = ""
        For j = 1 To lastCol
            If j > 1 Then lineText = lineText & ","
            lineText = lineText & """" & Replace(CStr(wsData.Cells(i, j).Value), """", """""") & """"
        Next j
        Print #fNum, lineText
    Next i
    
    Close #fNum
    
    MsgBox "Schedule exported to:" & vbCrLf & filePath, vbInformation
End Sub


Public Sub FindOptimalD0()
    '-----------------------------------------------------------
    ' Analyzes different D0 dates to find optimal schedule
    ' Avoids Shamal period and meets March deadline
    '-----------------------------------------------------------
    Dim testDate As Date
    Dim startDate As Date, endDate As Date
    Dim bestDate As Date, bestScore As Long
    Dim score As Long
    Dim results As String
    Dim sailDate As Date, sailConflicts As Long
    
    startDate = DateSerial(2026, 1, 5)
    endDate = DateSerial(2026, 1, 25)
    bestScore = -999
    
    results = "D0 Analysis (Sail conflicts / Days to March / Score):" & vbCrLf & String(50, "-") & vbCrLf
    
    For testDate = startDate To endDate
        ' Calculate sail dates and check conflicts
        sailConflicts = 0
        
        ' V1 Sail: D0+1
        If IsShamalDay(testDate + 1) Then sailConflicts = sailConflicts + 1
        ' V2 Sail: D0+5
        If IsShamalDay(testDate + 5) Then sailConflicts = sailConflicts + 1
        ' V3 Sail: D0+14
        If IsShamalDay(testDate + 14) Then sailConflicts = sailConflicts + 1
        ' V4 Sail: D0+19
        If IsShamalDay(testDate + 19) Then sailConflicts = sailConflicts + 1
        ' V5 Sail: D0+28
        If IsShamalDay(testDate + 28) Then sailConflicts = sailConflicts + 1
        ' V6 Sail: D0+33
        If IsShamalDay(testDate + 33) Then sailConflicts = sailConflicts + 1
        
        ' Days before March deadline
        Dim daysToMarch As Long
        daysToMarch = DateSerial(2026, 3, 1) - (testDate + 46)
        
        ' Score: Prefer fewer conflicts and more buffer to March
        score = (6 - sailConflicts) * 10 + daysToMarch
        
        results = results & Format(testDate, "mm/dd") & ": " & sailConflicts & " conflicts, " & _
                  daysToMarch & " days buffer, Score=" & score
        
        If score > bestScore Then
            bestScore = score
            bestDate = testDate
            results = results & " ★ BEST"
        End If
        results = results & vbCrLf
    Next testDate
    
    results = results & vbCrLf & "🎯 RECOMMENDED D0: " & Format(bestDate, "yyyy-mm-dd")
    
    MsgBox results, vbInformation, "D0 Optimization Analysis"
End Sub


Private Function IsShamalDay(checkDate As Date) As Boolean
    IsShamalDay = (checkDate >= DateSerial(2026, 1, 14) And checkDate <= DateSerial(2026, 1, 18))
End Function


Public Sub UpdateProgress()
    '-----------------------------------------------------------
    ' Updates task progress with a simple dialog
    ' Keyboard shortcut: Ctrl+Shift+P
    '-----------------------------------------------------------
    Dim wsData As Worksheet
    Dim taskID As String, newStatus As String
    Dim rng As Range
    
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    
    taskID = InputBox("Enter Task ID to update (e.g., LO-101):", "Update Progress")
    If taskID = "" Then Exit Sub
    
    Set rng = wsData.Range("A:A").Find(taskID, LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then
        MsgBox "Task ID not found: " & taskID, vbExclamation
        Exit Sub
    End If
    
    newStatus = InputBox("Enter new status:" & vbCrLf & _
                         "1 = Not Started" & vbCrLf & _
                         "2 = In Progress" & vbCrLf & _
                         "3 = Complete" & vbCrLf & _
                         "4 = Delayed", "Update Status")
    
    Select Case newStatus
        Case "1": wsData.Cells(rng.Row, 11).Value = "Not Started"
        Case "2": wsData.Cells(rng.Row, 11).Value = "In Progress"
        Case "3": wsData.Cells(rng.Row, 11).Value = "Complete"
        Case "4": wsData.Cells(rng.Row, 11).Value = "Delayed"
        Case Else: Exit Sub
    End Select
    
    MsgBox "Task " & taskID & " updated to: " & wsData.Cells(rng.Row, 11).Value, vbInformation
End Sub


Public Sub ShowTideInfo()
    '-----------------------------------------------------------
    ' Shows tide information for a specific date
    '-----------------------------------------------------------
    Dim checkDate As Date
    Dim dateInput As String
    Dim wsTide As Worksheet
    Dim rng As Range
    Dim info As String
    
    dateInput = InputBox("Enter date (yyyy-mm-dd):", "Tide Information", Format(Date, "yyyy-mm-dd"))
    If dateInput = "" Then Exit Sub
    
    On Error Resume Next
    checkDate = CDate(dateInput)
    If Err.Number <> 0 Then
        MsgBox "Invalid date format", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Set wsTide = ThisWorkbook.Sheets(TIDE_SHEET)
    Set rng = wsTide.Range("A:A").Find(checkDate, LookIn:=xlValues, LookAt:=xlWhole)
    
    If rng Is Nothing Then
        MsgBox "No tide data for " & dateInput, vbExclamation
    Else
        info = "🌊 Tide Info for " & Format(checkDate, "yyyy-mm-dd ddd") & vbCrLf & String(35, "-") & vbCrLf
        info = info & "High Tide Window: " & wsTide.Cells(rng.Row, 2).Value & vbCrLf
        info = info & "Max Height: " & wsTide.Cells(rng.Row, 3).Value & "m" & vbCrLf
        info = info & "Risk Level: " & wsTide.Cells(rng.Row, 4).Value & vbCrLf
        info = info & vbCrLf & "RoRo Requirement: ≥1.8m"
        MsgBox info, vbInformation, "Tide Information"
    End If
End Sub


' ═══════════════════════════════════════════════════════════════════════════
' AUTO-RUN ON WORKBOOK OPEN
' ═══════════════════════════════════════════════════════════════════════════

Public Sub SetupKeyboardShortcuts()
    '-----------------------------------------------------------
    ' Sets up keyboard shortcuts (call from Workbook_Open)
    '-----------------------------------------------------------
    Application.OnKey "^+u", "UpdateAllDates"     ' Ctrl+Shift+U
    Application.OnKey "^+b", "GenerateDailyBriefing" ' Ctrl+Shift+B
    Application.OnKey "^+e", "ExportToCSV"        ' Ctrl+Shift+E
    Application.OnKey "^+p", "UpdateProgress"     ' Ctrl+Shift+P
    Application.OnKey "^+t", "ShowTideInfo"       ' Ctrl+Shift+T
    Application.OnKey "^+o", "FindOptimalD0"      ' Ctrl+Shift+O
End Sub


Public Sub ClearKeyboardShortcuts()
    '-----------------------------------------------------------
    ' Clears keyboard shortcuts (call from Workbook_BeforeClose)
    '-----------------------------------------------------------
    Application.OnKey "^+u"
    Application.OnKey "^+b"
    Application.OnKey "^+e"
    Application.OnKey "^+p"
    Application.OnKey "^+t"
    Application.OnKey "^+o"
End Sub


' ═══════════════════════════════════════════════════════════════════════════
' ADD-ON: Automation Pack Utilities (General Excel/VBA Features)
' - LOG/SETTINGS sheet 기반 설정
' - 백업/내보내기/Pivot refresh/Python Runner 연동/검증 리포트
' ═══════════════════════════════════════════════════════════════════════════

Private Const LOG_SHEET As String = "LOG"
Private Const SETTINGS_SHEET As String = "SETTINGS"
Private Const VALID_SHEET As String = "VALIDATION"

'---------------------------
' Entry Point (Button)
'---------------------------
Public Sub RefreshAll()
    'One-click: Dates -> Gantt -> Conflicts -> Summary -> Pivots
    On Error GoTo EH
    Call SetupAutomationPack
    Call UpdateAllDates
    Call RefreshAllPivotsAndConnections
    MsgBox "RefreshAll completed.", vbInformation
    Exit Sub
EH:
    LogEvent "RefreshAll", "ERROR", Err.Description, CStr(Err.Number)
    MsgBox "RefreshAll failed: " & Err.Description, vbCritical
End Sub

Public Sub SetupAutomationPack()
    'Ensure LOG/SETTINGS/VALIDATION sheets exist and have headers
    On Error GoTo EH
    Dim ws As Worksheet

    Set ws = EnsureSheet(LOG_SHEET)
    If ws.Cells(1, 1).Value = "" Then
        ws.Range("A1:F1").Value = Array("Timestamp", "Level", "Procedure", "Message", "Details", "User")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 22
        ws.Columns("D").ColumnWidth = 60
        ws.Columns("E").ColumnWidth = 60
        ws.Columns("F").ColumnWidth = 18
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True
    End If

    Set ws = EnsureSheet(SETTINGS_SHEET)
    If ws.Cells(1, 1).Value = "" Then
        ws.Range("A1:C1").Value = Array("Key", "Value", "Description")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A").ColumnWidth = 28
        ws.Columns("B").ColumnWidth = 55
        ws.Columns("C").ColumnWidth = 60

        ws.Range("A2:C6").Value = Array( _
            Array("PYTHON_EXE", "py", "Python 실행 커맨드(예: py 또는 C:\Python311\python.exe)"), _
            Array("PY_SCRIPT", "C:\Path\agi_tr_runner.py", "Python Runner 스크립트 경로"), _
            Array("OUT_DIR", "C:\Temp\AGI_TR_Output", "Python 결과 저장 폴더"), _
            Array("BACKUP_DIR", "C:\Temp\AGI_TR_Backup", "백업 폴더(Workbook Copy)"), _
            Array("LOG_FILE", "C:\Temp\AGI_TR\ops.log", "텍스트 로그 파일(JSONL/line)") _
        )
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True
    End If

    Call LogEvent("SetupAutomationPack", "INFO", "OK", vbNullString)
    Exit Sub

EH:
    MsgBox "SetupAutomationPack failed: " & Err.Description, vbCritical
End Sub

'---------------------------
' Python Runner Bridge
'---------------------------
Public Sub RunPythonUpdate()
    RunPythonRunner "update"
End Sub

Public Sub RunPythonValidate()
    RunPythonRunner "validate"
End Sub

Public Sub RunPythonRunner(Optional ByVal mode As String = "update")
    'Shell out to python runner with workbook path
    'Settings from SETTINGS sheet:
    ' - PYTHON_EXE, PY_SCRIPT, OUT_DIR
    On Error GoTo EH

    Dim pyExe As String, pyScript As String, outDir As String
    Dim cmd As String

    pyExe = GetSetting("PYTHON_EXE", "py")
    pyScript = GetSetting("PY_SCRIPT", vbNullString)
    outDir = GetSetting("OUT_DIR", Environ$("TEMP"))

    If pyScript = vbNullString Then
        Err.Raise vbObjectError + 100, "RunPythonRunner", "SETTINGS!PY_SCRIPT is empty."
    End If

    cmd = "cmd /c """ & Quote(pyExe) & " " & Quote(pyScript) & _
          " --in " & Quote(ThisWorkbook.FullName) & _
          " --out " & Quote(outDir) & _
          " --mode " & mode & """"

    LogEvent "RunPythonRunner", "INFO", "EXEC", cmd
    Shell cmd, vbHide

    MsgBox "Python started (" & mode & "). Output: " & outDir, vbInformation
    Exit Sub

EH:
    LogEvent "RunPythonRunner", "ERROR", Err.Description, CStr(Err.Number)
    MsgBox "RunPythonRunner failed: " & Err.Description, vbCritical
End Sub

'---------------------------
' Backup & Export
'---------------------------
Public Sub BackupWorkbook()
    On Error GoTo EH

    Dim backupDir As String, fso As Object, stamp As String, dest As String

    backupDir = GetSetting("BACKUP_DIR", Environ$("TEMP") & "\AGI_TR_Backup")
    stamp = Format(Now, "yyyymmdd_hhnnss")
    dest = backupDir & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & stamp & ".xlsm"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(backupDir) Then fso.CreateFolder backupDir

    ThisWorkbook.SaveCopyAs dest
    LogEvent "BackupWorkbook", "INFO", "SavedCopyAs", dest
    MsgBox "Backup created: " & dest, vbInformation
    Exit Sub

EH:
    LogEvent "BackupWorkbook", "ERROR", Err.Description, CStr(Err.Number)
    MsgBox "BackupWorkbook failed: " & Err.Description, vbCritical
End Sub

Public Sub ExportVisibleSheetsToPDF()
    'Exports each visible sheet to individual PDF in OUT_DIR
    On Error GoTo EH

    Dim outDir As String, fso As Object
    Dim ws As Worksheet, pdfPath As String, base As String

    outDir = GetSetting("OUT_DIR", Environ$("TEMP"))
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

    base = Replace(ThisWorkbook.Name, ".xlsm", "")
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            pdfPath = outDir & "\" & base & "_" & ws.Name & ".pdf"
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
            LogEvent "ExportVisibleSheetsToPDF", "INFO", "Exported", pdfPath
        End If
    Next ws

    MsgBox "PDF export done: " & outDir, vbInformation
    Exit Sub

EH:
    LogEvent "ExportVisibleSheetsToPDF", "ERROR", Err.Description, CStr(Err.Number)
    MsgBox "ExportVisibleSheetsToPDF failed: " & Err.Description, vbCritical
End Sub

'---------------------------
' Data Validation Report (Excel-side)
'---------------------------
Public Sub ValidateScheduleData()
    'Quick validation in Excel (same spirit as python validate)
    On Error GoTo EH

    Dim wsData As Worksheet, wsV As Worksheet
    Dim lastRow As Long, r As Long, outR As Long
    Dim startDate As Variant, endDate As Variant, dur As Double
    Dim id As String

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsV = EnsureSheet(VALID_SHEET)

    wsV.Cells.Clear
    wsV.Range("A1:E1").Value = Array("Row", "ID", "Issue", "Start", "End")
    wsV.Rows(1).Font.Bold = True
    outR = 2

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    For r = DATA_START_ROW To lastRow
        id = CStr(wsData.Cells(r, 1).Value)
        If id <> vbNullString Then
            startDate = wsData.Cells(r, 7).Value
            endDate = wsData.Cells(r, 8).Value
            dur = Val(wsData.Cells(r, 9).Value)

            If IsEmpty(startDate) Or startDate = "" Then
                wsV.Cells(outR, 1).Value = r
                wsV.Cells(outR, 2).Value = id
                wsV.Cells(outR, 3).Value = "Start missing"
                outR = outR + 1
            ElseIf IsEmpty(endDate) Or endDate = "" Then
                wsV.Cells(outR, 1).Value = r
                wsV.Cells(outR, 2).Value = id
                wsV.Cells(outR, 3).Value = "End missing"
                wsV.Cells(outR, 4).Value = startDate
                outR = outR + 1
            ElseIf dur >= 1 And CDate(endDate) < CDate(startDate) Then
                wsV.Cells(outR, 1).Value = r
                wsV.Cells(outR, 2).Value = id
                wsV.Cells(outR, 3).Value = "End < Start (dur>=1)"
                wsV.Cells(outR, 4).Value = startDate
                wsV.Cells(outR, 5).Value = endDate
                outR = outR + 1
            End If
        End If
    Next r

    wsV.Columns("A:E").AutoFit
    LogEvent "ValidateScheduleData", "INFO", "Problems", CStr(outR - 2)
    MsgBox "Validation complete. Problems: " & (outR - 2), vbInformation
    Exit Sub

EH:
    LogEvent "ValidateScheduleData", "ERROR", Err.Description, CStr(Err.Number)
    MsgBox "ValidateScheduleData failed: " & Err.Description, vbCritical
End Sub

'---------------------------
' Pivot/Connection Refresh
'---------------------------
Public Sub RefreshAllPivotsAndConnections()
    On Error GoTo EH
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim pc As PivotCache
    Dim conn As WorkbookConnection

    For Each conn In wb.Connections
        On Error Resume Next
        conn.Refresh
        On Error GoTo EH
    Next conn

    For Each pc In wb.PivotCaches
        On Error Resume Next
        pc.Refresh
        On Error GoTo EH
    Next pc

    LogEvent "RefreshAllPivotsAndConnections", "INFO", "OK", vbNullString
    Exit Sub
EH:
    LogEvent "RefreshAllPivotsAndConnections", "ERROR", Err.Description, CStr(Err.Number)
End Sub

'---------------------------
' Helpers
'---------------------------
Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function GetSetting(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo Safe
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).Value), key, vbTextCompare) = 0 Then
            GetSetting = CStr(ws.Cells(r, 2).Value)
            If GetSetting = vbNullString Then GetSetting = defaultValue
            Exit Function
        End If
    Next r

Safe:
    GetSetting = defaultValue
End Function

Private Function Quote(ByVal s As String) As String
    Quote = """" & s & """"
End Function

Private Sub LogEvent(ByVal proc As String, ByVal level As String, ByVal msg As String, Optional ByVal details As String = "")
    'Writes to LOG sheet + optional text log file
    On Error Resume Next
    Dim ws As Worksheet, nr As Long, logFile As String
    Dim fso As Object, ts As Object

    Set ws = EnsureSheet(LOG_SHEET)
    nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nr, 1).Value = Now
    ws.Cells(nr, 2).Value = level
    ws.Cells(nr, 3).Value = proc
    ws.Cells(nr, 4).Value = msg
    ws.Cells(nr, 5).Value = details
    ws.Cells(nr, 6).Value = Environ$("USERNAME")

    logFile = GetSetting("LOG_FILE", vbNullString)
    If logFile <> vbNullString Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(fso.GetParentFolderName(logFile)) Then
            fso.CreateFolder fso.GetParentFolderName(logFile)
        End If
        Set ts = fso.OpenTextFile(logFile, 8, True, -1) 'ForAppending / Unicode
        ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & level & vbTab & proc & vbTab & msg & vbTab & details
        ts.Close
    End If
End Sub
