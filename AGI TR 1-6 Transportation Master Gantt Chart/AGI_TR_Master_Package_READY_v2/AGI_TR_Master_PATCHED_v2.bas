Attribute VB_Name = "AGI_TR_Master"
Option Explicit

'=====================================================================
' AGI HVDC Transformers TR1-TR6 - Master Schedule + Gantt Automation
' Scenario: 1-2-2-1 Voyages with 2 Jack-down Batches (3 units each)
'
' IMPORTANT PATCH (2026-01-07)
'   - Removed large "T = Array( _ ... )" literal which caused:
'       "Too many line continuations" (KOR: "연속된 행이 너무 많습니다")
'   - Tasks are now written by repeated AddTask() calls (no continuation overflow)
'=====================================================================

' === SHEET NAMES ===
Private Const CTRL_SHEET As String = "Control_Panel"
Private Const DATA_SHEET As String = "Schedule_Data"
Private Const GANTT_SHEET As String = "Gantt_Chart"
Private Const TIDE_SHEET As String = "Tide_Data"
Private Const WEATHER_SHEET As String = "Weather_Analysis"

' === CONTROL PANEL CELLS ===
Private Const D0_CELL As String = "C5"              ' D0 (Trip-1 Load-out)
Private Const SHAMAL_START_CELL As String = "C18"   ' Shamal HIGH start date
Private Const SHAMAL_END_CELL As String = "C19"     ' Shamal HIGH end date

' === SCHEDULE_DATA LAYOUT (row 5 headers, row 6+ data) ===
Private Const DATA_HEADER_ROW As Long = 5
Private Const DATA_START_ROW As Long = 6
Private Const COL_ID As Long = 1
Private Const COL_WBS As Long = 2
Private Const COL_TASK As Long = 3
Private Const COL_PHASE As Long = 4
Private Const COL_OWNER As Long = 5
Private Const COL_OFFSET As Long = 6
Private Const COL_START As Long = 7
Private Const COL_END As Long = 8
Private Const COL_DUR As Long = 9
Private Const COL_NOTES As Long = 10
Private Const COL_STATUS As Long = 11

' === GANTT LAYOUT ===
Private Const GANTT_HEADER_ROW As Long = 4
Private Const GANTT_DATE_START_COL As Long = 9        ' Column I
Private Const GANTT_DAYS As Long = 56                 ' D0-5 .. D0+50

'=====================================================================
' PUBLIC - PRIMARY ACTIONS
'=====================================================================

Public Sub RunAll()
    ' One-click: build scenario -> update schedule -> refresh gantt -> conflicts.
    EnsureSheetsAndLayout
    BuildScenario_1_2_2_1_JD3
    UpdateAllDates
End Sub

Public Sub BuildScenario_1_2_2_1_JD3()
    ' Scenario required by user:
    '   Trip 1: 1 unit (TR1)
    '   Trip 2: 2 units (TR2+TR3) -> Jack-down batch #1 (3 units)
    '   Trip 3: 2 units (TR4+TR5)
    '   Trip 4: 1 unit (TR6)     -> Jack-down batch #2 (3 units)

    Dim wsData As Worksheet
    Dim d0Date As Date
    Dim lastRow As Long
    Dim r As Long

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    d0Date = GetD0Date()

    Application.ScreenUpdating = False

    ' Titles (ASCII-only)
    wsData.Range("A1").Value = "AGI HVDC TR1-TR6 Transportation Schedule (Scenario: 1-2-2-1 / JD x2 @ 3 Units)"
    wsData.Range("A2").Value = "Trips: 4 (1-2-2-1) | Jack-down: 2 batches (3 units each) | Target: complete before 2026-03-01"
    wsData.Range("A3").Value = "Planning notes: Shamal window in Control_Panel; Tide: RoRo requires high tide; keep buffers."

    ' Clear old data rows
    lastRow = GetLastDataRow(wsData)
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
    wsData.Range(wsData.Cells(DATA_START_ROW, COL_ID), wsData.Cells(lastRow + 200, COL_STATUS)).ClearContents

    r = DATA_START_ROW

    ' --- One-time mobilization / prep ---
    AddTask wsData, r, d0Date, "MOB-001", "1.0", "MOBILIZATION (SPMT + marine equipment)", "MOBILIZATION", "Mammoet", -3, 1, "SPMT assembly + marine equipment mobilization"
    AddTask wsData, r, d0Date, "PREP-001", "1.1", "Beam replacement + deck prep (D-ring, steel sets, welding)", "DECK_PREP", "Mammoet", -2, 2, "One-time setup for entire campaign"

    ' --- Trip 1 (TR1 / 1 unit) ---
    AddTask wsData, r, d0Date, "V1", "2.0", "TRIP 1: TR1 transport (1 unit)", "MILESTONE", "All", 0, 0, "Start Trip 1"
    AddTask wsData, r, d0Date, "LO-101", "2.1", "TR1 load-out (RoRo) + stool-down", "LOADOUT", "Mammoet", 0, 1, "High tide required"
    AddTask wsData, r, d0Date, "SF-102", "2.2", "TR1 sea fastening + lashing", "SEAFAST", "Mammoet", 1, 1, "Sea fastening / lashing"
    AddTask wsData, r, d0Date, "SAIL-103", "2.3", "Trip 1 sail: MZP -> AGI (loaded)", "SAIL", "LCT Bushra", 2, 1, "Weather dependent"
    AddTask wsData, r, d0Date, "ARR-104", "2.4", "AGI load-in TR1 + store on jetty", "AGI_UNLOAD", "Mammoet", 3, 1, "RoRo load-in"
    AddTask wsData, r, d0Date, "RET-105", "2.5", "Trip 1 return: AGI -> MZP (empty)", "RETURN", "LCT Bushra", 4, 1, "Turnaround for Trip 2"

    ' --- Trip 2 (TR2+TR3 / 2 units) + Jack-down batch #1 (TR1-TR3) ---
    AddTask wsData, r, d0Date, "V2", "3.0", "TRIP 2: TR2+TR3 transport (2 units) + Batch JD #1 (TR1-TR3)", "MILESTONE", "All", 4, 0, "Start Trip 2"
    AddTask wsData, r, d0Date, "LO-201", "3.1", "TR2+TR3 load-out (RoRo) + sea fastening sequence", "LOADOUT", "Mammoet", 4, 3, "Two units sequential"
    AddTask wsData, r, d0Date, "SAIL-202", "3.2", "Trip 2 sail: MZP -> AGI (loaded)", "SAIL", "LCT Bushra", 7, 1, "Check Shamal window"
    AddTask wsData, r, d0Date, "ARR-203", "3.3", "AGI load-in TR3 + store on jetty", "AGI_UNLOAD", "Mammoet", 8, 1, "Unload sequence: TR3 first"
    AddTask wsData, r, d0Date, "ARR-204", "3.4", "AGI load-in TR2 + store on jetty", "AGI_UNLOAD", "Mammoet", 9, 1, "Unload TR2"
    AddTask wsData, r, d0Date, "JD1", "3.5", "BATCH JD #1: install TR1-TR3 (3 units)", "MILESTONE", "All", 10, 0, "Start installation when 3 units available"
    AddTask wsData, r, d0Date, "STG-205", "3.6", "Batch #1 staging + steel bridge setup (Jetty->Bay route)", "TURNING", "Mammoet", 10, 1, "Bridge + route prep"
    AddTask wsData, r, d0Date, "TURN-TR1", "3.7", "TR1 turning at bay", "TURNING", "Mammoet", 11, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR1", "3.8", "TR1 jack-down on temporary support", "JACKDOWN", "Mammoet", 14, 1, "TR1 installed"
    AddTask wsData, r, d0Date, "TURN-TR2", "3.9", "TR2 turning at bay", "TURNING", "Mammoet", 15, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR2", "3.10", "TR2 jack-down on temporary support", "JACKDOWN", "Mammoet", 18, 1, "TR2 installed"
    AddTask wsData, r, d0Date, "TURN-TR3", "3.11", "TR3 turning at bay", "TURNING", "Mammoet", 19, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR3", "3.12", "TR3 jack-down on temporary support", "JACKDOWN", "Mammoet", 22, 1, "TR3 installed"
    AddTask wsData, r, d0Date, "RET-TRIP2", "3.13", "Trip 2 return: AGI -> MZP (after Batch #1)", "RETURN", "LCT Bushra", 23, 1, "Reset for Trip 3"

    ' --- Trip 3 (TR4+TR5 / 2 units) ---
    AddTask wsData, r, d0Date, "V3", "4.0", "TRIP 3: TR4+TR5 transport (2 units)", "MILESTONE", "All", 24, 0, "Start Trip 3"
    AddTask wsData, r, d0Date, "LO-301", "4.1", "TR4+TR5 load-out (RoRo) + sea fastening sequence", "LOADOUT", "Mammoet", 24, 3, "Two units sequential"
    AddTask wsData, r, d0Date, "SAIL-302", "4.2", "Trip 3 sail: MZP -> AGI (loaded)", "SAIL", "LCT Bushra", 27, 1, "Weather dependent"
    AddTask wsData, r, d0Date, "ARR-303", "4.3", "AGI load-in TR5 + store on jetty", "AGI_UNLOAD", "Mammoet", 28, 1, "Unload sequence: TR5 first"
    AddTask wsData, r, d0Date, "ARR-304", "4.4", "AGI load-in TR4 + store on jetty", "AGI_UNLOAD", "Mammoet", 29, 1, "Unload TR4"
    AddTask wsData, r, d0Date, "RET-305", "4.5", "Trip 3 return: AGI -> MZP (empty)", "RETURN", "LCT Bushra", 30, 1, "Turnaround for Trip 4"

    ' --- Trip 4 (TR6 / 1 unit) + Jack-down batch #2 (TR4-TR6) ---
    AddTask wsData, r, d0Date, "V4", "5.0", "TRIP 4: TR6 transport (1 unit) + Batch JD #2 (TR4-TR6)", "MILESTONE", "All", 31, 0, "Start Trip 4"
    AddTask wsData, r, d0Date, "LO-401", "5.1", "TR6 load-out (RoRo) + sea fastening", "LOADOUT", "Mammoet", 31, 2, "One unit; tide window"
    AddTask wsData, r, d0Date, "SAIL-402", "5.2", "Trip 4 sail: MZP -> AGI (loaded)", "SAIL", "LCT Bushra", 33, 1, "Weather dependent"
    AddTask wsData, r, d0Date, "ARR-403", "5.3", "AGI load-in TR6 + store on jetty", "AGI_UNLOAD", "Mammoet", 34, 1, "Unload TR6"
    AddTask wsData, r, d0Date, "JD2", "5.4", "BATCH JD #2: install TR4-TR6 (3 units)", "MILESTONE", "All", 35, 0, "Start installation when 3 units available"
    AddTask wsData, r, d0Date, "STG-404", "5.5", "Batch #2 staging + steel bridge setup (Jetty->Bay route)", "TURNING", "Mammoet", 35, 1, "Bridge + route prep"
    AddTask wsData, r, d0Date, "TURN-TR4", "5.6", "TR4 turning at bay", "TURNING", "Mammoet", 36, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR4", "5.7", "TR4 jack-down on temporary support", "JACKDOWN", "Mammoet", 39, 1, "TR4 installed"
    AddTask wsData, r, d0Date, "TURN-TR5", "5.8", "TR5 turning at bay", "TURNING", "Mammoet", 40, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR5", "5.9", "TR5 jack-down on temporary support", "JACKDOWN", "Mammoet", 43, 1, "TR5 installed"
    AddTask wsData, r, d0Date, "TURN-TR6", "5.10", "TR6 turning at bay", "TURNING", "Mammoet", 44, 3, "Transport + precision turning"
    AddTask wsData, r, d0Date, "JD-TR6", "5.11", "TR6 jack-down on temporary support", "JACKDOWN", "Mammoet", 47, 1, "TR6 installed"
    AddTask wsData, r, d0Date, "RET-TRIP4", "5.12", "Trip 4 return: AGI -> MZP (after Batch #2)", "RETURN", "LCT Bushra", 48, 1, "Final return"
    AddTask wsData, r, d0Date, "BUF-END", "Z.1", "Campaign buffer - weather / tide / port congestion reserve", "BUFFER", "All", 49, 2, "Reserve days to protect Feb completion target"
    AddTask wsData, r, d0Date, "COMP", "99.0", "PROJECT COMPLETE - TR1-TR6 installed", "MILESTONE", "All", 50, 0, "Target: before 2026-03-01"

    Application.ScreenUpdating = True

    MsgBox "Scenario 1-2-2-1 / JD (3 units x2) written to Schedule_Data." & vbCrLf & _
           "Next: run UpdateAllDates (Ctrl+Shift+U) or RunAll (Ctrl+Shift+R).", vbInformation, "Scenario Loaded"
End Sub

Public Sub UpdateAllDates()
    ' Recalculate all task Start/End dates based on D0 and Offset/Duration.
    Dim wsData As Worksheet
    Dim d0Date As Date
    Dim lastRow As Long
    Dim i As Long
    Dim offset As Double, duration As Double
    Dim startDate As Date, endDate As Date

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    d0Date = GetD0Date()
    lastRow = GetLastDataRow(wsData)

    If lastRow < DATA_START_ROW Then
        MsgBox "No tasks found in Schedule_Data (starting row " & DATA_START_ROW & ").", vbExclamation, "UpdateAllDates"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) <> "" Then
            offset = CDbl(Val(wsData.Cells(i, COL_OFFSET).Value))
            duration = CDbl(Val(wsData.Cells(i, COL_DUR).Value))

            startDate = d0Date + offset
            endDate = CalcEndDate(startDate, duration)

            wsData.Cells(i, COL_START).Value = startDate
            wsData.Cells(i, COL_END).Value = endDate
        End If
    Next i

    UpdateControlPanelSummary
    RefreshGanttChart
    CheckScheduleConflicts

    Application.ScreenUpdating = True

    MsgBox "Schedule updated from D0: " & Format(d0Date, "yyyy-mm-dd"), vbInformation, "Update Complete"
End Sub

Public Sub RefreshGanttChart()
    ' Draws Gantt bars on Gantt_Chart based on Start/End (Schedule_Data).
    Dim wsGantt As Worksheet
    Dim wsData As Worksheet
    Dim wsCtrl As Worksheet

    Dim d0Date As Date
    Dim ganttStartDate As Date
    Dim chartDate As Date

    Dim i As Long, j As Long, lastRow As Long
    Dim startDate As Date, endDate As Date
    Dim phase As String
    Dim startCol As Long, endCol As Long
    Dim phaseColor As Long

    Dim shamalStart As Date, shamalEnd As Date

    EnsureSheetsAndLayout

    Set wsGantt = ThisWorkbook.Sheets(GANTT_SHEET)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)

    d0Date = GetD0Date()
    ganttStartDate = d0Date - 5

    ' Shamal window (editable in Control_Panel)
    shamalStart = GetOptionalDate(wsCtrl.Range(SHAMAL_START_CELL).Value, DateSerial(Year(d0Date), 1, 14))
    shamalEnd = GetOptionalDate(wsCtrl.Range(SHAMAL_END_CELL).Value, DateSerial(Year(d0Date), 1, 18))

    ' Update date header row (day cells)
    For j = 0 To GANTT_DAYS - 1
        chartDate = ganttStartDate + j
        With wsGantt.Cells(GANTT_HEADER_ROW, GANTT_DATE_START_COL + j)
            .Value = chartDate
            .NumberFormat = "d"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter

            ' Weekend shading
            If Weekday(chartDate, vbMonday) > 5 Then
                .Interior.Color = RGB(242, 242, 242)
                .Font.Color = RGB(0, 0, 0)
            Else
                .Interior.Color = RGB(31, 78, 121)
                .Font.Color = RGB(255, 255, 255)
            End If

            ' Shamal shading (HIGH)
            If chartDate >= shamalStart And chartDate <= shamalEnd Then
                .Interior.Color = RGB(255, 153, 0)
                .Font.Color = RGB(0, 0, 0)
            End If
        End With
    Next j

    lastRow = GetLastDataRow(wsData)
    If lastRow < DATA_START_ROW Then Exit Sub

    ' Clear old bars
    wsGantt.Range(wsGantt.Cells(DATA_START_ROW, GANTT_DATE_START_COL), _
                  wsGantt.Cells(lastRow, GANTT_DATE_START_COL + GANTT_DAYS - 1)).Interior.ColorIndex = xlNone

    ' Draw new bars
    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) <> "" Then
            startDate = wsData.Cells(i, COL_START).Value
            endDate = wsData.Cells(i, COL_END).Value
            phase = CStr(wsData.Cells(i, COL_PHASE).Value)

            If IsDate(startDate) And IsDate(endDate) Then
                startCol = GANTT_DATE_START_COL + CLng(DateValue(startDate) - ganttStartDate)
                endCol = GANTT_DATE_START_COL + CLng(DateValue(endDate) - ganttStartDate)

                ' Clip to chart range
                If startCol < GANTT_DATE_START_COL Then startCol = GANTT_DATE_START_COL
                If endCol > GANTT_DATE_START_COL + GANTT_DAYS - 1 Then endCol = GANTT_DATE_START_COL + GANTT_DAYS - 1

                phaseColor = GetPhaseColor(phase)
                wsGantt.Range(wsGantt.Cells(i, startCol), wsGantt.Cells(i, endCol)).Interior.Color = phaseColor

                ' Milestone: emphasize the start cell
                If UCase(Trim(phase)) = "MILESTONE" Then
                    wsGantt.Cells(i, startCol).Interior.Color = RGB(255, 0, 0)
                End If
            End If
        End If
    Next i
End Sub

Public Sub CheckScheduleConflicts()
    ' Flags Tide + Weather risks by coloring Schedule_Data ID cells.
    '   - LOADOUT / AGI_UNLOAD: Tide risk (max risk across Start..End)
    '   - SAIL: Shamal risk (per Start day)

    Dim wsData As Worksheet, wsCtrl As Worksheet
    Dim lastRow As Long, i As Long
    Dim startDate As Date, endDate As Date
    Dim phase As String, taskID As String
    Dim tideRisk As String, weatherRisk As String
    Dim conflictCount As Long
    Dim details As String

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)

    lastRow = GetLastDataRow(wsData)
    conflictCount = 0
    details = ""

    If lastRow < DATA_START_ROW Then Exit Sub

    ' Reset ID cell formatting
    wsData.Range(wsData.Cells(DATA_START_ROW, COL_ID), wsData.Cells(lastRow, COL_ID)).Interior.ColorIndex = xlNone

    For i = DATA_START_ROW To lastRow
        taskID = Trim(CStr(wsData.Cells(i, COL_ID).Value))
        If taskID <> "" Then
            phase = UCase(Trim(CStr(wsData.Cells(i, COL_PHASE).Value)))
            startDate = wsData.Cells(i, COL_START).Value
            endDate = wsData.Cells(i, COL_END).Value

            If phase = "LOADOUT" Or phase = "AGI_UNLOAD" Then
                tideRisk = GetMaxTideRisk(startDate, endDate)
                If tideRisk = "HIGH" Then
                    conflictCount = conflictCount + 1
                    wsData.Cells(i, COL_ID).Interior.Color = RGB(255, 0, 0)
                    details = details & "TIDE HIGH: " & taskID & " (" & Format(startDate, "mm-dd") & ")" & vbCrLf
                ElseIf tideRisk = "MEDIUM" Then
                    wsData.Cells(i, COL_ID).Interior.Color = RGB(255, 192, 0)
                    details = details & "TIDE MED: " & taskID & " (" & Format(startDate, "mm-dd") & ")" & vbCrLf
                End If
            End If

            If phase = "SAIL" Then
                weatherRisk = GetWeatherRisk(DateValue(startDate))
                If weatherRisk = "SHAMAL_HIGH" Then
                    conflictCount = conflictCount + 1
                    wsData.Cells(i, COL_ID).Interior.Color = RGB(255, 0, 0)
                    details = details & "SHAMAL HIGH: " & taskID & " (" & Format(startDate, "mm-dd") & ")" & vbCrLf
                ElseIf weatherRisk = "SHAMAL_MEDIUM" Then
                    wsData.Cells(i, COL_ID).Interior.Color = RGB(255, 192, 0)
                    details = details & "SHAMAL MED: " & taskID & " (" & Format(startDate, "mm-dd") & ")" & vbCrLf
                End If
            End If
        End If
    Next i

    wsCtrl.Range("C16").Value = conflictCount
    If details = "" Then details = "No conflicts detected."
    wsCtrl.Range("C17").Value = details
End Sub

Public Sub FindOptimalD0()
    ' Simple D0 search (Jan 1..Jan 20 of D0 year):
    '   - Maximize buffer to Mar 1
    '   - Penalize sailing during Shamal HIGH/MED
    '   - Penalize LOADOUT with Tide HIGH/MED

    Dim wsData As Worksheet
    Dim baseD0 As Date
    Dim testStart As Date, testEnd As Date, testDate As Date

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    baseD0 = GetD0Date()

    testStart = DateSerial(Year(baseD0), 1, 1)
    testEnd = DateSerial(Year(baseD0), 1, 20)

    Dim bestDate As Date
    Dim bestScore As Double
    bestScore = -1E+99

    Dim score As Double
    Dim projEnd As Date
    Dim bufToDeadline As Long
    Dim sailHigh As Long, sailMed As Long
    Dim tideHigh As Long, tideMed As Long

    For testDate = testStart To testEnd
        projEnd = ComputeProjectEndForD0(wsData, testDate)
        bufToDeadline = DateSerial(Year(testDate), 3, 1) - projEnd

        score = bufToDeadline * 2
        If bufToDeadline < 0 Then score = score - 2000

        sailHigh = CountWeatherConflicts(wsData, testDate, "SHAMAL_HIGH")
        sailMed = CountWeatherConflicts(wsData, testDate, "SHAMAL_MEDIUM")

        tideHigh = CountTideConflicts(wsData, testDate, "HIGH")
        tideMed = CountTideConflicts(wsData, testDate, "MEDIUM")

        score = score - sailHigh * 20 - sailMed * 10
        score = score - tideHigh * 5 - tideMed * 2

        If score > bestScore Then
            bestScore = score
            bestDate = testDate
        End If
    Next testDate

    Dim msg As String
    msg = "Best D0 found: " & Format(bestDate, "yyyy-mm-dd") & vbCrLf & _
          "Projected end: " & Format(ComputeProjectEndForD0(wsData, bestDate), "yyyy-mm-dd") & vbCrLf & _
          "Buffer to 2026-03-01: " & (DateSerial(Year(bestDate), 3, 1) - ComputeProjectEndForD0(wsData, bestDate)) & " days" & vbCrLf & vbCrLf & _
          "Apply this D0 to Control_Panel now?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Optimal D0") = vbYes Then
        ThisWorkbook.Sheets(CTRL_SHEET).Range(D0_CELL).Value = bestDate
        UpdateAllDates
    End If
End Sub

Public Sub GenerateDailyBriefing()
    ' Shows tasks starting today through next 7 days.
    Dim wsData As Worksheet
    Dim todayDate As Date
    Dim weekEnd As Date
    Dim lastRow As Long
    Dim i As Long
    Dim startDate As Date
    Dim taskName As String
    Dim briefing As String

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    todayDate = Date
    weekEnd = todayDate + 7

    briefing = "DAILY BRIEFING - " & Format(todayDate, "yyyy-mm-dd") & vbCrLf & vbCrLf

    lastRow = GetLastDataRow(wsData)
    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) <> "" Then
            startDate = wsData.Cells(i, COL_START).Value
            If IsDate(startDate) Then
                If startDate >= todayDate And startDate <= weekEnd Then
                    taskName = CStr(wsData.Cells(i, COL_TASK).Value)
                    briefing = briefing & Format(startDate, "mm-dd") & " | " & taskName & vbCrLf
                End If
            End If
        End If
    Next i

    MsgBox briefing, vbInformation, "Daily Briefing"
End Sub

Public Sub UpdateProgress()
    ' Update Schedule_Data!Status by Task ID.
    Dim wsData As Worksheet
    Dim taskID As String
    Dim status As String
    Dim rowIdx As Long

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)

    taskID = InputBox("Enter Task ID (e.g., LO-101, JD-TR3):", "Update Progress")
    If taskID = "" Then Exit Sub

    status = InputBox("Enter Status (Not Started / In Progress / Completed / Delayed):", "Update Progress")
    If status = "" Then Exit Sub

    rowIdx = FindTaskRow(wsData, taskID)
    If rowIdx = 0 Then
        MsgBox "Task ID not found: " & taskID, vbExclamation, "Update Progress"
        Exit Sub
    End If

    wsData.Cells(rowIdx, COL_STATUS).Value = status
    MsgBox "Updated " & taskID & " -> " & status, vbInformation, "Progress Updated"
End Sub

Public Sub ExportToCSV()
    ' Export Schedule_Data (A5:Klast) to CSV in the same folder.
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim filePath As String

    EnsureSheetsAndLayout

    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)
    lastRow = GetLastDataRow(wsData)

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook to disk first (so CSV export path is available).", vbExclamation, "ExportToCSV"
        Exit Sub
    End If

    filePath = ThisWorkbook.Path & "\AGI_TR_Schedule_" & Format(Now, "yyyymmdd_hhmm") & ".csv"

    wsData.Range("A" & DATA_HEADER_ROW & ":K" & lastRow).Copy
    Workbooks.Add
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ActiveWorkbook.SaveAs filePath, xlCSV
    ActiveWorkbook.Close False

    MsgBox "Exported to: " & filePath, vbInformation, "Export Complete"
End Sub

Public Sub ShowTideInfo()
    ' Show tide risk for a given date (from Tide_Data).
    Dim inputDate As String
    Dim queryDate As Date
    Dim risk As String

    EnsureSheetsAndLayout

    inputDate = InputBox("Enter date (YYYY-MM-DD):", "Tide Info")
    If inputDate = "" Then Exit Sub

    On Error GoTo ErrHandler
    queryDate = DateValue(inputDate)

    risk = GetTideRisk(queryDate)

    MsgBox "Date: " & Format(queryDate, "yyyy-mm-dd") & vbCrLf & _
           "Tide Risk: " & risk, vbInformation, "Tide Info"
    Exit Sub

ErrHandler:
    MsgBox "Invalid date format. Use YYYY-MM-DD.", vbExclamation, "Tide Info"
End Sub

'=====================================================================
' PUBLIC - KEYBOARD SHORTCUTS
'=====================================================================

Public Sub SetupKeyboardShortcuts()
    Application.OnKey "^+r", "AGI_TR_Master.RunAll"
    Application.OnKey "^+u", "AGI_TR_Master.UpdateAllDates"
    Application.OnKey "^+b", "AGI_TR_Master.GenerateDailyBriefing"
    Application.OnKey "^+o", "AGI_TR_Master.FindOptimalD0"
    Application.OnKey "^+t", "AGI_TR_Master.ShowTideInfo"
    Application.OnKey "^+p", "AGI_TR_Master.UpdateProgress"
    Application.OnKey "^+e", "AGI_TR_Master.ExportToCSV"
End Sub

Public Sub ClearKeyboardShortcuts()
    Application.OnKey "^+r"
    Application.OnKey "^+u"
    Application.OnKey "^+b"
    Application.OnKey "^+o"
    Application.OnKey "^+t"
    Application.OnKey "^+p"
    Application.OnKey "^+e"
End Sub

'=====================================================================
' PRIVATE - SCENARIO WRITER
'=====================================================================

Private Sub AddTask(ByVal wsData As Worksheet, ByRef rowPtr As Long, ByVal d0Date As Date, _
                    ByVal taskID As String, ByVal wbs As String, ByVal taskName As String, ByVal phase As String, _
                    ByVal owner As String, ByVal offset As Double, ByVal duration As Double, _
                    Optional ByVal notes As String = "", Optional ByVal status As String = "Not Started")

    Dim s As Date, e As Date

    wsData.Cells(rowPtr, COL_ID).Value = taskID
    wsData.Cells(rowPtr, COL_WBS).Value = wbs
    wsData.Cells(rowPtr, COL_TASK).Value = taskName
    wsData.Cells(rowPtr, COL_PHASE).Value = phase
    wsData.Cells(rowPtr, COL_OWNER).Value = owner
    wsData.Cells(rowPtr, COL_OFFSET).Value = CDbl(offset)
    wsData.Cells(rowPtr, COL_DUR).Value = CDbl(duration)
    wsData.Cells(rowPtr, COL_NOTES).Value = notes
    wsData.Cells(rowPtr, COL_STATUS).Value = status

    s = d0Date + CDbl(offset)
    e = CalcEndDate(s, CDbl(duration))
    wsData.Cells(rowPtr, COL_START).Value = s
    wsData.Cells(rowPtr, COL_END).Value = e

    rowPtr = rowPtr + 1
End Sub

'=====================================================================
' PRIVATE - CONTROL PANEL SUMMARY
'=====================================================================

Private Sub UpdateControlPanelSummary()
    Dim wsCtrl As Worksheet, wsData As Worksheet
    Dim d0Date As Date
    Dim projEnd As Date

    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET)

    d0Date = GetD0Date()
    projEnd = GetMaxEndDate(wsData)

    ' Key milestones (by ID; ignore if not present)
    wsCtrl.Range("C7").Value = GetTaskStartByID(wsData, "V2")
    wsCtrl.Range("C8").Value = GetTaskStartByID(wsData, "V3")
    wsCtrl.Range("C9").Value = GetTaskStartByID(wsData, "V4")
    wsCtrl.Range("C10").Value = projEnd

    ' Deadline buffer (Mar 1)
    wsCtrl.Range("C13").Value = DateSerial(Year(d0Date), 3, 1) - projEnd
End Sub

'=====================================================================
' PRIVATE - RISK LOGIC (TIDE / WEATHER)
'=====================================================================

Private Function GetTideRisk(ByVal targetDate As Date) As String
    ' Tide_Data assumed layout:
    '   - Col A: Date
    '   - Col D: Risk (LOW/MEDIUM/HIGH)
    Dim wsTide As Worksheet
    Dim lastRow As Long, i As Long

    On Error GoTo FailSafe
    Set wsTide = ThisWorkbook.Sheets(TIDE_SHEET)

    lastRow = wsTide.Cells(wsTide.Rows.Count, 1).End(xlUp).Row

    For i = 5 To lastRow
        If IsDate(wsTide.Cells(i, 1).Value) Then
            If DateValue(wsTide.Cells(i, 1).Value) = DateValue(targetDate) Then
                GetTideRisk = UCase(Trim(CStr(wsTide.Cells(i, 4).Value)))
                Exit Function
            End If
        End If
    Next i

FailSafe:
    If GetTideRisk = "" Then GetTideRisk = "UNKNOWN"
End Function

Private Function GetMaxTideRisk(ByVal startDate As Date, ByVal endDate As Date) As String
    ' Highest tide risk between startDate..endDate (inclusive).
    Dim d As Date
    Dim r As String
    Dim maxRank As Long

    maxRank = 0
    For d = DateValue(startDate) To DateValue(endDate)
        r = GetTideRisk(d)
        If RiskRank(r) > maxRank Then maxRank = RiskRank(r)
    Next d

    GetMaxTideRisk = RiskName(maxRank)
End Function

Private Function GetWeatherRisk(ByVal targetDate As Date) As String
    ' Simple Shamal model:
    '   - HIGH: between Control_Panel!C18..C19
    '   - MEDIUM: 2 days after HIGH end
    Dim wsCtrl As Worksheet
    Dim shStart As Date, shEnd As Date
    Dim medStart As Date, medEnd As Date

    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)

    shStart = GetOptionalDate(wsCtrl.Range(SHAMAL_START_CELL).Value, DateSerial(Year(targetDate), 1, 14))
    shEnd = GetOptionalDate(wsCtrl.Range(SHAMAL_END_CELL).Value, DateSerial(Year(targetDate), 1, 18))

    medStart = shEnd + 1
    medEnd = shEnd + 2

    If targetDate >= shStart And targetDate <= shEnd Then
        GetWeatherRisk = "SHAMAL_HIGH"
    ElseIf targetDate >= medStart And targetDate <= medEnd Then
        GetWeatherRisk = "SHAMAL_MEDIUM"
    Else
        GetWeatherRisk = "OK"
    End If
End Function

'=====================================================================
' PRIVATE - OPTIMIZER HELPERS
'=====================================================================

Private Function ComputeProjectEndForD0(ByVal wsData As Worksheet, ByVal d0 As Date) As Date
    Dim lastRow As Long, i As Long
    Dim offset As Double, dur As Double
    Dim s As Date, e As Date
    Dim maxE As Date

    lastRow = GetLastDataRow(wsData)
    maxE = d0

    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) <> "" Then
            offset = CDbl(Val(wsData.Cells(i, COL_OFFSET).Value))
            dur = CDbl(Val(wsData.Cells(i, COL_DUR).Value))
            s = d0 + offset
            e = CalcEndDate(s, dur)
            If e > maxE Then maxE = e
        End If
    Next i

    ComputeProjectEndForD0 = maxE
End Function

Private Function CountWeatherConflicts(ByVal wsData As Worksheet, ByVal d0 As Date, ByVal riskType As String) As Long
    Dim lastRow As Long, i As Long
    Dim offset As Double
    Dim phase As String
    Dim s As Date

    lastRow = GetLastDataRow(wsData)

    For i = DATA_START_ROW To lastRow
        phase = UCase(Trim(CStr(wsData.Cells(i, COL_PHASE).Value)))
        If phase = "SAIL" Then
            offset = CDbl(Val(wsData.Cells(i, COL_OFFSET).Value))
            s = d0 + offset
            If GetWeatherRisk(DateValue(s)) = riskType Then
                CountWeatherConflicts = CountWeatherConflicts + 1
            End If
        End If
    Next i
End Function

Private Function CountTideConflicts(ByVal wsData As Worksheet, ByVal d0 As Date, ByVal riskLevel As String) As Long
    Dim lastRow As Long, i As Long
    Dim offset As Double, dur As Double
    Dim phase As String
    Dim s As Date, e As Date
    Dim risk As String

    lastRow = GetLastDataRow(wsData)

    For i = DATA_START_ROW To lastRow
        phase = UCase(Trim(CStr(wsData.Cells(i, COL_PHASE).Value)))
        If phase = "LOADOUT" Or phase = "AGI_UNLOAD" Then
            offset = CDbl(Val(wsData.Cells(i, COL_OFFSET).Value))
            dur = CDbl(Val(wsData.Cells(i, COL_DUR).Value))
            s = d0 + offset
            e = CalcEndDate(s, dur)
            risk = GetMaxTideRisk(s, e)
            If risk = UCase(riskLevel) Then
                CountTideConflicts = CountTideConflicts + 1
            End If
        End If
    Next i
End Function

'=====================================================================
' PRIVATE - UTILITIES
'=====================================================================

Private Sub EnsureSheetsAndLayout()
    ' Creates missing sheets and minimal headers (non-destructive if already exists).
    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(CTRL_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = CTRL_SHEET
    End If
    Set ws = Nothing

    Set ws = ThisWorkbook.Sheets(DATA_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = DATA_SHEET
        ws.Range("A" & DATA_HEADER_ROW & ":K" & DATA_HEADER_ROW).Value = Array("ID", "WBS", "Task", "Phase", "Owner", "Offset", "Start", "End", "Duration", "Notes", "Status")
    End If
    Set ws = Nothing

    Set ws = ThisWorkbook.Sheets(GANTT_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = GANTT_SHEET
        ws.Cells(GANTT_HEADER_ROW, 1).Value = "ID"
        ws.Cells(GANTT_HEADER_ROW, 2).Value = "WBS"
        ws.Cells(GANTT_HEADER_ROW, 3).Value = "Task"
        ws.Cells(GANTT_HEADER_ROW, 4).Value = "Phase"
        ws.Cells(GANTT_HEADER_ROW, 5).Value = "Start"
        ws.Cells(GANTT_HEADER_ROW, 6).Value = "End"
        ws.Cells(GANTT_HEADER_ROW, 7).Value = "Dur"
        ws.Cells(GANTT_HEADER_ROW, 8).Value = "Owner"
    End If
    Set ws = Nothing

    Set ws = ThisWorkbook.Sheets(TIDE_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = TIDE_SHEET
        ws.Range("A4:D4").Value = Array("Date", "PeakTime", "Peak_m", "Risk")
    End If
    Set ws = Nothing

    On Error GoTo 0

    ' Ensure Control Panel key cells exist
    On Error Resume Next
    With ThisWorkbook.Sheets(CTRL_SHEET)
        If Not IsDate(.Range(D0_CELL).Value) Then .Range(D0_CELL).Value = Date
        If Trim(CStr(.Range(SHAMAL_START_CELL).Value)) = "" Then .Range(SHAMAL_START_CELL).Value = DateSerial(Year(.Range(D0_CELL).Value), 1, 14)
        If Trim(CStr(.Range(SHAMAL_END_CELL).Value)) = "" Then .Range(SHAMAL_END_CELL).Value = DateSerial(Year(.Range(D0_CELL).Value), 1, 18)
    End With
    On Error GoTo 0
End Sub

Private Function GetD0Date() As Date
    Dim wsCtrl As Worksheet
    Set wsCtrl = ThisWorkbook.Sheets(CTRL_SHEET)

    If IsDate(wsCtrl.Range(D0_CELL).Value) Then
        GetD0Date = DateValue(wsCtrl.Range(D0_CELL).Value)
    Else
        GetD0Date = Date
        wsCtrl.Range(D0_CELL).Value = GetD0Date
    End If
End Function

Private Function GetLastDataRow(ByVal wsData As Worksheet) As Long
    GetLastDataRow = wsData.Cells(wsData.Rows.Count, COL_ID).End(xlUp).Row
End Function

Private Function FindTaskRow(ByVal wsData As Worksheet, ByVal taskID As String) As Long
    Dim lastRow As Long, i As Long
    lastRow = GetLastDataRow(wsData)

    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) = taskID Then
            FindTaskRow = i
            Exit Function
        End If
    Next i

    FindTaskRow = 0
End Function

Private Function GetTaskStartByID(ByVal wsData As Worksheet, ByVal taskID As String) As Variant
    Dim r As Long
    r = FindTaskRow(wsData, taskID)
    If r = 0 Then
        GetTaskStartByID = ""
    Else
        GetTaskStartByID = wsData.Cells(r, COL_START).Value
    End If
End Function

Private Function GetMaxEndDate(ByVal wsData As Worksheet) As Date
    Dim lastRow As Long, i As Long
    Dim maxD As Date

    lastRow = GetLastDataRow(wsData)
    If lastRow < DATA_START_ROW Then
        GetMaxEndDate = Date
        Exit Function
    End If

    maxD = wsData.Cells(DATA_START_ROW, COL_END).Value

    For i = DATA_START_ROW To lastRow
        If Trim(CStr(wsData.Cells(i, COL_ID).Value)) <> "" Then
            If IsDate(wsData.Cells(i, COL_END).Value) Then
                If wsData.Cells(i, COL_END).Value > maxD Then maxD = wsData.Cells(i, COL_END).Value
            End If
        End If
    Next i

    GetMaxEndDate = maxD
End Function

Private Function CalcEndDate(ByVal startDate As Date, ByVal duration As Double) As Date
    ' Inclusive end date for day-based schedule:
    '   - duration <= 0 : same day
    '   - 0 < duration < 1 : same day
    '   - duration >= 1 : start + duration - 1
    If duration <= 0 Then
        CalcEndDate = startDate
    ElseIf duration < 1 Then
        CalcEndDate = startDate
    Else
        CalcEndDate = startDate + duration - 1
    End If
End Function

Private Function GetPhaseColor(ByVal phase As String) As Long
    Select Case UCase(Trim(phase))
        Case "MOBILIZATION": GetPhaseColor = RGB(142, 124, 195)
        Case "DECK_PREP": GetPhaseColor = RGB(111, 168, 220)
        Case "LOADOUT": GetPhaseColor = RGB(147, 196, 125)
        Case "SEAFAST": GetPhaseColor = RGB(118, 165, 175)
        Case "SAIL": GetPhaseColor = RGB(164, 194, 244)
        Case "AGI_UNLOAD": GetPhaseColor = RGB(246, 178, 107)
        Case "TURNING": GetPhaseColor = RGB(255, 217, 102)
        Case "JACKDOWN": GetPhaseColor = RGB(224, 102, 102)
        Case "RETURN": GetPhaseColor = RGB(153, 153, 153)
        Case "BUFFER": GetPhaseColor = RGB(217, 217, 217)
        Case "MILESTONE": GetPhaseColor = RGB(255, 0, 0)
        Case Else: GetPhaseColor = RGB(200, 200, 200)
    End Select
End Function

Private Function GetOptionalDate(ByVal valueInCell As Variant, ByVal fallback As Date) As Date
    If IsDate(valueInCell) Then
        GetOptionalDate = DateValue(valueInCell)
    Else
        GetOptionalDate = fallback
    End If
End Function

Private Function RiskRank(ByVal risk As String) As Long
    Select Case UCase(Trim(risk))
        Case "HIGH": RiskRank = 3
        Case "MEDIUM": RiskRank = 2
        Case "LOW": RiskRank = 1
        Case Else: RiskRank = 0
    End Select
End Function

Private Function RiskName(ByVal rank As Long) As String
    Select Case rank
        Case 3: RiskName = "HIGH"
        Case 2: RiskName = "MEDIUM"
        Case 1: RiskName = "LOW"
        Case Else: RiskName = "UNKNOWN"
    End Select
End Function
