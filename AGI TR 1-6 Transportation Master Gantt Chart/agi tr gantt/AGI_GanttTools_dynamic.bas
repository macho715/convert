Attribute VB_Name = "mod_GanttTools"
Option Explicit

'==========================================================
' AGI TR#1 | Dynamic Gantt Tools (VBA)
' Sheets expected:
'   - Inputs_Assumptions (named ranges: TIDE_THRESHOLD, GANTT_START, GANTT_END, GANTT_SLOT_HOURS)
'   - Schedule_Gantt
'   - FailSafe_Log
'
' How to use:
'   1) Save workbook as .xlsm
'   2) ALT+F11 -> File -> Import File... -> import this .bas module
'   3) (Optional) Add buttons on "Schedule_Gantt" and assign macros
'==========================================================

Private Const SHEET_GANTT As String = "Schedule_Gantt"
Private Const SHEET_LOG As String = "FailSafe_Log"

Private Const FIRST_TASK_ROW As Long = 6
Private Const COL_ID As Long = 1
Private Const COL_START As Long = 6
Private Const COL_END As Long = 7
Private Const COL_TIDE_START As Long = 9
Private Const COL_TIDE_END As Long = 10
Private Const COL_TIDE_CRIT As Long = 11
Private Const COL_TIDE_STATUS As Long = 12
Private Const COL_STATUS As Long = 13
Private Const COL_NOTES As Long = 14


Public Sub ShiftScheduleHours()
    ' Shift Start/End by user-entered hours.
    ' If user selects rows inside the task table, only those rows are shifted; otherwise all tasks are shifted.
    Dim hrsStr As String, hrs As Double
    hrsStr = InputBox("Shift Start/End by hours (can be negative). Example: 2 or -1.5", "Shift Schedule", "0")
    If Len(hrsStr) = 0 Then Exit Sub
    If Not IsNumeric(hrsStr) Then
        MsgBox "Invalid number.", vbExclamation
        Exit Sub
    End If
    hrs = CDbl(hrsStr)
    If hrs = 0 Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_GANTT)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < FIRST_TASK_ROW Then Exit Sub

    Dim rngRows As Range
    If TypeName(Selection) = "Range" Then
        If Not Intersect(Selection, ws.Range(ws.Cells(FIRST_TASK_ROW, COL_ID), ws.Cells(lastRow, COL_NOTES))) Is Nothing Then
            Set rngRows = Intersect(Selection.EntireRow, ws.Range(ws.Cells(FIRST_TASK_ROW, COL_ID), ws.Cells(lastRow, COL_NOTES))).Rows
        End If
    End If
    If rngRows Is Nothing Then
        Set rngRows = ws.Range(ws.Cells(FIRST_TASK_ROW, COL_ID), ws.Cells(lastRow, COL_NOTES)).Rows
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim r As Range
    For Each r In rngRows
        If Len(Trim$(ws.Cells(r.Row, COL_ID).Value)) > 0 Then
            If IsDate(ws.Cells(r.Row, COL_START).Value) Then
                ws.Cells(r.Row, COL_START).Value = ws.Cells(r.Row, COL_START).Value + hrs / 24#
            End If
            If IsDate(ws.Cells(r.Row, COL_END).Value) Then
                ws.Cells(r.Row, COL_END).Value = ws.Cells(r.Row, COL_END).Value + hrs / 24#
            End If
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True
End Sub


Public Sub ValidateTideCritical_Log()
    ' Checks tide-critical tasks (K="Y") and writes a Fail-safe log entry when Tide Status = "LOW".
    ' Does NOT overwrite formulas; it only appends issues to FailSafe_Log.
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_GANTT)
    Dim logWs As Worksheet: Set logWs = ThisWorkbook.Worksheets(SHEET_LOG)

    Dim threshold As Double
    On Error GoTo ThresholdErr
    threshold = CDbl(Range("TIDE_THRESHOLD").Value)
    On Error GoTo 0

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row
    If lastRow < FIRST_TASK_ROW Then Exit Sub

    Dim clearFirst As VbMsgBoxResult
    clearFirst = MsgBox("Clear existing FailSafe_Log entries (rows 4+)?", vbYesNoCancel + vbQuestion, "Fail-safe Log")
    If clearFirst = vbCancel Then Exit Sub
    If clearFirst = vbYes Then
        Call ClearFailSafeLog
    End If

    Application.ScreenUpdating = False
    Application.CalculateFull

    Dim nextLogRow As Long
    nextLogRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If nextLogRow < 4 Then nextLogRow = 3
    nextLogRow = nextLogRow + 1

    Dim i As Long
    For i = FIRST_TASK_ROW To lastRow
        Dim taskId As String
        taskId = Trim$(CStr(ws.Cells(i, COL_ID).Value))
        If Len(taskId) = 0 Then
            ' stop at first blank ID (optional). Comment out if you want to scan whole template.
            'Exit For
            GoTo NextRow
        End If

        If UCase$(Trim$(CStr(ws.Cells(i, COL_TIDE_CRIT).Value))) = "Y" Then
            Dim tideStatus As String
            tideStatus = UCase$(Trim$(CStr(ws.Cells(i, COL_TIDE_STATUS).Value)))

            If tideStatus = "LOW" Then
                Dim tStart As Variant, tEnd As Variant
                tStart = ws.Cells(i, COL_TIDE_START).Value
                tEnd = ws.Cells(i, COL_TIDE_END).Value

                logWs.Cells(nextLogRow, 1).Value = Now
                logWs.Cells(nextLogRow, 1).NumberFormat = "yyyy-mm-dd hh:mm"
                logWs.Cells(nextLogRow, 2).Value = taskId
                logWs.Cells(nextLogRow, 3).Value = "Tide below threshold for tide-critical task"
                logWs.Cells(nextLogRow, 4).Value = tStart
                logWs.Cells(nextLogRow, 5).Value = tEnd
                logWs.Cells(nextLogRow, 6).Value = threshold
                logWs.Cells(nextLogRow, 7).Value = "중단 / STOP: re-window to higher tide or revise method"
                nextLogRow = nextLogRow + 1
            End If
        End If

NextRow:
    Next i

    Application.ScreenUpdating = True
    MsgBox "Validation complete. Check 'FailSafe_Log'.", vbInformation
    Exit Sub

ThresholdErr:
    MsgBox "Named range 'TIDE_THRESHOLD' not found (or not numeric).", vbExclamation
End Sub


Public Sub ClearFailSafeLog()
    Dim logWs As Worksheet: Set logWs = ThisWorkbook.Worksheets(SHEET_LOG)
    Dim lastRow As Long: lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 4 Then
        logWs.Range("A4:G" & lastRow).ClearContents
    End If
End Sub


Public Sub ExportSchedulePDF()
    ' Exports Schedule_Gantt sheet as PDF (uses current print settings).
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_GANTT)

    Dim basePath As String
    If Len(ThisWorkbook.Path) > 0 Then
        basePath = ThisWorkbook.Path
    Else
        basePath = Environ$("USERPROFILE") & "\Desktop"
    End If

    Dim fileName As String
    fileName = basePath & "\Schedule_Gantt_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fileName, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    MsgBox "PDF exported: " & fileName, vbInformation
End Sub


Public Sub RefreshAll()
    Application.CalculateFull
End Sub
