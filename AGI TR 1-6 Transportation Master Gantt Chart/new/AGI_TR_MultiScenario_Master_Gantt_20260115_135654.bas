Option Explicit

' ============================================
' AGI TR Multi-Scenario Master Gantt - VBA Macros
' ============================================
' 사용법: Alt+F11 → Module 삽입 → 코드 붙여넣기
' ============================================

' === 통합 업데이트 함수 ===
Sub UpdateAllScenarios()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Sheets("Schedule_Data_Mammoet_Orig").Calculate
    Sheets("Gantt_Chart_Mammoet_Orig").Calculate
    Sheets("Schedule_Data_Mammoet_ScnA").Calculate
    Sheets("Gantt_Chart_Mammoet_ScnA").Calculate
    Sheets("Schedule_Data_Mammoet_Alt").Calculate
    Sheets("Gantt_Chart_Mammoet_Alt").Calculate
    Sheets("Tide_Data").Calculate
    On Error GoTo 0
    
    Sheets("Control_Panel").Calculate
    Sheets("Summary").Calculate
    Sheets("Weather_Analysis").Calculate
    
    Call RefreshAllGanttCharts
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "✅ 모든 시나리오 업데이트 완료!", vbInformation, "Update Complete"
End Sub

' === 모든 Gantt 차트 색상 갱신 ===
Sub RefreshAllGanttCharts()
    If Not ValidateControlPanelInputs() Then Exit Sub

    On Error Resume Next
    Call RefreshGanttChart_Mammoet_Original
    Call RefreshGanttChart_Mammoet_ScenarioA
    Call RefreshGanttChart_Mammoet_Alternative
    On Error GoTo 0
End Sub

' === Mammoet_Original Gantt 갱신 ===
Sub RefreshGanttChart_Mammoet_Original()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Dim ws As Worksheet, wsd As Worksheet
    Dim i As Long, j As Long, lastRow As Long, ganttRow As Long
    Dim startD As Date, endD As Date, projStart As Date, cellDate As Date
    Dim phase As String, dc As Long, lastCol As Long, maxJ As Long
    Dim shamalStart As Date, shamalEnd As Date

    Set ws = Sheets("Gantt_Chart_Mammoet_Orig")
    Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
    projStart = Sheets("Control_Panel").Range("B4").Value
    shamalStart = Sheets("Control_Panel").Range("H5").Value
    shamalEnd = Sheets("Control_Panel").Range("H6").Value
    dc = 8

    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    ws.Range(ws.Cells(5, dc), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone

    For j = 0 To maxJ
        ws.Cells(4, dc + j).Interior.Color = RGB(31, 78, 121)
        cellDate = projStart + j
        If cellDate >= shamalStart And cellDate <= shamalEnd Then
            ws.Cells(4, dc + j).Interior.Color = RGB(255, 152, 0)
        End If
    Next j

    For i = 6 To lastRow
        Dim activityId As String
        On Error Resume Next
        activityId = Trim(UCase(CStr(wsd.Cells(i, 1).Value)))
        On Error GoTo 0

        ' A1000 형식만 허용: A로 시작하고 두 번째 문자가 숫자여야 함
        ' (MOBILIZATION, SPMT, MARINE 등 텍스트 헤더는 제외)
        If activityId = "" Or Left(activityId, 1) <> "A" Then
            GoTo NextRow
        End If
        If Len(activityId) < 2 Or Not IsNumeric(Mid(activityId, 2, 1)) Then
            GoTo NextRow
        End If

        If IsDate(wsd.Cells(i, 6).Value) And wsd.Cells(i, 6).Value <> "" Then
            startD = wsd.Cells(i, 6).Value
            If IsDate(wsd.Cells(i, 7).Value) Then
                endD = wsd.Cells(i, 7).Value
            Else
                endD = startD
            End If
            phase = wsd.Cells(i, 4).Value

            ganttRow = i - 1

            For j = 0 To maxJ
                cellDate = projStart + j
                If cellDate >= startD And cellDate < endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                ElseIf cellDate = startD And startD = endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                    ws.Cells(ganttRow, dc + j).Value = Chr(9733)
                    ws.Cells(ganttRow, dc + j).HorizontalAlignment = xlCenter
                    ws.Cells(ganttRow, dc + j).Font.Size = 8
                End If
            Next j
        End If
NextRow:
    Next i

    Application.ScreenUpdating = True
End Sub

' === Mammoet_ScenarioA Gantt 갱신 ===
Sub RefreshGanttChart_Mammoet_ScenarioA()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Dim ws As Worksheet, wsd As Worksheet
    Dim i As Long, j As Long, lastRow As Long, ganttRow As Long
    Dim startD As Date, endD As Date, projStart As Date, cellDate As Date
    Dim phase As String, dc As Long, lastCol As Long, maxJ As Long
    Dim shamalStart As Date, shamalEnd As Date

    Set ws = Sheets("Gantt_Chart_Mammoet_ScnA")
    Set wsd = Sheets("Schedule_Data_Mammoet_ScnA")
    projStart = Sheets("Control_Panel").Range("B4").Value
    shamalStart = Sheets("Control_Panel").Range("H5").Value
    shamalEnd = Sheets("Control_Panel").Range("H6").Value
    dc = 8

    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    ws.Range(ws.Cells(5, dc), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone

    For j = 0 To maxJ
        ws.Cells(4, dc + j).Interior.Color = RGB(31, 78, 121)
        cellDate = projStart + j
        If cellDate >= shamalStart And cellDate <= shamalEnd Then
            ws.Cells(4, dc + j).Interior.Color = RGB(255, 152, 0)
        End If
    Next j

    For i = 6 To lastRow
        Dim activityId As String
        On Error Resume Next
        activityId = Trim(UCase(CStr(wsd.Cells(i, 1).Value)))
        On Error GoTo 0

        ' A1000 형식만 허용: A로 시작하고 두 번째 문자가 숫자여야 함
        ' (MOBILIZATION, SPMT, MARINE 등 텍스트 헤더는 제외)
        If activityId = "" Or Left(activityId, 1) <> "A" Then
            GoTo NextRow
        End If
        If Len(activityId) < 2 Or Not IsNumeric(Mid(activityId, 2, 1)) Then
            GoTo NextRow
        End If

        If IsDate(wsd.Cells(i, 6).Value) And wsd.Cells(i, 6).Value <> "" Then
            startD = wsd.Cells(i, 6).Value
            If IsDate(wsd.Cells(i, 7).Value) Then
                endD = wsd.Cells(i, 7).Value
            Else
                endD = startD
            End If
            phase = wsd.Cells(i, 4).Value

            ganttRow = i - 1

            For j = 0 To maxJ
                cellDate = projStart + j
                If cellDate >= startD And cellDate < endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                ElseIf cellDate = startD And startD = endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                    ws.Cells(ganttRow, dc + j).Value = Chr(9733)
                    ws.Cells(ganttRow, dc + j).HorizontalAlignment = xlCenter
                    ws.Cells(ganttRow, dc + j).Font.Size = 8
                End If
            Next j
        End If
NextRow:
    Next i

    Application.ScreenUpdating = True
End Sub

' === Mammoet_Alternative Gantt 갱신 ===
Sub RefreshGanttChart_Mammoet_Alternative()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Dim ws As Worksheet, wsd As Worksheet
    Dim i As Long, j As Long, lastRow As Long, ganttRow As Long
    Dim startD As Date, endD As Date, projStart As Date, cellDate As Date
    Dim phase As String, dc As Long, lastCol As Long, maxJ As Long
    Dim shamalStart As Date, shamalEnd As Date

    Set ws = Sheets("Gantt_Chart_Mammoet_Alt")
    Set wsd = Sheets("Schedule_Data_Mammoet_Alt")
    projStart = Sheets("Control_Panel").Range("B4").Value
    shamalStart = Sheets("Control_Panel").Range("H5").Value
    shamalEnd = Sheets("Control_Panel").Range("H6").Value
    dc = 8

    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    ws.Range(ws.Cells(5, dc), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone

    For j = 0 To maxJ
        ws.Cells(4, dc + j).Interior.Color = RGB(31, 78, 121)
        cellDate = projStart + j
        If cellDate >= shamalStart And cellDate <= shamalEnd Then
            ws.Cells(4, dc + j).Interior.Color = RGB(255, 152, 0)
        End If
    Next j

    For i = 6 To lastRow
        Dim activityId As String
        On Error Resume Next
        activityId = Trim(UCase(CStr(wsd.Cells(i, 1).Value)))
        On Error GoTo 0

        ' A1000 형식만 허용: A로 시작하고 두 번째 문자가 숫자여야 함
        ' (MOBILIZATION, SPMT, MARINE 등 텍스트 헤더는 제외)
        If activityId = "" Or Left(activityId, 1) <> "A" Then
            GoTo NextRow
        End If
        If Len(activityId) < 2 Or Not IsNumeric(Mid(activityId, 2, 1)) Then
            GoTo NextRow
        End If

        If IsDate(wsd.Cells(i, 6).Value) And wsd.Cells(i, 6).Value <> "" Then
            startD = wsd.Cells(i, 6).Value
            If IsDate(wsd.Cells(i, 7).Value) Then
                endD = wsd.Cells(i, 7).Value
            Else
                endD = startD
            End If
            phase = wsd.Cells(i, 4).Value

            ganttRow = i - 1

            For j = 0 To maxJ
                cellDate = projStart + j
                If cellDate >= startD And cellDate < endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                ElseIf cellDate = startD And startD = endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                    ws.Cells(ganttRow, dc + j).Value = Chr(9733)
                    ws.Cells(ganttRow, dc + j).HorizontalAlignment = xlCenter
                    ws.Cells(ganttRow, dc + j).Font.Size = 8
                End If
            Next j
        End If
NextRow:
    Next i

    Application.ScreenUpdating = True
End Sub

' === 조석 데이터 갱신 ===
Sub RefreshTideData()
    Dim ws As Worksheet
    Dim i As Long
    Dim tideThreshold As Double
    
    Set ws = Sheets("Tide_Data")
    tideThreshold = Sheets("Control_Panel").Range("H7").Value
    If tideThreshold = 0 Then tideThreshold = 1.9
    
    For i = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsNumeric(ws.Cells(i, 3).Value) Then
            If ws.Cells(i, 3).Value >= tideThreshold Then
                ws.Cells(i, 3).Font.Bold = True
                ws.Cells(i, 3).Font.Color = RGB(0, 102, 204)
                ws.Cells(i, 1).Interior.Color = RGB(227, 242, 253)
            End If
        End If
    Next i
    
    MsgBox "✅ 조석 데이터 강조 완료 (Tide ≥" & Format(tideThreshold, "0.00") & "m)", vbInformation
End Sub

' === 1. 전체 일정 업데이트 ===
Sub UpdateAllSchedules()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Sheets("Schedule_Data_Mammoet_Orig").Calculate
    Sheets("Gantt_Chart_Mammoet_Orig").Calculate
    Sheets("Control_Panel").Calculate
    Sheets("Summary").Calculate
    
    Call RefreshGanttChart
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "✅ 일정 업데이트 완료!" & vbCrLf & vbCrLf & _
           "프로젝트 시작: " & Format(Sheets("Control_Panel").Range("B4").Value, "YYYY-MM-DD") & vbCrLf & _
           "예상 완료: " & Format(Sheets("Control_Panel").Range("B9").Value, "YYYY-MM-DD"), _
           vbInformation, "Schedule Updated"
End Sub

' === 2. Gantt Chart 색상 갱신 ===
Sub RefreshGanttChart()
    If Not ValidateControlPanelInputs() Then Exit Sub

    Dim ws As Worksheet, wsd As Worksheet
    Dim i As Long, j As Long, lastRow As Long, ganttRow As Long
    Dim startD As Date, endD As Date, projStart As Date, cellDate As Date
    Dim phase As String, dc As Long, lastCol As Long, maxJ As Long
    Dim shamalStart As Date, shamalEnd As Date
    
    Set ws = Sheets("Gantt_Chart_Mammoet_Orig")
    Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
    projStart = Sheets("Control_Panel").Range("B4").Value
    shamalStart = Sheets("Control_Panel").Range("H5").Value
    shamalEnd = Sheets("Control_Panel").Range("H6").Value
    dc = 8 ' Date columns start at H
    
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' Clear existing colors in date columns
    ws.Range(ws.Cells(5, dc), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone
    
    ' Reset header colors + Shamal highlight
    For j = 0 To maxJ
        ws.Cells(4, dc + j).Interior.Color = RGB(31, 78, 121) ' HEADER color
        cellDate = projStart + j
        If cellDate >= shamalStart And cellDate <= shamalEnd Then
            ws.Cells(4, dc + j).Interior.Color = RGB(255, 152, 0) ' Orange
        End If
    Next j
    
    ' Apply Gantt bars
    For i = 6 To lastRow
        If IsDate(wsd.Cells(i, 6).Value) And wsd.Cells(i, 6).Value <> "" Then
            startD = wsd.Cells(i, 6).Value
            If IsDate(wsd.Cells(i, 7).Value) Then
                endD = wsd.Cells(i, 7).Value
            Else
                endD = startD
            End If
            phase = wsd.Cells(i, 4).Value
            
            ganttRow = i - 1
            
            For j = 0 To maxJ
                cellDate = projStart + j
                If cellDate >= startD And cellDate < endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                ElseIf cellDate = startD And startD = endD Then
                    ws.Cells(ganttRow, dc + j).Interior.Color = GetPhaseColor(phase)
                    ws.Cells(ganttRow, dc + j).Value = Chr(9733) ' Star
                    ws.Cells(ganttRow, dc + j).HorizontalAlignment = xlCenter
                    ws.Cells(ganttRow, dc + j).Font.Size = 8
                End If
            Next j
        End If
    Next i
    
    ' Highlight today
    For j = 0 To maxJ
        cellDate = projStart + j
        If cellDate = Date Then
            ws.Range(ws.Cells(4, dc + j), ws.Cells(lastRow, dc + j)).Borders(xlEdgeLeft).Color = RGB(255, 0, 0)
            ws.Range(ws.Cells(4, dc + j), ws.Cells(lastRow, dc + j)).Borders(xlEdgeLeft).Weight = xlThick
            Exit For
        End If
    Next j
    
    Application.ScreenUpdating = True
End Sub

' === Phase Color Helper ===
Function GetPhaseColor(phase As String) As Long
    Select Case phase
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
        Case Else: GetPhaseColor = RGB(255, 255, 255)
    End Select
End Function

' === 3. 프로젝트 리포트 생성 ===
Sub GenerateReport()
    Dim wsd As Worksheet
    Dim i As Long, total As Long, jdCount As Long, lastRow As Long
    Dim voyages As Long, milestones As Long
    
    Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    
    For i = 6 To lastRow
        If wsd.Cells(i, 1).Value <> "" Then
            total = total + 1
            If wsd.Cells(i, 4).Value = "JACKDOWN" Then jdCount = jdCount + 1
            If wsd.Cells(i, 4).Value = "MILESTONE" Then milestones = milestones + 1
            If Left(wsd.Cells(i, 1).Value, 1) = "V" And Len(wsd.Cells(i, 1).Value) = 2 Then voyages = voyages + 1
        End If
    Next i
    
    Dim rpt As String
    rpt = "╔══════════════════════════════════════╗" & vbCrLf & _
          "║   AGI HVDC TR Transportation Report  ║" & vbCrLf & _
          "╠══════════════════════════════════════╣" & vbCrLf & _
          "║ Report Date: " & Format(Now, "YYYY-MM-DD HH:NN") & "      ║" & vbCrLf & _
          "╠══════════════════════════════════════╣" & vbCrLf & _
          "║ PROJECT STATUS                       ║" & vbCrLf & _
          "║  Total Tasks: " & total & "                      ║" & vbCrLf & _
          "║  Voyages: " & voyages & "                          ║" & vbCrLf & _
          "║  Jack-down Events: " & jdCount & "                 ║" & vbCrLf & _
          "║  Milestones: " & milestones & "                       ║" & vbCrLf & _
          "╠══════════════════════════════════════╣" & vbCrLf & _
          "║ KEY DATES                            ║" & vbCrLf & _
          "║  Start: " & Format(Sheets("Control_Panel").Range("B4").Value, "YYYY-MM-DD") & "              ║" & vbCrLf & _
          "║  Target: " & Format(Sheets("Control_Panel").Range("B5").Value, "YYYY-MM-DD") & "             ║" & vbCrLf & _
          "║  Est.End: " & Format(Sheets("Control_Panel").Range("B9").Value, "YYYY-MM-DD") & "            ║" & vbCrLf & _
          "║  Status: " & Sheets("Control_Panel").Range("B11").Value & "               ║" & vbCrLf & _
          "╠══════════════════════════════════════╣" & vbCrLf & _
          "║ WEATHER RISK                         ║" & vbCrLf & _
          "║  Shamal: " & Format(Sheets("Control_Panel").Range("H5").Value, "MM/DD") & " - " & Format(Sheets("Control_Panel").Range("H6").Value, "MM/DD") & "           ║" & vbCrLf & _
          "╚══════════════════════════════════════╝"
    
    MsgBox rpt, vbInformation, "Project Report"
End Sub

' === 4. PDF 내보내기 ===
Sub ExportToPDF()
    Dim fp As String
    fp = ThisWorkbook.Path & "\AGI_TR_Gantt_" & Format(Date, "YYYYMMDD") & ".pdf"
    
    Sheets(Array("Schedule_Data_Mammoet_Orig", "Gantt_Chart_Mammoet_Orig", "Summary")).Select
    ActiveSheet.ExportAsFixedFormat xlTypePDF, fp, xlQualityStandard, True
    Sheets("Control_Panel").Select
    
    MsgBox "✅ PDF 저장 완료:" & vbCrLf & fp, vbInformation, "Export Complete"
End Sub

' === 5. 지연 시뮬레이션 ===
Sub SimulateDelay()
    Dim delayDays As Integer, origStart As Date
    Dim wsCtrl As Worksheet
    
    Set wsCtrl = Sheets("Control_Panel")
    origStart = wsCtrl.Range("B4").Value
    
    delayDays = InputBox("시뮬레이션할 지연 일수를 입력하세요:" & vbCrLf & _
                         "(현재 시작일: " & Format(origStart, "YYYY-MM-DD") & ")", _
                         "Delay Simulation", "7")
    
    If IsNumeric(delayDays) And delayDays <> 0 Then
        wsCtrl.Range("B4").Value = origStart + delayDays
        Call UpdateAllSchedules
        
        MsgBox "시뮬레이션 결과:" & vbCrLf & _
               "새 시작일: " & Format(wsCtrl.Range("B4").Value, "YYYY-MM-DD") & vbCrLf & _
               "새 완료일: " & Format(wsCtrl.Range("B9").Value, "YYYY-MM-DD") & vbCrLf & _
               "목표 대비: " & wsCtrl.Range("B11").Value, vbInformation, "Simulation Result"
        
        If MsgBox("원래 일정으로 복원하시겠습니까?", vbYesNo + vbQuestion, "Restore?") = vbYes Then
            wsCtrl.Range("B4").Value = origStart
            Call UpdateAllSchedules
        End If
    End If
End Sub

' === 6. Critical Path 강조 ===
Sub HighlightCritical()
    Dim wsd As Worksheet, i As Long, lastRow As Long
    
    Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    
    ' Reset
    wsd.Range(wsd.Cells(6, 1), wsd.Cells(lastRow, 9)).Font.Bold = False
    wsd.Range(wsd.Cells(6, 1), wsd.Cells(lastRow, 9)).Font.Color = RGB(0, 0, 0)
    
    ' Highlight Jack-down and Milestones
    For i = 6 To lastRow
        If wsd.Cells(i, 4).Value = "JACKDOWN" Then
            wsd.Range(wsd.Cells(i, 1), wsd.Cells(i, 9)).Font.Bold = True
            wsd.Range(wsd.Cells(i, 1), wsd.Cells(i, 9)).Font.Color = RGB(183, 28, 28)
        ElseIf wsd.Cells(i, 4).Value = "MILESTONE" Then
            wsd.Range(wsd.Cells(i, 1), wsd.Cells(i, 9)).Font.Bold = True
            wsd.Range(wsd.Cells(i, 1), wsd.Cells(i, 9)).Font.Color = RGB(21, 101, 192)
        End If
    Next i
    
    MsgBox "✅ Critical Path 강조 완료" & vbCrLf & _
           "🔴 빨강 = Jack-down (Critical)" & vbCrLf & _
           "🔵 파랑 = Milestone", vbInformation, "Critical Path"
End Sub

' === 7. 오늘 날짜 하이라이트 ===
Sub HighlightToday()
    Dim ws As Worksheet, j As Long, lastCol As Long, maxJ As Long, lastRow As Long
    Dim projStart As Date, dc As Long
    
    Set ws = Sheets("Gantt_Chart_Mammoet_Orig")
    projStart = Sheets("Control_Panel").Range("B4").Value
    dc = 8
    
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For j = 0 To maxJ
        If projStart + j = Date Then
            ws.Range(ws.Cells(4, dc + j), ws.Cells(lastRow, dc + j)).Interior.Color = RGB(255, 255, 200)
            ws.Cells(3, dc + j).Value = "TODAY"
            ws.Cells(3, dc + j).Font.Bold = True
            ws.Cells(3, dc + j).Font.Color = RGB(255, 0, 0)
            MsgBox "오늘 날짜 (" & Format(Date, "MM/DD") & ") 컬럼이 강조되었습니다.", vbInformation
            Exit For
        End If
    Next j
End Sub

' === 8. 날짜 변경 자동 트리거 (Control_Panel 시트에 추가) ===
' 아래 코드를 Control_Panel 시트의 코드 영역에 붙여넣으세요:
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Target.Address = "$B$4" Then
'         Call UpdateAllSchedules
'     End If
' End Sub

' === 9. 진행률 일괄 업데이트 ===
Sub BulkProgressUpdate()
    Dim wsd As Worksheet, i As Long, lastRow As Long
    Dim pctValue As Double
    
    pctValue = InputBox("일괄 적용할 진행률을 입력하세요 (0-100):", "Bulk Progress", "50")
    
    If IsNumeric(pctValue) Then
        pctValue = pctValue / 100
        Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
        lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
        
        ' Progress 컬럼이 없으면 추가
        If wsd.Cells(5, 10).Value <> "Progress" Then
            wsd.Cells(5, 10).Value = "Progress"
            wsd.Cells(5, 10).Font.Bold = True
            wsd.Cells(5, 10).Font.Color = RGB(255, 255, 255)
            wsd.Cells(5, 10).Interior.Color = RGB(31, 78, 121)
        End If
        
        For i = 6 To lastRow
            If wsd.Cells(i, 1).Value <> "" Then
                wsd.Cells(i, 10).Value = pctValue
                wsd.Cells(i, 10).NumberFormat = "0%"
            End If
        Next i
        
        MsgBox "진행률 " & Format(pctValue, "0%") & " 일괄 적용 완료", vbInformation
    End If
End Sub

' === 10. Shamal 위험 체크 ===
Sub CheckShamalRisk()
    Dim wsd As Worksheet, i As Long, lastRow As Long
    Dim taskDate As Date, shamalStart As Date, shamalEnd As Date
    Dim riskTasks As String, cnt As Long
    
    Set wsd = Sheets("Schedule_Data_Mammoet_Orig")
    shamalStart = Sheets("Control_Panel").Range("H5").Value
    shamalEnd = Sheets("Control_Panel").Range("H6").Value
    lastRow = wsd.Cells(wsd.Rows.Count, 1).End(xlUp).Row
    
    For i = 6 To lastRow
        If IsDate(wsd.Cells(i, 6).Value) Then
            taskDate = wsd.Cells(i, 6).Value
            If taskDate >= shamalStart And taskDate <= shamalEnd Then
                ' SAIL tasks are weather-critical
                If wsd.Cells(i, 4).Value = "SAIL" Or wsd.Cells(i, 4).Value = "LOADOUT" Then
                    cnt = cnt + 1
                    riskTasks = riskTasks & vbCrLf & "  ⚠️ " & wsd.Cells(i, 1).Value & ": " & wsd.Cells(i, 3).Value
                End If
            End If
        End If
    Next i
    
    If cnt > 0 Then
        MsgBox "⚠️ SHAMAL 위험 경고!" & vbCrLf & vbCrLf & _
               "Shamal 기간 (" & Format(shamalStart, "MM/DD") & "-" & Format(shamalEnd, "MM/DD") & ") 중 " & cnt & "개 기상 민감 작업 발견:" & vbCrLf & _
               riskTasks & vbCrLf & vbCrLf & _
               "일정 조정을 권장합니다.", vbExclamation, "Weather Risk Alert"
    Else
        MsgBox "✅ Shamal 기간 중 기상 민감 작업 없음" & vbCrLf & _
               "현재 일정은 안전합니다.", vbInformation, "Weather Check OK"
    End If
End Sub

' ============================================
' NEW: Control Panel Settings Reader Functions
' ============================================

' === Control Panel Input Validation ===
Function ValidateControlPanelInputs() As Boolean
    Dim ws As Worksheet
    Dim projStart As Variant, shStart As Variant, shEnd As Variant
    Dim tide As Variant, maintStart As Variant, maintEnd As Variant
    Dim voyage As String, earlyVal As String
    Dim issues As String

    On Error Resume Next
    Set ws = Sheets("Control_Panel")
    On Error GoTo 0
    If ws Is Nothing Then
        issues = issues & vbCrLf & "- Control_Panel 시트를 찾을 수 없습니다."
    End If

    If Not ws Is Nothing Then
        projStart = ws.Range("B4").Value
        If Not IsDate(projStart) Then
            issues = issues & vbCrLf & "- Control_Panel!B4(Project Start)가 유효한 날짜가 아닙니다."
        End If

        shStart = ws.Range("H5").Value
        shEnd = ws.Range("H6").Value
        If shStart <> "" And Not IsDate(shStart) Then
            issues = issues & vbCrLf & "- Control_Panel!H5(Shamal Start)가 유효한 날짜가 아닙니다."
        End If
        If shEnd <> "" And Not IsDate(shEnd) Then
            issues = issues & vbCrLf & "- Control_Panel!H6(Shamal End)가 유효한 날짜가 아닙니다."
        End If
        If IsDate(shStart) And IsDate(shEnd) Then
            If CDate(shStart) > CDate(shEnd) Then
                issues = issues & vbCrLf & "- Shamal 기간이 잘못되었습니다. 시작일이 종료일보다 클 수 없습니다."
            End If
        End If

        tide = ws.Range("H7").Value
        If tide <> "" And Not IsNumeric(tide) Then
            issues = issues & vbCrLf & "- Control_Panel!H7(Tide Threshold)가 숫자가 아닙니다."
        End If
        If IsNumeric(tide) Then
            If CDbl(tide) < 0 Then
                issues = issues & vbCrLf & "- Control_Panel!H7(Tide Threshold)는 0 이상이어야 합니다."
            End If
        End If

        maintStart = ws.Range("H10").Value
        maintEnd = ws.Range("H11").Value
        If maintStart <> "" And Not IsDate(maintStart) Then
            issues = issues & vbCrLf & "- Control_Panel!H10(LCT Maint Start)가 유효한 날짜가 아닙니다."
        End If
        If maintEnd <> "" And Not IsDate(maintEnd) Then
            issues = issues & vbCrLf & "- Control_Panel!H11(LCT Maint End)가 유효한 날짜가 아닙니다."
        End If
        If IsDate(maintStart) And IsDate(maintEnd) Then
            If CDate(maintStart) > CDate(maintEnd) Then
                issues = issues & vbCrLf & "- LCT Maintenance 기간이 잘못되었습니다. 시작일이 종료일보다 클 수 없습니다."
            End If
        End If

        voyage = Trim$(CStr(ws.Range("B6").Value))
        If voyage <> "" Then
            Select Case UCase$(voyage)
                Case "1-2-2-2", "2-2-2-1", "2-2-2-1_TWO_SPMT", "1X1X1X1X1X1X1"
                    ' ok
                Case Else
                    issues = issues & vbCrLf & "- Control_Panel!B6(Voyage Pattern)이 허용된 값이 아닙니다."
            End Select
        End If

        earlyVal = Trim$(CStr(ws.Range("B7").Value))
        If earlyVal <> "" Then
            Select Case UCase$(earlyVal)
                Case "TRUE", "FALSE", "YES", "NO", "1", "0"
                    ' ok
                Case Else
                    issues = issues & vbCrLf & "- Control_Panel!B7(Early Return)은 TRUE/FALSE, YES/NO, 1/0 중 하나여야 합니다."
            End Select
        End If
    End If

    If issues <> "" Then
        If MsgBox("다음 입력값 문제가 발견되었습니다:" & issues & vbCrLf & vbCrLf & _
                  "경고를 무시하고 계속 진행하시겠습니까?", vbExclamation + vbYesNo, "Validation Warning") = vbYes Then
            ValidateControlPanelInputs = True
        Else
            ValidateControlPanelInputs = False
        End If
        Exit Function
    End If

    ValidateControlPanelInputs = True
End Function

' === Get Voyage Pattern from Control Panel ===
Function GetVoyagePattern() As String
    ' Returns: "1-2-2-2", "2-2-2-1", "2-2-2-1_TWO_SPMT", or "1x1x1x1x1x1x1"
    GetVoyagePattern = Sheets("Control_Panel").Range("B6").Value
    If GetVoyagePattern = "" Then GetVoyagePattern = "1-2-2-2"
End Function

' === Check if Early Return is enabled ===
Function IsEarlyReturn() As Boolean
    ' TRUE = LCT returns after first JD in a pair
    ' FALSE = LCT returns after batch JD (both TRs)
    Dim val As String
    val = UCase(Trim(Sheets("Control_Panel").Range("B7").Value))
    IsEarlyReturn = (val = "TRUE" Or val = "YES" Or val = "1")
End Function

' === Get LCT Maintenance Start Date ===
Function GetLCTMaintStart() As Date
    On Error Resume Next
    GetLCTMaintStart = Sheets("Control_Panel").Range("H10").Value
    If Err.Number <> 0 Then GetLCTMaintStart = #1/1/2099#
    On Error GoTo 0
End Function

' === Get LCT Maintenance End Date ===
Function GetLCTMaintEnd() As Date
    On Error Resume Next
    GetLCTMaintEnd = Sheets("Control_Panel").Range("H11").Value
    If Err.Number <> 0 Then GetLCTMaintEnd = #1/1/2099#
    On Error GoTo 0
End Function

' === Highlight LCT Maintenance Period in Gantt ===
Sub HighlightLCTMaintenance()
    Dim ws As Worksheet
    Dim j As Long, lastCol As Long, maxJ As Long, lastRow As Long
    Dim projStart As Date, cellDate As Date, dc As Long
    Dim maintStart As Date, maintEnd As Date
    
    Set ws = Sheets("Gantt_Chart_Mammoet_Orig")
    projStart = Sheets("Control_Panel").Range("B4").Value
    maintStart = GetLCTMaintStart()
    maintEnd = GetLCTMaintEnd()
    dc = 8
    
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    maxJ = lastCol - dc
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Highlight maintenance period with gray
    For j = 0 To maxJ
        cellDate = projStart + j
        If cellDate >= maintStart And cellDate <= maintEnd Then
            ws.Range(ws.Cells(4, dc + j), ws.Cells(lastRow, dc + j)).Interior.Color = RGB(200, 200, 200)
            ws.Cells(3, dc + j).Value = "MAINT"
            ws.Cells(3, dc + j).Font.Bold = True
            ws.Cells(3, dc + j).Font.Size = 7
            ws.Cells(3, dc + j).Font.Color = RGB(128, 0, 0)
        End If
    Next j
    
    MsgBox "🔧 LCT Maintenance 기간 강조 완료:" & vbCrLf & _
           Format(maintStart, "YYYY-MM-DD") & " ~ " & Format(maintEnd, "YYYY-MM-DD"), _
           vbInformation, "LCT Maintenance"
End Sub

' === Display Current Control Panel Settings ===
Sub ShowControlPanelSettings()
    Dim msg As String
    
    msg = "📋 현재 Control Panel 설정:" & vbCrLf & vbCrLf & _
          "📅 Project Start: " & Format(Sheets("Control_Panel").Range("B4").Value, "YYYY-MM-DD") & vbCrLf & _
          "🎯 Target End: " & Format(Sheets("Control_Panel").Range("B5").Value, "YYYY-MM-DD") & vbCrLf & _
          "🚢 Voyage Pattern: " & GetVoyagePattern() & vbCrLf & _
          "🔄 Early Return: " & IIf(IsEarlyReturn(), "YES", "NO") & vbCrLf & vbCrLf & _
          "🌊 Shamal Period: " & Format(Sheets("Control_Panel").Range("H5").Value, "MM/DD") & _
          " ~ " & Format(Sheets("Control_Panel").Range("H6").Value, "MM/DD") & vbCrLf & _
          "🌊 Tide Threshold: " & Format(Sheets("Control_Panel").Range("H7").Value, "0.00") & "m" & vbCrLf & vbCrLf & _
          "🔧 LCT Maintenance: " & Format(GetLCTMaintStart(), "MM/DD") & _
          " ~ " & Format(GetLCTMaintEnd(), "MM/DD")
    
    MsgBox msg, vbInformation, "Control Panel Settings"
End Sub
