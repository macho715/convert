Option Explicit

' ============================================
' AGI TR Multi-Scenario Master Gantt - VBA (Generated)
' Notes:
' - Gantt is rendered by Conditional Formatting (no repaint needed).
' - Macros below are convenience utilities (recalc/export).
' ============================================

Private Function ScenarioPairs() As Variant
    ScenarioPairs = Array( _
        Array("Schedule_Data_Mammoet_Orig", "Gantt_Chart_Mammoet_Orig"), _
        Array("Schedule_Data_Mammoet_ScnA", "Gantt_Chart_Mammoet_ScnA"), _
        Array("Schedule_Data_Mammoet_Alt", "Gantt_Chart_Mammoet_Alt")
    )
End Function

Sub UpdateAll()
    Dim pairs As Variant, i As Long
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Sheets("Control_Panel").Calculate
    Sheets("Summary").Calculate
    Sheets("Scenario_KPIs").Calculate
    Sheets("Weather_Analysis").Calculate
    Sheets("Tide_Data").Calculate
    On Error GoTo 0

    pairs = ScenarioPairs()
    For i = LBound(pairs) To UBound(pairs)
        On Error Resume Next
        Sheets(CStr(pairs(i)(0))).Calculate
        Sheets(CStr(pairs(i)(1))).Calculate
        On Error GoTo 0
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Update complete", vbInformation
End Sub

Sub ExportToPDF()
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & "_export.pdf"
    On Error GoTo ErrHandler
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF exported: " & outPath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub

Sub ShowControlPanelSettings()
    Dim ws As Worksheet
    Set ws = Sheets("Control_Panel")
    MsgBox _
        "PROJECT_START: " & ws.Range("B4").Value & vbCrLf & _
        "TARGET_END: " & ws.Range("B5").Value & vbCrLf & _
        "SHAMAL: " & ws.Range("H5").Value & " ~ " & ws.Range("H6").Value & vbCrLf & _
        "TIDE_THRESHOLD: " & ws.Range("H7").Value & vbCrLf & _
        "CALENDAR_MODE: " & ws.Range("H12").Value & vbCrLf & _
        "WEEKEND_PATTERN: " & ws.Range("H13").Value, _
        vbInformation, "Control Panel Settings"
End Sub