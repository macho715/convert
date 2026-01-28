Attribute VB_Name = "modControlTower"
Option Explicit

'===============================================================================
' CONTROL TOWER - Single Entry Point for All Refresh Operations
' Version: 1.0
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
'===============================================================================
Public Sub RefreshAll_ControlTower()
    On Error GoTo EH

    Dim startTime As Double
    startTime = Timer

    AppStateGuard_Begin

    ' 1) DocGap operations
    Call DG_RefreshAll
    Call DG_StampUpdated

    ' 2) TR Tracker operations
    Call TR_ApplyStatusFormatting
    Call TR_CalculateProgress
    Call TR_CalculatePartyProgress
    Call TR_HighlightOverdue

    ' 3) Dashboard KPI update (D-7/D-3/D-1)
    Call UpdateDashboardKPIs
    Call UpdateDashboardTimestamp

    AppStateGuard_End

    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox "Control Tower Refresh completed" & vbCrLf & _
           "Elapsed time: " & Format(elapsed, "0.00") & " sec", vbInformation

    Exit Sub

EH:
    AppStateGuard_End
    MsgBox "Error occurred: " & Err.Description & vbCrLf & _
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

    Dim amberDays As Long: amberDays = CLng(GetCfgValue("Amber_Threshold_Days", 7))
    Dim redDays As Long: redDays = CLng(GetCfgValue("Red_Threshold_Days", 3))
    Dim criticalDays As Long: criticalDays = CLng(GetCfgValue("Critical_Threshold_Days", 1))

    wsDash.Range("B11").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & amberDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""""))"

    wsDash.Range("B12").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & redDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""""))"

    wsDash.Range("B13").Formula = "=SUMPRODUCT((tblTracker[SUBMISSION DATE]>=TODAY())*" & _
        "(tblTracker[SUBMISSION DATE]<=TODAY()+" & criticalDays & ")*" & _
        "(tblTracker[Status]<>""Approved"")*(tblTracker[Status]<>""Submitted"")*" & _
        "(tblTracker[Status]<>""Not Required"")*(tblTracker[SUBMISSION DATE]<>""""))"
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
' Helper: Get Config Value
'===============================================================================
Private Function GetCfgValue(ByVal key As String, ByVal defaultValue As Variant) As Variant
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

'===============================================================================
' External Refresh (Python) - Separate entry point with user prompt
'===============================================================================
Public Sub RefreshAll_WithPython()
    Dim response As VbMsgBoxResult
    response = MsgBox("Run Python to rebuild Document_Tracker." & vbCrLf & _
                     "You must reopen the workbook after it finishes." & vbCrLf & vbCrLf & _
                     "Proceed?", vbYesNo + vbQuestion, "External Refresh")

    If response = vbYes Then
        Call TR_Refresh_Document_Tracker
    End If
End Sub
