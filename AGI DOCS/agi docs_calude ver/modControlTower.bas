Attribute VB_Name = "modControlTower"
Option Explicit

'===============================================================================
' CONTROL TOWER - Single Entry Point for All Refresh Operations
' Version: 1.0 | Project: HVDC AGI TR Transportation
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
    
    Dim startTime As Double: startTime = Timer
    AppStateGuard_Begin
    
    ' 1) Recalculate all formulas
    Call DG_RefreshAll
    
    ' 2) TR Tracker operations
    Call TR_ApplyStatusFormatting
    
    ' 3) Dashboard timestamp update
    Call UpdateDashboardTimestamp
    
    AppStateGuard_End
    
    MsgBox "Control Tower Refresh completed" & vbCrLf & _
           "Elapsed time: " & Format(Timer - startTime, "0.00") & " sec", vbInformation
    Exit Sub
    
EH:
    AppStateGuard_End
    MsgBox "Error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

'===============================================================================
' Update Dashboard Timestamp
'===============================================================================
Private Sub UpdateDashboardTimestamp()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("D_Dashboard")
    If Not wsDash Is Nothing Then
        wsDash.Cells(3, 1).Value = "Last Updated: " & Format(Now, "YYYY-MM-DD HH:MM:SS")
    End If
    On Error GoTo 0
End Sub

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
