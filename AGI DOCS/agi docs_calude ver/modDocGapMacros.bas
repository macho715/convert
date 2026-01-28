Attribute VB_Name = "modDocGapMacros"
Option Explicit

'=====================
' Doc Gap Tracker Macros (v3.1)
' Version: 1.0 | Project: HVDC AGI TR Transportation
'=====================

Public Sub DG_RefreshAll()
    'Recalculate all formulas
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
    'Stamp last updated time in Dashboard
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("D_Dashboard")
    If Not wsDash Is Nothing Then
        wsDash.Range("A3").Value = "Last updated: " & Format(Now, "dd-mmm-yy hh:nn")
    End If
    On Error GoTo 0
End Sub

Public Sub DG_ApplyScenarioToManual()
    'Copy auto schedule dates to manual and set scenario to CUSTOM
    Dim wsInputs As Worksheet
    On Error Resume Next
    Set wsInputs = ThisWorkbook.Sheets("Inputs")
    If wsInputs Is Nothing Then Exit Sub
    
    With wsInputs
        .Range("C3").Value = .Range("B3").Value
        .Range("C4").Value = .Range("B4").Value
        .Range("C5").Value = .Range("B5").Value
        .Range("C6").Value = .Range("B6").Value
        .Range("B9").Value = "CUSTOM"
    End With
    Application.CalculateFull
    On Error GoTo 0
End Sub
