Option Explicit

'=====================
' Doc Gap Tracker Macros (v3.1)
'=====================

Sub RefreshAll()
    'Recalculate
    Application.CalculateFull
End Sub

Sub FilterMissing()
    'Filter current sheet by Status = Missing (assumes headers on row 2)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode = False Then Exit Sub
    ws.Range("A2").AutoFilter Field:=4, Criteria1:="Missing"
End Sub

Sub ClearFilters()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If
End Sub

Sub StampUpdated()
    'Stamp last updated time in Executive_Summary!C1
    With ThisWorkbook.Sheets("Executive_Summary")
        .Range("C1").Value = "Last updated: " & Format(Now, "dd-mmm-yy hh:nn")
    End With
End Sub

Sub ApplyScenarioToManual()
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
