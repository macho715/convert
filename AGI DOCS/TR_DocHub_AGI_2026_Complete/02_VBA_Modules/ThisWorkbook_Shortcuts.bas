Attribute VB_Name = "ThisWorkbook"
'===============================================================================
' ThisWorkbook - Keyboard Shortcuts & Workbook Events
' Version: 1.0
'===============================================================================

Option Explicit

'===============================================================================
' Workbook Open Event - Register Keyboard Shortcuts
'===============================================================================
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Register shortcuts on workbook open
    Application.OnKey "^+R", "RefreshAll_ControlTower"      ' Ctrl+Shift+R: Refresh All
    Application.OnKey "^+P", "EXP_ExportToPDF"              ' Ctrl+Shift+P: Export PDF
    Application.OnKey "^+E", "TR_Draft_Reminder_Emails"     ' Ctrl+Shift+E: Draft Emails
    
    ' Optional: Show welcome message
    ' MsgBox "TR DocHub AGI 2026 loaded." & vbCrLf & _
    '        "Shortcuts:" & vbCrLf & _
    '        "Ctrl+Shift+R: Refresh All" & vbCrLf & _
    '        "Ctrl+Shift+P: Export PDF" & vbCrLf & _
    '        "Ctrl+Shift+E: Draft Reminder Emails", vbInformation
    
    On Error GoTo 0
End Sub

'===============================================================================
' Workbook Before Close Event - Unregister Shortcuts
'===============================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    ' Unregister shortcuts on close to prevent conflicts
    Application.OnKey "^+R"
    Application.OnKey "^+P"
    Application.OnKey "^+E"
    
    On Error GoTo 0
End Sub

'===============================================================================
' Optional: Workbook Activate Event - Refresh on activation
'===============================================================================
Private Sub Workbook_Activate()
    ' Optional: Auto-refresh on activation (comment out if not needed)
    ' Call RefreshAll_ControlTower
End Sub
