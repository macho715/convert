Attribute VB_Name = "ThisWorkbook"
'===============================================================================
' ThisWorkbook - Keyboard Shortcuts & Workbook Events
' Version: 1.0 | Project: HVDC AGI TR Transportation
'
' INSTALLATION:
' 1. Open VBA Editor (Alt+F11)
' 2. Double-click "ThisWorkbook" in Project Explorer
' 3. Copy and paste this code
' 4. Save workbook as .xlsm
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
    Application.OnKey "^+G", "GenerateTrackerRows"          ' Ctrl+Shift+G: Generate Rows
    Application.OnKey "^+D", "RecalcDeadlines"              ' Ctrl+Shift+D: Recalc Deadlines
    Application.OnKey "^+V", "ValidateBeforeExport"         ' Ctrl+Shift+V: Validate
    
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
    Application.OnKey "^+G"
    Application.OnKey "^+D"
    Application.OnKey "^+V"
    
    On Error GoTo 0
End Sub

'===============================================================================
' Optional: Workbook Activate Event
'===============================================================================
Private Sub Workbook_Activate()
    ' Auto-refresh on activation (uncomment if needed)
    ' Call RefreshAll_ControlTower
End Sub

'===============================================================================
' Workbook Sheet Change Event - Auto-format on status change
'===============================================================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    
    ' Auto-apply status formatting when T_Tracker Status column changes
    If Sh.Name = "T_Tracker" And Target.Column = 8 Then
        Call TR_ApplyStatusFormatting
    End If
    
    On Error GoTo 0
End Sub
