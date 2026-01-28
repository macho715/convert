Attribute VB_Name = "TR_DocTracker_Module"
Option Explicit

'===============================================================================
' HVDC TR TRANSPORTATION - DOCUMENT TRACKER VBA MODULE
' Version: 1.0 | Project: HVDC AGI TR Transportation
'===============================================================================

'===============================================================================
' 1. STATUS CONDITIONAL FORMATTING
'===============================================================================
Public Sub TR_ApplyStatusFormatting()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("T_Tracker")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 5 Then lastRow = 5
    
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(5, 8), ws.Cells(lastRow, 8))
        Select Case cell.Value
            Case "Accepted", "Submitted"
                cell.Interior.Color = RGB(144, 238, 144)
                cell.Font.Color = RGB(0, 100, 0)
                cell.Font.Bold = True
            Case "In Progress"
                cell.Interior.Color = RGB(255, 255, 153)
                cell.Font.Color = RGB(128, 128, 0)
                cell.Font.Bold = False
            Case "Not Started"
                cell.Interior.Color = RGB(255, 204, 204)
                cell.Font.Color = RGB(139, 0, 0)
                cell.Font.Bold = False
            Case "Rejected"
                cell.Interior.Color = RGB(255, 100, 100)
                cell.Font.Color = RGB(100, 0, 0)
                cell.Font.Bold = True
            Case "On Hold", "Waived"
                cell.Interior.Color = RGB(211, 211, 211)
                cell.Font.Color = RGB(128, 128, 128)
                cell.Font.Bold = False
            Case Else
                cell.Interior.ColorIndex = xlNone
                cell.Font.ColorIndex = xlAutomatic
                cell.Font.Bold = False
        End Select
    Next cell
    
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' 2. HIGHLIGHT OVERDUE ITEMS
'===============================================================================
Public Sub TR_HighlightOverdue()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("T_Tracker")
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long
    For r = 5 To lastRow
        Dim dueDate As Variant: dueDate = ws.Cells(r, 7).Value
        Dim status As String: status = CStr(ws.Cells(r, 8).Value)
        
        If IsDate(dueDate) And status <> "Accepted" And status <> "Waived" Then
            If CDate(dueDate) < Date Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Interior.Color = RGB(255, 200, 200)
            ElseIf CDate(dueDate) <= Date + 7 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Interior.Color = RGB(255, 255, 200)
            End If
        End If
    Next r
    
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' 3. EXPORT TO PDF
'===============================================================================
Public Sub EXP_ExportToPDF()
    On Error GoTo EH
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ws.Name & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, _
                            Quality:=xlQualityStandard, IncludeDocProperties:=True
    
    MsgBox "PDF exported: " & filePath, vbInformation
    Exit Sub
    
EH:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

'===============================================================================
' 4. QUICK STATUS UPDATE
'===============================================================================
Public Sub TR_QuickStatusUpdate()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim newStatus As String
    
    Set ws = ActiveSheet
    Set selectedCell = Selection
    
    If ws.Name <> "T_Tracker" Then
        MsgBox "Please run this from T_Tracker sheet.", vbExclamation
        Exit Sub
    End If
    
    If selectedCell.Column <> 8 Then
        MsgBox "Please select a Status cell (Column H).", vbExclamation
        Exit Sub
    End If
    
    newStatus = InputBox("Enter new status:" & vbCrLf & _
                         "1 = Accepted" & vbCrLf & _
                         "2 = In Progress" & vbCrLf & _
                         "3 = Not Started" & vbCrLf & _
                         "4 = Submitted" & vbCrLf & _
                         "5 = Rejected" & vbCrLf & _
                         "6 = On Hold" & vbCrLf & _
                         "7 = Waived", "Quick Status Update", "1")
    
    Select Case newStatus
        Case "1": selectedCell.Value = "Accepted"
        Case "2": selectedCell.Value = "In Progress"
        Case "3": selectedCell.Value = "Not Started"
        Case "4": selectedCell.Value = "Submitted"
        Case "5": selectedCell.Value = "Rejected"
        Case "6": selectedCell.Value = "On Hold"
        Case "7": selectedCell.Value = "Waived"
        Case Else
            MsgBox "Invalid input.", vbExclamation
            Exit Sub
    End Select
    
    Call TR_ApplyStatusFormatting
End Sub

'===============================================================================
' 5. DRAFT REMINDER EMAILS (Placeholder)
'===============================================================================
Public Sub TR_Draft_Reminder_Emails()
    MsgBox "Email draft feature requires Outlook integration." & vbCrLf & _
           "Please configure Party_Contacts sheet with email addresses.", vbInformation
End Sub

'===============================================================================
' 6. FILTER BY STATUS
'===============================================================================
Public Sub TR_FilterByStatus()
    Dim ws As Worksheet
    Dim statusFilter As String
    Dim statusChoice As String
    
    Set ws = ThisWorkbook.Sheets("T_Tracker")
    
    statusChoice = InputBox("Select status to filter:" & vbCrLf & _
                          "1 = Not Started" & vbCrLf & _
                          "2 = In Progress" & vbCrLf & _
                          "3 = Submitted" & vbCrLf & _
                          "4 = Accepted" & vbCrLf & _
                          "5 = Rejected" & vbCrLf & _
                          "0 = Show All", "Filter by Status", "0")
    
    Select Case statusChoice
        Case "1": statusFilter = "Not Started"
        Case "2": statusFilter = "In Progress"
        Case "3": statusFilter = "Submitted"
        Case "4": statusFilter = "Accepted"
        Case "5": statusFilter = "Rejected"
        Case "0"
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ws.Range("A4").AutoFilter
            MsgBox "Filter cleared.", vbInformation
            Exit Sub
        Case Else
            MsgBox "Invalid input.", vbExclamation
            Exit Sub
    End Select
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A4").AutoFilter
    ws.Range("A4").AutoFilter Field:=8, Criteria1:=statusFilter
    
    MsgBox "Filtered by: " & statusFilter, vbInformation
End Sub
