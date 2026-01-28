Attribute VB_Name = "TR_DocTracker_Module"
'===============================================================================
' HVDC TR TRANSPORTATION - DOCUMENT TRACKER VBA MODULE
' Version: 1.0
' Date: 2026-01-19
' Author: Samsung C&T - Project Team
'
' 사용법:
' 1. Excel에서 Alt+F11로 VBA 편집기 열기
' 2. 파일 > 가져오기... 선택
' 3. 이 .bas 파일 선택하여 가져오기
' 4. 매크로 실행: Alt+F8 → 원하는 매크로 선택 → 실행
'===============================================================================

Option Explicit

'===============================================================================
' 1. STATUS CONDITIONAL FORMATTING - 상태별 색상 자동 적용
'===============================================================================
Public Sub TR_ApplyStatusFormatting()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Document_Tracker 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Status columns: H, K, N, Q (columns 8, 11, 14, 17)
    Dim statusCols As Variant
    statusCols = Array(8, 11, 14, 17)
    
    Dim col As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For Each col In statusCols
        For Each cell In ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
            Select Case cell.Value
                Case "Complete"
                    cell.Interior.Color = RGB(144, 238, 144)  ' Light Green
                    cell.Font.Color = RGB(0, 100, 0)
                    cell.Font.Bold = True
                Case "In Progress"
                    cell.Interior.Color = RGB(255, 255, 153)  ' Light Yellow
                    cell.Font.Color = RGB(128, 128, 0)
                    cell.Font.Bold = False
                Case "Not Started"
                    cell.Interior.Color = RGB(255, 204, 204)  ' Light Red
                    cell.Font.Color = RGB(139, 0, 0)
                    cell.Font.Bold = False
                Case "Pending Review"
                    cell.Interior.Color = RGB(173, 216, 230)  ' Light Blue
                    cell.Font.Color = RGB(0, 0, 139)
                    cell.Font.Bold = False
                Case "N/A"
                    cell.Interior.Color = RGB(211, 211, 211)  ' Light Gray
                    cell.Font.Color = RGB(128, 128, 128)
                    cell.Font.Bold = False
                Case Else
                    cell.Interior.ColorIndex = xlNone
                    cell.Font.ColorIndex = xlAutomatic
                    cell.Font.Bold = False
            End Select
        Next cell
    Next col
    
    Application.ScreenUpdating = True
    MsgBox "상태별 색상 서식이 적용되었습니다.", vbInformation
End Sub

'===============================================================================
' 2. PROGRESS CALCULATOR - 진행률 계산
'===============================================================================
Public Sub TR_CalculateProgress()
    Dim ws As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim statusCols As Variant
    Dim col As Variant
    Dim voyageNum As Integer
    Dim totalDocs As Integer
    Dim completeDocs As Integer
    Dim inProgressDocs As Integer
    Dim notStartedDocs As Integer
    Dim naDocs As Integer
    Dim cell As Range
    Dim progressPct As Double
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    statusCols = Array(8, 11, 14, 17)  ' V1, V2, V3, V4
    
    ' Calculate for each voyage
    voyageNum = 1
    For Each col In statusCols
        totalDocs = 0
        completeDocs = 0
        inProgressDocs = 0
        notStartedDocs = 0
        naDocs = 0
        
        For Each cell In ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
            If cell.Value <> "" Then
                totalDocs = totalDocs + 1
                Select Case cell.Value
                    Case "Complete"
                        completeDocs = completeDocs + 1
                    Case "In Progress"
                        inProgressDocs = inProgressDocs + 1
                    Case "Not Started"
                        notStartedDocs = notStartedDocs + 1
                    Case "N/A"
                        naDocs = naDocs + 1
                End Select
            End If
        Next cell
        
        ' Calculate progress percentage (excluding N/A)
        If (totalDocs - naDocs) > 0 Then
            progressPct = (completeDocs / (totalDocs - naDocs)) * 100
        Else
            progressPct = 0
        End If
        
        ' Update Dashboard - Voyage status (row 7-10, column 9)
        wsDash.Cells(6 + voyageNum, 9).Value = Format(progressPct, "0.0") & "% (" & completeDocs & "/" & (totalDocs - naDocs) & ")"
        
        ' Color code the progress
        If progressPct >= 100 Then
            wsDash.Cells(6 + voyageNum, 9).Interior.Color = RGB(144, 238, 144)
        ElseIf progressPct >= 75 Then
            wsDash.Cells(6 + voyageNum, 9).Interior.Color = RGB(173, 216, 230)
        ElseIf progressPct >= 50 Then
            wsDash.Cells(6 + voyageNum, 9).Interior.Color = RGB(255, 255, 153)
        Else
            wsDash.Cells(6 + voyageNum, 9).Interior.Color = RGB(255, 204, 204)
        End If
        
        voyageNum = voyageNum + 1
    Next col
    
    MsgBox "진행률이 업데이트되었습니다." & vbCrLf & _
           "Dashboard 시트에서 확인하세요.", vbInformation
End Sub

'===============================================================================
' 3. PARTY-WISE PROGRESS CALCULATOR - 파티별 진행률 계산
'===============================================================================
Public Sub TR_CalculatePartyProgress()
    Dim ws As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim partyCol As Integer
    Dim cell As Range
    Dim parties As Collection
    Dim partyStats As Object
    Dim party As Variant
    Dim statusVal As String
    Dim dashRow As Integer
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set partyStats = CreateObject("Scripting.Dictionary")
    
    partyCol = 6  ' Column F - Responsible Party
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize party statistics
    Dim partyList As Variant
    partyList = Array("Samsung C&T", "Mammoet", "OFCO Agency", "ADNOC L&S", _
                      "Vessel Owner (KFS)", "DSV Solutions", "MWS (Sterling)")
    
    For Each party In partyList
        partyStats(party) = Array(0, 0, 0, 0, 0)  ' Total, Complete, InProgress, NotStarted, NA
    Next party
    
    ' Count documents by party and status (using V1 status as reference)
    Dim r As Long
    For r = 5 To lastRow
        party = ws.Cells(r, partyCol).Value
        statusVal = ws.Cells(r, 8).Value  ' V1 Status column
        
        If partyStats.Exists(party) Then
            Dim stats As Variant
            stats = partyStats(party)
            stats(0) = stats(0) + 1  ' Total
            
            Select Case statusVal
                Case "Complete"
                    stats(1) = stats(1) + 1
                Case "In Progress"
                    stats(2) = stats(2) + 1
                Case "Not Started"
                    stats(3) = stats(3) + 1
                Case "N/A"
                    stats(4) = stats(4) + 1
            End Select
            
            partyStats(party) = stats
        End If
    Next r
    
    ' Update Dashboard - Party Progress (starting row 14)
    dashRow = 14
    For Each party In partyList
        stats = partyStats(party)
        
        wsDash.Cells(dashRow, 1).Value = party
        wsDash.Cells(dashRow, 2).Value = stats(0)  ' Total
        wsDash.Cells(dashRow, 3).Value = stats(1)  ' Complete
        wsDash.Cells(dashRow, 4).Value = stats(2)  ' In Progress
        wsDash.Cells(dashRow, 5).Value = stats(3)  ' Not Started
        wsDash.Cells(dashRow, 6).Value = stats(4)  ' N/A
        
        ' Progress %
        If (stats(0) - stats(4)) > 0 Then
            wsDash.Cells(dashRow, 7).Value = Format((stats(1) / (stats(0) - stats(4))) * 100, "0.0") & "%"
        Else
            wsDash.Cells(dashRow, 7).Value = "N/A"
        End If
        
        dashRow = dashRow + 1
    Next party
    
    MsgBox "파티별 진행률이 업데이트되었습니다.", vbInformation
End Sub

'===============================================================================
' 4. HIGHLIGHT OVERDUE DOCUMENTS - 마감일 초과 서류 강조
'===============================================================================
Public Sub TR_HighlightOverdue()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dateCols As Variant
    Dim col As Variant
    Dim deadlines As Variant
    Dim voyageIdx As Integer
    Dim docDeadline As Date
    Dim statusVal As String
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Date columns and corresponding deadlines (2026)
    dateCols = Array(9, 12, 15, 18)  ' V1, V2, V3, V4 date columns
    deadlines = Array(#1/23/2026#, #2/3/2026#, #2/12/2026#, #2/20/2026#)
    
    Application.ScreenUpdating = False
    
    ' Clear existing highlights first
    ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, 19)).Interior.ColorIndex = xlNone
    
    voyageIdx = 0
    For Each col In dateCols
        docDeadline = deadlines(voyageIdx)
        
        For r = 5 To lastRow
            statusVal = ws.Cells(r, col - 1).Value  ' Status is one column before date
            
            ' If not complete and date is empty or past deadline
            If statusVal <> "Complete" And statusVal <> "N/A" Then
                If Date > docDeadline Then
                    ' Highlight the entire row for this voyage columns
                    ws.Range(ws.Cells(r, col - 1), ws.Cells(r, col + 1)).Interior.Color = RGB(255, 182, 193)  ' Light pink
                End If
            End If
        Next r
        
        voyageIdx = voyageIdx + 1
    Next col
    
    Application.ScreenUpdating = True
    MsgBox "마감일 초과 서류가 강조되었습니다." & vbCrLf & _
           "(분홍색으로 표시)", vbInformation
End Sub

'===============================================================================
' 5. EXPORT TO PDF - PDF 내보내기
'===============================================================================
Public Sub EXP_ExportToPDF()
    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String
    
    Set ws = ActiveSheet
    
    filePath = ThisWorkbook.Path
    If filePath = "" Then filePath = Environ("USERPROFILE") & "\Documents"
    
    fileName = filePath & "\TR_Document_Tracker_" & Format(Now, "YYYYMMDD_HHMMSS") & ".pdf"
    
    On Error GoTo ErrHandler
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           fileName:=fileName, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=True
    
    MsgBox "PDF가 생성되었습니다:" & vbCrLf & fileName, vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "PDF 생성 중 오류 발생: " & Err.Description, vbExclamation
End Sub

'===============================================================================
' 6. SEND EMAIL REMINDER - 이메일 알림 (Outlook 연동)
'===============================================================================
Public Sub EXP_SendEmailReminder()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim overdueList As String
    Dim partyEmail As Object
    Dim outlook As Object
    Dim mail As Object
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    Set partyEmail = CreateObject("Scripting.Dictionary")
    
    ' Party email mapping
    partyEmail("OFCO Agency") = "nkk@ofco-int.com"
    partyEmail("Mammoet") = "Yulia.Frolova@mammoet.com"
    partyEmail("ADNOC L&S") = "moda@adnoc.ae"
    partyEmail("Vessel Owner (KFS)") = "lct.bushra@khalidfarajshipping.com"
    partyEmail("DSV Solutions") = "jay.manaloto@dsv.com"
    partyEmail("Samsung C&T") = ""
    partyEmail("MWS (Sterling)") = ""
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Build overdue document list
    overdueList = "HVDC TR Transportation - Pending Documents:" & vbCrLf & vbCrLf
    
    For r = 5 To lastRow
        If ws.Cells(r, 8).Value = "Not Started" Or ws.Cells(r, 8).Value = "In Progress" Then
            overdueList = overdueList & "• " & ws.Cells(r, 3).Value & _
                         " (" & ws.Cells(r, 6).Value & ")" & vbCrLf
        End If
    Next r
    
    ' Create Outlook email
    On Error GoTo OutlookError
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)
    
    With mail
        .Subject = "[HVDC TR] Document Submission Reminder - " & Format(Date, "YYYY-MM-DD")
        .Body = overdueList & vbCrLf & vbCrLf & _
                "Please submit pending documents before the deadline." & vbCrLf & _
                "Contact: Samsung C&T Project Team"
        .Display  ' Use .Send to send directly
    End With
    
    MsgBox "이메일 초안이 생성되었습니다.", vbInformation
    Exit Sub
    
OutlookError:
    MsgBox "Outlook 연결 실패: " & Err.Description, vbExclamation
End Sub

'===============================================================================
' 7. AUTO-REFRESH ALL - 모든 기능 자동 실행
'===============================================================================
Public Sub TR_AutoRefreshAll()
    Call TR_ApplyStatusFormatting
    Call TR_CalculateProgress
    Call TR_CalculatePartyProgress
    Call TR_HighlightOverdue
    
    ' Update timestamp on Dashboard
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    wsDash.Cells(3, 1).Value = "Last Updated: " & Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    MsgBox "모든 데이터가 업데이트되었습니다!", vbInformation
End Sub

'===============================================================================
' 8. CREATE VOYAGE SUMMARY REPORT - 항차별 요약 리포트 생성
'===============================================================================
Public Sub TR_CreateVoyageSummaryReport()
    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim voyageNum As Integer
    Dim statusCols As Variant
    Dim startRow As Integer
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    
    ' Create or clear Report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Summary_Report")
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsReport.Name = "Summary_Report"
    Else
        wsReport.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Title
    wsReport.Range("A1:G1").Merge
    wsReport.Cells(1, 1).Value = "HVDC TR Transportation - Document Status Summary Report"
    wsReport.Cells(1, 1).Font.Bold = True
    wsReport.Cells(1, 1).Font.Size = 14
    
    wsReport.Cells(2, 1).Value = "Generated: " & Format(Now, "YYYY-MM-DD HH:MM")
    
    statusCols = Array(8, 11, 14, 17)
    Dim voyageNames As Variant
    voyageNames = Array("Voyage 1 (TR 1-2)", "Voyage 2 (TR 3-4)", "Voyage 3 (TR 5-6)", "Voyage 4 (TR 7)")
    
    startRow = 4
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For voyageNum = 0 To 3
        ' Voyage header
        wsReport.Cells(startRow, 1).Value = voyageNames(voyageNum)
        wsReport.Cells(startRow, 1).Font.Bold = True
        wsReport.Cells(startRow, 1).Interior.Color = RGB(30, 58, 95)
        wsReport.Cells(startRow, 1).Font.Color = RGB(255, 255, 255)
        wsReport.Range(wsReport.Cells(startRow, 1), wsReport.Cells(startRow, 5)).Merge
        
        ' Headers
        startRow = startRow + 1
        wsReport.Cells(startRow, 1).Value = "Document"
        wsReport.Cells(startRow, 2).Value = "Responsible"
        wsReport.Cells(startRow, 3).Value = "Priority"
        wsReport.Cells(startRow, 4).Value = "Status"
        wsReport.Cells(startRow, 5).Value = "Remarks"
        
        Dim col As Integer
        For col = 1 To 5
            wsReport.Cells(startRow, col).Font.Bold = True
            wsReport.Cells(startRow, col).Interior.Color = RGB(200, 200, 200)
        Next col
        
        ' Document list
        For r = 5 To lastRow
            Dim statusVal As String
            statusVal = ws.Cells(r, statusCols(voyageNum)).Value
            
            If statusVal <> "N/A" And statusVal <> "" Then
                startRow = startRow + 1
                wsReport.Cells(startRow, 1).Value = ws.Cells(r, 3).Value  ' Document name
                wsReport.Cells(startRow, 2).Value = ws.Cells(r, 6).Value  ' Responsible
                wsReport.Cells(startRow, 3).Value = ws.Cells(r, 5).Value  ' Priority
                wsReport.Cells(startRow, 4).Value = statusVal              ' Status
                wsReport.Cells(startRow, 5).Value = ws.Cells(r, statusCols(voyageNum) + 2).Value  ' Remarks
                
                ' Color code status
                Select Case statusVal
                    Case "Complete"
                        wsReport.Cells(startRow, 4).Interior.Color = RGB(144, 238, 144)
                    Case "In Progress"
                        wsReport.Cells(startRow, 4).Interior.Color = RGB(255, 255, 153)
                    Case "Not Started"
                        wsReport.Cells(startRow, 4).Interior.Color = RGB(255, 204, 204)
                    Case "Pending Review"
                        wsReport.Cells(startRow, 4).Interior.Color = RGB(173, 216, 230)
                End Select
            End If
        Next r
        
        startRow = startRow + 3
    Next voyageNum
    
    ' Auto-fit columns
    wsReport.Columns("A:E").AutoFit
    
    MsgBox "Summary Report가 생성되었습니다.", vbInformation
    wsReport.Activate
End Sub

'===============================================================================
' 9. QUICK STATUS UPDATE - 빠른 상태 업데이트
'===============================================================================
Public Sub TR_QuickStatusUpdate()
    Dim ws As Worksheet
    Dim selectedCell As Range
    Dim newStatus As String
    
    Set ws = ActiveSheet
    Set selectedCell = Selection
    
    If ws.Name <> "Document_Tracker" Then
        MsgBox "Document_Tracker 시트에서 실행해주세요.", vbExclamation
        Exit Sub
    End If
    
    ' Check if selected cell is in status column
    If selectedCell.Column <> 8 And selectedCell.Column <> 11 And _
       selectedCell.Column <> 14 And selectedCell.Column <> 17 Then
        MsgBox "상태 열(V1/V2/V3/V4 Status)을 선택하세요.", vbExclamation
        Exit Sub
    End If
    
    newStatus = InputBox("새 상태 입력:" & vbCrLf & _
                         "1 = Complete" & vbCrLf & _
                         "2 = In Progress" & vbCrLf & _
                         "3 = Not Started" & vbCrLf & _
                         "4 = Pending Review" & vbCrLf & _
                         "5 = N/A", "Quick Status Update", "1")
    
    Select Case newStatus
        Case "1"
            selectedCell.Value = "Complete"
        Case "2"
            selectedCell.Value = "In Progress"
        Case "3"
            selectedCell.Value = "Not Started"
        Case "4"
            selectedCell.Value = "Pending Review"
        Case "5"
            selectedCell.Value = "N/A"
        Case Else
            MsgBox "유효하지 않은 입력입니다.", vbExclamation
            Exit Sub
    End Select
    
    ' Auto-apply formatting
    Call TR_ApplyStatusFormatting
End Sub

'===============================================================================
' 10. FILTER BY PARTY - 파티별 필터
'===============================================================================
Public Sub TR_FilterByParty()
    Dim ws As Worksheet
    Dim partyFilter As String
    Dim partyChoice As String
    
    Set ws = ThisWorkbook.Sheets("Document_Tracker")
    
    partyChoice = InputBox("파티 선택:" & vbCrLf & _
                          "1 = Samsung C&T" & vbCrLf & _
                          "2 = Mammoet" & vbCrLf & _
                          "3 = OFCO Agency" & vbCrLf & _
                          "4 = ADNOC L&S" & vbCrLf & _
                          "5 = Vessel Owner (KFS)" & vbCrLf & _
                          "6 = DSV Solutions" & vbCrLf & _
                          "7 = MWS (Sterling)" & vbCrLf & _
                          "0 = 모두 보기", "Filter by Party", "0")
    
    Select Case partyChoice
        Case "1"
            partyFilter = "Samsung C&T"
        Case "2"
            partyFilter = "Mammoet"
        Case "3"
            partyFilter = "OFCO Agency"
        Case "4"
            partyFilter = "ADNOC L&S"
        Case "5"
            partyFilter = "Vessel Owner (KFS)"
        Case "6"
            partyFilter = "DSV Solutions"
        Case "7"
            partyFilter = "MWS (Sterling)"
        Case "0"
            ws.AutoFilterMode = False
            ws.Range("A4").AutoFilter
            MsgBox "필터가 해제되었습니다.", vbInformation
            Exit Sub
        Case Else
            MsgBox "유효하지 않은 입력입니다.", vbExclamation
            Exit Sub
    End Select
    
    ' Apply filter
    ws.AutoFilterMode = False
    ws.Range("A4").AutoFilter
    ws.Range("A4").AutoFilter Field:=6, Criteria1:=partyFilter
    
    MsgBox partyFilter & " 서류만 표시됩니다.", vbInformation
End Sub

'===============================================================================
' WORKSHEET EVENTS (이 코드는 ThisWorkbook 또는 Sheet 모듈에 복사)
'===============================================================================
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' Auto-apply formatting when status changes
'     If Target.Column = 8 Or Target.Column = 11 Or _
'        Target.Column = 14 Or Target.Column = 17 Then
'         Call TR_ApplyStatusFormatting
'     End If
' End Sub

'===============================================================================
' KEYBOARD SHORTCUTS (Workbook_Open 이벤트에 추가)
'===============================================================================
' Private Sub Workbook_Open()
'     ' Ctrl+Shift+R = Refresh All
'     Application.OnKey "^+R", "TR_AutoRefreshAll"
'     ' Ctrl+Shift+S = Quick Status Update
'     Application.OnKey "^+S", "TR_QuickStatusUpdate"
'     ' Ctrl+Shift+P = Export to PDF
'     Application.OnKey "^+P", "EXP_ExportToPDF"
' End Sub
'
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     ' Remove shortcuts
'     Application.OnKey "^+R"
'     Application.OnKey "^+S"
'     Application.OnKey "^+P"
' End Sub
