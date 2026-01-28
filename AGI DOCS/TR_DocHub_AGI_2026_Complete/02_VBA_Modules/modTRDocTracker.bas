Attribute VB_Name = "modTRDocTracker"
Option Explicit

'========================================================
' TR Document Tracker - VBA Utilities (Office LTSC 2021)
' 목적:
' 1) Excel에서 버튼 클릭 → Python Refresh 실행
' 2) Overdue/Due Soon 문서 리마인더 메일(Outlook) Draft 생성
' 3) 실행 로그를 LOG 시트에 기록
'========================================================

Private Const SCRIPT_FILE As String = "create_tr_document_tracker_v2.py"

Private Function GetPythonExe() As String
    ' 가장 호환성 높은 방식: Windows Python Launcher
    GetPythonExe = "py"
End Function

Private Function GetScriptPath() As String
    Dim basePath As String
    basePath = ThisWorkbook.Path
    GetScriptPath = basePath & Application.PathSeparator & ".." & Application.PathSeparator & _
                    "01_Python_Builders" & Application.PathSeparator & SCRIPT_FILE
End Function

Private Function GetCfgValue(ByVal keyName As String, Optional ByVal defaultValue As Variant) As Variant
    ' Config 시트: A열 Key / B열 Value
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Config")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 4 To lastRow
        If CStr(ws.Cells(r, 1).Value) = keyName Then
            GetCfgValue = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
    GetCfgValue = defaultValue
    Exit Function
EH:
    GetCfgValue = defaultValue
End Function

Private Sub WriteLog(ByVal actionName As String, ByVal message As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("LOG")
    Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = actionName
    ws.Cells(nextRow, 3).Value = Environ$("username")
    ws.Cells(nextRow, 4).Value = message
End Sub

Public Sub TR_Refresh_Document_Tracker()
    On Error GoTo EH

    Dim py As String: py = GetPythonExe()
    Dim script As String: script = GetScriptPath()
    Dim wbFile As String: wbFile = ThisWorkbook.FullName

    If Dir(script) = "" Then
        MsgBox "Python script not found: " & script & vbCrLf & _
               "→ 동일 폴더에 " & SCRIPT_FILE & " 파일을 저장하세요.", vbCritical
        Exit Sub
    End If

    Dim cmd As String
    cmd = py & " " & """" & script & """" & " --refresh " & """" & wbFile & """"

    WriteLog "REFRESH", "Run: " & cmd

    ' Python 실행 (숨김)
    Shell cmd, vbHide

    MsgBox "Refresh started." & vbCrLf & _
           "Python 완료 후, 파일을 다시 열면 Document_Tracker가 갱신됩니다.", vbInformation
    Exit Sub

EH:
    WriteLog "REFRESH_ERROR", Err.Description
    MsgBox "Refresh failed: " & Err.Description, vbCritical
End Sub

Public Sub TR_Draft_Reminder_Emails()
    On Error GoTo EH

    Dim dueSoonDays As Long
    dueSoonDays = CLng(GetCfgValue("DueSoon_Threshold_Days", 7))

    Dim wsT As Worksheet: Set wsT = ThisWorkbook.Worksheets("Document_Tracker")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets("Party_Contacts")

    Dim lo As ListObject: Set lo = wsT.ListObjects("tblTracker")

    Dim colParty As Long: colParty = lo.ListColumns("Responsible Party").Index
    Dim colDue As Long: colDue = lo.ListColumns("SUBMISSION DATE").Index
    Dim colStatus As Long: colStatus = lo.ListColumns("Status").Index
    Dim colDoc As Long: colDoc = lo.ListColumns("Document Name").Index
    Dim colVoy As Long: colVoy = lo.ListColumns("Voyage").Index
    Dim colRemark As Long: colRemark = lo.ListColumns("Remarks").Index

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim dictMail As Object: Set dictMail = CreateObject("Scripting.Dictionary") ' party -> email

    ' Load contacts
    Dim loC As ListObject: Set loC = wsC.ListObjects("tblContacts")
    Dim rC As ListRow
    For Each rC In loC.ListRows
        Dim p As String: p = CStr(rC.Range.Cells(1, 1).Value)
        Dim em As String: em = CStr(rC.Range.Cells(1, 2).Value)
        If Len(p) > 0 Then dictMail(p) = em
    Next rC

    Dim r As ListRow
    For Each r In lo.ListRows
        Dim party As String: party = CStr(r.Range.Cells(1, colParty).Value)
        Dim dueVal As Variant: dueVal = r.Range.Cells(1, colDue).Value
        Dim st As String: st = CStr(r.Range.Cells(1, colStatus).Value)

        If Len(party) = 0 Then GoTo NextRow
        If IsDate(dueVal) = False Then GoTo NextRow

        ' Skip completed / not required
        If st = "Approved" Or st = "Submitted" Or st = "Not Required" Then GoTo NextRow

        Dim dueDt As Date: dueDt = CDate(dueVal)
        Dim daysLeft As Long: daysLeft = DateDiff("d", Date, dueDt)

        If daysLeft < 0 Or daysLeft <= dueSoonDays Then
            Dim line As String
            line = party & vbTab & _
                   CStr(r.Range.Cells(1, colVoy).Value) & vbTab & _
                   CStr(r.Range.Cells(1, colDoc).Value) & vbTab & _
                   Format(dueDt, "yyyy-mm-dd") & vbTab & _
                   st & vbTab & _
                   CStr(r.Range.Cells(1, colRemark).Value)

            If dict.Exists(party) Then
                dict(party) = dict(party) & vbCrLf & line
            Else
                dict(party) = line
            End If
        End If

NextRow:
    Next r

    If dict.Count = 0 Then
        MsgBox "No overdue/due soon items found.", vbInformation
        Exit Sub
    End If

    ' Create Outlook drafts
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")

    Dim k As Variant
    For Each k In dict.Keys
        Dim mail As Object
        Set mail = olApp.CreateItem(0) ' 0 = MailItem

        Dim toAddr As String
        If dictMail.Exists(CStr(k)) Then
            toAddr = CStr(dictMail(CStr(k)))
        Else
            toAddr = "[MASK]@company.com"
        End If

        mail.To = toAddr
        mail.Subject = "[TR Doc Tracker] Due Soon / Overdue Items - " & CStr(k)

        mail.Body = "Hi " & CStr(k) & "," & vbCrLf & vbCrLf & _
                    "Below items are overdue or due within " & dueSoonDays & " days." & vbCrLf & _
                    "(Party | Voyage | Document | Due Date | Status | Remarks)" & vbCrLf & vbCrLf & _
                    dict(k) & vbCrLf & vbCrLf & _
                    "Please update status in the Document_Tracker after submission." & vbCrLf & _
                    "Thanks."

        mail.Display   ' Draft only (no auto-send)
        WriteLog "EMAIL_DRAFT", "Draft created for " & CStr(k) & " (" & toAddr & ")"
    Next k

    MsgBox "Email drafts created: " & dict.Count, vbInformation
    Exit Sub

EH:
    WriteLog "EMAIL_ERROR", Err.Description
    MsgBox "Email draft failed: " & Err.Description, vbCritical
End Sub
