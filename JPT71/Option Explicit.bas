Option Explicit

'==========================================================
' JPT71 - One Button: Refresh + Create FINAL sheets (values only)
'
' Output:
' 1) _REFRESHED.xlsx : refreshed workbook (may contain formulas)
'    + FINAL sheets inside the same workbook (values only)
'
' IMPORTANT:
' - Put jpt71_refresh_export_final.py in the SAME folder as this workbook
' - Install: pip install pandas==2.33 openpyxl==3.1.5 pywin32
'==========================================================

Private Const PY_EXE As String = "py"  ' 또는 "C:\Path\python.exe"
Private Const TIMEOUT_SECONDS As Long = 300  ' 5분 타임아웃

Private Function GetScriptPath() As String
    On Error GoTo EH

    ' 우선순위 1: jpt71_refresh_export_final.py
    Dim p As String
    p = ThisWorkbook.Path & Application.PathSeparator & "jpt71_refresh_export_final.py"
    If Dir(p) <> "" Then
        GetScriptPath = p
        Exit Function
    End If

    ' 우선순위 2: jpt71_refresh_export_final_22.py (호환성)
    p = ThisWorkbook.Path & Application.PathSeparator & "jpt71_refresh_export_final_22.py"
    If Dir(p) <> "" Then
        GetScriptPath = p
        Exit Function
    End If

    ' fallback: 사용자에게 파일 선택
    Dim picked As Variant
    picked = Application.GetOpenFilename("Python script (*.py),*.py", , "Select jpt71_refresh_export_final.py")
    If picked = False Then
        GetScriptPath = ""
        Exit Function
    End If

    GetScriptPath = CStr(picked)
    Exit Function

EH:
    GetScriptPath = ""
End Function

Public Sub Refresh_And_Export_Final_OneButton()
    On Error GoTo EH

    Dim scriptPath As String
    scriptPath = GetScriptPath()
    If scriptPath = "" Or Dir(scriptPath) = "" Then
        MsgBox "Python 스크립트를 찾을 수 없습니다." & vbCrLf & _
               "jpt71_refresh_export_final.py (또는 jpt71_refresh_export_final_22.py)를 엑셀과 같은 폴더에 두세요.", vbExclamation
        Exit Sub
    End If

    Dim inPath As String, outRef As String
    inPath = ThisWorkbook.FullName
    outRef = Left(inPath, InStrRev(inPath, ".") - 1) & "_REFRESHED.xlsx"

    ' 기존 파일이 열려있으면 닫기 (선택사항)
    On Error Resume Next
    Application.Workbooks(Dir(outRef)).Close SaveChanges:=False
    On Error GoTo EH

    ThisWorkbook.Save

    ' 진행 상황 표시
    Application.StatusBar = "Python 스크립트 실행 중... (잠시만 기다려주세요)"
    Application.ScreenUpdating = False

    ' Python 스크립트 호출 (2개 인자: 입력 파일, 출력 파일)
    Dim cmd As String
    cmd = """" & PY_EXE & """ """ & scriptPath & """ """ & inPath & """ """ & outRef & """"

    Dim sh As Object, ex As Object
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(cmd)

    ' 타임아웃 처리
    Dim startTime As Double
    startTime = Timer
    Dim elapsed As Double

    Do While ex.Status = 0
        elapsed = Timer - startTime
        If elapsed > TIMEOUT_SECONDS Then
            Application.StatusBar = False
            Application.ScreenUpdating = True
            MsgBox "Python 스크립트 실행이 타임아웃되었습니다." & vbCrLf & _
                   "시간 제한: " & TIMEOUT_SECONDS & "초", vbExclamation
            Exit Sub
        End If
        DoEvents
    Loop

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 에러 확인
    If ex.ExitCode <> 0 Then
        Dim errMsg As String
        errMsg = ex.StdErr.ReadAll
        If errMsg = "" Then
            errMsg = ex.StdOut.ReadAll
        End If
        MsgBox "Python 실행 실패:" & vbCrLf & vbCrLf & errMsg, vbExclamation
        Exit Sub
    End If

    ' 파일 생성 확인
    If Dir(outRef) = "" Then
        MsgBox "_REFRESHED 파일이 생성되지 않았습니다." & vbCrLf & outRef, vbExclamation
        Exit Sub
    End If

    ' REFRESHED 파일 열기 (FINAL 시트 포함)
    Application.Workbooks.Open outRef
    MsgBox "완료: REFRESHED 파일을 생성/오픈했습니다." & vbCrLf & vbCrLf & _
           "생성된 파일:" & vbCrLf & _
           "• " & Dir(outRef) & vbCrLf & vbCrLf & _
           "내부 시트:" & vbCrLf & _
           "• Shipping_List_FINAL (값만)" & vbCrLf & _
           "• Mail_Draft_FINAL (값만)" & vbCrLf & _
           "• Cross_Gantt_FINAL (값만)", vbInformation
    Exit Sub

EH:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Refresh_And_Export_Final Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbExclamation
End Sub
