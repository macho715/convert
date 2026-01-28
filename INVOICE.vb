Option Explicit
'==========================
' Module: modApplyFormula (Step 2 - 최종 완성 버전)
'==========================

'--- 헤더 상수 ---
Private Const HDR_REMARK As String = "REMARK"
Private Const HDR_TOTAL As String = "TOTAL (USD)"
Private Const HDR_RATE As String = "RATE"
Private Const HDR_QTY As String = "Q'TY"
Private Const HDR_REV1 As String = "REV RATE"
Private Const HDR_REV2 As String = "REV TOTAL"
Private Const HDR_DIFF As String = "DIFFERENCE"
Private Const EXCLUDE_SHEET1 As String = "SUMMARY"
Private Const EXCLUDE_SHEET2 As String = "DEC"
Private Const EXCLUDE_SHEET3 As String = "MasterData"

'--- 보조 함수 ---
Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Rows(headerRow).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If Not c Is Nothing Then FindHeaderCol = c.Column
End Function

'--- 메인 실행 프로시저 ---
Public Sub ApplyFormula_ByDynamicRemark_ExactTotal_Safe()
    Dim t0 As Single: t0 = Timer
    On Error GoTo ErrH
    AppBegin "ApplyFormula"
    LogActionSafe "ApplyFormula", "BEGIN"

    ApplyFormula_Impl

    LogActionSafe "ApplyFormula", "END " & Format(Timer - t0, "0.00s")
Done:
    AppEnd
    Exit Sub
ErrH:
    LogActionSafe "ApplyFormula", "ERR: " & Err.description & " (" & Err.Number & ")"
    Resume Done
End Sub

'--- 실제 로직 구현부 ---
Private Sub ApplyFormula_Impl()
    Dim ws As Worksheet, remarkCell As Range, totalCell As Range
    Dim headerRow As Long, revCol As Long
    Dim lLastRow As Long
    Dim rateCol As Long, qtyCol As Long, totalCol As Long, descCol As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And UCase(ws.Name) <> EXCLUDE_SHEET1 And UCase(ws.Name) <> EXCLUDE_SHEET2 And UCase(ws.Name) <> EXCLUDE_SHEET3 Then

            Set remarkCell = SafeFind(ws, HDR_REMARK, True)
            If remarkCell Is Nothing Then GoTo NextWs

            headerRow = remarkCell.Row
            
            ' [GEMINI FIX] 데이터 존재 판단 기준을 'DESCRIPTION' 열로 변경
            descCol = FindHeaderCol(ws, headerRow, "DESCRIPTION")
            If descCol = 0 Then GoTo NextWs ' DESCRIPTION 헤더가 없으면 스킵
            
            lLastRow = lastDataRow(ws, descCol)
            
            If lLastRow <= headerRow Then GoTo NextWs

            rateCol = FindHeaderCol(ws, headerRow, HDR_RATE)
            qtyCol = FindHeaderCol(ws, headerRow, HDR_QTY)
            totalCol = FindHeaderCol(ws, headerRow, HDR_TOTAL)
            
            If rateCol = 0 Or qtyCol = 0 Or totalCol = 0 Then GoTo NextWs

            revCol = remarkCell.Column + 1
            ws.Cells(headerRow, revCol).Value = HDR_REV1
            ws.Cells(headerRow, revCol + 1).Value = HDR_REV2
            ws.Cells(headerRow, revCol + 2).Value = HDR_DIFF
            ws.Range(ws.Cells(headerRow, revCol), ws.Cells(headerRow, revCol + 2)).Font.Bold = True

            ClearRange ws.Range(ws.Cells(headerRow + 1, revCol), ws.Cells(lLastRow, revCol + 2))

            ws.Range(ws.Cells(headerRow + 1, revCol), ws.Cells(lLastRow, revCol)).FormulaR1C1 = "=ROUND(RC" & rateCol & ",2)"
            ws.Range(ws.Cells(headerRow + 1, revCol + 1), ws.Cells(lLastRow, revCol + 1)).FormulaR1C1 = "=RC[-1]*RC" & qtyCol
            ws.Range(ws.Cells(headerRow + 1, revCol + 2), ws.Cells(lLastRow, revCol + 2)).FormulaR1C1 = "=RC[-1]-RC" & totalCol

            Set totalCell = SafeFind(ws, "TOTAL", True)
            
            If Not totalCell Is Nothing Then
                With ws.Cells(totalCell.Row, revCol + 1)
                    .FormulaR1C1 = "=SUM(R" & (headerRow + 1) & "C:R[-1]C)"
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                End With
                With ws.Cells(totalCell.Row, revCol + 2)
                    .FormulaR1C1 = "=SUM(R" & (headerRow + 1) & "C:R[-1]C)"
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Bold = True
                End With
            End If
            
            Dim formatEndRow As Long
            If totalCell Is Nothing Then
                formatEndRow = lLastRow
            Else
                formatEndRow = totalCell.Row
            End If
            
            ws.Range(ws.Cells(headerRow + 1, revCol), ws.Cells(formatEndRow, revCol + 2)).NumberFormat = "#,##0.00"
            
        End If
NextWs:
    Next ws
End Sub


Option Explicit

'================================================================
' Module: modCompileMaster
' Purpose: 모든 시트를 순회하며 '데이터 행'만 필터링하여 마스터 시트에 취합
'================================================================

'--- 메인 실행 프로시저 ---
Public Sub CompileAllSheets()
    Dim t0 As Single: t0 = Timer
    
    '--- 1. 환경 설정 및 출력 시트 준비 ---
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    
    Dim outWS As Worksheet
    On Error Resume Next
    Set outWS = ThisWorkbook.Worksheets("MasterData")
    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        outWS.Name = "MasterData"
    End If
    On Error GoTo ErrH
    outWS.Cells.Clear
    
    '--- 2. 마스터 시트에 헤더 작성 ---
    Dim headers As Variant
    headers = Array("CWI Job Number", "Order Ref. Number", "S/No", "RATE SOURCE", "DESCRIPTION", "RATE", "Formula", "Q'TY", "TOTAL (USD)", "REMARK", "REV RATE", "REV TOTAL", "DIFFERENCE")
    outWS.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    outWS.Rows(1).Font.Bold = True
    
    '--- 3. 모든 시트를 순회하며 데이터 추출 ---
    Dim ws As Worksheet
    Dim writeRow As Long: writeRow = 2
    
    For Each ws In ThisWorkbook.Worksheets
        ' 마스터 시트이거나 숨겨진 시트는 건너뛰기
        If ws.Name <> outWS.Name And ws.Visible = xlSheetVisible Then
            
            '--- 4. 시트 상단의 헤더 정보 추출 ---
            Dim jobNum As String, orderRef As String
            jobNum = GetValueFromLabel(ws, "CW1 Job Number")
            orderRef = GetValueFromLabel(ws, "Order Ref. Number")
            
            '--- 5. [로직 변경] 시트의 모든 행을 검사하여 데이터 행만 추출 ---
            Dim firstHeaderRow As Long, lastUsedRow As Long
            Dim snCol As Long, lastCol As Long
            Dim r As Long
            
            ' S/No 헤더를 찾아 기준 열로 설정
            firstHeaderRow = FindHeaderRow(ws, "S/No")
            If firstHeaderRow > 0 Then
                snCol = FindCol(ws, firstHeaderRow, "S/No")
                lastCol = ws.Cells(firstHeaderRow, ws.Columns.Count).End(xlToLeft).Column
                lastUsedRow = ws.Cells(ws.Rows.Count, snCol).End(xlUp).Row
                
                ' 헤더 아래부터 마지막 사용된 행까지 순회
                For r = firstHeaderRow + 1 To lastUsedRow
                    ' S/No 열에 숫자가 있는 행만 데이터 행으로 간주
                    If IsNumeric(ws.Cells(r, snCol).Value) And Not IsEmpty(ws.Cells(r, snCol).Value) Then
                        ' 헤더 정보 쓰기
                        outWS.Cells(writeRow, 1).Value = jobNum
                        outWS.Cells(writeRow, 2).Value = orderRef
                        
                        ' 데이터 행 전체 복사
                        ws.Range(ws.Cells(r, snCol), ws.Cells(r, lastCol)).Copy
                        outWS.Cells(writeRow, 3).PasteSpecial xlPasteValues
                        
                        writeRow = writeRow + 1
                    End If
                Next r
            End If
        End If
    Next ws
    
    '--- 6. 마무리 ---
    Application.CutCopyMode = False
    outWS.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "모든 시트의 데이터 행 취합 완료!", vbInformation, "작업 완료"
    Exit Sub

ErrH:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "오류 발생: " & Err.description, vbCritical, "오류"
End Sub


'--- 보조 함수 (Helper Functions) ---

Private Function GetValueFromLabel(ws As Worksheet, labelText As String) As String
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.UsedRange.Find(What:=labelText, LookIn:=xlValues, LookAt:=xlPart)
    On Error GoTo 0
    
    If Not foundCell Is Nothing Then
        GetValueFromLabel = CStr(foundCell.Offset(0, 1).Value)
    Else
        GetValueFromLabel = ""
    End If
End Function

Private Function FindHeaderRow(ws As Worksheet, headerText As String) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.UsedRange.Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not foundCell Is Nothing Then
        FindHeaderRow = foundCell.Row
    Else
        FindHeaderRow = 0
    End If
End Function

Private Function FindCol(ws As Worksheet, r As Long, headerText As String) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Rows(r).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not foundCell Is Nothing Then
        FindCol = foundCell.Column
    Else
        FindCol = 0
    End If
End Function

Option Explicit
'==========================
' Module: modExtractFormulas (Step 1)
'==========================

Private Const HDR_RATE As String = "RATE"
Private Const HDR_FORM As String = "Formula"
Private Const EXCLUDE_SHEET1 As String = "FEB"
Private Const EXCLUDE_SHEET2 As String = "InvoiceData"
Private Const EXCLUDE_SHEET3 As String = "SUMMARY"

Public Sub ExtractFormulasWithExclusion()
    Dim t0 As Single: t0 = Timer
    On Error GoTo ErrH
    AppBegin "ExtractFormulas"
    LogActionSafe "ExtractFormulas", "BEGIN"

    ExtractFormulas_Impl

    LogActionSafe "ExtractFormulas", "END " & Format(Timer - t0, "0.00s")
Done:
    AppEnd
    Exit Sub
ErrH:
    LogActionSafe "ExtractFormulas", "ERR: " & Err.description & " (" & Err.Number & ")"
    Resume Done
End Sub

Private Sub ExtractFormulas_Impl()
    Dim ws As Worksheet
    Dim headerCell As Range
    Dim rateCol As Long, formCol As Long
    Dim firstDataRow As Long, lLastRow As Long
    Dim srcRng As Range, arrF As Variant, arrOut() As String
    Dim r As Long, nRows As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And UCase(ws.Name) <> EXCLUDE_SHEET1 And _
           UCase(ws.Name) <> EXCLUDE_SHEET2 And UCase(ws.Name) <> EXCLUDE_SHEET3 Then

            Set headerCell = SafeFind(ws, HDR_RATE, True)
            If headerCell Is Nothing Then GoTo NextWs

            rateCol = headerCell.Column
            firstDataRow = headerCell.Row + 1
            lLastRow = lastDataRow(ws, rateCol)
            If lLastRow < firstDataRow Then GoTo NextWs

            formCol = rateCol + 1
            If ws.Cells(headerCell.Row, formCol).Value2 <> HDR_FORM Then
                ws.Columns(formCol).Insert Shift:=xlToRight
                ws.Cells(headerCell.Row, formCol).Value2 = HDR_FORM
                ws.Cells(headerCell.Row, formCol).Font.Bold = True
            End If

            With ws.Range(ws.Cells(firstDataRow, formCol), ws.Cells(lLastRow, formCol))
                .NumberFormat = "@"
                .ClearContents
            End With

            Set srcRng = ws.Range(ws.Cells(firstDataRow, rateCol), ws.Cells(lLastRow, rateCol))
            If srcRng.Rows.Count = 1 Then
                ReDim arrF(1 To 1, 1 To 1)
                arrF(1, 1) = srcRng.formula
            Else
                arrF = srcRng.formula
            End If
            
            nRows = UBound(arrF, 1)
            ReDim arrOut(1 To nRows, 1 To 1)

            For r = 1 To nRows
                If Len(CStr(arrF(r, 1))) > 1 Then
                    If Left$(CStr(arrF(r, 1)), 1) = "=" Then
                        arrOut(r, 1) = "'" & CStr(arrF(r, 1))
                    Else
                        arrOut(r, 1) = vbNullString
                    End If
                End If
            Next r

            ws.Range(ws.Cells(firstDataRow, formCol), ws.Cells(lLastRow, formCol)).Value = arrOut
        End If
NextWs:
    Next ws
End Sub

Option Explicit
'==========================
' Module: modHelpers
' Desc: 공용 헬퍼, 로깅, 환경 제어 유틸리티
'==========================

Public Sub LogAction(ByVal tag As String, ByVal msg As String)
    Dim wsLog As Worksheet, nr As Long
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "LOG"
        wsLog.Range("A1:D1").Value = Array("TIMESTAMP", "TAG", "MESSAGE", "USER")
        wsLog.Rows(1).Font.Bold = True
    End If
    On Error GoTo 0
    
    With wsLog
        nr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(nr, 1).Value = Now
        .Cells(nr, 2).Value = tag
        .Cells(nr, 3).Value = msg
        .Cells(nr, 4).Value = Environ$("Username")
    End With
End Sub

Public Sub LogActionSafe(ByVal tag As String, ByVal msg As String)
    On Error Resume Next
    LogAction tag, msg
    If Err.Number <> 0 Then
        Debug.Print Now, tag, msg
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub AppBegin(ByVal tag As String)
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Application.StatusBar = "Running: " & tag & " ..."
    On Error GoTo 0
End Sub

Public Sub AppEnd()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Public Function SafeFind(ByVal ws As Worksheet, ByVal whatText As String, _
                         ByVal exactMatch As Boolean, Optional ByVal inRange As Range) As Range
    Dim findRng As Range, lookAtMode As XlLookAt
    On Error Resume Next
    If inRange Is Nothing Then Set findRng = ws.UsedRange Else Set findRng = inRange
    On Error GoTo 0
    If findRng Is Nothing Then Exit Function
    
    lookAtMode = IIf(exactMatch, xlWhole, xlPart)
    Set SafeFind = findRng.Find(What:=whatText, After:=findRng.Cells(1, 1), LookIn:=xlValues, _
                                LookAt:=lookAtMode, SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, MatchCase:=False)
End Function

Public Function lastDataRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    If col <= 0 Then
        lastDataRow = 0
        Exit Function
    End If
    
    If Application.WorksheetFunction.CountA(ws.Columns(col)) = 0 Then
        lastDataRow = 0
    Else
        lastDataRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    End If
End Function

Public Sub ClearRange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    On Error Resume Next
    target.ClearContents
    On Error GoTo 0
End Sub

Public Function ProcExists(ByVal ProcName As String) As Boolean
    On Error Resume Next
    Application.Run ProcName
    Select Case Err.Number
        Case 0: ProcExists = True
        Case 1004, 438: ProcExists = False
        Case Else: ProcExists = True
    End Select
    Err.Clear
    On Error GoTo 0
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    If Err.Number <> 0 Then
        Err.Clear
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
    On Error GoTo 0
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    If IsError(v) Then
        Nz = defaultValue
    ElseIf IsNumeric(v) Then
        If Not IsEmpty(v) And Not IsNull(v) Then Nz = CDbl(v) Else Nz = defaultValue
    Else
        Nz = defaultValue
    End If
End Function

Option Explicit
'==========================
' Module: modPipeline (최종 재구성 버전)
'==========================

Public Sub START_PIPELINE()
    RunPipeline_Final
End Sub

Public Sub RunPipeline_Final(Optional ByVal ShowMsg As Boolean = True)
    Dim t0 As Single: t0 = Timer
    AppBegin "Final Pipeline"
    On Error GoTo ErrH
    LogActionSafe "PIPELINE", "START"

    ' 1단계: Formula 추출
    SafeRun "ExtractFormulasWithExclusion"
    
    ' 2단계: REV RATE 등 계산
    SafeRun "ApplyFormula_ByDynamicRemark_ExactTotal_Safe"
    
    ' 3단계: 최종 취합
    SafeRun "CompileAllSheets"

Done:
    LogActionSafe "PIPELINE", "END " & Format(Timer - t0, "0.00s")
    AppEnd
    If ShowMsg Then MsgBox "모든 파이프라인 작업 완료!", vbInformation, "Pipeline Complete"
    Exit Sub

ErrH:
    LogActionSafe "PIPELINE", "FATAL ERR: " & Err.description & " (" & Err.Number & ")"
    AppEnd
    If ShowMsg Then MsgBox "파이프라인 중단: " & vbCrLf & Err.description, vbCritical, "Pipeline Error"
End Sub

Private Sub SafeRun(ByVal ProcName As String)
    On Error GoTo ErrH
    
    If Not ProcExists(ProcName) Then
        Err.Raise 10001, , "프로시저 없음: " & ProcName
    End If

    Application.StatusBar = "Running: " & ProcName & " ..."
    Application.Run ProcName
    LogActionSafe ProcName, "OK"
    Exit Sub

ErrH:
    ' SafeRun에서 오류 발생 시, 상위 프로시저(RunPipeline_Final)의 오류 처리기로 넘김
    Err.Raise Err.Number, "Error in " & ProcName, Err.description
End Sub

Option Explicit
' Minimal buttons: Export_To_CSV, Run_Python_Audit, Refresh_All (user can paste real code provided earlier)
