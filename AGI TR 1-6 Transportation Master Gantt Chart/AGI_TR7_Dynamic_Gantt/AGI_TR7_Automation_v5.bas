Attribute VB_Name = "AGI_TR7_Automation_v4"
Option Explicit

' ============================================================
' AGI HVDC Transformers (TR1..TR7) – Dynamic Gantt Automation
' Version: v4 (2026-01-07)
'
' PURPOSE (ops-oriented; no PDF export)
' 1) Start date in Inputs!B5 drives the whole workbook (dates are formula-based).
'    This VBA adds convenience tools (refresh, gates, lookahead, QC).
' 2) Weather ↔ schedule idea: pull daily forecast (wind & wave) and flag NO-GO days.
' 3) Practical tools: target-date check, 14-day lookahead, task risk flagging.
'
' IMPORTANT
' - This module is delivered as .bas. Import into your workbook (Developer ▸ VB Editor ▸ File ▸ Import File…)
' - Macros require "Trust access" enabled and network access for Open-Meteo.
' ============================================================

' ----------------------------
' Sheet names
' ----------------------------
Private Const SH_INPUTS As String = "Inputs"
Private Const SH_WX As String = "Weather_Forecast"
Private Const SH_PLAN_A As String = "Plan_A_Realistic"
Private Const SH_PLAN_B As String = "Plan_B_Fast"

' ----------------------------
' Inputs addresses
' ----------------------------
Private Const ADR_LO_START As String = "B5"       ' LO Commencement (TR1 LO start @MZP)
Private Const ADR_TARGET As String = "B6"         ' Target complete by

' Weather config block (Inputs, cols D:E)
Private Const ROW_W_MZP_LAT As Long = 20
Private Const ROW_W_MZP_LON As Long = 21
Private Const ROW_W_AGI_LAT As Long = 22
Private Const ROW_W_AGI_LON As Long = 23
Private Const ROW_W_WIND_MAX As Long = 24
Private Const ROW_W_WAVE_MAX As Long = 25
Private Const ROW_W_TZ As Long = 26
Private Const ROW_W_HORIZON As Long = 27

' Plan sheet layout
Private Const PLAN_HEADER_ROW As Long = 5
Private Const PLAN_FIRST_ROW As Long = 7
Private Const COL_PHASE As Long = 4
Private Const COL_START As Long = 7
Private Const COL_FINISH As Long = 8
Private Const COL_NOTES As Long = 12

' ----------------------------
' Public entrypoints
' ----------------------------
Public Sub UpdateAll(Optional ByVal RefreshWeather As Boolean = False)
    ' One-click ops refresh:
    ' - Full recalc
    ' - Optional weather pull
    ' - Flag NO-GO start dates
    ' - Target date check
    ' - 14-day lookahead sheet
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Application.CalculateFull

    If RefreshWeather Then
        RefreshWeather_OpenMeteo
    End If

    FlagNoGoTasks SH_PLAN_A
    FlagNoGoTasks SH_PLAN_B
    CheckTargetDate SH_PLAN_A
    CheckTargetDate SH_PLAN_B
    CreateLookahead_14D

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "UpdateAll failed: " & Err.Description, vbExclamation
End Sub


Public Sub RefreshWeather_OpenMeteo()
    ' Pulls daily forecast (next N days) for wind & wave:
    ' - Wind: open-meteo forecast API (daily winds..._10m_max)
    ' - Wave: open-meteo marine API (daily wave_height_max)
    '
    ' Output: Weather_Forecast sheet
    ' Columns:
    '   A Date
    '   B MZP wind max (kn)
    '   C MZP wave max (m)
    '   D AGI wind max (kn)
    '   E AGI wave max (m)
    '   F Gate (GO/NO-GO)
    '   G Notes
    '
    ' NOTE: This is operational convenience. Always validate with the latest approved MS / Port Control.

    On Error GoTo EH

    Dim wsI As Worksheet, wsW As Worksheet
    Set wsI = ThisWorkbook.Worksheets(SH_INPUTS)
    Set wsW = ThisWorkbook.Worksheets(SH_WX)

    Dim mzpLat As Double, mzpLon As Double, agiLat As Double, agiLon As Double
    Dim windMax As Double, waveMax As Double
    Dim tz As String, horizon As Long

    mzpLat = CDbl(wsI.Cells(ROW_W_MZP_LAT, "E").Value)
    mzpLon = CDbl(wsI.Cells(ROW_W_MZP_LON, "E").Value)
    agiLat = CDbl(wsI.Cells(ROW_W_AGI_LAT, "E").Value)
    agiLon = CDbl(wsI.Cells(ROW_W_AGI_LON, "E").Value)
    windMax = CDbl(wsI.Cells(ROW_W_WIND_MAX, "E").Value)
    waveMax = CDbl(wsI.Cells(ROW_W_WAVE_MAX, "E").Value)
    tz = CStr(wsI.Cells(ROW_W_TZ, "E").Value)
    horizon = CLng(wsI.Cells(ROW_W_HORIZON, "E").Value)
    If horizon < 3 Then horizon = 16

    Dim dates() As String
    Dim mzpWind() As Double, agiWind() As Double
    Dim mzpWave() As Double, agiWave() As Double

    ' Wind (knots) – parse CSV
    Call FetchDailyCsvWindKn(mzpLat, mzpLon, tz, horizon, dates, mzpWind)
    Call FetchDailyCsvWindKn(agiLat, agiLon, tz, horizon, dates, agiWind)

    ' Wave (m) – parse CSV
    Call FetchDailyCsvWaveM(mzpLat, mzpLon, tz, horizon, dates, mzpWave)
    Call FetchDailyCsvWaveM(agiLat, agiLon, tz, horizon, dates, agiWave)

    ' Write sheet
    Dim r As Long, i As Long
    ' Validate arrays (avoid LBound/UBound crash if CSV parsing returned 0 rows)
    Dim lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(dates): ub = UBound(dates)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo EH
        MsgBox "Weather pull returned no usable rows (dates array not initialized)." & vbCrLf & _
               "Check: Internet access / Open-Meteo reachable / Inputs timezone (IANA, e.g., Asia/Dubai) / variable names.", vbExclamation
        Exit Sub
    End If
    On Error GoTo EH

    ' Validate the other series arrays and align length
    Dim ubW1 As Long, ubW2 As Long, ubS1 As Long, ubS2 As Long
    On Error Resume Next
    ubW1 = UBound(mzpWind)
    ubW2 = UBound(agiWind)
    ubS1 = UBound(mzpWave)
    ubS2 = UBound(agiWave)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo EH
        MsgBox "Weather pull succeeded but one or more series could not be parsed (wind/wave array not initialized)." & vbCrLf & _
               "Check: Open-Meteo CSV format, variable names, or corporate proxy blocking marine-api.", vbExclamation
        Exit Sub
    End If
    On Error GoTo EH

    ' Align loop length to shortest series to prevent out-of-range
    If ubW1 < ub Then ub = ubW1
    If ubW2 < ub Then ub = ubW2
    If ubS1 < ub Then ub = ubS1
    If ubS2 < ub Then ub = ubS2

    wsW.Range("A2:G1000").ClearContents
    For i = lb To ub
        r = 2 + (i - lb)
        wsW.Cells(r, 1).Value = CDate(dates(i))
        wsW.Cells(r, 1).NumberFormat = "dd-mmm-yy"
        wsW.Cells(r, 2).Value = Round(mzpWind(i), 2)
        wsW.Cells(r, 3).Value = Round(mzpWave(i), 2)
        wsW.Cells(r, 4).Value = Round(agiWind(i), 2)
        wsW.Cells(r, 5).Value = Round(agiWave(i), 2)

        Dim gate As String
        If (mzpWind(i) <= windMax And agiWind(i) <= windMax And mzpWave(i) <= waveMax And agiWave(i) <= waveMax) Then
            gate = "GO"
        Else
            gate = "NO-GO"
        End If
        wsW.Cells(r, 6).Value = gate
        wsW.Cells(r, 7).Value = "wind<= " & windMax & "kn; wave<= " & waveMax & "m"
    Next i

    Exit Sub
EH:
    MsgBox "RefreshWeather_OpenMeteo failed: " & Err.Description, vbExclamation
End Sub


Public Sub CreateLookahead_14D()
    ' Builds a lightweight 14-day lookahead list (both plans) based on Start/Finish windows.
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Lookahead_14D")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.Name = "Lookahead_14D"
    End If

    wsOut.Cells.Clear
    wsOut.Range("A1:F1").Value = Array("Plan", "ID", "Task", "Phase", "Start", "Finish")
    wsOut.Range("A1:F1").Font.Bold = True

    Dim d0 As Date, d1 As Date
    d0 = Date
    d1 = DateAdd("d", 13, d0)

    Dim outR As Long
    outR = 2

    outR = AppendLookahead(wsOut, outR, SH_PLAN_A, d0, d1, "Plan A")
    outR = AppendLookahead(wsOut, outR, SH_PLAN_B, d0, d1, "Plan B")

    wsOut.Columns("A:F").AutoFit
End Sub


' ----------------------------
' QC / Helpers
' ----------------------------
Private Sub CheckTargetDate(ByVal planSheet As String)
    ' Flags whether plan finishes after Inputs!B6.
    Dim wsI As Worksheet, wsP As Worksheet
    Set wsI = ThisWorkbook.Worksheets(SH_INPUTS)
    Set wsP = ThisWorkbook.Worksheets(planSheet)

    Dim target As Date
    target = CDate(wsI.Range(ADR_TARGET).Value)

    Dim lastFinish As Date
    lastFinish = GetMaxFinish(wsP)

    Dim msg As String
    If lastFinish > target Then
        msg = planSheet & ": finish " & Format(lastFinish, "dd-mmm-yy") & " (AFTER target " & Format(target, "dd-mmm-yy") & ")"
        wsP.Range("A2").Value = msg
    Else
        msg = planSheet & ": finish " & Format(lastFinish, "dd-mmm-yy") & " (ON/B4 target " & Format(target, "dd-mmm-yy") & ")"
        wsP.Range("A2").Value = msg
    End If
End Sub


Private Sub FlagNoGoTasks(ByVal planSheet As String)
    ' For each task row, if Start date has Weather_Forecast gate = NO-GO, add note flag.
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets(planSheet)

    Dim wsW As Worksheet
    Set wsW = ThisWorkbook.Worksheets(SH_WX)

    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    If lastRow < PLAN_FIRST_ROW Then Exit Sub

    Dim r As Long
    For r = PLAN_FIRST_ROW To lastRow
        Dim ph As String
        ph = CStr(wsP.Cells(r, COL_PHASE).Value)
        If ph = "SUMMARY" Or ph = "" Then GoTo NextR

        Dim d As Variant
        d = wsP.Cells(r, COL_START).Value
        If Not IsDate(d) Then GoTo NextR

        Dim gate As String
        gate = GetGateForDate(wsW, CDate(d))
        If gate = "NO-GO" Then
            If InStr(1, CStr(wsP.Cells(r, COL_NOTES).Value), "WX:NO-GO") = 0 Then
                wsP.Cells(r, COL_NOTES).Value = Trim(CStr(wsP.Cells(r, COL_NOTES).Value) & " | WX:NO-GO")
            End If
        Else
            ' Optional: remove NO-GO tag if now GO
        End If
NextR:
    Next r
End Sub


Private Function GetGateForDate(ByVal wsW As Worksheet, ByVal d As Date) As String
    Dim rng As Range
    Dim f As Range
    Set rng = wsW.Range("A:A")
    Set f = rng.Find(What:=d, LookAt:=xlWhole)
    If f Is Nothing Then
        GetGateForDate = ""
    Else
        GetGateForDate = CStr(wsW.Cells(f.Row, 6).Value)
    End If
End Function


Private Function AppendLookahead(ByVal wsOut As Worksheet, ByVal outR As Long, ByVal planSheet As String, ByVal d0 As Date, ByVal d1 As Date, ByVal label As String) As Long
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Worksheets(planSheet)

    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = PLAN_FIRST_ROW To lastRow
        Dim st As Variant, fn As Variant
        st = wsP.Cells(r, COL_START).Value
        fn = wsP.Cells(r, COL_FINISH).Value
        If IsDate(st) And IsDate(fn) Then
            If CDate(fn) >= d0 And CDate(st) <= d1 Then
                wsOut.Cells(outR, 1).Value = label
                wsOut.Cells(outR, 2).Value = wsP.Cells(r, 1).Value
                wsOut.Cells(outR, 3).Value = wsP.Cells(r, 3).Value
                wsOut.Cells(outR, 4).Value = wsP.Cells(r, COL_PHASE).Value
                wsOut.Cells(outR, 5).Value = CDate(st)
                wsOut.Cells(outR, 6).Value = CDate(fn)
                wsOut.Cells(outR, 5).NumberFormat = "dd-mmm-yy"
                wsOut.Cells(outR, 6).NumberFormat = "dd-mmm-yy"
                outR = outR + 1
            End If
        End If
    Next r
    AppendLookahead = outR
End Function


Private Function GetMaxFinish(ByVal wsP As Worksheet) As Date
    Dim lastRow As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, m As Date
    m = DateSerial(2000, 1, 1)
    For r = PLAN_FIRST_ROW To lastRow
        If IsDate(wsP.Cells(r, COL_FINISH).Value) Then
            If CDate(wsP.Cells(r, COL_FINISH).Value) > m Then m = CDate(wsP.Cells(r, COL_FINISH).Value)
        End If
    Next r
    GetMaxFinish = m
End Function


' ----------------------------
' Open-Meteo CSV fetch + parse
' ----------------------------
Private Sub FetchDailyCsvWindKn(ByVal lat As Double, ByVal lon As Double, ByVal tz As String, ByVal horizon As Long, ByRef dates() As String, ByRef values() As Double)
    ' Forecast endpoint (CSV)
    Dim url As String
    url = "https://api.open-meteo.com/v1/forecast" & _
          "?latitude=" & CStr(lat) & "&longitude=" & CStr(lon) & _
          "&daily=wind_speed_10m_max" & _
          "&wind_speed_unit=kn" & _
          "&timezone=" & UrlEncode(tz) & _
          "&forecast_days=" & CStr(horizon) & _
          "&format=csv"
    Dim csv As String
    csv = HttpGet(url)
    ParseDailyCsv csv, "wind_speed_10m_max", dates, values
    ' Convert m/s -> knots if Open-Meteo returns m/s (depends on default units).
    ' Open-Meteo default is km/h for windspeed; to avoid ambiguity, request "wind_speed_unit=kn".
    ' If you want strict knots, change URL to add: &wind_speed_unit=kn
End Sub


Private Sub FetchDailyCsvWaveM(ByVal lat As Double, ByVal lon As Double, ByVal tz As String, ByVal horizon As Long, ByRef dates() As String, ByRef values() As Double)
    Dim url As String
    url = "https://marine-api.open-meteo.com/v1/marine" & _
          "?latitude=" & CStr(lat) & "&longitude=" & CStr(lon) & _
          "&daily=wave_height_max" & _
          "&timezone=" & UrlEncode(tz) & _
          "&forecast_days=" & CStr(horizon) & _
          "&format=csv"
    Dim csv As String
    csv = HttpGet(url)
    ParseDailyCsv csv, "wave_height_max", dates, values
End Sub


Private Sub ParseDailyCsv(ByVal csv As String, ByVal seriesName As String, ByRef dates() As String, ByRef values() As Double)
    ' Robust CSV parser for Open-Meteo "format=csv" responses.
    ' Handles the common case where metadata lines come before the actual data table.
    ' Expected data section header starts with: time,<seriesName>

    On Error GoTo EH

    Erase dates
    Erase values

    Dim txt As String
    txt = Replace(csv, vbCr, "")

    Dim lines() As String
    lines = Split(txt, vbLf)
    If UBound(lines) < 1 Then Exit Sub

    Dim hdr() As String
    Dim headerIdx As Long
    headerIdx = -1

    Dim i As Long, ln As String
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextHeader
        If Left$(ln, 1) = "#" Then GoTo NextHeader

        hdr = Split(ln, ",")
        If UBound(hdr) >= 1 Then
            If LCase$(Trim$(hdr(0))) = "time" Or LCase$(Trim$(hdr(0))) = "date" Then
                headerIdx = i
                Exit For
            End If
        End If
NextHeader:
    Next i

    If headerIdx < 0 Then Exit Sub

    Dim idx As Long
    idx = FindIndex(hdr, seriesName)
    If idx < 0 Then
        ' Try alias without underscores (fallback compatibility)
        idx = FindIndex(hdr, Replace(seriesName, "_", ""))
        If idx < 0 Then Exit Sub
    End If

    ' Count data rows after header (only rows that look like ISO date YYYY-MM-DD)
    Dim n As Long
    n = 0
    For i = headerIdx + 1 To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextCount
        If Left$(ln, 1) = "#" Then GoTo NextCount

        Dim parts() As String
        parts = Split(ln, ",")
        If UBound(parts) >= 0 Then
            If LooksLikeISODate(parts(0)) Then n = n + 1
        End If
NextCount:
    Next i

    If n = 0 Then Exit Sub

    ReDim dates(0 To n - 1)
    ReDim values(0 To n - 1)

    Dim r As Long
    r = 0
    For i = headerIdx + 1 To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextRow
        If Left$(ln, 1) = "#" Then GoTo NextRow

        Dim parts2() As String
        parts2 = Split(ln, ",")
        If UBound(parts2) >= 0 Then
            If LooksLikeISODate(parts2(0)) Then
                dates(r) = parts2(0)
                If idx <= UBound(parts2) Then
                    values(r) = Val(parts2(idx))
                Else
                    values(r) = 0
                End If
                r = r + 1
                If r >= n Then Exit For
            End If
        End If
NextRow:
    Next i

    Exit Sub

EH:
    ' Leave arrays erased on failure
End Sub


Private Function LooksLikeISODate(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    LooksLikeISODate = (Len(t) >= 10 And Mid$(t, 5, 1) = "-" And Mid$(t, 8, 1) = "-")
End Function



Private Function FindIndex(ByRef arr() As String, ByVal key As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If LCase(Trim(arr(i))) = LCase(Trim(key)) Then
            FindIndex = i
            Exit Function
        End If
    Next i
    FindIndex = -1
End Function


Private Function HttpGet(ByVal url As String) As String
    Dim xhr As Object
    Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
    xhr.Open "GET", url, False
    xhr.SetRequestHeader "User-Agent", "AGI_TR7_Gantt/1.0"
    xhr.Send
    If xhr.Status <> 200 Then
        Err.Raise vbObjectError + 101, , "HTTP " & xhr.Status & " for " & url
    End If
    HttpGet = CStr(xhr.ResponseText)
End Function


Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122
                out = out & ch
            Case 45, 46, 95
                out = out & ch
            Case Else
                out = out & "%" & Right$("0" & Hex(AscW(ch)), 2)
        End Select
    Next i
    UrlEncode = out
End Function
