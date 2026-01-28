Attribute VB_Name = "modAGI_TR_Suite"
Option Explicit

'===============================================================================
' AGI HVDC Transformer Transportation — Master Planner VBA Suite
'-------------------------------------------------------------------------------
' Scope (AGI Site):
'   - Scenario/Custom schedule generation (voyages + install batches)
'   - Tide/Weather risk assessment + auto-hold insertion (optional)
'   - Gantt build (timeline header + bars + critical highlight)
'   - Monte Carlo (P50/P80) finish forecasting
'   - D0 optimizer for deadline probability
'   - Baseline freeze + variance/change log
'   - Reports (executive summary, milestones, daily/weekly lookahead)
'   - Exports (CSV + PDF)
'
' IMPORTANT:
'   This module is delivered as .BAS for import.
'   1) Save workbook as .XLSM
'   2) Import this module
'   3) Run SetupWorkbook once (creates buttons/validations/named ranges)
'
' Version: 3.1.1
' Date: 2026-01-07 (patched v3.1.1)
'===============================================================================

'========================
' Sheets
'========================
Private Const SH_CTRL As String = "Control_Panel"
Private Const SH_SCEN As String = "Scenario_Library"
Private Const SH_PAT As String = "Pattern_Tasks"
Private Const SH_DATA As String = "Schedule_Data"
Private Const SH_GANTT As String = "Gantt_Chart"
Private Const SH_TIDE As String = "Tide_Data"
Private Const SH_WEATHER As String = "Weather_Risk"
Private Const SH_DASH As String = "Dashboard"
Private Const SH_DOCS As String = "Docs_Checklist"
Private Const SH_EVID As String = "Evidence_Checklist"
Private Const SH_BASE As String = "Baseline"
Private Const SH_CHG As String = "Change_Log"
Private Const SH_REP As String = "Reports"
Private Const SH_LOG As String = "Logs"
Private Const SH_EXP As String = "Exports"

'========================
' Control_Panel cells (template v3)
'========================
Private Const C_D0 As String = "C5"
Private Const C_SCEN As String = "C6"
Private Const C_TRIPPLAN As String = "C7"
Private Const C_BATCHES As String = "C8"
Private Const C_JACKS As String = "C9"
Private Const C_WBUF As String = "C10"
Private Const C_THOLD As String = "C11"
Private Const C_TARGET As String = "C12"
Private Const C_MC_RUNS As String = "C13"
Private Const C_MC_CONF As String = "C14"
Private Const C_DEADLINE As String = "C15"

' Outputs (template v3)
Private Const O_FINISH As String = "C18"
Private Const O_P50 As String = "C19"
Private Const O_P80 As String = "C20"
Private Const O_MEET As String = "C21"
Private Const O_CONFLICTS As String = "C22"
Private Const O_CPLEN As String = "C23"
Private Const O_NOTES As String = "C24"

'========================
' Schedule_Data columns (Row 5 header)
'========================
Private Const HDR_ROW As Long = 5
Private Const ROW0 As Long = 6

Private Const COL_ID As Long = 1
Private Const COL_WBS As Long = 2
Private Const COL_TASK As Long = 3
Private Const COL_PHASE As Long = 4
Private Const COL_OWNER As Long = 5
Private Const COL_OFFSET As Long = 6
Private Const COL_START As Long = 7
Private Const COL_END As Long = 8
Private Const COL_DUR As Long = 9
Private Const COL_TRLIST As Long = 10
Private Const COL_VOY As Long = 11
Private Const COL_BATCH As Long = 12
Private Const COL_TIDERISK As Long = 13
Private Const COL_WEATHERRISK As Long = 14
Private Const COL_CRIT As Long = 15
Private Const COL_BS As Long = 16
Private Const COL_BE As Long = 17
Private Const COL_AS As Long = 18
Private Const COL_AE As Long = 19
Private Const COL_PCT As Long = 20
Private Const COL_STATUS As Long = 21
Private Const COL_NOTES As Long = 22

'========================
' Internal task record
'========================
Private Type TaskRec
    ID As String
    WBS As String
    Task As String
    Phase As String
    Owner As String
    Offset As Double
    Duration As Double
    TR_List As String
    Voyage As Long
    Batch As Long
    TideSensitive As Boolean
    WeatherSensitive As Boolean
    KeyTag As String
    OrderInGroup As Long
    GroupKey As String
End Type

Private mTasks() As TaskRec
Private mCount As Long

'===============================================================================
' Public UI Entrypoints
'===============================================================================

Public Sub SetupWorkbook()
    ' One-time initializer: ensures sheets/headers exist, applies validations,
    ' creates action buttons, sets keyboard shortcuts.
    On Error GoTo ErrH
    Application.ScreenUpdating = False
    EnsureAllSheetsExist
    EnsureHeaders
    ApplyValidations
    CreateButtons
    SetupKeyboardShortcuts
    LogMsg "INFO", "SetupWorkbook", "Workbook setup complete."
    MsgBox "Setup complete. Save as .XLSM and use Ctrl+Shift+U to Run All.", vbInformation
Clean:
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    LogMsg "ERROR", "SetupWorkbook", Err.Description
    Application.ScreenUpdating = True
    MsgBox "SetupWorkbook failed: " & Err.Description, vbCritical
End Sub

Public Sub SelfTest()
    ' Lightweight sanity checks (no external dependencies).
    ' Writes results to Logs sheet and displays a summary.
    On Error GoTo ErrH

    Dim missing As String, ok As Long, fail As Long
    missing = ""

    ' Required sheets
    CheckSheetExists "Control_Panel", missing
    CheckSheetExists "Scenario_Library", missing
    CheckSheetExists "Pattern_Tasks", missing
    CheckSheetExists "Schedule_Data", missing
    CheckSheetExists "Gantt_Chart", missing
    CheckSheetExists "Tide_Data", missing
    CheckSheetExists "Weather_Risk", missing
    CheckSheetExists "Logs", missing
    CheckSheetExists "Exports", missing

    If Len(missing) > 0 Then
        LogMsg "ERROR", "SelfTest", "Missing sheets: " & missing
        MsgBox "SelfTest FAILED. Missing sheets: " & missing, vbCritical
        Exit Sub
    End If

    ' Required headers (Pattern_Tasks)
    Dim ws As Worksheet
    Set ws = Sheets("Pattern_Tasks")
    If UCase$(CStr(ws.Cells(1, 1).Value)) <> "TEMPLATE" Then fail = fail + 1 Else ok = ok + 1
    If UCase$(CStr(ws.Cells(1, 8).Value)) <> "DURATION_DAYS" Then fail = fail + 1 Else ok = ok + 1
    If UCase$(CStr(ws.Cells(1, 9).Value)) <> "TIDESENSITIVE" Then fail = fail + 1 Else ok = ok + 1
    If UCase$(CStr(ws.Cells(1, 10).Value)) <> "WEATHERSENSITIVE" Then fail = fail + 1 Else ok = ok + 1

    ' Control panel essentials
    Dim scen As String
    scen = CStr(Sheets("Control_Panel").Range("C6").Value)
    If Len(Trim$(scen)) = 0 Then fail = fail + 1 Else ok = ok + 1

    LogMsg "INFO", "SelfTest", "OK=" & ok & " | FAIL=" & fail
    If fail = 0 Then
        MsgBox "SelfTest PASSED. Checks: " & ok, vbInformation
    Else
        MsgBox "SelfTest WARNING. OK=" & ok & " FAIL=" & fail & " (see Logs sheet).", vbExclamation
    End If
    Exit Sub

ErrH:
    LogMsg "ERROR", "SelfTest", Err.Description
    MsgBox "SelfTest failed: " & Err.Description, vbCritical
End Sub

Private Sub CheckSheetExists(ByVal shName As String, ByRef missing As String)
    On Error GoTo NotFound
    Dim ws As Worksheet
    Set ws = Sheets(shName)
    Exit Sub
NotFound:
    If Len(missing) > 0 Then missing = missing & ", "
    missing = missing & shName
End Sub



Public Sub RunAll(Optional ByVal Silent As Boolean = False)
    ' Master routine:
    '   1) Read inputs + validate
    '   2) Build schedule (voyages + install batches)
    '   3) Write Schedule_Data
    '   4) Compute dates
    '   5) Risk assess + auto-holds (if enabled by nonzero hold settings)
    '   6) Recompute dates (post-hold)
    '   7) Critical path mark + Gantt rebuild + Dashboard + Reports
    '   8) Optional Monte Carlo summary
    Dim d0 As Date, scen As String
    Dim runs As Long
    Dim conf As Double
    Dim msg As String

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    d0 = ReadDate(SH_CTRL, C_D0, Date)
    scen = CStr(Sheets(SH_CTRL).Range(C_SCEN).Value)

    ValidateInputs msg
    If Len(msg) > 0 Then
        LogMsg "WARN", "RunAll", msg
        If Not Silent Then MsgBox msg, vbExclamation, "Validation Notes"
    End If

    BuildScheduleFromControl d0, scen
    WriteSchedule
    UpdateDates d0

    ' Risk evaluation + auto-holds (only if hold days > 0)
    EvaluateRisks
    If ReadLong(SH_CTRL, C_THOLD, 0) > 0 Or ReadLong(SH_CTRL, C_WBUF, 0) > 0 Then
        ApplyAutoHolds d0
        UpdateDates d0
        EvaluateRisks
    End If

    MarkCriticalPath
    RefreshGantt
    UpdateDashboardAndOutputs d0

    ' Monte Carlo
    runs = ReadLong(SH_CTRL, C_MC_RUNS, 0)
    conf = ReadDouble(SH_CTRL, C_MC_CONF, 0.8)
    If runs > 0 Then
        MonteCarloFinish d0, runs, conf
    End If

    BuildReports d0

    If Not Silent Then
        MsgBox "Run All complete." & vbCrLf & _
               "Finish: " & Sheets(SH_CTRL).Range(O_FINISH).Value & vbCrLf & _
               "P50/P80: " & Sheets(SH_CTRL).Range(O_P50).Value & " / " & Sheets(SH_CTRL).Range(O_P80).Value, vbInformation
    End If

Clean:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrH:
    LogMsg "ERROR", "RunAll", Err.Description
    If Not Silent Then MsgBox "RunAll failed: " & Err.Description, vbCritical
    Resume Clean
End Sub


Public Sub OptimizeD0()
    ' Finds earliest D0 meeting deadline at requested confidence (Monte Carlo)
    ' Searches around current D0 in a bounded range.
    Dim baseD0 As Date, deadline As Date, conf As Double
    Dim runs As Long
    Dim d As Date, best As Date
    Dim ok As Boolean
    Dim searchStart As Date, searchEnd As Date

    On Error GoTo ErrH
    baseD0 = ReadDate(SH_CTRL, C_D0, Date)
    deadline = ReadDate(SH_CTRL, C_DEADLINE, DateSerial(2026, 3, 1))
    conf = ReadDouble(SH_CTRL, C_MC_CONF, 0.8)
    runs = ReadLong(SH_CTRL, C_MC_RUNS, 500)

    searchStart = baseD0 - 14
    searchEnd = baseD0 + 30

    best = DateSerial(2099, 12, 31)
    For d = searchStart To searchEnd
        Sheets(SH_CTRL).Range(C_D0).Value = d
        RunAll True
        ok = (CDate(Sheets(SH_CTRL).Range(O_P80).Value) <= deadline)
        If ok Then
            best = d
            Exit For
        End If
    Next d

    If best < DateSerial(2099, 12, 31) Then
        Sheets(SH_CTRL).Range(C_D0).Value = best
        RunAll True
        MsgBox "Optimal D0 found: " & Format(best, "yyyy-mm-dd") & vbCrLf & _
               "P80 Finish: " & Format(CDate(Sheets(SH_CTRL).Range(O_P80).Value), "yyyy-mm-dd"), vbInformation
    Else
        MsgBox "No D0 in search window met the deadline at the selected confidence." & vbCrLf & _
               "Consider increasing resources, reducing buffers, or extending the deadline.", vbExclamation
    End If

    Exit Sub
ErrH:
    LogMsg "ERROR", "OptimizeD0", Err.Description
    MsgBox "OptimizeD0 failed: " & Err.Description, vbCritical
End Sub


Public Sub FreezeBaseline()
    ' Copies current Start/End into baseline columns and Baseline sheet snapshot.
    Dim ws As Worksheet, wsB As Worksheet
    Dim lastR As Long, r As Long
    On Error GoTo ErrH

    Set ws = Sheets(SH_DATA)
    Set wsB = Sheets(SH_BASE)

    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    ' Write baseline columns
    For r = ROW0 To lastR
        ws.Cells(r, COL_BS).Value = ws.Cells(r, COL_START).Value
        ws.Cells(r, COL_BE).Value = ws.Cells(r, COL_END).Value
        ws.Cells(r, COL_BS).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, COL_BE).NumberFormat = "yyyy-mm-dd"
    Next r

    ' Snapshot to Baseline sheet
    wsB.Cells.Clear
    wsB.Range("A1").Value = "Baseline Snapshot (" & Format(Now, "yyyy-mm-dd hh:nn:ss") & ")"
    wsB.Range("A1").Font.Bold = True
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(lastR, COL_NOTES)).Copy wsB.Range("A3")

    LogChange "FreezeBaseline", "Baseline frozen for " & (lastR - ROW0 + 1) & " tasks."
    MsgBox "Baseline frozen.", vbInformation
    Exit Sub

ErrH:
    LogMsg "ERROR", "FreezeBaseline", Err.Description
    MsgBox "FreezeBaseline failed: " & Err.Description, vbCritical
End Sub


Public Sub CompareToBaseline()
    ' Writes variances into Change_Log and flags schedule deltas.
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim id As String
    Dim bs As Variant, be As Variant, s As Variant, e As Variant
    Dim deltaS As Long, deltaE As Long
    Dim any As Boolean

    On Error GoTo ErrH
    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    any = False
    For r = ROW0 To lastR
        id = CStr(ws.Cells(r, COL_ID).Value)
        bs = ws.Cells(r, COL_BS).Value
        be = ws.Cells(r, COL_BE).Value
        s = ws.Cells(r, COL_START).Value
        e = ws.Cells(r, COL_END).Value

        If IsDate(bs) And IsDate(s) Then
            deltaS = DateDiff("d", CDate(bs), CDate(s))
        Else
            deltaS = 0
        End If

        If IsDate(be) And IsDate(e) Then
            deltaE = DateDiff("d", CDate(be), CDate(e))
        Else
            deltaE = 0
        End If

        If deltaS <> 0 Or deltaE <> 0 Then
            any = True
            LogChange "Variance", "Task " & id & " baseline delta: Start " & deltaS & "d, End " & deltaE & "d"
        End If
    Next r

    If any Then
        MsgBox "Baseline variances recorded in Change_Log.", vbInformation
    Else
        MsgBox "No variances detected vs baseline.", vbInformation
    End If
    Exit Sub

ErrH:
    LogMsg "ERROR", "CompareToBaseline", Err.Description
    MsgBox "CompareToBaseline failed: " & Err.Description, vbCritical
End Sub


Public Sub ExportPDFandCSV()
    ' Exports:
    '   - Schedule_Data to CSV
    '   - Gantt_Chart to PDF
    On Error GoTo ErrH
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook as .XLSM first (File → Save As) before exporting.", vbExclamation
        Exit Sub
    End If


    ExportScheduleCSV
    ExportGanttPDF

    MsgBox "Exports complete. See 'Exports' sheet for file paths.", vbInformation
    Exit Sub
ErrH:
    LogMsg "ERROR", "ExportPDFandCSV", Err.Description
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub


Public Sub DailyBriefing()
    ' Displays today's active tasks + next 7 days starting tasks
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim today As Date, nextW As Date
    Dim s As Date, e As Date
    Dim msgA As String, msgU As String

    On Error GoTo ErrH
    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    today = Date
    nextW = today + 7

    msgA = "ACTIVE TODAY (" & Format(today, "yyyy-mm-dd ddd") & ")" & vbCrLf & String(60, "-") & vbCrLf
    msgU = vbCrLf & "STARTING NEXT 7 DAYS" & vbCrLf & String(60, "-") & vbCrLf

    For r = ROW0 To lastR
        If IsDate(ws.Cells(r, COL_START).Value) Then
            s = CDate(ws.Cells(r, COL_START).Value)
            e = IIf(IsDate(ws.Cells(r, COL_END).Value), CDate(ws.Cells(r, COL_END).Value), s)
            If s <= today And today <= e Then
                msgA = msgA & "• " & ws.Cells(r, COL_ID).Value & " | " & ws.Cells(r, COL_TASK).Value & vbCrLf
            End If
            If s > today And s <= nextW Then
                msgU = msgU & "• " & Format(s, "mm/dd") & " | " & ws.Cells(r, COL_ID).Value & " | " & ws.Cells(r, COL_TASK).Value & vbCrLf
            End If
        End If
    Next r

    MsgBox msgA & msgU, vbInformation, "AGI TR Daily Briefing"
    Exit Sub
ErrH:
    LogMsg "ERROR", "DailyBriefing", Err.Description
    MsgBox "DailyBriefing failed: " & Err.Description, vbCritical
End Sub

'===============================================================================
' Keyboard Shortcuts
'===============================================================================

Public Sub Auto_Open()
    ' Enables shortcuts automatically when workbook opens.
    SetupKeyboardShortcuts
End Sub

Public Sub Auto_Close()
    ' Clears shortcuts when workbook closes.
    ClearKeyboardShortcuts
End Sub


Public Sub SetupKeyboardShortcuts()
    On Error Resume Next
    Application.OnKey "^+U", "modAGI_TR_Suite.RunAll"
    Application.OnKey "^+O", "modAGI_TR_Suite.OptimizeD0"
    Application.OnKey "^+M", "modAGI_TR_Suite.RunMonteCarloOnly"
    Application.OnKey "^+R", "modAGI_TR_Suite.ExportPDFandCSV"
    Application.OnKey "^+B", "modAGI_TR_Suite.DailyBriefing"
    Application.OnKey "^+S", "modAGI_TR_Suite.FreezeBaseline"
    Application.OnKey "^+D", "modAGI_TR_Suite.CompareToBaseline"
End Sub

Public Sub ClearKeyboardShortcuts()
    On Error Resume Next
    Application.OnKey "^+U"
    Application.OnKey "^+O"
    Application.OnKey "^+M"
    Application.OnKey "^+R"
    Application.OnKey "^+B"
    Application.OnKey "^+S"
    Application.OnKey "^+D"
End Sub

Public Sub RunMonteCarloOnly()
    Dim d0 As Date, runs As Long, conf As Double
    d0 = ReadDate(SH_CTRL, C_D0, Date)
    runs = ReadLong(SH_CTRL, C_MC_RUNS, 500)
    conf = ReadDouble(SH_CTRL, C_MC_CONF, 0.8)
    MonteCarloFinish d0, runs, conf
    MsgBox "Monte Carlo updated: P50/P80 outputs refreshed.", vbInformation
End Sub

'===============================================================================
' Core Build / Write / Dates
'===============================================================================

Private Sub BuildScheduleFromControl(ByVal d0 As Date, ByVal scen As String)
    ' Reads scenario settings and builds tasks list in memory.
    Dim tripPlan As String, batches As String
    Dim jacks As Long, wBuf As Long, tideHold As Long
    Dim sTrip As String, sBatch As String

    tripPlan = CStr(Sheets(SH_CTRL).Range(C_TRIPPLAN).Value)
    batches = CStr(Sheets(SH_CTRL).Range(C_BATCHES).Value)
    jacks = ReadLong(SH_CTRL, C_JACKS, 3)
    wBuf = ReadLong(SH_CTRL, C_WBUF, 1)
    tideHold = ReadLong(SH_CTRL, C_THOLD, 1)

    ' If scenario is not custom, load defaults from Scenario_Library (and override CP cells if blank)
    If Left$(scen, 2) = "S1" Or Left$(scen, 2) = "S2" Then
        If GetScenarioDefaults(scen, sTrip, sBatch, jacks, wBuf, tideHold) Then
            If Len(Trim$(tripPlan)) = 0 Then tripPlan = sTrip
            If Len(Trim$(batches)) = 0 Then batches = sBatch
            Sheets(SH_CTRL).Range(C_TRIPPLAN).Value = tripPlan
            Sheets(SH_CTRL).Range(C_BATCHES).Value = batches
            Sheets(SH_CTRL).Range(C_JACKS).Value = jacks
            Sheets(SH_CTRL).Range(C_WBUF).Value = wBuf
            Sheets(SH_CTRL).Range(C_THOLD).Value = tideHold
        End If
    End If

    BuildSchedule tripPlan, batches, jacks, wBuf, tideHold
End Sub


Private Sub BuildSchedule(ByVal tripPlan As String, ByVal batchPlan As String, _
                          ByVal parallelJacks As Long, ByVal weatherBuffer As Long, ByVal tideHold As Long)
    ' Constructs schedule for:
    '   - One-time mobilization
    '   - Voyages (1TR / 2TR using Pattern_Tasks templates)
    '   - Install batches (parallelized by jacks)
    ' Two-track scheduling:
    '   - Voyage track (LCT_Bushra): sequential
    '   - Install track (3대 잭다운 lines): can overlap voyages

    Dim trips() As Long, batches() As Long
    Dim i As Long, v As Long, b As Long
    Dim trCounter As Long
    Dim curVoy As Double, curIns As Double
    Dim deliveredOffset As Double
    Dim cumDelivered As Long
    Dim batchIdx As Long, batchSize As Long

    Erase mTasks
    mCount = 0

    trips = ParseLongList(tripPlan)
    batches = ParseLongList(batchPlan)

    ' One-time: Mobilization + deck prep (planning-level)
    curVoy = 0
    AddTask "MOB-001", "1.0", "Mobilization (SPMT/Marine/Steelworks) + Function Test", "MOBILIZATION", "Mammoet", curVoy, 1, "", 0, 0, False, False, "", 1, "MOB"
    curVoy = curVoy + 1
    AddTask "PREP-001", "1.1", "Deck Preparations (D-ring, markings, steel sets, welding)", "DECK_PREP", "Mammoet", curVoy, 2, "", 0, 0, False, False, "", 2, "MOB"
    curVoy = curVoy + 2

    trCounter = 1
    cumDelivered = 0
    batchIdx = 0
    curIns = 0

    For v = LBound(trips) To UBound(trips)
        Dim loadN As Long
        Dim voyKey As String
        Dim voyStart As Double
        Dim voyEnd As Double

        loadN = trips(v)
        If loadN <= 0 Then GoTo NextV

        voyKey = "VOY" & CStr(v + 1)
        voyStart = curVoy

        deliveredOffset = BuildVoyageTasks(v + 1, loadN, trCounter, voyStart, weatherBuffer, tideHold)
        voyEnd = GetGroupEndOffset(voyKey)
        curVoy = voyEnd

        ' update TR counters
        cumDelivered = cumDelivered + loadN
        trCounter = trCounter + loadN

        ' When cumulative delivered reaches next batch threshold, schedule an install batch
        If batchIdx <= UBound(batches) Then
            If cumDelivered >= SumFirstN(batches, batchIdx + 1) Then
                batchIdx = batchIdx + 1
                batchSize = batches(batchIdx - 1)
                If batchSize > 0 Then
                    curIns = BuildInstallBatch(batchIdx, batchSize, parallelJacks, _
                                               MaxD(curIns, deliveredOffset + 0.5), tideHold)
                End If
            End If
        End If

NextV:
    Next v

    ' Project completion milestone
    Dim finishOff As Double
    finishOff = MaxD(GetGroupEndOffset("VOY" & CStr(UBound(trips) + 1)), curIns)
    AddTask "COMP", "99.0", "PROJECT COMPLETE — All TRs installed", "MILESTONE", "All", finishOff, 0, "", 0, 0, False, False, "PROJECT_END", 999, "MILESTONE"
End Sub


Private Function BuildVoyageTasks(ByVal voyNo As Long, ByVal loadN As Long, ByVal firstTR As Long, _
                                 ByVal voyStart As Double, ByVal weatherBuffer As Long, ByVal tideHold As Long) As Double
    ' Builds voyage tasks using Pattern_Tasks (VOYAGE_1TR/VOYAGE_2TR).
    ' Returns deliveredOffset = end of DELIVERY_READY task (jetty-ready).
    Dim tpl As String
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim seq As Long
    Dim code As String, tName As String, phase As String, owner As String
    Dim dur As Double
    Dim tideS As Boolean, weathS As Boolean
    Dim tag As String
    Dim offset As Double
    Dim id As String, wbs As String, trList As String
    Dim delivered As Double
    Dim order As Long, gKey As String
    Dim trA As Long, trB As Long

    tpl = IIf(loadN = 1, "VOYAGE_1TR", "VOYAGE_2TR")
    Set ws = Sheets(SH_PAT)
    lastR = LastRow(ws, 1, 2)

    gKey = "VOY" & CStr(voyNo)
    offset = voyStart
    order = 0
    delivered = -1

    ' TR assignment
    trA = firstTR
    trB = firstTR + 1
    If loadN = 1 Then
        trList = "TR" & trA
    Else
        trList = "TR" & trA & ",TR" & trB
    End If

    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = tpl Then
            seq = CLng(Nz(ws.Cells(r, 2).Value, 0))
            code = CStr(ws.Cells(r, 3).Value)
            tName = CStr(ws.Cells(r, 4).Value)
            phase = CStr(ws.Cells(r, 5).Value)
            owner = CStr(ws.Cells(r, 6).Value)
            dur = CDbl(Nz(ws.Cells(r, 8).Value, 0))
            tideS = (CLng(Nz(ws.Cells(r, 9).Value, 0)) = 1)
            weathS = (CLng(Nz(ws.Cells(r, 10).Value, 0)) = 1)
            tag = CStr(ws.Cells(r, 11).Value)

            order = order + 1
            id = "V" & Format(voyNo, "00") & "-" & code
            wbs = CStr(voyNo + 1) & "." & Format(order, "00")

            ' Expand labels for 2TR template
            If loadN = 2 Then
                If code = "LO1" Or code = "SF1" Then
                    tName = Replace(tName, "TR-A", "TR" & trA)
                ElseIf code = "LO2" Or code = "SF2" Or code = "UN2" Then
                    tName = Replace(tName, "TR-B", "TR" & trB)
                ElseIf code = "UN1" Then
                    tName = Replace(tName, "TR-A", "TR" & trA)
                End If
            End If

            AddTask id, wbs, tName, phase, owner, offset, dur, trList, voyNo, 0, tideS, weathS, tag, order, gKey

            offset = offset + dur

            ' capture delivered milestone (jetty-ready)
            If UCase$(tag) = "DELIVERY_READY" Then
                delivered = offset
            End If
        End If
    Next r

    ' voyage-level weather buffer (days) applied at end
    If weatherBuffer > 0 Then
        order = order + 1
        AddTask "V" & Format(voyNo, "00") & "-WBUF", CStr(voyNo + 1) & "." & Format(order, "00"), _
                "Weather/Operational Buffer (voyage)", "BUFFER", "All", offset, weatherBuffer, trList, voyNo, 0, False, False, "", order, gKey
        offset = offset + weatherBuffer
    End If

    ' Returned delivered offset; if missing, use end of UNLD/UN1 task heuristics
    If delivered < 0 Then delivered = voyStart + 4
    BuildVoyageTasks = delivered
End Function


Private Function BuildInstallBatch(ByVal batchNo As Long, ByVal batchSize As Long, ByVal parallelJacks As Long, _
                                  ByVal startOffset As Double, ByVal tideHold As Long) As Double
    ' Builds installation batch with parallel lanes.
    ' Returns end offset of batch.
    Dim lane() As Double
    Dim i As Long, laneIdx As Long
    Dim trStart As Long, trNo As Long
    Dim baseTR As Long
    Dim order As Long
    Dim gKey As String
    Dim bStart As Double, bEnd As Double
    Dim id As String, wbs As String
    Dim trList As String

    If parallelJacks <= 0 Then parallelJacks = 1
    ReDim lane(1 To parallelJacks)

    bStart = startOffset
    gKey = "BATCH" & CStr(batchNo)
    order = 0

    ' Batch setup bridge/access
    order = order + 1
    id = "B" & Format(batchNo, "00") & "-BR1"
    wbs = "I." & batchNo & ".01"
    AddTask id, wbs, "Batch " & batchNo & " — Steel bridge/access prep", "BRIDGE", "Mammoet", bStart, 0.5, "", 0, batchNo, False, False, "", order, gKey

    For i = 1 To parallelJacks
        lane(i) = bStart + 0.5
    Next i

    ' Determine TR numbers in this batch based on delivery order:
    ' Batch1: TR1..TRbatchSize
    ' Batch2: next TRs, etc.
    baseTR = GetNextTRStartForBatch(batchNo)
    trList = "TR" & baseTR & "–TR" & (baseTR + batchSize - 1)

    For i = 0 To batchSize - 1
        trNo = baseTR + i
        laneIdx = (i Mod parallelJacks) + 1

        ' Transport
        order = order + 1
        id = "B" & Format(batchNo, "00") & "-TR" & trNo & "-TRNS"
        wbs = "I." & batchNo & "." & Format(order, "00")
        AddTask id, wbs, "TR" & trNo & " — Load on SPMT + Transport to Bay", "TRANSPORT", "Mammoet", lane(laneIdx), 0.5, "TR" & trNo, 0, batchNo, False, False, "", order, gKey
        lane(laneIdx) = lane(laneIdx) + 0.5

        ' Turning
        order = order + 1
        id = "B" & Format(batchNo, "00") & "-TR" & trNo & "-TURN"
        wbs = "I." & batchNo & "." & Format(order, "00")
        AddTask id, wbs, "TR" & trNo & " — Turning operation (90°)", "TURNING", "Mammoet", lane(laneIdx), 3, "TR" & trNo, 0, batchNo, False, False, "", order, gKey
        lane(laneIdx) = lane(laneIdx) + 3

        ' Jackdown
        order = order + 1
        id = "B" & Format(batchNo, "00") & "-TR" & trNo & "-JD"
        wbs = "I." & batchNo & "." & Format(order, "00")
        AddTask id, wbs, "TR" & trNo & " — Jacking down on temporary support (Install Complete)", "JACKDOWN", "Mammoet", lane(laneIdx), 1, "TR" & trNo, 0, batchNo, False, False, "INSTALL_COMPLETE", order, gKey
        lane(laneIdx) = lane(laneIdx) + 1
    Next i

    bEnd = lane(1)
    For i = 2 To parallelJacks
        If lane(i) > bEnd Then bEnd = lane(i)
    Next i

    ' Batch close bridge/restore
    order = order + 1
    id = "B" & Format(batchNo, "00") & "-BR2"
    wbs = "I." & batchNo & "." & Format(order, "00")
    AddTask id, wbs, "Batch " & batchNo & " — Bridge relocation/restore", "BRIDGE", "Mammoet", bEnd, 0.5, trList, 0, batchNo, False, False, "", order, gKey
    bEnd = bEnd + 0.5

    BuildInstallBatch = bEnd
End Function


Private Sub AddTask(ByVal id As String, ByVal wbs As String, ByVal tName As String, ByVal phase As String, _
                    ByVal owner As String, ByVal off As Double, ByVal dur As Double, ByVal trList As String, _
                    ByVal voy As Long, ByVal batch As Long, ByVal tideS As Boolean, ByVal weathS As Boolean, _
                    ByVal tag As String, ByVal orderInGroup As Long, ByVal groupKey As String)
    mCount = mCount + 1
    ReDim Preserve mTasks(1 To mCount)
    With mTasks(mCount)
        .ID = id
        .WBS = wbs
        .Task = tName
        .Phase = phase
        .Owner = owner
        .Offset = off
        .Duration = dur
        .TR_List = trList
        .Voyage = voy
        .Batch = batch
        .TideSensitive = tideS
        .WeatherSensitive = weathS
        .KeyTag = tag
        .OrderInGroup = orderInGroup
        .GroupKey = groupKey
    End With
End Sub


Private Sub WriteSchedule()
    Dim ws As Worksheet
    Dim r As Long

    Set ws = Sheets(SH_DATA)
    ws.Rows(ROW0 & ":" & ws.Rows.Count).ClearContents

    For r = 1 To mCount
        With ws
            .Cells(ROW0 + r - 1, COL_ID).Value = mTasks(r).ID
            .Cells(ROW0 + r - 1, COL_WBS).Value = mTasks(r).WBS
            .Cells(ROW0 + r - 1, COL_TASK).Value = mTasks(r).Task
            .Cells(ROW0 + r - 1, COL_PHASE).Value = mTasks(r).Phase
            .Cells(ROW0 + r - 1, COL_OWNER).Value = mTasks(r).Owner
            .Cells(ROW0 + r - 1, COL_OFFSET).Value = mTasks(r).Offset
            .Cells(ROW0 + r - 1, COL_DUR).Value = mTasks(r).Duration
            .Cells(ROW0 + r - 1, COL_TRLIST).Value = mTasks(r).TR_List
            .Cells(ROW0 + r - 1, COL_VOY).Value = IIf(mTasks(r).Voyage = 0, "", mTasks(r).Voyage)
            .Cells(ROW0 + r - 1, COL_BATCH).Value = IIf(mTasks(r).Batch = 0, "", mTasks(r).Batch)
            .Cells(ROW0 + r - 1, COL_STATUS).Value = "Not Started"
            .Cells(ROW0 + r - 1, COL_PCT).Value = 0
        End With
    Next r

    FormatScheduleSheet
End Sub


Private Sub UpdateDates(ByVal d0 As Date)
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim off As Double, dur As Double
    Dim s As Date, e As Date

    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    For r = ROW0 To lastR
        off = CDbl(Nz(ws.Cells(r, COL_OFFSET).Value, 0))
        dur = CDbl(Nz(ws.Cells(r, COL_DUR).Value, 0))

        s = DateAdd("d", Fix(off), d0)
        e = CalcEndDate(s, dur)

        ws.Cells(r, COL_START).Value = s
        ws.Cells(r, COL_END).Value = e
        ws.Cells(r, COL_START).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, COL_END).NumberFormat = "yyyy-mm-dd"
    Next r
End Sub


Private Function CalcEndDate(ByVal startDate As Date, ByVal durDays As Double) As Date
    ' Uses ceiling for partial days (planning-level)
    Dim d As Long
    If durDays <= 0 Then
        CalcEndDate = startDate
    Else
        d = CLng(Application.WorksheetFunction.RoundUp(durDays, 0)) - 1
        If d < 0 Then d = 0
        CalcEndDate = DateAdd("d", d, startDate)
    End If
End Function

'===============================================================================
' Risks + Holds
'===============================================================================

Private Sub EvaluateRisks()
    ' Writes TideRisk/WeatherRisk columns and counts conflicts.
    Dim ws As Worksheet, wsCtrl As Worksheet
    Dim lastR As Long, r As Long
    Dim s As Date
    Dim tideR As String, weathR As String
    Dim conf As Long

    Set ws = Sheets(SH_DATA)
    Set wsCtrl = Sheets(SH_CTRL)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    conf = 0
    For r = ROW0 To lastR
        If IsDate(ws.Cells(r, COL_START).Value) Then
            s = CDate(ws.Cells(r, COL_START).Value)

            tideR = ""
            weathR = ""

            If IsTideSensitivePhase(CStr(ws.Cells(r, COL_PHASE).Value)) Then
                tideR = GetTideRisk(s)
                ws.Cells(r, COL_TIDERISK).Value = tideR
                If tideR = "HIGH" Then conf = conf + 1
            Else
                ws.Cells(r, COL_TIDERISK).Value = ""
            End If

            If IsWeatherSensitivePhase(CStr(ws.Cells(r, COL_PHASE).Value)) Then
                weathR = GetWeatherRisk(s, CStr(ws.Cells(r, COL_PHASE).Value))
                ws.Cells(r, COL_WEATHERRISK).Value = weathR
                If weathR = "HIGH" Then conf = conf + 1
            Else
                ws.Cells(r, COL_WEATHERRISK).Value = ""
            End If
        End If
    Next r

    wsCtrl.Range(O_CONFLICTS).Value = conf
End Sub


Private Sub ApplyAutoHolds(ByVal d0 As Date)
    ' Adds hold days to durations and shifts downstream offsets within each GroupKey.
    ' - Tide hold applies to LOADOUT/AGI_UNLOAD when TideRisk=HIGH
    ' - Weather-induced hold is handled via voyage-level Weather Buffer already.
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim holdT As Long
    Dim phase As String
    Dim tideR As String
    Dim groupKey As String
    Dim order As Long
    Dim addHold As Double

    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    holdT = ReadLong(SH_CTRL, C_THOLD, 0)
    If holdT <= 0 Then Exit Sub

    ' First pass: mark holds per task
    For r = ROW0 To lastR
        phase = UCase$(Trim$(CStr(ws.Cells(r, COL_PHASE).Value)))
        tideR = UCase$(Trim$(CStr(ws.Cells(r, COL_TIDERISK).Value)))
        If (phase = "LOADOUT" Or phase = "AGI_UNLOAD") And tideR = "HIGH" Then
            ws.Cells(r, COL_NOTES).Value = AppendNote(CStr(ws.Cells(r, COL_NOTES).Value), "AUTO-HOLD(TIDE): +" & holdT & "d")
            ws.Cells(r, COL_DUR).Value = CDbl(Nz(ws.Cells(r, COL_DUR).Value, 0)) + holdT
            ' shift downstream in same group based on task order index (derived by row order)
            groupKey = InferGroupKeyFromID(CStr(ws.Cells(r, COL_ID).Value))
            If Len(groupKey) > 0 Then
                ShiftGroupOffsets ws, groupKey, r, holdT
            End If
        End If
    Next r
End Sub


Private Sub ShiftGroupOffsets(ByVal ws As Worksheet, ByVal groupKey As String, ByVal triggerRow As Long, ByVal holdDays As Long)
    ' Shifts offsets for tasks in same group located after triggerRow.
    Dim lastR As Long, r As Long
    lastR = LastRow(ws, COL_ID, ROW0)
    For r = triggerRow + 1 To lastR
        If InferGroupKeyFromID(CStr(ws.Cells(r, COL_ID).Value)) = groupKey Then
            ws.Cells(r, COL_OFFSET).Value = CDbl(Nz(ws.Cells(r, COL_OFFSET).Value, 0)) + holdDays
        End If
    Next r
End Sub


Private Function InferGroupKeyFromID(ByVal id As String) As String
    ' VOY group: V##-...
    ' BATCH group: B##-...
    If Left$(id, 1) = "V" And InStr(1, id, "-", vbTextCompare) > 0 Then
        InferGroupKeyFromID = "VOY" & CStr(CLng(Mid$(id, 2, 2)))
    ElseIf Left$(id, 1) = "B" And InStr(1, id, "-", vbTextCompare) > 0 Then
        InferGroupKeyFromID = "BATCH" & CStr(CLng(Mid$(id, 2, 2)))
    Else
        InferGroupKeyFromID = ""
    End If
End Function


Private Function IsTideSensitivePhase(ByVal phase As String) As Boolean
    phase = UCase$(Trim$(phase))
    IsTideSensitivePhase = (phase = "LOADOUT" Or phase = "AGI_UNLOAD")
End Function

Private Function IsWeatherSensitivePhase(ByVal phase As String) As Boolean
    phase = UCase$(Trim$(phase))
    IsWeatherSensitivePhase = (phase = "SAIL" Or phase = "RETURN")
End Function


Private Function GetTideRisk(ByVal d As Date) As String
    ' Lookup from Tide_Data (Date in column A, Risk in column C).
    Dim ws As Worksheet
    Dim rng As Range, f As Range
    Set ws = Sheets(SH_TIDE)
    Set rng = ws.Range("A:A")
    Set f = rng.Find(What:=CLng(d), LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        GetTideRisk = UCase$(Trim$(CStr(ws.Cells(f.Row, 3).Value)))
    Else
        GetTideRisk = "LOW"
    End If
End Function


Private Function GetWeatherRisk(ByVal d As Date, ByVal phase As String) As String
    ' Lookup Weather_Risk by date range; return LOW/MED/HIGH.
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim s As Date, e As Date, lvl As String, applies As String
    Set ws = Sheets(SH_WEATHER)
    lastR = LastRow(ws, 1, 2)
    For r = 2 To lastR
        If IsDate(ws.Cells(r, 1).Value) And IsDate(ws.Cells(r, 2).Value) Then
            s = CDate(ws.Cells(r, 1).Value)
            e = CDate(ws.Cells(r, 2).Value)
            lvl = UCase$(Trim$(CStr(ws.Cells(r, 3).Value)))
            applies = UCase$(Trim$(CStr(ws.Cells(r, 7).Value)))
            If d >= s And d <= e Then
                If InStr(1, applies, UCase$(phase), vbTextCompare) > 0 Then
                    GetWeatherRisk = lvl
                    Exit Function
                End If
            End If
        End If
    Next r
    GetWeatherRisk = "LOW"
End Function

'===============================================================================
' Critical Path (planning-level)
'===============================================================================

Private Sub MarkCriticalPath()
    ' Planning-level: mark as critical any task in the chain that determines project finish.
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim finish As Date
    Dim maxEnd As Date, maxRow As Long
    Dim critGroup As String

    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    maxEnd = DateSerial(1900, 1, 1)
    maxRow = ROW0
    For r = ROW0 To lastR
        If IsDate(ws.Cells(r, COL_END).Value) Then
            If CDate(ws.Cells(r, COL_END).Value) > maxEnd Then
                maxEnd = CDate(ws.Cells(r, COL_END).Value)
                maxRow = r
            End If
        End If
    Next r

    critGroup = InferGroupKeyFromID(CStr(ws.Cells(maxRow, COL_ID).Value))

    For r = ROW0 To lastR
        If InferGroupKeyFromID(CStr(ws.Cells(r, COL_ID).Value)) = critGroup Then
            ws.Cells(r, COL_CRIT).Value = "Y"
        Else
            ws.Cells(r, COL_CRIT).Value = ""
        End If
    Next r

    Sheets(SH_CTRL).Range(O_CPLEN).Value = DateDiff("d", ReadDate(SH_CTRL, C_D0, Date), maxEnd)
End Sub

'===============================================================================
' Gantt
'===============================================================================

Private Sub RefreshGantt()
    Dim wsD As Worksheet, wsG As Worksheet
    Dim lastR As Long, r As Long, gRow As Long
    Dim minS As Date, maxE As Date
    Dim startT As Date, totalDays As Long
    Dim col0 As Long, c As Long
    Dim d As Date
    Dim s As Date, e As Date
    Dim crit As String, phase As String

    Set wsD = Sheets(SH_DATA)
    Set wsG = Sheets(SH_GANTT)

    lastR = LastRow(wsD, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    minS = DateSerial(2099, 12, 31)
    maxE = DateSerial(1900, 1, 1)
    For r = ROW0 To lastR
        If IsDate(wsD.Cells(r, COL_START).Value) Then
            s = CDate(wsD.Cells(r, COL_START).Value)
            e = IIf(IsDate(wsD.Cells(r, COL_END).Value), CDate(wsD.Cells(r, COL_END).Value), s)
            If s < minS Then minS = s
            If e > maxE Then maxE = e
        End If
    Next r

    startT = DateAdd("d", -2, minS)
    totalDays = DateDiff("d", startT, DateAdd("d", 2, maxE)) + 1
    If totalDays < 30 Then totalDays = 30
    If totalDays > 140 Then totalDays = 140

    col0 = 13 ' timeline starts at M

    ' Clear prior
    wsG.Range(wsG.Cells(5, 1), wsG.Cells(wsG.Rows.Count, col0 + 200)).ClearContents
    wsG.Range(wsG.Cells(5, col0), wsG.Cells(wsG.Rows.Count, col0 + 200)).Interior.ColorIndex = xlNone

    ' Header timeline row 4
    For c = 0 To totalDays - 1
        d = DateAdd("d", c, startT)
        wsG.Cells(4, col0 + c).Value = d
        wsG.Cells(4, col0 + c).NumberFormat = "d"
        If Weekday(d, vbMonday) > 5 Then
            wsG.Cells(4, col0 + c).Interior.Color = RGB(242, 242, 242)
        End If
        If GetWeatherRisk(d, "SAIL") = "HIGH" Then
            wsG.Cells(4, col0 + c).Interior.Color = RGB(255, 230, 153)
        End If
    Next c

    ' Write rows + bars
    For r = ROW0 To lastR
        gRow = 5 + (r - ROW0)

        wsG.Cells(gRow, 1).Value = wsD.Cells(r, COL_ID).Value
        wsG.Cells(gRow, 2).Value = wsD.Cells(r, COL_WBS).Value
        wsG.Cells(gRow, 3).Value = wsD.Cells(r, COL_TASK).Value
        wsG.Cells(gRow, 4).Value = wsD.Cells(r, COL_PHASE).Value
        wsG.Cells(gRow, 5).Value = wsD.Cells(r, COL_START).Value
        wsG.Cells(gRow, 6).Value = wsD.Cells(r, COL_END).Value
        wsG.Cells(gRow, 7).Value = wsD.Cells(r, COL_DUR).Value
        wsG.Cells(gRow, 8).Value = wsD.Cells(r, COL_OWNER).Value
        wsG.Cells(gRow, 9).Value = wsD.Cells(r, COL_TRLIST).Value
        wsG.Cells(gRow, 10).Value = wsD.Cells(r, COL_VOY).Value
        wsG.Cells(gRow, 11).Value = wsD.Cells(r, COL_BATCH).Value
        wsG.Cells(gRow, 12).Value = wsD.Cells(r, COL_CRIT).Value

        wsG.Cells(gRow, 5).NumberFormat = "yyyy-mm-dd"
        wsG.Cells(gRow, 6).NumberFormat = "yyyy-mm-dd"

        If IsDate(wsD.Cells(r, COL_START).Value) Then
            s = CDate(wsD.Cells(r, COL_START).Value)
            e = IIf(IsDate(wsD.Cells(r, COL_END).Value), CDate(wsD.Cells(r, COL_END).Value), s)
            phase = UCase$(CStr(wsD.Cells(r, COL_PHASE).Value))
            crit = CStr(wsD.Cells(r, COL_CRIT).Value)
            DrawBar wsG, gRow, col0, startT, s, e, phase, (crit = "Y")
        End If
    Next r

    wsG.Activate
    wsG.Range(wsG.Cells(5, col0), wsG.Cells(5, col0)).Select
    ActiveWindow.FreezePanes = True
End Sub


Private Sub DrawBar(ByVal ws As Worksheet, ByVal row As Long, ByVal col0 As Long, ByVal startT As Date, _
                    ByVal s As Date, ByVal e As Date, ByVal phase As String, ByVal isCrit As Boolean)
    Dim cStart As Long, cEnd As Long, c As Long
    Dim fill As Long

    cStart = col0 + DateDiff("d", startT, s)
    cEnd = col0 + DateDiff("d", startT, e)

    fill = PhaseColor(phase)
    For c = cStart To cEnd
        ws.Cells(row, c).Interior.Color = fill
        If isCrit Then
            ws.Cells(row, c).Borders(xlEdgeBottom).Color = RGB(192, 0, 0)
            ws.Cells(row, c).Borders(xlEdgeBottom).Weight = xlThick
        End If
    Next c
End Sub


Private Function PhaseColor(ByVal phase As String) As Long
    phase = UCase$(Trim$(phase))
    Select Case phase
        Case "MOBILIZATION": PhaseColor = RGB(142, 124, 195)
        Case "DECK_PREP": PhaseColor = RGB(111, 168, 220)
        Case "LOADOUT": PhaseColor = RGB(147, 196, 125)
        Case "SEAFAST": PhaseColor = RGB(118, 165, 175)
        Case "SAIL": PhaseColor = RGB(164, 194, 244)
        Case "AGI_UNLOAD": PhaseColor = RGB(246, 178, 107)
        Case "TURNING": PhaseColor = RGB(255, 217, 102)
        Case "JACKDOWN": PhaseColor = RGB(224, 102, 102)
        Case "RETURN": PhaseColor = RGB(153, 153, 153)
        Case "BRIDGE": PhaseColor = RGB(215, 227, 188)
        Case "TRANSPORT": PhaseColor = RGB(180, 205, 230)
        Case "BUFFER": PhaseColor = RGB(217, 217, 217)
        Case "MILESTONE": PhaseColor = RGB(192, 0, 0)
        Case Else: PhaseColor = RGB(200, 200, 200)
    End Select
End Function

'===============================================================================
' Monte Carlo
'===============================================================================

Private Sub MonteCarloFinish(ByVal d0 As Date, ByVal runs As Long, ByVal conf As Double)
    ' Simulates finish date based on:
    '   - Weather_Risk triangular delays for SAIL/RETURN/LOADOUT/AGI_UNLOAD
    '   - Tide_Data risk delays for LOADOUT/AGI_UNLOAD
    ' Planning-level: apply random delays at voyage level + risk tasks.
    Dim wsD As Worksheet
    Dim lastR As Long
    Dim fin() As Double
    Dim i As Long, fOff As Double
    Dim p50 As Date, pX As Date

    On Error GoTo ErrH
    If runs < 50 Then runs = 50
    If conf < 0.5 Then conf = 0.5
    If conf > 0.95 Then conf = 0.95

    Set wsD = Sheets(SH_DATA)
    lastR = LastRow(wsD, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    ReDim fin(1 To runs)

    Randomize
    For i = 1 To runs
        fOff = SimulateFinishOffset(d0)
        fin(i) = fOff
    Next i

    SortDoubles fin

    p50 = DateAdd("d", CLng(fin(Application.WorksheetFunction.RoundUp(runs * 0.5, 0))), d0)
    pX = DateAdd("d", CLng(fin(Application.WorksheetFunction.RoundUp(runs * conf, 0))), d0)

    Sheets(SH_CTRL).Range(O_P50).Value = Format(p50, "yyyy-mm-dd")
    Sheets(SH_CTRL).Range(O_P80).Value = Format(pX, "yyyy-mm-dd")
    Exit Sub
ErrH:
    LogMsg "ERROR", "MonteCarloFinish", Err.Description
End Sub


Private Function SimulateFinishOffset(ByVal d0 As Date) As Double
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim off As Double, dur As Double
    Dim s As Date, phase As String
    Dim dly As Double
    Dim endOff As Double

    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    endOff = 0

    For r = ROW0 To lastR
        off = CDbl(Nz(ws.Cells(r, COL_OFFSET).Value, 0))
        dur = CDbl(Nz(ws.Cells(r, COL_DUR).Value, 0))
        phase = UCase$(Trim$(CStr(ws.Cells(r, COL_PHASE).Value)))

        s = DateAdd("d", CLng(off), d0)
        dly = 0

        ' Weather
        If IsWeatherSensitivePhase(phase) Or phase = "LOADOUT" Or phase = "AGI_UNLOAD" Then
            dly = dly + SampleWeatherDelay(s, phase)
        End If

        ' Tide
        If IsTideSensitivePhase(phase) Then
            dly = dly + SampleTideDelay(s)
        End If

        If off + dur + dly > endOff Then endOff = off + dur + dly
    Next r

    SimulateFinishOffset = endOff
End Function


Private Function SampleWeatherDelay(ByVal d As Date, ByVal phase As String) As Double
    ' Triangular delay (min/mode/max) by date window & phase applicability.
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim s As Date, e As Date, lvl As String, applies As String
    Dim dMin As Double, dMode As Double, dMax As Double

    Set ws = Sheets(SH_WEATHER)
    lastR = LastRow(ws, 1, 2)

    For r = 2 To lastR
        If IsDate(ws.Cells(r, 1).Value) And IsDate(ws.Cells(r, 2).Value) Then
            s = CDate(ws.Cells(r, 1).Value)
            e = CDate(ws.Cells(r, 2).Value)
            applies = UCase$(CStr(ws.Cells(r, 7).Value))
            If d >= s And d <= e Then
                If InStr(1, applies, UCase$(phase), vbTextCompare) > 0 Or InStr(1, applies, "SAIL", vbTextCompare) > 0 Then
                    dMin = CDbl(Nz(ws.Cells(r, 4).Value, 0))
                    dMode = CDbl(Nz(ws.Cells(r, 5).Value, dMin))
                    dMax = CDbl(Nz(ws.Cells(r, 6).Value, dMode))
                    SampleWeatherDelay = Triangular(dMin, dMode, dMax)
                    Exit Function
                End If
            End If
        End If
    Next r

    SampleWeatherDelay = 0
End Function


Private Function SampleTideDelay(ByVal d As Date) As Double
    Dim lvl As String
    lvl = GetTideRisk(d)
    Select Case lvl
        Case "HIGH": SampleTideDelay = Triangular(0, 1, 2)
        Case "MED": SampleTideDelay = Triangular(0, 0, 1)
        Case Else: SampleTideDelay = 0
    End Select
End Function


Private Function Triangular(ByVal a As Double, ByVal c As Double, ByVal b As Double) As Double
    ' Returns random value from triangular distribution.
    Dim u As Double, f As Double
    If b <= a Then Triangular = a: Exit Function
    u = Rnd
    f = (c - a) / (b - a)
    If u < f Then
        Triangular = a + Sqr(u * (b - a) * (c - a))
    Else
        Triangular = b - Sqr((1 - u) * (b - a) * (b - c))
    End If
End Function


Private Sub SortDoubles(ByRef arr() As Double)
    Dim i As Long, j As Long
    Dim tmp As Double
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'===============================================================================
' Dashboard + Outputs + Reports
'===============================================================================

Private Sub UpdateDashboardAndOutputs(ByVal d0 As Date)
    Dim wsCtrl As Worksheet
    Dim finish As Date
    Dim deadline As Date
    Dim ok As Boolean

    Set wsCtrl = Sheets(SH_CTRL)
    finish = GetProjectFinish()
    deadline = ReadDate(SH_CTRL, C_DEADLINE, DateSerial(2026, 3, 1))

    wsCtrl.Range(O_FINISH).Value = Format(finish, "yyyy-mm-dd")
    wsCtrl.Range(O_P50).Value = "(run MC)"
    wsCtrl.Range(O_P80).Value = "(run MC)"
    ok = (finish <= deadline)
    wsCtrl.Range(O_MEET).Value = IIf(ok, "YES", "NO")
    wsCtrl.Range(O_NOTES).Value = BuildNotesText(finish, deadline)

    ' Minimal dashboard
    Sheets(SH_DASH).Range("A1").Value = "Dashboard (auto-updated) — " & Format(Now, "yyyy-mm-dd hh:nn")
    Sheets(SH_DASH).Range("A5").Value = "Deterministic Finish"
    Sheets(SH_DASH).Range("B5").Value = Format(finish, "yyyy-mm-dd")
    Sheets(SH_DASH).Range("A6").Value = "Deadline"
    Sheets(SH_DASH).Range("B6").Value = Format(deadline, "yyyy-mm-dd")
    Sheets(SH_DASH).Range("A7").Value = "Conflicts (Tide/Weather HIGH)"
    Sheets(SH_DASH).Range("B7").Value = wsCtrl.Range(O_CONFLICTS).Value
End Sub


Private Function BuildNotesText(ByVal finish As Date, ByVal deadline As Date) As String
    Dim s As String
    s = "Plan finish: " & Format(finish, "yyyy-mm-dd") & vbCrLf & _
        "Deadline: " & Format(deadline, "yyyy-mm-dd") & vbCrLf & _
        "Actions:" & vbCrLf & _
        "1) If NO, run OptimizeD0 (Ctrl+Shift+O) and/or reduce buffers." & vbCrLf & _
        "2) Replace Tide_Data with official tide table for final plan." & vbCrLf & _
        "3) Maintain separate resources: Shuttle vs Install (3대 잭다운) to preserve overlap."
    BuildNotesText = s
End Function


Private Sub BuildReports(ByVal d0 As Date)
    Dim ws As Worksheet
    Dim finish As Date
    finish = GetProjectFinish()

    Set ws = Sheets(SH_REP)
    ws.Cells.Clear
    ws.Range("A1").Value = "AGI TR Transportation — Executive Report"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3").Value = "Generated:"
    ws.Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn")
    ws.Range("A4").Value = "D0:"
    ws.Range("B4").Value = Format(d0, "yyyy-mm-dd")
    ws.Range("A5").Value = "Deterministic Finish:"
    ws.Range("B5").Value = Format(finish, "yyyy-mm-dd")
    ws.Range("A6").Value = "Deadline:"
    ws.Range("B6").Value = Format(ReadDate(SH_CTRL, C_DEADLINE, DateSerial(2026, 3, 1)), "yyyy-mm-dd")
    ws.Range("A7").Value = "Conflicts (HIGH):"
    ws.Range("B7").Value = Sheets(SH_CTRL).Range(O_CONFLICTS).Value

    ws.Range("A9").Value = "Milestones"
    ws.Range("A9").Font.Bold = True
    WriteMilestones ws, 10

    ws.Range("A20").Value = "Next 14 Days Lookahead"
    ws.Range("A20").Font.Bold = True
    WriteLookahead ws, 21, d0, 14
End Sub


Private Sub WriteMilestones(ByVal wsRep As Worksheet, ByVal startRow As Long)
    Dim wsD As Worksheet
    Dim lastR As Long, r As Long, o As Long
    Set wsD = Sheets(SH_DATA)
    lastR = LastRow(wsD, COL_ID, ROW0)

    wsRep.Cells(startRow, 1).Value = "ID"
    wsRep.Cells(startRow, 2).Value = "Task"
    wsRep.Cells(startRow, 3).Value = "Date"
    wsRep.Range(wsRep.Cells(startRow, 1), wsRep.Cells(startRow, 3)).Font.Bold = True

    o = startRow + 1
    For r = ROW0 To lastR
        If UCase$(Trim$(CStr(wsD.Cells(r, COL_PHASE).Value))) = "MILESTONE" Or _
           InStr(1, UCase$(CStr(wsD.Cells(r, COL_TASK).Value)), "INSTALL COMPLETE", vbTextCompare) > 0 Then
            wsRep.Cells(o, 1).Value = wsD.Cells(r, COL_ID).Value
            wsRep.Cells(o, 2).Value = wsD.Cells(r, COL_TASK).Value
            wsRep.Cells(o, 3).Value = Format(CDate(wsD.Cells(r, COL_END).Value), "yyyy-mm-dd")
            o = o + 1
        End If
    Next r
End Sub


Private Sub WriteLookahead(ByVal wsRep As Worksheet, ByVal startRow As Long, ByVal d0 As Date, ByVal daysAhead As Long)
    Dim wsD As Worksheet
    Dim lastR As Long, r As Long, o As Long
    Dim today As Date, cut As Date
    Dim s As Date

    today = Date
    cut = today + daysAhead

    Set wsD = Sheets(SH_DATA)
    lastR = LastRow(wsD, COL_ID, ROW0)

    wsRep.Cells(startRow, 1).Value = "Start"
    wsRep.Cells(startRow, 2).Value = "ID"
    wsRep.Cells(startRow, 3).Value = "Task"
    wsRep.Range(wsRep.Cells(startRow, 1), wsRep.Cells(startRow, 3)).Font.Bold = True

    o = startRow + 1
    For r = ROW0 To lastR
        If IsDate(wsD.Cells(r, COL_START).Value) Then
            s = CDate(wsD.Cells(r, COL_START).Value)
            If s >= today And s <= cut Then
                wsRep.Cells(o, 1).Value = Format(s, "yyyy-mm-dd")
                wsRep.Cells(o, 2).Value = wsD.Cells(r, COL_ID).Value
                wsRep.Cells(o, 3).Value = wsD.Cells(r, COL_TASK).Value
                o = o + 1
            End If
        End If
    Next r
End Sub

'===============================================================================
' Exports
'===============================================================================

Private Sub ExportScheduleCSV()
    Dim ws As Worksheet, wsE As Worksheet
    Dim lastR As Long, lastC As Long
    Dim fp As String
    Dim r As Long, c As Long
    Dim line As String
    Dim fNum As Integer

    Set ws = Sheets(SH_DATA)
    Set wsE = Sheets(SH_EXP)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub
    lastC = COL_NOTES

    fp = ThisWorkbook.Path & "\AGI_TR_Schedule_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    fNum = FreeFile
    Open fp For Output As #fNum

    line = ""
    For c = 1 To lastC
        If c > 1 Then line = line & ","
        line = line & """" & Replace(CStr(ws.Cells(HDR_ROW, c).Value), """", """""") & """"
    Next c
    Print #fNum, line

    For r = ROW0 To lastR
        line = ""
        For c = 1 To lastC
            If c > 1 Then line = line & ","
            line = line & """" & Replace(CStr(ws.Cells(r, c).Value), """", """""") & """"
        Next c
        Print #fNum, line
    Next r
    Close #fNum

    wsE.Range("A5").Value = "CSV"
    wsE.Range("B5").Value = fp
    LogMsg "INFO", "ExportScheduleCSV", fp
End Sub


Private Sub ExportGanttPDF()
    Dim ws As Worksheet, wsE As Worksheet
    Dim fp As String
    Set ws = Sheets(SH_GANTT)
    Set wsE = Sheets(SH_EXP)

    fp = ThisWorkbook.Path & "\AGI_TR_Gantt_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fp, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    wsE.Range("A6").Value = "PDF"
    wsE.Range("B6").Value = fp
    LogMsg "INFO", "ExportGanttPDF", fp
End Sub

'===============================================================================
' Buttons (optional UI)
'===============================================================================

Private Sub CreateButtons()
    ' Creates simple buttons on Control_Panel.
    Dim ws As Worksheet
    Dim leftPos As Double, topPos As Double
    Dim btn As Shape
    On Error Resume Next
    Set ws = Sheets(SH_CTRL)

    ' Delete existing buttons created by this macro
    Dim s As Shape
    For Each s In ws.Shapes
        If Left$(s.Name, 7) = "AGIbtn_" Then s.Delete
    Next s

    leftPos = ws.Range("B27").Left
    topPos = ws.Range("B27").Top + 22

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 160, 26)
    btn.Name = "AGIbtn_RunAll"
    btn.TextFrame2.TextRange.Text = "Run All (Ctrl+Shift+U)"
    btn.OnAction = "modAGI_TR_Suite.RunAll"

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 170, topPos, 160, 26)
    btn.Name = "AGIbtn_Optimize"
    btn.TextFrame2.TextRange.Text = "Optimize D0 (Ctrl+Shift+O)"
    btn.OnAction = "modAGI_TR_Suite.OptimizeD0"

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 340, topPos, 160, 26)
    btn.Name = "AGIbtn_MC"
    btn.TextFrame2.TextRange.Text = "Monte Carlo (Ctrl+Shift+M)"
    btn.OnAction = "modAGI_TR_Suite.RunMonteCarloOnly"

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos + 34, 160, 26)
    btn.Name = "AGIbtn_Export"
    btn.TextFrame2.TextRange.Text = "Export PDF/CSV (Ctrl+Shift+R)"
    btn.OnAction = "modAGI_TR_Suite.ExportPDFandCSV"

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 170, topPos + 34, 160, 26)
    btn.Name = "AGIbtn_Baseline"
    btn.TextFrame2.TextRange.Text = "Freeze Baseline (Ctrl+Shift+S)"
    btn.OnAction = "modAGI_TR_Suite.FreezeBaseline"

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 340, topPos + 34, 160, 26)
    btn.Name = "AGIbtn_Compare"
    btn.TextFrame2.TextRange.Text = "Compare Baseline (Ctrl+Shift+D)"
    btn.OnAction = "modAGI_TR_Suite.CompareToBaseline"

    ' Style
    For Each s In ws.Shapes
        If Left$(s.Name, 7) = "AGIbtn_" Then
            s.Fill.ForeColor.RGB = RGB(47, 85, 151)
            s.Line.ForeColor.RGB = RGB(31, 78, 121)
            s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            s.TextFrame2.TextRange.Font.Size = 10
        End If
    Next s
End Sub

'===============================================================================
' Validations / Formatting
'===============================================================================

Private Sub EnsureAllSheetsExist()
    ' No-op if sheets exist; creates missing sheets.
    Dim names As Variant, i As Long
    names = Array(SH_CTRL, SH_SCEN, SH_PAT, SH_DATA, SH_GANTT, SH_TIDE, SH_WEATHER, SH_DASH, SH_DOCS, SH_EVID, SH_BASE, SH_CHG, SH_REP, SH_LOG, SH_EXP)
    For i = LBound(names) To UBound(names)
        If Not SheetExists(CStr(names(i))) Then
            ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = CStr(names(i))
        End If
    Next i
End Sub

Private Sub EnsureHeaders()
    ' Ensures Schedule_Data header row exists.
    Dim ws As Worksheet
    Set ws = Sheets(SH_DATA)
    If Len(Trim$(CStr(ws.Cells(HDR_ROW, 1).Value))) = 0 Then
        ws.Cells(HDR_ROW, 1).Value = "ID"
    End If
End Sub

Private Sub ApplyValidations()
    ' Applies scenario dropdown if missing.
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Sheets(SH_CTRL)
    ws.Range(C_SCEN).Validation.Delete
    ws.Range(C_SCEN).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="S1_6Voy_1TR,S2_4Voy_1-2-2-1,S3_Custom"
End Sub

Private Sub FormatScheduleSheet()
    Dim ws As Worksheet
    Dim lastR As Long
    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    If lastR < ROW0 Then Exit Sub

    ws.Columns(COL_START).NumberFormat = "yyyy-mm-dd"
    ws.Columns(COL_END).NumberFormat = "yyyy-mm-dd"
    ws.Columns(COL_BS).NumberFormat = "yyyy-mm-dd"
    ws.Columns(COL_BE).NumberFormat = "yyyy-mm-dd"
End Sub

'===============================================================================
' Scenario defaults
'===============================================================================

Private Function GetScenarioDefaults(ByVal scenID As String, ByRef tripPlan As String, ByRef batchPlan As String, _
                                    ByRef jacks As Long, ByRef wBuf As Long, ByRef tideHold As Long) As Boolean
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Set ws = Sheets(SH_SCEN)
    lastR = LastRow(ws, 1, 2)
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = scenID Then
            tripPlan = CStr(ws.Cells(r, 3).Value)
            batchPlan = CStr(ws.Cells(r, 4).Value)
            jacks = CLng(Nz(ws.Cells(r, 5).Value, jacks))
            wBuf = CLng(Nz(ws.Cells(r, 6).Value, wBuf))
            tideHold = CLng(Nz(ws.Cells(r, 7).Value, tideHold))
            GetScenarioDefaults = True
            Exit Function
        End If
    Next r
    GetScenarioDefaults = False
End Function

'===============================================================================
' Validation
'===============================================================================

Private Sub ValidateInputs(ByRef msgOut As String)
    Dim tripPlan As String, batchPlan As String
    Dim trips() As Long, batches() As Long
    Dim s As String

    msgOut = ""
    tripPlan = CStr(Sheets(SH_CTRL).Range(C_TRIPPLAN).Value)
    batchPlan = CStr(Sheets(SH_CTRL).Range(C_BATCHES).Value)

    trips = ParseLongList(tripPlan)
    batches = ParseLongList(batchPlan)

    If SumArray(trips) <> 6 Then
        msgOut = msgOut & "- Trip plan sum is " & SumArray(trips) & " (expected 6 TR units)." & vbCrLf
    End If
    If SumArray(batches) <> 6 Then
        msgOut = msgOut & "- Install batch sum is " & SumArray(batches) & " (expected 6 TR units)." & vbCrLf
    End If
    If ReadLong(SH_CTRL, C_JACKS, 3) < 1 Then
        msgOut = msgOut & "- Parallel jacks must be >= 1." & vbCrLf
    End If
    If Not IsDate(Sheets(SH_CTRL).Range(C_D0).Value) Then
        msgOut = msgOut & "- D0 must be a valid date." & vbCrLf
    End If
End Sub

'===============================================================================
' Utilities
'===============================================================================

Private Function ParseLongList(ByVal s As String) As Long()
    ' Parses "1,2,2,1" -> array of longs
    Dim parts() As String, tmp() As Long
    Dim i As Long, n As Long, v As Long
    s = Replace(Replace(Trim$(s), ";", ","), " ", "")
    If Len(s) = 0 Then
        ReDim tmp(0 To 0): tmp(0) = 0
        ParseLongList = tmp
        Exit Function
    End If
    parts = Split(s, ",")
    ReDim tmp(0 To UBound(parts))
    n = 0
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            v = CLng(Val(parts(i)))
            tmp(n) = v
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim tmp(0 To 0): tmp(0) = 0
    ElseIf n - 1 < UBound(tmp) Then
        ReDim Preserve tmp(0 To n - 1)
    End If
    ParseLongList = tmp
End Function

Private Function SumArray(ByRef arr() As Long) As Long
    Dim i As Long, s As Long
    s = 0
    For i = LBound(arr) To UBound(arr)
        s = s + arr(i)
    Next i
    SumArray = s
End Function

Private Function SumFirstN(ByRef arr() As Long, ByVal n As Long) As Long
    Dim i As Long, s As Long
    s = 0
    For i = LBound(arr) To UBound(arr)
        If i + 1 <= n Then s = s + arr(i)
    Next i
    SumFirstN = s
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function Nz(ByVal v As Variant, ByVal def As Variant) As Variant
    If IsError(v) Then
        Nz = def
    ElseIf IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        Nz = def
    Else
        Nz = v
    End If
End Function

Private Function LastRow(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < startRow Then r = startRow - 1
    LastRow = r
End Function

Private Function ReadDate(ByVal sheetName As String, ByVal addr As String, ByVal def As Date) As Date
    On Error GoTo ErrH
    Dim v As Variant
    v = ThisWorkbook.Worksheets(sheetName).Range(addr).Value
    If IsDate(v) Then
        ReadDate = CDate(v)
    Else
        ReadDate = def
    End If
    Exit Function
ErrH:
    ReadDate = def
End Function

Private Function ReadLong(ByVal sheetName As String, ByVal addr As String, ByVal def As Long) As Long
    On Error GoTo ErrH
    Dim v As Variant
    v = ThisWorkbook.Worksheets(sheetName).Range(addr).Value
    ReadLong = CLng(Nz(v, def))
    Exit Function
ErrH:
    ReadLong = def
End Function

Private Function ReadDouble(ByVal sheetName As String, ByVal addr As String, ByVal def As Double) As Double
    On Error GoTo ErrH
    Dim v As Variant
    v = ThisWorkbook.Worksheets(sheetName).Range(addr).Value
    ReadDouble = CDbl(Nz(v, def))
    Exit Function
ErrH:
    ReadDouble = def
End Function

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    SheetExists = Not ws Is Nothing
End Function

Private Function AppendNote(ByVal base As String, ByVal add As String) As String
    If Len(Trim$(base)) = 0 Then
        AppendNote = add
    Else
        AppendNote = base & " | " & add
    End If
End Function

Private Function GetGroupEndOffset(ByVal groupKey As String) As Double
    ' Gets maximum Offset+Duration for tasks in memory list whose GroupKey matches.
    Dim i As Long, endO As Double
    endO = 0
    For i = 1 To mCount
        If mTasks(i).GroupKey = groupKey Then
            If mTasks(i).Offset + mTasks(i).Duration > endO Then endO = mTasks(i).Offset + mTasks(i).Duration
        End If
    Next i
    GetGroupEndOffset = endO
End Function

Private Function GetProjectFinish() As Date
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim maxE As Date
    Set ws = Sheets(SH_DATA)
    lastR = LastRow(ws, COL_ID, ROW0)
    maxE = DateSerial(1900, 1, 1)
    For r = ROW0 To lastR
        If IsDate(ws.Cells(r, COL_END).Value) Then
            If CDate(ws.Cells(r, COL_END).Value) > maxE Then maxE = CDate(ws.Cells(r, COL_END).Value)
        End If
    Next r
    GetProjectFinish = maxE
End Function


Private Function GetNextTRStartForBatch(ByVal batchNo As Long) As Long
    ' Determines TR numbering for each install batch based on already-added tasks.
    ' If no prior batch install tasks exist, returns 1.
    Dim i As Long
    Dim maxTR As Long
    Dim t As String, n As Long
    maxTR = 0

    For i = 1 To mCount
        If UCase$(mTasks(i).KeyTag) = "INSTALL_COMPLETE" Then
            t = UCase$(mTasks(i).TR_List) ' e.g., "TR4"
            n = ExtractFirstNumber(t)
            If n > maxTR Then maxTR = n
        End If
    Next i

    If maxTR = 0 Then
        GetNextTRStartForBatch = 1
    Else
        GetNextTRStartForBatch = maxTR + 1
    End If
End Function

Private Function ExtractFirstNumber(ByVal s As String) As Long
    ' Extracts first integer from string; returns 0 if none.
    Dim i As Long, ch As String, num As String
    num = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            num = num & ch
        ElseIf Len(num) > 0 Then
            Exit For
        End If
    Next i
    If Len(num) = 0 Then
        ExtractFirstNumber = 0
    Else
        ExtractFirstNumber = CLng(num)
    End If
End Function


'===============================================================================
' Logging
'===============================================================================

Private Sub LogMsg(ByVal level As String, ByVal moduleName As String, ByVal message As String)
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = Sheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = level
    ws.Cells(r, 3).Value = moduleName
    ws.Cells(r, 4).Value = message
End Sub

Private Sub LogChange(ByVal action As String, ByVal details As String)
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = Sheets(SH_CHG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = Environ$("USERNAME")
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = details
End Sub