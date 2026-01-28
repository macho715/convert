Attribute VB_Name = "modOperations"
Option Explicit

'===============================================================================
' OPERATIONS MODULE - Minimum Essential Functions for TR_DocHub_AGI
' Version: 1.0 | Project: HVDC AGI TR Transportation
'===============================================================================

'------------------------------------------------------------------------------
' 1. InitializeWorkbook()
'------------------------------------------------------------------------------
Public Sub InitializeWorkbook()
    On Error GoTo EH
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Dim required_sheets As Variant
    required_sheets = Array("S_Voyages", "M_DocCatalog", "M_Parties", "R_DeadlineRules", "T_Tracker", "D_Dashboard")
    
    Dim sh As Variant
    For Each sh In required_sheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sh))
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Required sheet missing: " & sh, vbCritical
            Exit Sub
        End If
        Set ws = Nothing
    Next sh
    
    Application.ScreenUpdating = True
    MsgBox "Workbook initialized successfully.", vbInformation
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    MsgBox "Initialize failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' 2. GenerateTrackerRows(voyageID or All)
'------------------------------------------------------------------------------
Public Sub GenerateTrackerRows(Optional voyageID As String = "ALL")
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim ws_voy As Worksheet, ws_doc As Worksheet, ws_tracker As Worksheet
    Dim lo_voy As ListObject, lo_doc As ListObject, lo_tracker As ListObject
    
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_doc = ThisWorkbook.Sheets("M_DocCatalog")
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_doc = ws_doc.ListObjects("tbl_DocCatalog")
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    
    Dim existing_keys As Object
    Set existing_keys = CreateObject("Scripting.Dictionary")
    
    If Not lo_tracker.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To lo_tracker.DataBodyRange.Rows.Count
            Dim v_id As String, d_code As String
            v_id = CStr(lo_tracker.DataBodyRange.Cells(r, 1).Value)
            d_code = CStr(lo_tracker.DataBodyRange.Cells(r, 2).Value)
            If v_id <> "" And d_code <> "" Then
                existing_keys(v_id & "|" & d_code) = True
            End If
        Next r
    End If
    
    Dim new_rows As Long: new_rows = 0
    Dim voy_row As ListRow, doc_row As ListRow
    
    For Each voy_row In lo_voy.ListRows
        Dim v_id_val As String
        v_id_val = CStr(voy_row.Range(1).Value)
        If voyageID <> "ALL" And v_id_val <> voyageID Then GoTo NextVoyage
        
        For Each doc_row In lo_doc.ListRows
            Dim d_code_val As String, req_flag As String, active_flag As String, default_party As String
            d_code_val = CStr(doc_row.Range(1).Value)
            req_flag = UCase(CStr(doc_row.Range(5).Value))
            active_flag = UCase(CStr(doc_row.Range(7).Value))
            default_party = CStr(doc_row.Range(4).Value)
            
            If req_flag <> "Y" Or active_flag <> "Y" Then GoTo NextDoc
            
            Dim key As String: key = v_id_val & "|" & d_code_val
            If existing_keys.Exists(key) Then GoTo NextDoc
            
            Dim new_row As ListRow
            Set new_row = lo_tracker.ListRows.Add
            new_row.Range(1).Value = v_id_val
            new_row.Range(2).Value = d_code_val
            new_row.Range(3).Value = default_party
            new_row.Range(8).Value = "Not Started"
            
            existing_keys(key) = True
            new_rows = new_rows + 1
NextDoc:
        Next doc_row
NextVoyage:
    Next voy_row
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Generated " & new_rows & " new tracker rows.", vbInformation
    Exit Sub
    
EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "GenerateTrackerRows failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' 3. RecalcDeadlines()
'------------------------------------------------------------------------------
Public Sub RecalcDeadlines()
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws_tracker As Worksheet, ws_voy As Worksheet, ws_rules As Worksheet
    Dim lo_tracker As ListObject, lo_voy As ListObject, lo_rules As ListObject
    
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_rules = ThisWorkbook.Sheets("R_DeadlineRules")
    
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_rules = ws_rules.ListObjects("tbl_RuleDeadline")
    
    If lo_tracker.DataBodyRange Is Nothing Then
        MsgBox "No tracker data found. Run GenerateTrackerRows first.", vbExclamation
        GoTo Cleanup
    End If
    
    Dim updated_count As Long: updated_count = 0
    Dim tr_row As ListRow
    
    For Each tr_row In lo_tracker.ListRows
        Dim vid As String, dcode As String
        vid = CStr(tr_row.Range(1).Value)
        dcode = CStr(tr_row.Range(2).Value)
        
        If vid = "" Or dcode = "" Then GoTo NextTrackerRow
        
        Dim best_rule As ListRow: Set best_rule = Nothing
        Dim best_priority As Long: best_priority = 9999
        Dim rule_row As ListRow
        
        For Each rule_row In lo_rules.ListRows
            Dim rule_doc As String, rule_active As String, rule_priority As Long
            rule_doc = CStr(rule_row.Range(2).Value)
            rule_active = UCase(CStr(rule_row.Range(7).Value))
            On Error Resume Next
            rule_priority = CLng(rule_row.Range(6).Value)
            On Error GoTo 0
            
            If rule_doc = dcode And rule_active = "Y" And rule_priority < best_priority Then
                Set best_rule = rule_row
                best_priority = rule_priority
            End If
        Next rule_row
        
        If best_rule Is Nothing Then GoTo NextTrackerRow
        
        Dim anchor_field As String, offset_days As Long, cal_type As String
        anchor_field = CStr(best_rule.Range(3).Value)
        offset_days = CLng(best_rule.Range(4).Value)
        cal_type = CStr(best_rule.Range(5).Value)
        
        tr_row.Range(4).Value = anchor_field
        tr_row.Range(6).Value = offset_days
        
        Dim voy_match As ListRow: Set voy_match = Nothing
        For Each voy_match In lo_voy.ListRows
            If CStr(voy_match.Range(1).Value) = vid Then Exit For
        Next voy_match
        
        If Not voy_match Is Nothing Then
            Dim anchor_col_idx As Long: anchor_col_idx = 0
            Dim col_idx As Long
            For col_idx = 1 To lo_voy.ListColumns.Count
                If lo_voy.ListColumns(col_idx).Name = anchor_field Then
                    anchor_col_idx = col_idx
                    Exit For
                End If
            Next col_idx
            
            If anchor_col_idx > 0 Then
                Dim anchor_date As Variant
                anchor_date = voy_match.Range(anchor_col_idx).Value
                If IsDate(anchor_date) Then
                    tr_row.Range(5).Value = anchor_date
                    
                    Dim due_date As Date
                    If UCase(cal_type) = "WD" Then
                        due_date = WorksheetFunction.WorkDay(CDate(anchor_date), offset_days)
                    Else
                        due_date = CDate(anchor_date) + offset_days
                    End If
                    tr_row.Range(7).Value = due_date
                    updated_count = updated_count + 1
                End If
            End If
        End If
NextTrackerRow:
    Next tr_row
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Updated " & updated_count & " deadline(s).", vbInformation
    Exit Sub
    
EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "RecalcDeadlines failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' 4. ValidateBeforeExport()
'------------------------------------------------------------------------------
Public Sub ValidateBeforeExport()
    On Error GoTo EH
    
    Dim errors As Object: Set errors = CreateObject("Scripting.Dictionary")
    Dim ws_tracker As Worksheet, ws_voy As Worksheet, ws_doc As Worksheet
    Dim lo_tracker As ListObject, lo_voy As ListObject, lo_doc As ListObject
    
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_doc = ThisWorkbook.Sheets("M_DocCatalog")
    
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_doc = ws_doc.ListObjects("tbl_DocCatalog")
    
    ' Check Voyage anchor fields
    Dim voy_row As ListRow
    For Each voy_row In lo_voy.ListRows
        Dim vid As String: vid = CStr(voy_row.Range(1).Value)
        If vid = "" Then GoTo NextVoyCheck
        
        Dim required_anchors As Variant
        required_anchors = Array(5, 6, 7, 8)
        
        Dim a As Variant
        For Each a In required_anchors
            If IsEmpty(voy_row.Range(CLng(a)).Value) Then
                errors.Add errors.Count + 1, "Voyage " & vid & ": Missing anchor column " & lo_voy.ListColumns(CLng(a)).Name
            End If
        Next a
NextVoyCheck:
    Next voy_row
    
    ' Check Tracker data
    If lo_tracker.DataBodyRange Is Nothing Then
        errors.Add errors.Count + 1, "T_Tracker has no data rows"
    Else
        Dim tr_row As ListRow
        For Each tr_row In lo_tracker.ListRows
            Dim status_val As String, evid_link As String, dcode_val As String
            status_val = CStr(tr_row.Range(8).Value)
            evid_link = CStr(tr_row.Range(11).Value)
            dcode_val = CStr(tr_row.Range(2).Value)
            
            If status_val = "Submitted" And evid_link = "" Then
                Dim doc_row As ListRow
                For Each doc_row In lo_doc.ListRows
                    If CStr(doc_row.Range(1).Value) = dcode_val Then
                        If UCase(CStr(doc_row.Range(6).Value)) = "Y" Then
                            errors.Add errors.Count + 1, "Tracker: " & tr_row.Range(1).Value & "|" & dcode_val & " - Submitted but EvidenceLink missing"
                        End If
                        Exit For
                    End If
                Next doc_row
            End If
            
            If IsEmpty(tr_row.Range(7).Value) Or tr_row.Range(7).Value = "" Then
                errors.Add errors.Count + 1, "Tracker: " & tr_row.Range(1).Value & "|" & tr_row.Range(2).Value & " - DueDate empty"
            End If
        Next tr_row
    End If
    
    ' Report
    If errors.Count > 0 Then
        Dim msg As String: msg = "Validation found " & errors.Count & " issue(s):" & vbCrLf & vbCrLf
        Dim k As Variant
        For Each k In errors.Keys
            msg = msg & errors(k) & vbCrLf
        Next k
        MsgBox msg, vbExclamation, "Validation Errors"
    Else
        MsgBox "Validation passed. No errors found.", vbInformation
    End If
    Exit Sub
    
EH:
    MsgBox "ValidateBeforeExport failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' 5. ExportVoyagePack(voyageID)
'------------------------------------------------------------------------------
Public Sub ExportVoyagePack(Optional voyageID As String = "")
    On Error GoTo EH
    
    If voyageID = "" Then
        voyageID = InputBox("Enter VoyageID to export (e.g., V01):", "Export Voyage Pack", "V01")
        If voyageID = "" Then Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim base_path As String: base_path = ThisWorkbook.Path
    If base_path = "" Then base_path = Environ("USERPROFILE") & "\Documents"
    
    Dim export_folder As String
    export_folder = base_path & "\TR_DocHub_Export_" & voyageID & "_" & Format(Now, "YYYYMMDD_HHMMSS")
    
    On Error Resume Next
    MkDir export_folder
    On Error GoTo EH
    
    Dim ws_dash As Worksheet: Set ws_dash = ThisWorkbook.Sheets("D_Dashboard")
    ws_dash.ExportAsFixedFormat Type:=xlTypePDF, _
                                 Filename:=export_folder & "\Dashboard_" & voyageID & ".pdf", _
                                 Quality:=xlQualityStandard, _
                                 IncludeDocProperties:=True, _
                                 IgnorePrintAreas:=False, _
                                 OpenAfterPublish:=False
    
    Application.ScreenUpdating = True
    MsgBox "Export Pack created:" & vbCrLf & export_folder, vbInformation
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    MsgBox "ExportVoyagePack failed: " & Err.Description, vbCritical
End Sub
