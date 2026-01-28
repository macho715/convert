Attribute VB_Name = "modOperations"
Option Explicit

'===============================================================================
' OPERATIONS MODULE - Minimum Essential Functions for TR_DocHub_AGI
'===============================================================================

'------------------------------------------------------------------------------
' 1. InitializeWorkbook()
'------------------------------------------------------------------------------
Public Sub InitializeWorkbook()
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    
    ' 테이블 존재 검증
    Dim ws As Worksheet
    Dim required_sheets As Variant
    required_sheets = Array("S_Voyages", "M_DocCatalog", "M_Parties", "R_DeadlineRules", "T_Tracker", "D_Dashboard")
    
    Dim sh As Variant
    For Each sh In required_sheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sh)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Required sheet missing: " & sh, vbCritical
            Exit Sub
        End If
    Next sh
    
    ' Named Ranges 재생성 (필요 시)
    ' 드롭다운/조건부서식은 Python 빌더에서 이미 설정됨
    
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
    
    Dim ws_voy As Worksheet
    Dim ws_doc As Worksheet
    Dim ws_tracker As Worksheet
    Dim lo_voy As ListObject
    Dim lo_doc As ListObject
    Dim lo_tracker As ListObject
    
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_doc = ThisWorkbook.Sheets("M_DocCatalog")
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_doc = ws_doc.ListObjects("tbl_DocCatalog")
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    
    ' Collect existing keys to avoid duplicates
    Dim existing_keys As Object
    Set existing_keys = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To lo_tracker.DataBodyRange.Rows.Count
        Dim v_id As String, d_code As String
        v_id = lo_tracker.DataBodyRange.Cells(r, 1).Value
        d_code = lo_tracker.DataBodyRange.Cells(r, 2).Value
        If v_id <> "" And d_code <> "" Then
            existing_keys(v_id & "|" & d_code) = True
        End If
    Next r
    
    ' Generate rows: Voyages × Docs (Active+Required)
    Dim new_rows As Long
    new_rows = 0
    
    Dim voy_row As ListRow
    Dim doc_row As ListRow
    
    For Each voy_row In lo_voy.ListRows
        Dim v_id_val As String
        v_id_val = CStr(voy_row.Range(1).Value)  ' VoyageID
        
        If voyageID <> "ALL" And v_id_val <> voyageID Then
            GoTo NextVoyage
        End If
        
        For Each doc_row In lo_doc.ListRows
            Dim d_code_val As String
            Dim req_flag As String
            Dim active_flag As String
            Dim default_party As String
            
            d_code_val = CStr(doc_row.Range(1).Value)  ' DocCode
            req_flag = UCase(CStr(doc_row.Range(5).Value))  ' RequiredFlag
            active_flag = UCase(CStr(doc_row.Range(7).Value))  ' ActiveFlag
            default_party = CStr(doc_row.Range(4).Value)  ' DefaultResponsiblePartyID
            
            ' Only Active + Required docs
            If req_flag <> "Y" Or active_flag <> "Y" Then
                GoTo NextDoc
            End If
            
            ' Check duplicate
            Dim key As String
            key = v_id_val & "|" & d_code_val
            If existing_keys.Exists(key) Then
                GoTo NextDoc
            End If
            
            ' Create new row
            Dim new_row As ListRow
            Set new_row = lo_tracker.ListRows.Add
            new_row.Range(1).Value = v_id_val  ' VoyageID
            new_row.Range(2).Value = d_code_val  ' DocCode
            new_row.Range(3).Value = default_party  ' ResponsiblePartyID
            new_row.Range(8).Value = "Not Started"  ' Status
            
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
    
    Dim ws_tracker As Worksheet
    Dim ws_voy As Worksheet
    Dim ws_rules As Worksheet
    Dim lo_tracker As ListObject
    Dim lo_voy As ListObject
    Dim lo_rules As ListObject
    
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_rules = ThisWorkbook.Sheets("R_DeadlineRules")
    
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_rules = ws_rules.ListObjects("tbl_RuleDeadline")
    
    Dim updated_count As Long
    updated_count = 0
    
    Dim tr_row As ListRow
    For Each tr_row In lo_tracker.ListRows
        Dim v_id As String, d_code As String
        v_id = CStr(tr_row.Range(1).Value)  ' VoyageID
        d_code = CStr(tr_row.Range(2).Value)  ' DocCode
        
        If v_id = "" Or d_code = "" Then
            GoTo NextTrackerRow
        End If
        
        ' Find matching rule (DocCode, ActiveFlag=Y, Priority 최소)
        Dim best_rule As ListRow
        Set best_rule = Nothing
        Dim best_priority As Long
        best_priority = 9999
        
        Dim rule_row As ListRow
        For Each rule_row In lo_rules.ListRows
            Dim rule_doc As String, rule_active As String, rule_priority As Long
            rule_doc = CStr(rule_row.Range(2).Value)  ' DocCode
            rule_active = UCase(CStr(rule_row.Range(7).Value))  ' ActiveFlag
            rule_priority = CLng(rule_row.Range(6).Value)  ' Priority
            
            If rule_doc = d_code And rule_active = "Y" And rule_priority < best_priority Then
                Set best_rule = rule_row
                best_priority = rule_priority
            End If
        Next rule_row
        
        If best_rule Is Nothing Then
            GoTo NextTrackerRow
        End If
        
        ' Get rule values
        Dim anchor_field As String, offset_days As Long, cal_type As String
        anchor_field = CStr(best_rule.Range(3).Value)  ' AnchorField
        offset_days = CLng(best_rule.Range(4).Value)  ' OffsetDays
        cal_type = CStr(best_rule.Range(5).Value)  ' CalendarType
        
        ' Find Voyage row
        Dim voy_row As ListRow
        Set voy_row = Nothing
        For Each voy_row In lo_voy.ListRows
            If CStr(voy_row.Range(1).Value) = v_id Then
                Exit For
            End If
        Next voy_row
        
        If voy_row Is Nothing Then
            GoTo NextTrackerRow
        End If
        
        ' Get AnchorDate from Voyage (find column by header)
        Dim anchor_date As Date
        anchor_date = 0
        Dim col_idx As Long
        For col_idx = 1 To lo_voy.ListColumns.Count
            If lo_voy.ListColumns(col_idx).Name = anchor_field Then
                Dim anchor_val As Variant
                anchor_val = voy_row.Range(col_idx).Value
                If IsDate(anchor_val) Then
                    anchor_date = CDate(anchor_val)
                End If
                Exit For
            End If
        Next col_idx
        
        If anchor_date = 0 Then
            GoTo NextTrackerRow
        End If
        
        ' Calculate DueDate
        Dim due_date As Date
        If cal_type = "WD" Then
            ' WORKDAY.INTL (simplified - assumes weekend pattern "0000011")
            due_date = Application.WorksheetFunction.WorkDay_Intl(anchor_date, offset_days, "0000011")
        Else
            due_date = anchor_date + offset_days
        End If
        
        ' Update tracker row
        tr_row.Range(4).Value = anchor_field  ' AnchorField
        tr_row.Range(5).Value = anchor_date  ' AnchorDate
        tr_row.Range(6).Value = offset_days  ' OffsetDays
        tr_row.Range(7).Value = due_date  ' DueDate
        
        ' Calculate RAG
        Dim rag As String
        If due_date < Date Then
            rag = "Overdue"
        ElseIf due_date <= Date + 7 Then
            rag = "DueSoon"
        Else
            rag = "OK"
        End If
        tr_row.Range(15).Value = rag  ' RAG
        
        updated_count = updated_count + 1
        
NextTrackerRow:
    Next tr_row
    
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Application.ScreenUpdating = True
    
    MsgBox "Recalculated " & updated_count & " deadline(s).", vbInformation
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
    
    Dim errors As Object
    Set errors = CreateObject("Scripting.Dictionary")
    
    Dim ws_tracker As Worksheet
    Dim ws_voy As Worksheet
    Dim ws_doc As Worksheet
    Dim lo_tracker As ListObject
    Dim lo_voy As ListObject
    Dim lo_doc As ListObject
    
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    Set ws_voy = ThisWorkbook.Sheets("S_Voyages")
    Set ws_doc = ThisWorkbook.Sheets("M_DocCatalog")
    
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    Set lo_voy = ws_voy.ListObjects("tbl_Voyage")
    Set lo_doc = ws_doc.ListObjects("tbl_DocCatalog")
    
    ' Check 1: Anchor 필드 누락
    Dim voy_row As ListRow
    For Each voy_row In lo_voy.ListRows
        Dim v_id As String
        v_id = CStr(voy_row.Range(1).Value)
        If v_id = "" Then GoTo NextVoyCheck
        
        ' Check required anchors
        Dim required_anchors As Variant
        required_anchors = Array("MZP Arrival", "Load-out", "MZP Departure", "AGI Arrival")
        
        Dim anchor As Variant
        For Each anchor In required_anchors
            Dim col_idx As Long
            For col_idx = 1 To lo_voy.ListColumns.Count
                If lo_voy.ListColumns(col_idx).Name = anchor Then
                    If IsEmpty(voy_row.Range(col_idx).Value) Then
                        errors.Add errors.Count + 1, "Voyage " & v_id & ": Missing " & anchor
                    End If
                    Exit For
                End If
            Next col_idx
        Next anchor
        
NextVoyCheck:
    Next voy_row
    
    ' Check 2: Status=Submitted AND EvidenceRequiredFlag=Y AND EvidenceLink=""
    Dim tr_row As ListRow
    For Each tr_row In lo_tracker.ListRows
        Dim status_val As String, evid_link As String, d_code_val As String
        status_val = CStr(tr_row.Range(8).Value)  ' Status
        evid_link = CStr(tr_row.Range(11).Value)  ' EvidenceLink
        d_code_val = CStr(tr_row.Range(2).Value)  ' DocCode
        
        If status_val = "Submitted" And evid_link = "" Then
            ' Check EvidenceRequiredFlag
            Dim doc_row As ListRow
            For Each doc_row In lo_doc.ListRows
                If CStr(doc_row.Range(1).Value) = d_code_val Then
                    Dim evid_req As String
                    evid_req = UCase(CStr(doc_row.Range(6).Value))  ' EvidenceRequiredFlag
                    If evid_req = "Y" Then
                        errors.Add errors.Count + 1, "Tracker: " & tr_row.Range(1).Value & "|" & d_code_val & " - Submitted but EvidenceLink missing"
                    End If
                    Exit For
                End If
            Next doc_row
        End If
    Next tr_row
    
    ' Check 3: DueDate="" (룰 매칭 실패)
    For Each tr_row In lo_tracker.ListRows
        Dim due_date_val As Variant
        due_date_val = tr_row.Range(7).Value  ' DueDate
        If IsEmpty(due_date_val) Or due_date_val = "" Then
            errors.Add errors.Count + 1, "Tracker: " & tr_row.Range(1).Value & "|" & tr_row.Range(2).Value & " - DueDate empty (rule matching failed)"
        End If
    Next tr_row
    
    ' Report errors
    If errors.Count > 0 Then
        Dim msg As String
        msg = "Validation found " & errors.Count & " issue(s):" & vbCrLf & vbCrLf
        Dim key As Variant
        For Each key In errors.Keys
            msg = msg & errors(key) & vbCrLf
        Next key
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
    
    Dim base_path As String
    base_path = ThisWorkbook.Path
    If base_path = "" Then base_path = Environ("USERPROFILE") & "\Documents"
    
    Dim export_folder As String
    export_folder = base_path & "\TR_DocHub_Export_" & voyageID & "_" & Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Create folder
    On Error Resume Next
    MkDir export_folder
    On Error GoTo EH
    
    ' 1. Dashboard PDF
    Dim ws_dash As Worksheet
    Set ws_dash = ThisWorkbook.Sheets("D_Dashboard")
    ws_dash.ExportAsFixedFormat Type:=xlTypePDF, _
                                 fileName:=export_folder & "\Dashboard_" & voyageID & ".pdf", _
                                 Quality:=xlQualityStandard, _
                                 IncludeDocProperties:=True, _
                                 IgnorePrintAreas:=False, _
                                 OpenAfterPublish:=False
    
    ' 2. Tracker CSV (filtered by VoyageID)
    ' Note: CSV export requires creating a temporary workbook
    Dim wb_temp As Workbook
    Set wb_temp = Workbooks.Add
    Dim ws_temp As Worksheet
    Set ws_temp = wb_temp.Sheets(1)
    
    ' Copy filtered Tracker data
    Dim ws_tracker As Worksheet
    Dim lo_tracker As ListObject
    Set ws_tracker = ThisWorkbook.Sheets("T_Tracker")
    Set lo_tracker = ws_tracker.ListObjects("tbl_Tracker")
    
    ' Copy headers
    lo_tracker.HeaderRowRange.Copy
    ws_temp.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ' Copy data (filtered by VoyageID)
    Dim tr_row As ListRow
    Dim dest_row As Long
    dest_row = 2
    For Each tr_row In lo_tracker.ListRows
        If CStr(tr_row.Range(1).Value) = voyageID Then
            tr_row.Range.Copy
            ws_temp.Range("A" & dest_row).PasteSpecial Paste:=xlPasteValues
            dest_row = dest_row + 1
        End If
    Next tr_row
    
    Application.CutCopyMode = False
    
    ' Save as CSV
    wb_temp.SaveAs fileName:=export_folder & "\Tracker_" & voyageID & ".csv", _
                    FileFormat:=xlCSV, _
                    CreateBackup:=False
    wb_temp.Close SaveChanges:=False
    
    ' 3. Evidence Index CSV
    Set wb_temp = Workbooks.Add
    Set ws_temp = wb_temp.Sheets(1)
    ws_temp.Range("A1").Value = "VoyageID"
    ws_temp.Range("B1").Value = "DocCode"
    ws_temp.Range("C1").Value = "EvidenceLink"
    ws_temp.Range("D1").Value = "EvidenceNote"
    
    dest_row = 2
    For Each tr_row In lo_tracker.ListRows
        If CStr(tr_row.Range(1).Value) = voyageID Then
            Dim evid_link_val As String
            evid_link_val = CStr(tr_row.Range(11).Value)  ' EvidenceLink
            If evid_link_val <> "" Then
                ws_temp.Range("A" & dest_row).Value = tr_row.Range(1).Value
                ws_temp.Range("B" & dest_row).Value = tr_row.Range(2).Value
                ws_temp.Range("C" & dest_row).Value = evid_link_val
                ws_temp.Range("D" & dest_row).Value = tr_row.Range(12).Value
                dest_row = dest_row + 1
            End If
        End If
    Next tr_row
    
    wb_temp.SaveAs fileName:=export_folder & "\EvidenceIndex_" & voyageID & ".csv", _
                    FileFormat:=xlCSV, _
                    CreateBackup:=False
    wb_temp.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "Export Pack created:" & vbCrLf & export_folder, vbInformation
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    If Not wb_temp Is Nothing Then
        On Error Resume Next
        wb_temp.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "ExportVoyagePack failed: " & Err.Description, vbCritical
End Sub
