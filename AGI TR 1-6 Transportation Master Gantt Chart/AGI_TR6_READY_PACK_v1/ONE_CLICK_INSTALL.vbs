' AGI_TR6 One-Click VBA Installer (Enhanced)
' - Opens base workbook (xlsx), imports VBA module, injects ThisWorkbook events, saves as xlsm
' Requirement: Excel "Trust access to the VBA project object model" enabled.
Option Explicit

Dim fso, basePath, wbXlsx, basFile, twFile, outXlsm, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
basePath = fso.GetParentFolderName(WScript.ScriptFullName)

' Normalize paths
wbXlsx   = fso.BuildPath(basePath, "AGI_TR6_VBA_Enhanced_AUTOMATION.xlsx")
basFile  = fso.BuildPath(basePath, "AGI_TR_AutomationPack.bas")
twFile   = fso.BuildPath(basePath, "ThisWorkbook_EventCode.txt")
outXlsm  = fso.BuildPath(basePath, "AGI_TR6_VBA_Enhanced_AUTOMATION_READY.xlsm")
logFile  = fso.BuildPath(basePath, "INSTALL_LOG.txt")

Sub LogLine(msg)
    Dim ts
    On Error Resume Next
    Set ts = fso.OpenTextFile(logFile, 8, True, 0) 'ForAppending
    ts.WriteLine Now & " | " & msg
    ts.Close
    On Error GoTo 0
End Sub

Function ReadAllText(path)
    Dim ts
    Set ts = fso.OpenTextFile(path, 1, False, 0)
    ReadAllText = ts.ReadAll
    ts.Close
End Function

' Pre-flight checks
On Error Resume Next
If Not fso.FileExists(wbXlsx) Then
    WScript.Echo "ERROR: Base workbook not found: " & wbXlsx & vbCrLf & _
                 "Please ensure AGI_TR6_VBA_Enhanced_AUTOMATION.xlsx exists in the same folder."
    WScript.Quit 2
End If
If Not fso.FileExists(basFile) Then
    WScript.Echo "ERROR: VBA module not found: " & basFile
    WScript.Quit 2
End If
If Not fso.FileExists(twFile) Then
    WScript.Echo "ERROR: ThisWorkbook event code not found: " & twFile
    WScript.Quit 2
End If
On Error GoTo 0

Dim xl, wb, vbproj, comps, c, code, tw, savedOk
savedOk = False

On Error GoTo EH
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

LogLine "=== AGI_TR6 Install Started ==="
LogLine "Excel started (version: " & xl.Version & ")"

Set wb = xl.Workbooks.Open(wbXlsx, ReadOnly:=False)
LogLine "Workbook opened: " & wb.Name

' Access VBProject (requires Trust Center setting)
On Error Resume Next
Set vbproj = wb.VBProject
If Err.Number <> 0 Then
    LogLine "ERROR: Cannot access VBProject. Trust Center setting required."
    WScript.Echo "ERROR: Cannot access VBA project." & vbCrLf & _
                 "Please enable: Excel > File > Options > Trust Center > " & _
                 "Trust Center Settings > Macro Settings > " & _
                 "[x] Trust access to the VBA project object model"
    GoTo EH
End If
On Error GoTo EH

Set comps = vbproj.VBComponents

' Remove previous standard module if exists (avoid duplicates)
Dim i, removedCount
removedCount = 0
For i = comps.Count To 1 Step -1
    Set c = comps.Item(i)
    If LCase(c.Type) = 1 Then ' vbext_ct_StdModule = 1
        If LCase(c.Name) = LCase("AGI_TR_AutomationPack") Or _
           LCase(c.Name) = LCase("Module1") Or _
           LCase(c.Name) = LCase("AGI_TR6") Then
            comps.Remove c
            LogLine "Removed existing module: " & c.Name
            removedCount = removedCount + 1
        End If
    End If
Next
If removedCount > 0 Then
    LogLine "Cleaned up " & removedCount & " existing module(s)"
End If

' Import BAS
comps.Import basFile
LogLine "Imported module: " & fso.GetFileName(basFile)

' Inject ThisWorkbook code
Set tw = comps.Item("ThisWorkbook").CodeModule
tw.DeleteLines 1, tw.CountOfLines
code = ReadAllText(twFile)
tw.AddFromString code
LogLine "Injected ThisWorkbook event code (" & tw.CountOfLines & " lines)"

' Save as XLSM
If fso.FileExists(outXlsm) Then
    fso.DeleteFile outXlsm, True
    LogLine "Deleted existing READY.xlsm"
End If
wb.SaveAs outXlsm, 52 ' xlOpenXMLWorkbookMacroEnabled = 52
LogLine "Saved as: " & fso.GetFileName(outXlsm)
savedOk = True

wb.Close False
xl.Quit

If savedOk Then
    LogLine "=== Install Completed Successfully ==="
    WScript.Echo "OK. Created: " & fso.GetFileName(outXlsm) & vbCrLf & _
                 "You can now open it and enable macros."
    WScript.Quit 0
Else
    WScript.Echo "Failed to create xlsm."
    WScript.Quit 3
End If

EH:
    LogLine "ERROR " & Err.Number & " : " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    WScript.Echo "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
                 "Tip: Excel > File > Options > Trust Center > " & _
                 "Trust Center Settings > Macro Settings > " & _
                 "[x] Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
                 "See INSTALL_LOG.txt for details."
    WScript.Quit 1
