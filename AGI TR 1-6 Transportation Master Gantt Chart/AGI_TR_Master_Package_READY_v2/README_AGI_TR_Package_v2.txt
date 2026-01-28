AGI TR Master Package (AGI Site) - VBA Patched v2

Issue fixed
  - VBA error: "연속된 행이 너무 많습니다" (Too many line continuations)
  - Root cause: a single statement used too many line continuations (T = Array( _ ... ) over 24+ lines)
  - Fix: Scenario tasks are written using repeated AddTask() calls (no continuation overflow)

Quick start (manual import)
  1) Open AGI_TR_Master_RELEASE_v2.xlsx
  2) Save As -> Excel Macro-Enabled Workbook (*.xlsm)
  3) Alt+F11 -> VBA editor
  4) File -> Import File... -> select AGI_TR_Master_PATCHED_v2.bas
  5) In VBA editor, open "ThisWorkbook" and paste:

        Private Sub Workbook_Open()
            On Error Resume Next
            AGI_TR_Master.SetupKeyboardShortcuts
        End Sub

  6) Back to Excel -> Developer -> Macros -> Run 'AGI_TR_Master.RunAll'

Keyboard shortcuts
  Ctrl+Shift+R : RunAll
  Ctrl+Shift+U : UpdateAllDates
  Ctrl+Shift+O : FindOptimalD0
  Ctrl+Shift+B : GenerateDailyBriefing
  Ctrl+Shift+T : ShowTideInfo
  Ctrl+Shift+P : UpdateProgress
  Ctrl+Shift+E : ExportToCSV

Build script (optional)
  - build_agi_tr_master_release_v2.py can update D0 and (optionally) embed VBA into .xlsm
  - Embedding VBA requires Windows + Excel + pywin32 and Trust Center setting.
