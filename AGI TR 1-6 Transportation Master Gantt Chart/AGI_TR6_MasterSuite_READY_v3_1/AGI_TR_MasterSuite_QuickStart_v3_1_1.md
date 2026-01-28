# AGI TR MasterSuite (v3.1.1) — QuickStart

## Package contents
- **AGI_TR6_MasterSuite_READY_v3_1_1.xlsm** : Ready workbook (schedule+gantt prefilled).  
- **AGI_TR_MasterSuite_v3_1_1.bas** : VBA module (import into the workbook).  
- **build_agi_tr_mastersuite_v3_1_1.py** : Python builder (rebuild workbook).  
- **test_mastersuite_v3_1_1.py** : Python sanity tests.

## Immediate use (Excel)
1) Open **AGI_TR6_MasterSuite_READY_v3_1_1.xlsm**
2) Enable Editing + Enable Content (Macros)
3) Press **ALT+F11** (VBA editor) → **File → Import File...**
4) Import **AGI_TR_MasterSuite_v3_1_1.bas**
5) Back in Excel: Developer → Macros → run **SetupWorkbook** (once)
6) Run **SelfTest** (once) to confirm installation
7) Run **RunAll** to regenerate schedule/gantt based on Control_Panel inputs

### Shortcuts (after importing VBA)
- Ctrl+Shift+U : RunAll
- Ctrl+Shift+O : OptimizeD0
- Ctrl+Shift+M : Monte Carlo only
- Ctrl+Shift+R : Export PDF+CSV
- Ctrl+Shift+B : Daily briefing
- Ctrl+Shift+S : Freeze baseline
- Ctrl+Shift+D : Compare to baseline

## Important data notes
- **Tide_Data** is a *placeholder plan* (must be replaced with official tide table).
- **Weather_Risk** windows are editable; keep voyage buffer for conservative planning.

## Python (optional)
Rebuild workbook:
```bash
python build_agi_tr_mastersuite_v3_1_1.py --out AGI_TR6_MasterSuite_READY_v3_1_1.xlsx
```

Run sanity tests:
```bash
python test_mastersuite_v3_1_1.py AGI_TR6_MasterSuite_READY_v3_1_1.xlsx
```
