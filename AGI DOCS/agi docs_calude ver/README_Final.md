# TR_DocHub_AGI_2026 Final Package
## HVDC AGI TR Transportation - Document Tracker System

### 📦 Package Contents

```
TR_DocHub_AGI_2026_Final/
├── TR_DocHub_AGI_2026_Final_YYYYMMDD_HHMMSS.xlsx   # Main Excel Template
├── VBA_Modules/                                     # VBA Modules for Import
│   ├── modOperations.bas                           # Core operations (Init/Generate/Recalc/Export)
│   ├── modControlTower.bas                         # Single refresh entry point
│   ├── TR_DocTracker_Module.bas                    # Status formatting & utilities
│   ├── modDocGapMacros.bas                         # DocGap integration
│   └── ThisWorkbook.cls                            # Keyboard shortcuts & events
├── build_tr_dochub_final.py                        # Python builder script
└── README.md                                        # This file
```

---

### 🚀 Quick Start Guide

#### Step 1: Open Excel File
```
Open: TR_DocHub_AGI_2026_Final_YYYYMMDD_HHMMSS.xlsx
```

#### Step 2: Save as Macro-Enabled Workbook
```
File > Save As > Excel Macro-Enabled Workbook (*.xlsm)
```

#### Step 3: Import VBA Modules
```
1. Press Alt+F11 (Open VBA Editor)
2. File > Import File...
3. Import all .bas files from VBA_Modules folder:
   - modOperations.bas
   - modControlTower.bas
   - TR_DocTracker_Module.bas
   - modDocGapMacros.bas
4. Double-click "ThisWorkbook" in Project Explorer
5. Copy/paste contents of ThisWorkbook.cls
6. Save and close VBA Editor
```

#### Step 4: Initialize and Generate
```
1. Close and reopen the workbook
2. Run: InitializeWorkbook() (Alt+F8 > InitializeWorkbook > Run)
3. Run: GenerateTrackerRows() (Alt+F8 > GenerateTrackerRows > Run)
4. Run: RecalcDeadlines() (Alt+F8 > RecalcDeadlines > Run)
```

---

### ⌨️ Keyboard Shortcuts

| Shortcut | Function | Description |
|----------|----------|-------------|
| **Ctrl+Shift+R** | RefreshAll_ControlTower | Full refresh (formulas + formatting) |
| **Ctrl+Shift+P** | EXP_ExportToPDF | Export current sheet to PDF |
| **Ctrl+Shift+E** | TR_Draft_Reminder_Emails | Draft reminder emails (Outlook) |
| **Ctrl+Shift+G** | GenerateTrackerRows | Generate tracker rows from Voyages × Docs |
| **Ctrl+Shift+D** | RecalcDeadlines | Recalculate all deadlines |
| **Ctrl+Shift+V** | ValidateBeforeExport | Validate data before export |

---

### 📊 Sheet Structure

| Sheet | Purpose | Key Tables |
|-------|---------|------------|
| **D_Dashboard** | KPI Summary & Quick Actions | - |
| **S_Voyages** | Voyage Schedule (TR 1-7) | tbl_Voyage |
| **M_DocCatalog** | Document Requirements Master | tbl_DocCatalog |
| **M_Parties** | Responsible Party Master | tbl_Party |
| **R_DeadlineRules** | DueDate Calculation Rules | tbl_RuleDeadline |
| **T_Tracker** | Main Tracker (Transaction) | tbl_Tracker |
| **Holidays** | UAE/Project Holidays | - |
| **Party_Contacts** | Party Email Contacts | tbl_Contacts |
| **C_Config** | System Configuration | - |
| **LOG** | Execution Log | - |
| **VBA_Pasteboard** | VBA Code Reference | - |
| **Instructions** | Usage Guide | - |
| **Lists** | Dropdown Sources (Hidden) | - |

---

### 🔄 Workflow

```
1. S_Voyages: Enter voyage dates (MZP Arrival, Load-out, Departure, AGI Arrival)
                    ↓
2. M_DocCatalog: Define required documents (RequiredFlag=Y, ActiveFlag=Y)
                    ↓
3. R_DeadlineRules: Set DueDate rules (DocCode → AnchorField + OffsetDays)
                    ↓
4. GenerateTrackerRows(): Creates T_Tracker rows (Voyages × Docs)
                    ↓
5. RecalcDeadlines(): Calculates DueDates based on rules
                    ↓
6. T_Tracker: Update Status/SubmittedDate/AcceptedDate
                    ↓
7. D_Dashboard: Monitor KPIs (Overdue, DueSoon, Completion %)
                    ↓
8. ExportVoyagePack(): Export voyage-specific PDF/CSV
```

---

### 📋 DueDate Calculation Logic

```
DueDate = AnchorDate + OffsetDays

Where:
- AnchorField: MZP Arrival | Load-out | MZP Departure | AGI Arrival | Doc Deadline | Land Permit By
- OffsetDays: Positive (after) or Negative (before) days
- CalendarType: CAL (calendar days) or WD (working days via WORKDAY.INTL)
- Priority: Lower number = Higher priority (multiple rules per DocCode)
```

---

### ⚠️ Important Notes

1. **Composite Key**: VoyageID + DocCode = Unique (no duplicates in T_Tracker)
2. **Run GenerateTrackerRows** after adding new voyages or documents
3. **EvidenceLink**: Use file path or hyperlink format
4. **Holidays**: Add project-specific holidays for WORKDAY.INTL calculations
5. **Party_Contacts**: Configure email addresses for reminder drafts

---

### 🛠️ Troubleshooting

| Issue | Solution |
|-------|----------|
| Shortcuts not working | Close and reopen workbook after VBA import |
| GenerateTrackerRows fails | Check if tbl_Voyage and tbl_DocCatalog exist |
| DueDate empty | Verify R_DeadlineRules has matching DocCode with ActiveFlag=Y |
| Validation errors | Fix missing anchor dates in S_Voyages |

---

### 📌 Version Info

- **Version**: 1.0 Final
- **Date**: 2026-01-19
- **Project**: HVDC AGI TR Transportation
- **Excel**: 2021 LTSC / Microsoft 365
- **Python**: 3.11+

---

### 📞 Support

- Check `Instructions` sheet for usage guide
- Check `VBA_Pasteboard` sheet for VBA code reference
- Review `LOG` sheet for execution history

---

**© 2026 Samsung C&T - HVDC Project Team**
