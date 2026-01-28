#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGI TR 1-6 Transportation Master Gantt Chart Generator
6 Voyages (1 TR per voyage), 3 Jack-down Events (when 2 TRs arrive at AGI)
Weather-integrated schedule with Winter Shamal risk assessment
Target: Complete before March 2026
"""

import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import Rule, FormulaRule
from openpyxl.styles.differential import DifferentialStyle

# === COLOR SCHEME ===
COLORS = {
    "HEADER": "1F4E79",
    "MOBILIZATION": "8E7CC3",
    "DECK_PREP": "6FA8DC",
    "LOADOUT": "93C47D",
    "SEAFAST": "76A5AF",
    "SAIL": "A4C2F4",
    "AGI_UNLOAD": "F6B26B",
    "TURNING": "FFD966",
    "JACKDOWN": "E06666",
    "RETURN": "999999",
    "BUFFER": "D9D9D9",
    "WEATHER_RISK": "FF9900",
    "MILESTONE": "FF0000",
    "WEEKEND": "F2F2F2",
    "SHAMAL_RISK": "FFE0B2",
}

BORDER_THIN = Side(style="thin", color="A6A6A6")

def thin_border():
    return Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)

def create_gantt_workbook():
    wb = Workbook()
    
    # === SHEET 1: SCHEDULE DATA ===
    ws_sched = wb.active
    ws_sched.title = "Schedule_Data"
    
    # Project start date
    project_start = dt.date(2026, 1, 6)
    
    # Define tasks with durations (in days)
    # Pattern: Each voyage cycle
    tasks = []
    
    # === MOBILIZATION (One-time, Day 1) ===
    tasks.append({
        "id": "MOB-001", "wbs": "1.0", "task": "MOBILIZATION", 
        "phase": "MOBILIZATION", "owner": "Mammoet",
        "start": project_start, "duration": 1,
        "notes": "SPMT Assembly + Marine Equipment Mobilization"
    })
    tasks.append({
        "id": "PREP-001", "wbs": "1.1", "task": "Deck Preparations (D-ring, steel sets, welding)",
        "phase": "DECK_PREP", "owner": "Mammoet",
        "start": project_start + dt.timedelta(days=1), "duration": 2,
        "notes": "One-time setup for all voyages"
    })
    
    # === VOYAGE CYCLES ===
    # Key timing parameters per voyage (days)
    T_LOADOUT = 1.0        # Load TR onto SPMT + RoRo load-out
    T_SEAFAST = 0.5        # Sea fastening
    T_MWS = 0.5            # MWS + MPI + Final prep
    T_SAIL_TO_AGI = 1.0    # Marine transportation MZP → AGI
    T_UNLOAD = 1.0         # RoRo load-in at AGI
    T_TURNING = 3.0        # Turning of TR at bay area
    T_JACKDOWN = 1.0       # Jacking down on temporary support
    T_RETURN = 1.0         # LCT return to MZP
    T_RESET = 0.5          # Reset equipment for next voyage
    T_BUFFER = 0.5         # Buffer for weather/delays
    
    # Voyage cycle total (excluding jack-down for odd trips)
    T_VOYAGE_BASIC = T_LOADOUT + T_SEAFAST + T_MWS + T_SAIL_TO_AGI + T_UNLOAD + T_RETURN + T_RESET
    
    # Weather risk periods (Winter Shamal)
    shamal_risk_start = dt.date(2026, 1, 14)
    shamal_risk_end = dt.date(2026, 1, 21)
    
    # Voyage start dates (optimized to avoid peak Shamal)
    # Strategy: Complete V1+V2 before Shamal peak, V3-V6 after
    voyage_starts = [
        dt.date(2026, 1, 9),    # V1: TR1
        dt.date(2026, 1, 13),   # V2: TR2 + JD1 (tight, before peak)
        dt.date(2026, 1, 22),   # V3: TR3 (after Shamal)
        dt.date(2026, 1, 27),   # V4: TR4 + JD2
        dt.date(2026, 2, 2),    # V5: TR5
        dt.date(2026, 2, 7),    # V6: TR6 + JD3
    ]
    
    task_id = 100
    for v in range(6):
        voyage_num = v + 1
        tr_num = v + 1
        vs = voyage_starts[v]
        is_jackdown_voyage = (voyage_num % 2 == 0)  # V2, V4, V6
        
        # Weather risk flag
        in_shamal = vs >= shamal_risk_start and vs <= shamal_risk_end
        weather_note = "⚠️ Winter Shamal Risk Zone" if in_shamal else "✅ Good Weather Window"
        
        # Voyage Summary
        tasks.append({
            "id": f"V{voyage_num}", "wbs": f"{voyage_num+1}.0", 
            "task": f"VOYAGE {voyage_num}: TR{tr_num} Transport{'+ Jack-down' if is_jackdown_voyage else ''}",
            "phase": "MILESTONE", "owner": "All",
            "start": vs, "duration": 0,
            "notes": weather_note
        })
        
        # Load-out at MZP
        task_id += 1
        lo_start = vs
        tasks.append({
            "id": f"LO-{task_id}", "wbs": f"{voyage_num+1}.1",
            "task": f"TR{tr_num} Load-out on SPMT + RoRo to LCT",
            "phase": "LOADOUT", "owner": "Mammoet",
            "start": lo_start, "duration": T_LOADOUT,
            "notes": f"High tide required - Check Tide Table"
        })
        
        # Sea Fastening
        task_id += 1
        sf_start = lo_start + dt.timedelta(days=T_LOADOUT)
        tasks.append({
            "id": f"SF-{task_id}", "wbs": f"{voyage_num+1}.2",
            "task": f"TR{tr_num} Sea Fastening + Lashing",
            "phase": "SEAFAST", "owner": "Mammoet",
            "start": sf_start, "duration": T_SEAFAST,
            "notes": "12-point lashing per DAS standard"
        })
        
        # MWS Approval
        task_id += 1
        mws_start = sf_start + dt.timedelta(days=T_SEAFAST)
        tasks.append({
            "id": f"MWS-{task_id}", "wbs": f"{voyage_num+1}.3",
            "task": f"MWS + MPI + Final Preparations",
            "phase": "BUFFER", "owner": "Aries/Captain",
            "start": mws_start, "duration": T_MWS,
            "notes": "Marine Warranty Surveyor approval"
        })
        
        # Sail to AGI
        task_id += 1
        sail_start = mws_start + dt.timedelta(days=T_MWS)
        tasks.append({
            "id": f"SAIL-{task_id}", "wbs": f"{voyage_num+1}.4",
            "task": f"V{voyage_num} Sail-away: MZP → AGI Site",
            "phase": "SAIL", "owner": "LCT Bushra",
            "start": sail_start, "duration": T_SAIL_TO_AGI,
            "notes": weather_note
        })
        
        # AGI Arrival + Unload
        task_id += 1
        arr_start = sail_start + dt.timedelta(days=T_SAIL_TO_AGI)
        tasks.append({
            "id": f"ARR-{task_id}", "wbs": f"{voyage_num+1}.5",
            "task": f"AGI Arrival + TR{tr_num} RoRo Load-in",
            "phase": "AGI_UNLOAD", "owner": "Mammoet",
            "start": arr_start, "duration": T_UNLOAD,
            "notes": f"AGI FWD Draft ≤ 2.70m required"
        })
        
        # Store on Jetty (for odd voyages) or immediate transport
        unload_end = arr_start + dt.timedelta(days=T_UNLOAD)
        
        if is_jackdown_voyage:
            # Transport both TRs to bay + Turning + Jack-down
            prev_tr = tr_num - 1
            
            # Transport TR(n-1) to bay
            task_id += 1
            tasks.append({
                "id": f"TRNS-{task_id}", "wbs": f"{voyage_num+1}.6",
                "task": f"TR{prev_tr} Load on SPMT + Transport to Bay {(voyage_num//2)+2}",
                "phase": "TURNING", "owner": "Mammoet",
                "start": unload_end, "duration": 1.0,
                "notes": "Steel bridge installation + SPMT transport"
            })
            
            # Turning TR(n-1)
            task_id += 1
            turn1_start = unload_end + dt.timedelta(days=1)
            tasks.append({
                "id": f"TURN-{task_id}", "wbs": f"{voyage_num+1}.7",
                "task": f"TR{prev_tr} Turning (90°) - 10t Forklift",
                "phase": "TURNING", "owner": "Mammoet",
                "start": turn1_start, "duration": T_TURNING,
                "notes": "Precision turning at TR Bay"
            })
            
            # Jack-down TR(n-1)
            task_id += 1
            jd1_start = turn1_start + dt.timedelta(days=T_TURNING)
            tasks.append({
                "id": f"JD-{task_id}", "wbs": f"{voyage_num+1}.8",
                "task": f"★ TR{prev_tr} Jack-Down on Temporary Support",
                "phase": "JACKDOWN", "owner": "Mammoet",
                "start": jd1_start, "duration": T_JACKDOWN,
                "notes": f"MILESTONE: TR{prev_tr} Installation Complete"
            })
            
            # Transport TR(n) to bay
            task_id += 1
            tasks.append({
                "id": f"TRNS-{task_id}", "wbs": f"{voyage_num+1}.9",
                "task": f"TR{tr_num} Load on SPMT + Transport to Bay {(voyage_num//2)+3}",
                "phase": "TURNING", "owner": "Mammoet",
                "start": jd1_start, "duration": 1.0,
                "notes": "Parallel operation"
            })
            
            # Turning TR(n)
            task_id += 1
            turn2_start = jd1_start + dt.timedelta(days=T_JACKDOWN)
            tasks.append({
                "id": f"TURN-{task_id}", "wbs": f"{voyage_num+1}.10",
                "task": f"TR{tr_num} Turning (90°) - 10t Forklift",
                "phase": "TURNING", "owner": "Mammoet",
                "start": turn2_start, "duration": T_TURNING,
                "notes": "Precision turning at TR Bay"
            })
            
            # Jack-down TR(n)
            task_id += 1
            jd2_start = turn2_start + dt.timedelta(days=T_TURNING)
            tasks.append({
                "id": f"JD-{task_id}", "wbs": f"{voyage_num+1}.11",
                "task": f"★ TR{tr_num} Jack-Down on Temporary Support",
                "phase": "JACKDOWN", "owner": "Mammoet",
                "start": jd2_start, "duration": T_JACKDOWN,
                "notes": f"MILESTONE: TR{tr_num} Installation Complete"
            })
            
            # Return voyage
            task_id += 1
            ret_start = jd2_start + dt.timedelta(days=T_JACKDOWN)
            tasks.append({
                "id": f"RET-{task_id}", "wbs": f"{voyage_num+1}.12",
                "task": f"V{voyage_num} LCT Return: AGI → MZP",
                "phase": "RETURN", "owner": "LCT Bushra",
                "start": ret_start, "duration": T_RETURN,
                "notes": "SPMT demob if last voyage"
            })
        else:
            # Odd voyage - just store on jetty and return
            task_id += 1
            tasks.append({
                "id": f"STORE-{task_id}", "wbs": f"{voyage_num+1}.6",
                "task": f"TR{tr_num} Stored on AGI Jetty (await V{voyage_num+1})",
                "phase": "BUFFER", "owner": "Mammoet",
                "start": unload_end, "duration": 0.5,
                "notes": "Awaiting pair TR for joint jack-down"
            })
            
            # Return voyage
            task_id += 1
            ret_start = unload_end + dt.timedelta(days=0.5)
            tasks.append({
                "id": f"RET-{task_id}", "wbs": f"{voyage_num+1}.7",
                "task": f"V{voyage_num} LCT Return: AGI → MZP",
                "phase": "RETURN", "owner": "LCT Bushra",
                "start": ret_start, "duration": T_RETURN,
                "notes": "Quick turnaround for next TR"
            })
        
        # Buffer / Reset
        task_id += 1
        last_task = tasks[-1]
        buf_start = last_task["start"] + dt.timedelta(days=last_task["duration"])
        tasks.append({
            "id": f"BUF-{task_id}", "wbs": f"{voyage_num+1}.99",
            "task": f"V{voyage_num} Buffer / Equipment Reset",
            "phase": "BUFFER", "owner": "All",
            "start": buf_start, "duration": T_BUFFER,
            "notes": "Weather contingency + equipment check"
        })
    
    # === PROJECT COMPLETION MILESTONE ===
    last_task = tasks[-1]
    completion_date = last_task["start"] + dt.timedelta(days=last_task["duration"] + 1)
    tasks.append({
        "id": "COMP", "wbs": "99.0",
        "task": "★★★ PROJECT COMPLETE - All 6 TRs Installed ★★★",
        "phase": "MILESTONE", "owner": "All",
        "start": completion_date, "duration": 0,
        "notes": f"Target: Before March 2026 | Actual: {completion_date.strftime('%Y-%m-%d')}"
    })
    
    # === WRITE SCHEDULE DATA ===
    headers = ["ID", "WBS", "Task", "Phase", "Owner", "Start", "End", "Duration", "Notes"]
    
    # Title rows
    ws_sched.merge_cells("A1:I1")
    ws_sched["A1"] = "AGI HVDC TR 1-6 Transportation Master Schedule"
    ws_sched["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws_sched["A1"].fill = PatternFill("solid", fgColor=COLORS["HEADER"])
    ws_sched["A1"].alignment = Alignment(horizontal="center")
    
    ws_sched.merge_cells("A2:I2")
    ws_sched["A2"] = f"6 Voyages | 3 Jack-down Events | Target: Before March 2026 | Generated: {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_sched["A2"].fill = PatternFill("solid", fgColor=COLORS["HEADER"])
    ws_sched["A2"].font = Font(color="FFFFFF")
    
    ws_sched.merge_cells("A3:I3")
    ws_sched["A3"] = "⚠️ Winter Shamal Risk: Jan 14-21, 2026 | Schedule optimized to minimize weather exposure"
    ws_sched["A3"].fill = PatternFill("solid", fgColor=COLORS["SHAMAL_RISK"])
    ws_sched["A3"].font = Font(bold=True)
    
    # Headers
    header_row = 5
    for col, h in enumerate(headers, 1):
        cell = ws_sched.cell(header_row, col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COLORS["HEADER"])
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border()
    
    # Data rows
    for row_idx, t in enumerate(tasks, header_row + 1):
        end_date = t["start"] + dt.timedelta(days=t["duration"]) if t["duration"] > 0 else t["start"]
        row_data = [
            t["id"], t["wbs"], t["task"], t["phase"], t["owner"],
            t["start"], end_date, t["duration"], t["notes"]
        ]
        
        phase_color = COLORS.get(t["phase"], "FFFFFF")
        
        for col, val in enumerate(row_data, 1):
            cell = ws_sched.cell(row_idx, col, value=val)
            cell.border = thin_border()
            
            if col in (6, 7):  # Dates
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center")
            elif col == 8:  # Duration
                cell.number_format = "0.0"
                cell.alignment = Alignment(horizontal="center")
            
            # Color coding
            if t["phase"] == "MILESTONE":
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=COLORS["MILESTONE"])
            elif col == 3:  # Task column
                cell.fill = PatternFill("solid", fgColor=phase_color)
    
    # Column widths
    col_widths = {"A": 12, "B": 8, "C": 55, "D": 14, "E": 15, "F": 12, "G": 12, "H": 10, "I": 45}
    for col, w in col_widths.items():
        ws_sched.column_dimensions[col].width = w
    
    # Freeze panes
    ws_sched.freeze_panes = "A6"
    
    # === SHEET 2: GANTT CHART ===
    ws_gantt = wb.create_sheet("Gantt_Chart")
    
    # Determine date range
    all_starts = [t["start"] for t in tasks]
    all_ends = [t["start"] + dt.timedelta(days=max(t["duration"], 1)) for t in tasks]
    min_date = min(all_starts) - dt.timedelta(days=1)
    max_date = max(all_ends) + dt.timedelta(days=2)
    
    days = []
    cur = min_date
    while cur <= max_date:
        days.append(cur)
        cur += dt.timedelta(days=1)
    
    # Metadata columns
    META_COLS = ["ID", "WBS", "Task", "Phase", "Start", "End", "Dur"]
    gantt_start_col = len(META_COLS) + 1
    
    # Title
    ws_gantt.merge_cells(start_row=1, start_column=1, end_row=1, end_column=gantt_start_col + len(days) - 1)
    ws_gantt["A1"] = "AGI HVDC TR 1-6 Master Gantt Chart (Jan-Feb 2026)"
    ws_gantt["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws_gantt["A1"].fill = PatternFill("solid", fgColor=COLORS["HEADER"])
    
    ws_gantt.merge_cells(start_row=2, start_column=1, end_row=2, end_column=gantt_start_col + len(days) - 1)
    ws_gantt["A2"] = "Orange Zone = Winter Shamal Risk Period (Jan 14-21) | Red = Jack-down Milestones | Target: Complete before Mar 2026"
    ws_gantt["A2"].fill = PatternFill("solid", fgColor=COLORS["SHAMAL_RISK"])
    ws_gantt["A2"].font = Font(bold=True)
    
    # Header row
    header_row = 4
    for col, h in enumerate(META_COLS, 1):
        cell = ws_gantt.cell(header_row, col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COLORS["HEADER"])
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border()
    
    # Date headers
    for j, d in enumerate(days):
        col = gantt_start_col + j
        cell = ws_gantt.cell(header_row, col)
        cell.value = d.day
        cell.font = Font(bold=True, size=9, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=COLORS["HEADER"])
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border()
        
        # Weekend shading
        if d.weekday() >= 5:
            cell.fill = PatternFill("solid", fgColor=COLORS["WEEKEND"])
            cell.font = Font(bold=True, size=9, color="666666")
        
        # Shamal risk period
        if shamal_risk_start <= d <= shamal_risk_end:
            cell.fill = PatternFill("solid", fgColor=COLORS["WEATHER_RISK"])
            cell.font = Font(bold=True, size=9)
    
    # Month row
    ws_gantt.cell(3, gantt_start_col).value = "Jan 2026"
    ws_gantt.cell(3, gantt_start_col).font = Font(bold=True)
    
    # Find Feb start
    for j, d in enumerate(days):
        if d.month == 2 and d.day == 1:
            ws_gantt.cell(3, gantt_start_col + j).value = "Feb 2026"
            ws_gantt.cell(3, gantt_start_col + j).font = Font(bold=True)
            break
    
    # Data rows
    first_data_row = header_row + 1
    for row_idx, t in enumerate(tasks, first_data_row):
        end_date = t["start"] + dt.timedelta(days=max(t["duration"], 0.5))
        row_data = [t["id"], t["wbs"], t["task"], t["phase"], t["start"], end_date, t["duration"]]
        
        for col, val in enumerate(row_data, 1):
            cell = ws_gantt.cell(row_idx, col, value=val)
            cell.border = thin_border()
            if col in (5, 6):
                cell.number_format = "DD-MMM"
                cell.alignment = Alignment(horizontal="center")
            elif col == 7:
                cell.number_format = "0.0"
                cell.alignment = Alignment(horizontal="center")
            
            if t["phase"] == "MILESTONE":
                cell.font = Font(bold=True)
        
        # Gantt bars
        phase_color = COLORS.get(t["phase"], "D9D9D9")
        
        for j, d in enumerate(days):
            col = gantt_start_col + j
            cell = ws_gantt.cell(row_idx, col)
            cell.border = thin_border()
            
            # Check if this date is within task range
            task_start = t["start"]
            task_end = end_date
            
            if task_start <= d < task_end or (task_start == task_end == d):
                cell.fill = PatternFill("solid", fgColor=phase_color)
                if t["phase"] == "MILESTONE":
                    cell.value = "★"
                    cell.alignment = Alignment(horizontal="center")
                    cell.font = Font(bold=True, color="FFFFFF")
    
    # Column widths
    meta_widths = {"A": 10, "B": 6, "C": 45, "D": 12, "E": 10, "F": 10, "G": 5}
    for col, w in meta_widths.items():
        ws_gantt.column_dimensions[col].width = w
    
    for j in range(len(days)):
        ws_gantt.column_dimensions[get_column_letter(gantt_start_col + j)].width = 3.5
    
    # Freeze panes
    ws_gantt.freeze_panes = ws_gantt.cell(first_data_row, gantt_start_col)
    
    # === SHEET 3: WEATHER ANALYSIS ===
    ws_weather = wb.create_sheet("Weather_Analysis")
    
    ws_weather["A1"] = "UAE Winter Weather Analysis (January-February 2026)"
    ws_weather["A1"].font = Font(bold=True, size=14)
    ws_weather.merge_cells("A1:F1")
    
    weather_data = [
        ["Parameter", "Jan 1-10", "Jan 11-20", "Jan 21-31", "Feb 1-15", "Notes"],
        ["Avg Wind (kt)", "11-13", "16-21", "13-15", "12-14", "Peak Shamal: Jan 14-18"],
        ["Gust Risk (kt)", "15-18", "22-28", "16-20", "15-18", "NO-GO if >22kt sustained"],
        ["Wave Height (m)", "0.4-0.6", "0.8-1.2", "0.5-0.7", "0.4-0.6", "HOLD if >0.8m"],
        ["Visibility (km)", "6-8", "2-5", "5-7", "6-8", "Dust/sand reduction"],
        ["Sailing Window", "✅ GO", "⚠️ HOLD/NO-GO", "✅ GO", "✅ GO", "Schedule optimized"],
        ["Risk Level", "LOW", "HIGH", "MEDIUM", "LOW", "Plan A avoids Jan 14-21"],
    ]
    
    for row_idx, row_data in enumerate(weather_data, 3):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_weather.cell(row_idx, col_idx, value=val)
            cell.border = thin_border()
            if row_idx == 3:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor=COLORS["HEADER"])
            elif "NO-GO" in str(val) or "HIGH" in str(val):
                cell.fill = PatternFill("solid", fgColor="FFE0E0")
            elif "HOLD" in str(val) or "MEDIUM" in str(val):
                cell.fill = PatternFill("solid", fgColor=COLORS["SHAMAL_RISK"])
            elif "GO" in str(val) or "LOW" in str(val):
                cell.fill = PatternFill("solid", fgColor="E0FFE0")
    
    # Column widths
    ws_weather.column_dimensions["A"].width = 18
    for col in ["B", "C", "D", "E"]:
        ws_weather.column_dimensions[col].width = 14
    ws_weather.column_dimensions["F"].width = 28
    
    # === SHEET 4: SUMMARY ===
    ws_summary = wb.create_sheet("Summary")
    
    summary_data = [
        ["AGI HVDC Transformer Transportation - Project Summary"],
        [""],
        ["Key Parameters"],
        ["Total Transformers", "6 units (TR1-TR6)"],
        ["Weight per TR", "217-271 tons"],
        ["Total Voyages", "6 (1 TR per voyage)"],
        ["Jack-down Events", "3 (after V2, V4, V6)"],
        ["Vessel", "LCT BUSHRA"],
        ["Route", "Mina Zayed Port ↔ AGI Site"],
        [""],
        ["Schedule Summary"],
        ["Project Start", "2026-01-06"],
        ["Mobilization", "Jan 6-8, 2026"],
        ["Voyage 1 (TR1)", "Jan 9-12, 2026"],
        ["Voyage 2 (TR2 + JD1)", "Jan 13-25, 2026"],
        ["Voyage 3 (TR3)", "Jan 22-25, 2026"],
        ["Voyage 4 (TR4 + JD2)", "Jan 27 - Feb 6, 2026"],
        ["Voyage 5 (TR5)", "Feb 2-5, 2026"],
        ["Voyage 6 (TR6 + JD3)", "Feb 7-18, 2026"],
        ["Project Complete", f"~{completion_date.strftime('%Y-%m-%d')}"],
        [""],
        ["Weather Risk Management"],
        ["Winter Shamal Period", "Jan 14-21, 2026 (HIGH RISK)"],
        ["Mitigation Strategy", "Schedule V1+V2 before peak; V3-V6 after"],
        ["Weather Windows", "9 GO/HOLD decision points"],
        ["Buffer Days", "0.5 day per voyage"],
        [""],
        ["Critical Constraints"],
        ["AGI FWD Draft Limit", "≤ 2.70m"],
        ["Ramp Angle (AGI)", "≤ 10°"],
        ["GM Requirement", "≥ 1.50m"],
        ["PTW Lead Time (AGI)", "24 hours"],
    ]
    
    for row_idx, row_data in enumerate(summary_data, 1):
        if isinstance(row_data, list) and len(row_data) == 1:
            ws_summary.merge_cells(f"A{row_idx}:B{row_idx}")
            cell = ws_summary.cell(row_idx, 1, value=row_data[0])
            cell.font = Font(bold=True, size=12 if row_idx == 1 else 11)
            if row_idx == 1:
                cell.fill = PatternFill("solid", fgColor=COLORS["HEADER"])
                cell.font = Font(bold=True, size=14, color="FFFFFF")
        elif isinstance(row_data, list) and len(row_data) == 2:
            ws_summary.cell(row_idx, 1, value=row_data[0]).font = Font(bold=True)
            ws_summary.cell(row_idx, 2, value=row_data[1])
    
    ws_summary.column_dimensions["A"].width = 25
    ws_summary.column_dimensions["B"].width = 35
    
    return wb

def main():
    print("Generating AGI TR 6-Voyage Gantt Chart...")
    wb = create_gantt_workbook()
    output_path = "/home/claude/AGI_TR_6Voyage_Master_Gantt.xlsx"
    wb.save(output_path)
    print(f"✅ Generated: {output_path}")
    return output_path

if __name__ == "__main__":
    main()