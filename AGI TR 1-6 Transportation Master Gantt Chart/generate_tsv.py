import csv
import datetime as dt

# Project Start Date (from Untitled-1.py logic)
PROJECT_START = dt.date(2026, 1, 18)

# Updated Task List from Untitled-1.py
default_tasks = [
    # Mobilization
    ("MOB-001", "1.0", "MOBILIZATION", "MOBILIZATION", "Mammoet", 0, "DUR_MOB", "SPMT Assembly + Marine Equipment Mobilization"),
    ("PREP-001", "1.1", "Deck Preparations", "DECK_PREP", "Mammoet", 1, "DUR_DECK", "One-time setup for all voyages"),
    
    # Voyage 1: LO 01-18, SAIL 01-20, ARR 01-22
    ("V1", "2.0", "VOYAGE 1: TR1 Transport", "MILESTONE", "All", 0, 0, "✅ Tide ≥1.90m (2.05m) | Good Weather Window"),
    ("LO-101", "2.1", "TR1 Load-out on LCT", "LOADOUT", "Mammoet", 0, "DUR_LO", "Tide ≥1.90m (2.05m) required"),
    ("SF-102", "2.2", "TR1 Sea Fastening", "SEAFAST", "Mammoet", 0, "DUR_SF", "12-point lashing"),
    ("MWS-103", "2.3", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 0, "DUR_MWS", "Marine Warranty Surveyor"),
    ("SAIL-104", "2.4", "V1 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 2, "DUR_SAIL", "✅ Good Weather Window"),
    ("ARR-105", "2.5", "AGI Arrival + TR1 RORO Unload", "AGI_UNLOAD", "Mammoet", 4, "DUR_UL", "Tide ≥1.90m (1.91m) | AGI FWD Draft ≤ 2.70m"),
    ("STORE-106", "2.6", "TR1 Stored on AGI Laydown", "BUFFER", "Mammoet", 4, "DUR_BUF", "Awaiting pair TR2"),
    ("RET-107", "2.7", "V1 LCT Return: AGI→MZP", "RETURN", "LCT Bushra", 4, "DUR_RET", "Quick turnaround"),
    ("BUF-108", "2.99", "V1 Buffer / Equipment Reset", "BUFFER", "All", 5, "DUR_BUF", "Weather contingency"),
    
    # Voyage 2: LO 01-26, SAIL 01-27, ARR 01-29
    ("V2", "3.0", "VOYAGE 2: TR2 Transport + JD-1", "MILESTONE", "All", 8, 0, "✅ Tide ≥1.90m (1.91m) | Good Weather Window (before Shamal)"),
    ("LO-109", "3.1", "TR2 Load-out on LCT", "LOADOUT", "Mammoet", 8, "DUR_LO", "Tide ≥1.90m (1.91m) required"),
    ("SF-110", "3.2", "TR2 Sea Fastening", "SEAFAST", "Mammoet", 8, "DUR_SF", "12-point lashing"),
    ("MWS-110A", "3.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 8, "DUR_MWS", "Pre-sail verification"),
    ("SAIL-111", "3.3", "V2 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 9, "DUR_SAIL", "✅ Good Weather Window"),
    ("ARR-112", "3.4", "AGI Arrival + TR2 RORO Unload", "AGI_UNLOAD", "Mammoet", 11, "DUR_UL", "Tide ≥1.90m (2.03m) | AGI FWD Draft ≤ 2.70m"),
    ("TRN-113", "3.5", "TR1 Transport to Bay-1", "TURNING", "Mammoet", 12, 1, "Steel bridge install"),
    ("TURN-114", "3.6", "TR1 Turning (90° rotation)", "TURNING", "Mammoet", 12, "DUR_TURN", "10t Forklift required"),
    ("TRN-116", "3.8", "TR2 Transport to Bay-2", "TURNING", "Mammoet", 12, 1, ""),
    ("TURN-117", "3.9", "TR2 Turning (90° rotation)", "TURNING", "Mammoet", 12, "DUR_TURN", ""),
    ("JD-120A", "3.95", "JD-1 Jack-Down TR1", "JACKDOWN", "Mammoet", 14, "DUR_JD", "MILESTONE: TR1 complete | 02-01"),
    ("RET-119", "3.11", "V2 LCT Return: AGI->MZP", "RETURN", "LCT Bushra", 15, "DUR_RET", "Return after first JD (SPMT reuse)"),
    ("JD-120B", "3.96", "JD-1 Jack-Down TR2", "JACKDOWN", "Mammoet", 16, "DUR_JD", "MILESTONE: TR2 complete | 02-02"),
    ("BUF-120", "3.99", "V2 Buffer / Shamal Recovery", "BUFFER", "All", 17, "DUR_BUF", "Post-Shamal weather check"),
    
    # Voyage 3: LO 01-31, SAIL 02-02, ARR 02-03
    ("V3", "4.0", "VOYAGE 3: TR3 Transport", "MILESTONE", "All", 13, 0, "✅ Tide ≥1.90m (2.07m) | Post-Shamal Window"),
    ("LO-121", "4.1", "TR3 Load-out on LCT", "LOADOUT", "Mammoet", 13, "DUR_LO", "Tide ≥1.90m (2.07m)"),
    ("SF-122", "4.2", "TR3 Sea Fastening", "SEAFAST", "Mammoet", 13, "DUR_SF", ""),
    ("MWS-122A", "4.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 13, "DUR_MWS", ""),
    ("SAIL-123", "4.3", "V3 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 15, "DUR_SAIL", "Good weather"),
    ("ARR-124", "4.4", "AGI Arrival + TR3 RORO Unload", "AGI_UNLOAD", "Mammoet", 16, "DUR_UL", "Tide ≥1.90m (2.04m)"),
    ("STORE-125", "4.5", "TR3 Stored on AGI Laydown", "BUFFER", "Mammoet", 16, "DUR_BUF", "Awaiting pair TR4"),
    ("RET-126", "4.6", "V3 LCT Return: AGI→MZP", "RETURN", "LCT Bushra", 17, "DUR_RET", ""),
    ("BUF-127", "4.99", "V3 Buffer", "BUFFER", "All", 17, "DUR_BUF", ""),
    
    # Voyage 4: LO 02-15, SAIL 02-16, ARR 02-18
    ("V4", "5.0", "VOYAGE 4: TR4 Transport + JD-2", "MILESTONE", "All", 28, 0, "✅ Tide ≥1.90m (1.90m) | Shamal 종료 직후"),
    ("LO-128", "5.1", "TR4 Load-out on LCT", "LOADOUT", "Mammoet", 28, "DUR_LO", "Tide ≥1.90m (1.90m)"),
    ("SF-129", "5.2", "TR4 Sea Fastening", "SEAFAST", "Mammoet", 28, "DUR_SF", ""),
    ("MWS-129A", "5.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 28, "DUR_MWS", ""),
    ("SAIL-130", "5.3", "V4 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 29, "DUR_SAIL", ""),
    ("ARR-131", "5.4", "AGI Arrival + TR4 RORO Unload", "AGI_UNLOAD", "Mammoet", 31, "DUR_UL", "Tide ≥1.90m (1.96m)"),
    ("TRN-132", "5.5", "TR3 Transport to Bay-3", "TURNING", "Mammoet", 31, 1, ""),
    ("TURN-133", "5.6", "TR3 Turning (90° rotation)", "TURNING", "Mammoet", 31, "DUR_TURN", ""),
    ("TRN-135", "5.8", "TR4 Transport to Bay-4", "TURNING", "Mammoet", 31, 1, ""),
    ("TURN-136", "5.9", "TR4 Turning (90° rotation)", "TURNING", "Mammoet", 31, "DUR_TURN", ""),
    ("JD-139A", "5.95", "JD-2 Jack-Down TR3", "JACKDOWN", "Mammoet", 33, "DUR_JD", "MILESTONE: TR3 complete | 02-20"),
    ("RET-138", "5.11", "V4 LCT Return: AGI->MZP", "RETURN", "LCT Bushra", 34, "DUR_RET", "Return after first JD (SPMT reuse)"),
    ("JD-139B", "5.96", "JD-2 Jack-Down TR4", "JACKDOWN", "Mammoet", 35, "DUR_JD", "MILESTONE: TR4 complete | 02-21"),
    ("BUF-140", "5.99", "V4 Buffer", "BUFFER", "All", 36, "DUR_BUF", ""),
    
    # Voyage 5: LO 02-23, SAIL 02-23, ARR 02-24 (Fast-turn)
    ("V5", "6.0", "VOYAGE 5: TR5 Transport", "MILESTONE", "All", 36, 0, "✅ Tide ≥1.90m (1.99m) | Fast-turn"),
    ("LO-140", "6.1", "TR5 Load-out on LCT", "LOADOUT", "Mammoet", 36, "DUR_LO", "Tide ≥1.90m (1.99m)"),
    ("SF-141", "6.2", "TR5 Sea Fastening", "SEAFAST", "Mammoet", 36, "DUR_SF", ""),
    ("MWS-141A", "6.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 36, "DUR_MWS", ""),
    ("SAIL-142", "6.3", "V5 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 36, "DUR_SAIL", "Fast-turn"),
    ("ARR-143", "6.4", "AGI Arrival + TR5 RORO Unload", "AGI_UNLOAD", "Mammoet", 37, "DUR_UL", "Tide ≥1.90m (2.01m)"),
    ("STORE-144", "6.5", "TR5 Stored on AGI Laydown", "BUFFER", "Mammoet", 37, "DUR_BUF", "Awaiting pair TR6"),
    ("RET-145", "6.6", "V5 LCT Return: AGI→MZP", "RETURN", "LCT Bushra", 37, "DUR_RET", ""),
    ("BUF-146", "6.99", "V5 Buffer", "BUFFER", "All", 37, "DUR_BUF", ""),
    
    # Voyage 6: LO 02-25, SAIL 02-25, ARR 02-26 (Fast-turn)
    ("V6", "7.0", "VOYAGE 6: TR6 Transport + JD-3", "MILESTONE", "All", 38, 0, "✅ Tide ≥1.90m (2.01m) | Fast-turn"),
    ("LO-147", "7.1", "TR6 Load-out on LCT", "LOADOUT", "Mammoet", 38, "DUR_LO", "Tide ≥1.90m (2.01m)"),
    ("SF-148", "7.2", "TR6 Sea Fastening", "SEAFAST", "Mammoet", 38, "DUR_SF", ""),
    ("MWS-148A", "7.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 38, "DUR_MWS", ""),
    ("SAIL-149", "7.3", "V6 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 38, "DUR_SAIL", "Fast-turn"),
    ("ARR-150", "7.4", "AGI Arrival + TR6 RORO Unload", "AGI_UNLOAD", "Mammoet", 39, "DUR_UL", "Tide ≥1.90m (1.98m)"),
    ("TRN-151", "7.5", "TR5 Transport to Bay-5", "TURNING", "Mammoet", 39, 1, ""),
    ("TURN-152", "7.6", "TR5 Turning (90° rotation)", "TURNING", "Mammoet", 39, "DUR_TURN", ""),
    ("TRN-154", "7.8", "TR6 Transport to Bay-6", "TURNING", "Mammoet", 39, 1, ""),
    ("TURN-155", "7.9", "TR6 Turning (90° rotation)", "TURNING", "Mammoet", 39, "DUR_TURN", ""),
    ("JD-157A", "7.95", "JD-3 Jack-Down TR5", "JACKDOWN", "Mammoet", 40, "DUR_JD", "MILESTONE: TR5 complete | 02-27"),
    ("RET-158", "7.11", "V6 LCT Return: AGI->MZP", "RETURN", "LCT Bushra", 41, "DUR_RET", "Return after first JD (SPMT reuse)"),
    ("JD-157B", "7.96", "JD-3 Jack-Down TR6", "JACKDOWN", "Mammoet", 42, "DUR_JD", "MILESTONE: TR6 complete | 02-28"),
    ("BUF-159", "7.99", "V6 Buffer / Reset for V7", "BUFFER", "All", 43, "DUR_BUF", ""),
    
    # Voyage 7: LO 02-27, SAIL 02-27, ARR 02-28 (Final)
    ("V7", "8.0", "VOYAGE 7: TR7 Transport + JD-4", "MILESTONE", "All", 40, 0, "✅ Tide ≥1.90m (1.92m) | Final unit"),
    ("LO-201", "8.1", "TR7 Load-out on LCT", "LOADOUT", "Mammoet", 40, "DUR_LO", "Tide ≥1.90m (1.92m) required"),
    ("SF-202", "8.2", "TR7 Sea Fastening", "SEAFAST", "Mammoet", 40, "DUR_SF", "12-point lashing"),
    ("MWS-202A", "8.25", "MWS + MPI + Final Check", "BUFFER", "Aries/Captain", 40, "DUR_MWS", ""),
    ("SAIL-203", "8.3", "V7 Sail-away: MZP→AGI", "SAIL", "LCT Bushra", 40, "DUR_SAIL", "Weather window required"),
    ("ARR-204", "8.4", "AGI Arrival + TR7 RORO Unload", "AGI_UNLOAD", "Mammoet", 41, "DUR_UL", "Tide ≥1.90m (1.93m) | AGI FWD Draft ≤ 2.70m"),
    ("TRN-205", "8.5", "TR7 Transport to Bay-7", "TURNING", "Mammoet", 41, 1, "Steel bridge install"),
    ("TURN-206", "8.6", "TR7 Turning (90° rotation)", "TURNING", "Mammoet", 41, "DUR_TURN", "10t Forklift required"),
    ("JD-207", "8.7", "★ JD-4 Jack-Down (TR7)", "JACKDOWN", "Mammoet", 41, "DUR_JD", "MILESTONE: TR7 Complete | 02-28"),
    ("RET-208", "8.8", "V7 LCT Final Return: AGI→MZP", "RETURN", "LCT Bushra", 41, "DUR_RET", "Final return"),
    
    # Demobilization
    ("DEMOB", "9.0", "DEMOBILIZATION", "MOBILIZATION", "Mammoet", 42, "DUR_MOB", "Equipment return"),
    ("END", "99.0", "★★★ PROJECT COMPLETE ★★★", "MILESTONE", "All", 42, 0, "All 7 TRs Installed | Jan-Feb 2026 Complete"),
]

# Phase Mapping (Python Phase -> TSV Phase)
phase_mapping_rev = {
    "MOBILIZATION": "Mobilization",
    "DECK_PREP": "Deck Prep",
    "LOADOUT": "MZP Loadout",
    "SEAFAST": "Sea Fastening",
    "BUFFER": "Buffer",
    "SAIL": "Sea Passage",
    "AGI_UNLOAD": "AGI Arrival",
    "TURNING": "Onshore SPMT",
    "JACKDOWN": "Jackdown",
    "RETURN": "Return",
    "MILESTONE": "Marine Transport",
}

# Duration Mapping (Duration_Ref -> Duration_days)
dur_mapping = {
    "DUR_MOB": 1.0,
    "DUR_DECK": 3.0,
    "DUR_LO": 1.0,
    "DUR_SF": 0.5,
    "DUR_MWS": 0.5,
    "DUR_SAIL": 1.0,
    "DUR_UL": 1.0,
    "DUR_TURN": 3.0,
    "DUR_JD": 1.0,
    "DUR_RET": 1.0,
    "DUR_BUF": 0.5,
}

output_file = "AGI_TR_Schedule_Updated.tsv"

with open(output_file, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f, delimiter='\t')
    # Header
    writer.writerow(["ID", "WBS", "Task", "Phase", "Owner", "Start", "Duration_days", "Notes"])
    
    for t in default_tasks:
        tid, wbs, task, phase, owner, offset, dur_ref, notes = t
        
        # Calculate Start Date
        start_date = PROJECT_START + dt.timedelta(days=offset)
        start_str = start_date.strftime("%Y-%m-%d")
        
        # Map Phase
        tsv_phase = phase_mapping_rev.get(phase, phase)
        
        # Map Duration
        if isinstance(dur_ref, str):
            duration = dur_mapping.get(dur_ref, 0)
        else:
            duration = dur_ref
            
        writer.writerow([tid, wbs, task, tsv_phase, owner, start_str, duration, notes])

print(f"✅ Generated {output_file} with {len(default_tasks)} tasks.")
