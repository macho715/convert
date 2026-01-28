# -*- coding: utf-8 -*-
"""
Apply Mammoet NOTES logic to schedule TSV exports (Primavera/P6-like).

- Input: TSV with columns:
  Activity ID, Activity Name, Original Duration, Planned Start, Planned Finish, Actual Start, Actual Finish
- Output: same TSV with an added column "Notes" (or overwritten).

Logic is derived from "Original (2-2-2-1 / Single SPMT)" NOTES convention:
- Tide gate: "TIDE>=1.90 required (Loadout start)"
- Weather gate: "WX gate"
- RoRo: "RORO + ramp"
- Seafastening/MWS: "Lashing + survey"
- Turning: "3.0d/unit"
- Jackdown: "1.0d/unit"
- Return & crew demob: "After final JD"
- Mobilization: "SPMT + grillage in MZP"
- Deck preparations: "MWS pre-check ready"
- Beam replacement (extension): "One-time setup (same LCT)"

Usage:
  python apply_notes_logic.py --in MFA2.tsv --out MFA2_with_NOTES.tsv
"""

from __future__ import annotations
import argparse
import re
import pandas as pd


def is_leaf_activity(activity_id: str) -> bool:
    """Leaf activity rows are assumed to be 'A' + 4 digits (e.g., A1074)."""
    if activity_id is None:
        return False
    return bool(re.match(r"^A\d{4}$", str(activity_id).strip()))


def apply_notes_logic(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Notes" not in df.columns:
        df["Notes"] = ""

    names = df["Activity Name"].astype(str)
    ids = df["Activity ID"].astype(str)

    out_notes = []
    for aid, name in zip(ids, names):
        note = ""
        if not is_leaf_activity(aid):
            out_notes.append(note)
            continue

        n = name.strip()

        # Priority matters (first match wins)
        if re.search(r"\bBeam Replacement\b", n, flags=re.I):
            note = "One-time setup (same LCT)"

        elif re.search(r"Mobilization of SPMT", n, flags=re.I):
            note = "SPMT + grillage in MZP"

        elif re.search(r"\bDeck Preparations?\b", n, flags=re.I):
            note = "MWS pre-check ready"

        elif re.search(r"\bLoad-out\b", n, flags=re.I) or re.search(r"Loading on LCT", n, flags=re.I) or re.search(r"tide allows", n, flags=re.I):
            note = "TIDE>=1.90 required (Loadout start)"

        elif re.search(r"Seafastening|Sea fastening", n, flags=re.I) or re.search(r"\bMWS\b", n, flags=re.I):
            note = "Lashing + survey"

        elif re.search(r"Sail-away back", n, flags=re.I) or re.search(r"\bback to Mina Zayed\b", n, flags=re.I):
            note = "After final JD"

        elif re.search(r"Crew Demobilization|Crew Demob", n, flags=re.I):
            note = "After final JD"

        elif re.search(r"\bSail-away\b", n, flags=re.I):
            note = "WX gate"

        elif re.search(r"\bLoad-in\b", n, flags=re.I) or re.search(r"\bRoRo\b", n, flags=re.I) or re.search(r"Berthing", n, flags=re.I):
            note = "RORO + ramp"

        elif re.search(r"\bTurning\b", n, flags=re.I):
            note = "3.0d/unit"

        elif re.search(r"Jacking down|Jackdown", n, flags=re.I):
            note = "1.0d/unit"

        elif re.search(r"Buffer|reset", n, flags=re.I):
            note = "contingency"

        out_notes.append(note)

    df["Notes"] = out_notes
    return df


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input TSV path")
    ap.add_argument("--out", dest="out", required=True, help="Output TSV path")
    args = ap.parse_args()

    df = pd.read_csv(args.inp, sep="\t")
    df2 = apply_notes_logic(df)
    df2.to_csv(args.out, sep="\t", index=False)


if __name__ == "__main__":
    main()
