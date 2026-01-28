#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd


ARABIC_PREFIXES = {"BIN", "BINT", "AL", "ABU", "ABUL", "IBN"}
QUALIFIER_RE = re.compile(r"\b(new|old|visa|eid)\b", re.IGNORECASE)
SN_PATTERNS = {
    "s.n.",
    "s.n",
    "sn",
    "s/no",
    "s/no.",
    "s/ no",
    "no.",
    "no",
    "number",
    "num",
    "#",
    "serial",
    "serial no",
    "serial number",
    "seq",
    "sequence",
}
NAME_PATTERNS = {
    "name",
    "employee name",
    "full name",
    "이름",
    "staff name",
    "personnel name",
    "person name",
    "worker name",
    "crew name",
}
POSITION_PATTERNS = {
    "position",
    "job",
    "title",
    "role",
    "designation",
    "직책",
    "직위",
    "담당",
}


def normalize_name_advanced(value: str) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = re.sub(r"[^\w\s\-.]", "", text)
    text = re.sub(r"\s+", " ", text)
    parts = []
    for part in text.split():
        if part in ARABIC_PREFIXES:
            parts.append(part)
        else:
            parts.append(part)
    return " ".join(parts)


def extract_first_last_name(value: str) -> tuple[str, str]:
    normalized = normalize_name_advanced(value)
    parts = normalized.split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]


def similarity_advanced(name1: str, name2: str) -> float:
    if not name1 or not name2:
        return 0.0
    norm1 = normalize_name_advanced(name1)
    norm2 = normalize_name_advanced(name2)
    full_sim = SequenceMatcher(None, norm1, norm2).ratio()
    first1, last1 = extract_first_last_name(name1)
    first2, last2 = extract_first_last_name(name2)
    first_sim = SequenceMatcher(None, first1, first2).ratio() if first1 and first2 else 0.0
    last_sim = SequenceMatcher(None, last1, last2).ratio() if last1 and last2 else 0.0
    combined = (full_sim * 0.5) + (first_sim * 0.25) + (last_sim * 0.25)
    return max(full_sim, combined)


def normalize_folder_name(folder_name: str) -> str:
    text = re.sub(r"^\d+\.\s*", "", folder_name)
    text = re.sub(r"^[A-Z\s]+-\s*", "", text)
    text = re.sub(r"\s*-\s*(new|old)\s+visa.*$", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*-\s*eid.*$", "", text, flags=re.IGNORECASE)
    return normalize_name_advanced(text)


def to_sn(value) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"\d+", text)
    if not match:
        return None
    return int(match.group(0))


def load_tsv(tsv_path: Path) -> pd.DataFrame:
    return pd.read_csv(tsv_path, sep="\t", dtype=str).fillna("")


def find_sn_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_lower = str(col).lower().strip()
        col_norm = col_lower.replace(" ", "").replace(".", "").replace("/", "")
        for pattern in SN_PATTERNS:
            pat_norm = pattern.replace(" ", "").replace(".", "").replace("/", "")
            if col_norm == pat_norm or pat_norm in col_norm:
                return col
    return None


def find_name_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if "name" in col_lower and "number" not in col_lower and "no" not in col_lower:
            return col
        if any(pat in col_lower for pat in NAME_PATTERNS):
            return col
    for col in df.columns:
        if df[col].dtype == "object":
            sample = df[col].dropna()
            if not sample.empty:
                text = str(sample.iloc[0])
                if re.search(r"[A-Za-z가-힣]", text) and " " in text and 5 < len(text) < 50:
                    return col
    return None


def find_position_column(df: pd.DataFrame) -> str | None:
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(pat in col_lower for pat in POSITION_PATTERNS):
            return col
    return None


def filter_excel_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all")
    name_col = find_name_column(df)
    if name_col:
        df = df[df[name_col].astype(str).str.strip() != ""]
    sn_col = find_sn_column(df)
    if sn_col:
        df = df[pd.to_numeric(df[sn_col], errors="coerce").notna()]
    return df.reset_index(drop=True)


def load_excel(excel_path: Path, sheet: str | None = None) -> pd.DataFrame:
    sheet_to_read = sheet if sheet is not None else 0
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=None, engine="openpyxl")
    header_row = None
    for idx in range(min(10, len(df_raw))):
        row_vals = [str(v).strip() for v in df_raw.iloc[idx].tolist()]
        if "S.N." in row_vals or "S.N" in row_vals:
            header_row = idx
            break
    if header_row is None:
        df = pd.read_excel(excel_path, sheet_name=sheet_to_read, engine="openpyxl")
        return filter_excel_rows(df.fillna(""))
    headers = [str(v).strip() for v in df_raw.iloc[header_row].tolist()]
    df = df_raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.loc[:, df.columns.notna()]
    return filter_excel_rows(df.fillna(""))


@dataclass
class FolderEntry:
    sn: int | None
    raw_name: str
    person: str
    person_norm: str
    role: str
    path: Path


def parse_folder_entry(folder: Path) -> FolderEntry:
    name = folder.name
    match = re.match(r"^(\d+)\.\s*(.+)$", name)
    sn = int(match.group(1)) if match else None
    rest = match.group(2) if match else name
    parts = [p.strip() for p in rest.split(" - ") if p.strip()]
    role = parts[0] if parts else ""
    person = ""
    for part in parts[1:]:
        if not QUALIFIER_RE.search(part):
            person = part
            break
    if not person and len(parts) >= 2:
        person = parts[1]
    if not person:
        person = rest
    person = re.sub(r"\s*-\s*(new|old)\s+visa.*$", "", person, flags=re.IGNORECASE).strip()
    person = re.sub(r"\s*-\s*eid.*$", "", person, flags=re.IGNORECASE).strip()
    person_norm = normalize_name_advanced(person)
    return FolderEntry(sn=sn, raw_name=name, person=person, person_norm=person_norm, role=role, path=folder)


def build_folder_map(base_dir: Path) -> dict[int, FolderEntry]:
    folder_map: dict[int, FolderEntry] = {}
    for entry in sorted(base_dir.iterdir()):
        if entry.is_dir():
            info = parse_folder_entry(entry)
            if info.sn is not None:
                folder_map[info.sn] = info
    return folder_map


def summarize_folder_files(folder: Path) -> dict[str, int]:
    files = list(folder.glob("*"))
    return {
        "pdf": sum(1 for f in files if f.suffix.lower() == ".pdf"),
        "image": sum(1 for f in files if f.suffix.lower() in {".jpg", ".jpeg", ".png"}),
        "total": len(files),
    }


def find_best_match(
    name: str,
    candidates: list[dict[str, Any]],
    name_key: str,
    threshold: float,
) -> tuple[dict[str, Any] | None, float]:
    best = None
    best_sim = 0.0
    for cand in candidates:
        sim = similarity_advanced(name, cand[name_key])
        if sim > best_sim:
            best_sim = sim
            best = cand
    if best is not None and best_sim >= threshold:
        return best, best_sim
    return None, best_sim


def validate(
    tsv_path: Path,
    excel_path: Path,
    folder_root: Path,
    sheet: str | None = None,
    excel_threshold: float = 0.7,
    folder_threshold: float = 0.6,
    report_file: Path | None = None,
) -> int:
    tsv_df = load_tsv(tsv_path)
    excel_df = load_excel(excel_path, sheet=sheet)

    tsv_sn_col = find_sn_column(tsv_df)
    tsv_name_col = find_name_column(tsv_df)
    tsv_position_col = find_position_column(tsv_df)
    if not tsv_name_col:
        raise ValueError("TSV name column not found")

    tsv_records = []
    for idx, row in tsv_df.iterrows():
        sn_value = None
        if tsv_sn_col:
            sn_value = to_sn(row.get(tsv_sn_col, ""))
        if sn_value is None:
            sn_value = idx + 1
        name_value = str(row.get(tsv_name_col, "")).strip()
        position_value = str(row.get(tsv_position_col, "")).strip() if tsv_position_col else ""
        tsv_records.append(
            {
                "sn": sn_value,
                "name": name_value,
                "position": position_value,
                "name_norm": normalize_name_advanced(name_value),
            }
        )

    excel_records = []
    name_col = find_name_column(excel_df)
    excel_sn_col = find_sn_column(excel_df)
    if name_col:
        for idx, row in excel_df.iterrows():
            name = str(row.get(name_col, "")).strip()
            if name:
                excel_records.append(
                    {
                        "id": idx,
                        "row_num": idx + 2,
                        "sn": row.get(excel_sn_col, "") if excel_sn_col else "",
                        "name": name,
                        "name_norm": normalize_name_advanced(name),
                    }
                )

    folder_map = build_folder_map(folder_root)
    folder_records = []
    for entry in folder_map.values():
        counts = summarize_folder_files(entry.path)
        folder_records.append(
            {
                "sn": entry.sn,
                "raw_name": entry.raw_name,
                "person": entry.person,
                "name_norm": entry.person_norm or normalize_folder_name(entry.raw_name),
                "pdf": counts["pdf"],
                "image": counts["image"],
            }
        )

    tsv_to_excel = []
    tsv_to_folder = []
    sn_mismatches_folder = []
    matched_excel_ids = set()
    matched_folder_sn = set()

    for tsv_rec in tsv_records:
        best_excel, sim_excel = find_best_match(tsv_rec["name"], excel_records, "name", excel_threshold)
        if best_excel:
            matched_excel_ids.add(best_excel["id"])
            tsv_to_excel.append(
                {
                    "tsv_sn": tsv_rec["sn"],
                    "tsv_name": tsv_rec["name"],
                    "excel_row": best_excel["row_num"],
                    "excel_sn": best_excel["sn"],
                    "excel_name": best_excel["name"],
                    "similarity": round(sim_excel, 3),
                }
            )

        best_folder, sim_folder = find_best_match(tsv_rec["name"], folder_records, "person", folder_threshold)
        if best_folder:
            matched_folder_sn.add(best_folder["sn"])
            sn_mismatch = (
                tsv_rec["sn"] is not None
                and best_folder["sn"] is not None
                and tsv_rec["sn"] != best_folder["sn"]
            )
            tsv_to_folder.append(
                {
                    "tsv_sn": tsv_rec["sn"],
                    "tsv_name": tsv_rec["name"],
                    "folder_sn": best_folder["sn"],
                    "folder_name": best_folder["raw_name"],
                    "similarity": round(sim_folder, 3),
                    "pdf": best_folder["pdf"],
                    "image": best_folder["image"],
                    "sn_mismatch": sn_mismatch,
                }
            )
            if sn_mismatch:
                sn_mismatches_folder.append(
                    {
                        "tsv_sn": tsv_rec["sn"],
                        "tsv_name": tsv_rec["name"],
                        "folder_sn": best_folder["sn"],
                        "folder_name": best_folder["raw_name"],
                        "similarity": round(sim_folder, 3),
                    }
                )

    unmatched_tsv = [r for r in tsv_records if r["sn"] is not None and all(m["tsv_sn"] != r["sn"] for m in tsv_to_excel)]
    unmatched_excel = [r for r in excel_records if r["id"] not in matched_excel_ids]
    unmatched_folder = [r for r in folder_records if r["sn"] not in matched_folder_sn]

    sn_mismatches = []
    for match in tsv_to_excel:
        if match["excel_sn"] and str(match["excel_sn"]).strip() != str(match["tsv_sn"]).strip():
            sn_mismatches.append(match)

    lines = []
    lines.append("TSV rows: %d" % len(tsv_df))
    lines.append("Excel rows (filtered): %d" % len(excel_df))
    lines.append("Folder entries: %d" % len(folder_records))
    lines.append("TSV -> Excel matches: %d/%d" % (len(tsv_to_excel), len(tsv_records)))
    lines.append("TSV -> Folder matches: %d/%d" % (len(tsv_to_folder), len(tsv_records)))
    if unmatched_tsv:
        lines.append("Unmatched TSV (name to Excel): %d" % len(unmatched_tsv))
        for item in unmatched_tsv:
            lines.append("  - %s (%s)" % (item["name"], item["position"]))
    if unmatched_excel:
        lines.append("Unmatched Excel (name to TSV): %d" % len(unmatched_excel))
        for item in unmatched_excel[:10]:
            lines.append("  - row %s: %s" % (item["row_num"], item["name"]))
        if len(unmatched_excel) > 10:
            lines.append("  ... %d more" % (len(unmatched_excel) - 10))
    if unmatched_folder:
        lines.append("Unmatched folders (name to TSV): %d" % len(unmatched_folder))
        for item in unmatched_folder:
            lines.append("  - %s" % item["raw_name"])
    if sn_mismatches_folder:
        lines.append("S.N. mismatches between TSV and Folder: %d" % len(sn_mismatches_folder))
        for item in sn_mismatches_folder:
            lines.append(
                "  - TSV %s vs Folder %s: %s"
                % (item["tsv_sn"], item["folder_sn"], item["tsv_name"])
            )
    if sn_mismatches:
        lines.append("S.N. mismatches between TSV and Excel: %d" % len(sn_mismatches))
        for item in sn_mismatches:
            lines.append(
                "  - TSV %s vs Excel %s: %s"
                % (item["tsv_sn"], item["excel_sn"], item["tsv_name"])
            )

    print("\n".join(lines))

    if report_file:
        summary_rows = [
            {"metric": "tsv_rows", "value": len(tsv_df)},
            {"metric": "excel_rows", "value": len(excel_df)},
            {"metric": "folder_entries", "value": len(folder_records)},
            {"metric": "tsv_to_excel_matches", "value": len(tsv_to_excel)},
            {"metric": "tsv_to_folder_matches", "value": len(tsv_to_folder)},
            {"metric": "unmatched_tsv", "value": len(unmatched_tsv)},
            {"metric": "unmatched_excel", "value": len(unmatched_excel)},
            {"metric": "unmatched_folder", "value": len(unmatched_folder)},
            {"metric": "sn_mismatches_folder", "value": len(sn_mismatches_folder)},
            {"metric": "sn_mismatches", "value": len(sn_mismatches)},
        ]
        report_file.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(report_file, engine="openpyxl") as writer:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="summary", index=False)
            pd.DataFrame(tsv_to_excel).to_excel(writer, sheet_name="tsv_to_excel", index=False)
            pd.DataFrame(tsv_to_folder).to_excel(writer, sheet_name="tsv_to_folder", index=False)
            pd.DataFrame(unmatched_tsv).to_excel(writer, sheet_name="unmatched_tsv", index=False)
            pd.DataFrame(unmatched_excel).to_excel(writer, sheet_name="unmatched_excel", index=False)
            pd.DataFrame(unmatched_folder).to_excel(writer, sheet_name="unmatched_folder", index=False)
            pd.DataFrame(sn_mismatches_folder).to_excel(writer, sheet_name="sn_mismatches_folder", index=False)
            pd.DataFrame(sn_mismatches).to_excel(writer, sheet_name="sn_mismatches", index=False)

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate Mammoet TSV vs Excel vs folder structure")
    parser.add_argument("--tsv", default="mammoet/S.N.tsv", help="Path to TSV file")
    parser.add_argument(
        "--excel",
        default="mammoet/15111578 - Samsung HVDC - Mina Zayed Manpower - 2026.xlsx",
        help="Path to Excel file",
    )
    parser.add_argument(
        "--folder",
        default="mammoet/Mammoet Mina Zayed Manpower - 2026 - Part 1",
        help="Folder root to validate",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet name (default: first sheet)")
    parser.add_argument("--excel-threshold", type=float, default=0.7, help="Name similarity threshold for Excel")
    parser.add_argument("--folder-threshold", type=float, default=0.6, help="Name similarity threshold for folders")
    parser.add_argument("--report-file", default="", help="Write single Excel report file")
    args = parser.parse_args()

    report_file = Path(args.report_file) if args.report_file else None

    if report_file is None:
        report_file = Path(args.tsv).parent / "mammoet_validation_report.xlsx"

    return validate(
        Path(args.tsv),
        Path(args.excel),
        Path(args.folder),
        sheet=args.sheet,
        excel_threshold=args.excel_threshold,
        folder_threshold=args.folder_threshold,
        report_file=report_file,
    )


if __name__ == "__main__":
    raise SystemExit(main())
