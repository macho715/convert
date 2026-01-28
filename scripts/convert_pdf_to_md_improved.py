#!/usr/bin/env python3
"""
Improved PDF to Markdown converter with cleaner table extraction.
"""

import argparse
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is not installed. Install it with: pip install pdfplumber")
    sys.exit(1)


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def normalize_cell(value: Any) -> str:
    return str(value or "").replace("\n", " ").strip()


def is_numeric_cell(text: str) -> bool:
    cleaned = text.replace(",", "").replace("%", "").strip()
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = cleaned[1:-1].strip()
    if not cleaned:
        return False
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def is_unit_like(text: str) -> bool:
    if not text:
        return False
    if "%" in text:
        return True
    lowered = text.lower()
    if "filled" in lowered and len(text) <= 12:
        return True
    if text.startswith("(") and text.endswith(")") and len(text) <= 6:
        return True
    return len(text) <= 4 and text.isalpha()


def is_label_text(text: str) -> bool:
    """Check if text is a label that should be removed from data rows."""
    if not text:
        return False
    lowered = text.lower().strip()
    # Common labels that appear in data rows but should be removed
    labels = {
        "maximum", "allowable", "% filled", "filled", 
        "remark", "ok", "nok", "stage", "no.", "no"
    }
    return lowered in labels or lowered.startswith("% filled") or lowered == "%"


def is_stage_index(text: str) -> bool:
    """Check if text is a stage index number (1, 2, 3, etc.) that should be removed."""
    if not text:
        return False
    text = text.strip()
    # Single digit or two-digit number that might be a stage index
    if text.isdigit():
        num = int(text)
        return 1 <= num <= 99  # Likely stage numbers
    return False


def fix_sno_column(rows: List[List[str]]) -> List[List[str]]:
    """Fix S.No. column positioning issues."""
    if not rows or len(rows) < 2:
        return rows
    
    header = rows[0]
    data_rows = rows[1:]
    
    # Check if first data column has sequential numbers (1, 2, 3...)
    if len(data_rows) > 0 and len(data_rows[0]) > 0:
        first_col_vals = [normalize_cell(row[0]) for row in data_rows[:5] if len(row) > 0 and normalize_cell(row[0])]
        
        # Check if first column has sequential numbers starting from 1
        if len(first_col_vals) >= 3:
            try:
                nums = [int(v) for v in first_col_vals if v.isdigit()]
                if len(nums) >= 3 and nums[0] == 1:
                    # Check if it's sequential
                    is_sequential = nums == list(range(1, len(nums) + 1))
                    
                    # First column is likely S.No.
                    # Check if header needs to be updated
                    header_first = normalize_cell(header[0]) if len(header) > 0 else ""
                    
                    # If header doesn't have S.No. in first column, update it
                    if not ("s.no" in header_first.lower() or "no." in header_first.lower()):
                        # Check if second column has "S.No." text
                        if len(data_rows[0]) > 1:
                            second_col_vals = [normalize_cell(row[1]) for row in data_rows[:3] if len(row) > 1]
                            second_col_has_sno = any("s.no" in v.lower() or ("no." in v.lower() and len(v) <= 5) for v in second_col_vals if v)
                            
                            if second_col_has_sno:
                                # Update header: first column becomes S.No., remove "S.No." from second
                                fixed_header = ["S.No."]
                                for i, cell in enumerate(header[1:], start=1):
                                    cell_val = normalize_cell(cell)
                                    if not ("s.no" in cell_val.lower() or ("no." in cell_val.lower() and len(cell_val) <= 5)):
                                        fixed_header.append(cell)
                                fixed_rows = [fixed_header]
                                
                                for row in data_rows:
                                    if len(row) > 1:
                                        # Keep first column (S.No. value), remove "S.No." from second column
                                        second_cell = normalize_cell(row[1])
                                        if "s.no" in second_cell.lower() or ("no." in second_cell.lower() and len(second_cell) <= 5):
                                            new_row = [row[0]] + list(row[2:])
                                        else:
                                            new_row = [row[0]] + list(row[1:])
                                        fixed_rows.append(new_row)
                                    else:
                                        fixed_rows.append(row)
                                return fixed_rows
                    else:
                        # Header already has S.No., just ensure data rows keep first column
                        return rows
            except (ValueError, IndexError):
                pass
    
    return rows


def clean_data_row_labels(row: List[str], header: List[str]) -> List[str]:
    """Remove label text and stage indices from data rows."""
    if not row or not header:
        return row
    
    cleaned_row = []
    prev_was_stage_index = False
    
    for i, cell in enumerate(row):
        cell_val = normalize_cell(cell)
        
        # Skip if it's a label text
        if is_label_text(cell_val):
            cleaned_row.append("")
            prev_was_stage_index = False
            continue
        
        # Check if this is a stage index that appears between data values
        if is_stage_index(cell_val):
            # Check if previous cell was a number (likely a data value)
            if i > 0 and len(cleaned_row) > 0:
                prev_cell = cleaned_row[-1] if cleaned_row else ""
                if is_numeric_cell(prev_cell):
                    # This stage index is between data values, skip it
                    prev_was_stage_index = True
                    continue
            
            # Check if next cell is also a number
            if i < len(row) - 1:
                next_cell = normalize_cell(row[i + 1] if i + 1 < len(row) else "")
                if is_numeric_cell(next_cell):
                    # Stage index between two data values, skip it
                    prev_was_stage_index = True
                    continue
        
        # Check if this column's header suggests it should contain only numbers
        if i < len(header):
            header_text = header[i].lower()
            # If header contains "stage" and cell is a stage index, skip it
            if "stage" in header_text and is_stage_index(cell_val) and not header_text.startswith("stage"):
                cleaned_row.append("")
                prev_was_stage_index = True
                continue
            # If header is about values (not stage numbers), remove stage indices
            if any(keyword in header_text for keyword in ["value", "moment", "force", "height", "ratio", "area", "angle", "lever", "required"]):
                if is_stage_index(cell_val):
                    cleaned_row.append("")
                    prev_was_stage_index = True
                    continue
        
        cleaned_row.append(cell_val)
        prev_was_stage_index = False
    
    return cleaned_row


def fill_merged_cells_final_pass(rows: List[List[str]]) -> List[List[str]]:
    """Final pass to fill remaining empty cells that should have repeated values."""
    if not rows or len(rows) < 3:
        return rows
    
    num_cols = max(len(row) for row in rows)
    header = rows[0]
    data_rows = rows[1:]
    
    # Identify columns that likely have repeated values
    # Look for columns with "Allowable" in header and mostly empty data cells
    repeated_cols = {}
    for col_idx in range(min(num_cols, len(header))):
        header_text = header[col_idx].lower() if col_idx < len(header) else ""
        if "allowable" in header_text or "remark" in header_text:
            # Count filled vs empty in this column
            filled = sum(1 for row in data_rows if col_idx < len(row) and normalize_cell(row[col_idx]))
            empty = sum(1 for row in data_rows if col_idx >= len(row) or not normalize_cell(row[col_idx]))
            
            # If first row has value but others are empty, likely repeated
            if len(data_rows) > 0 and col_idx < len(data_rows[0]):
                first_val = normalize_cell(data_rows[0][col_idx])
                if first_val and empty > filled:
                    repeated_cols[col_idx] = first_val
    
    # Fill the identified columns
    filled_data = [header]
    for row in data_rows:
        filled_row = list(row)
        while len(filled_row) < num_cols:
            filled_row.append("")
        
        for col_idx in range(num_cols):
            if col_idx in repeated_cols and not normalize_cell(filled_row[col_idx]):
                filled_row[col_idx] = repeated_cols[col_idx]
        
        filled_data.append(filled_row)
    
    return filled_data


def fill_merged_cells(rows: List[List[str]]) -> List[List[str]]:
    """Fill empty cells with values from previous row (handles merged cells)."""
    if not rows or len(rows) < 2:
        return rows
    
    filled_rows = [rows[0].copy()]  # Keep header as is
    num_cols = max(len(row) for row in rows)
    
    # First pass: identify which columns have repeated values
    # by analyzing the pattern of filled vs empty cells
    column_patterns = {}
    for col_idx in range(num_cols):
        filled_values = []
        empty_count = 0
        for row_idx in range(1, len(rows)):
            if col_idx < len(rows[row_idx]):
                val = normalize_cell(rows[row_idx][col_idx])
                if val:
                    filled_values.append(val)
                else:
                    empty_count += 1
        
        # If column has few filled values but many empty, likely merged cells
        if len(filled_values) <= 2 and empty_count > len(filled_values):
            # Get the first filled value as the repeated value
            if filled_values:
                column_patterns[col_idx] = filled_values[0]
    
    # Second pass: fill empty cells based on patterns
    for row_idx in range(1, len(rows)):
        current_row = rows[row_idx].copy()
        previous_row = filled_rows[-1]
        
        # Pad rows to same length
        while len(current_row) < num_cols:
            current_row.append("")
        while len(previous_row) < num_cols:
            previous_row.append("")
        
        # Fill empty cells with previous row's value
        filled_row = []
        for col_idx in range(num_cols):
            current_val = normalize_cell(current_row[col_idx] if col_idx < len(current_row) else "")
            prev_val = normalize_cell(previous_row[col_idx] if col_idx < len(previous_row) else "")
            
            # If current cell is empty
            if not current_val:
                # Priority 1: Use pattern value if this column has a repeated pattern
                if col_idx in column_patterns:
                    filled_row.append(column_patterns[col_idx])
                # Priority 2: Use previous row's value if it's numeric or common repeated value
                elif prev_val and (is_numeric_cell(prev_val) or prev_val in ["OK", "NOK"]):
                    # Additional check: make sure this column pattern suggests repetition
                    empty_in_col = sum(1 for r in rows[1:] if col_idx < len(r) and not normalize_cell(r[col_idx] if col_idx < len(r) else ""))
                    filled_in_col = sum(1 for r in rows[1:] if col_idx < len(r) and normalize_cell(r[col_idx] if col_idx < len(r) else ""))
                    
                    if empty_in_col > filled_in_col:
                        filled_row.append(prev_val)
                    else:
                        filled_row.append("")
                else:
                    filled_row.append("")
            else:
                filled_row.append(current_val)
        
        filled_rows.append(filled_row)
    
    return filled_rows


def clean_table_rows(table_rows: List[List[Any]]) -> Tuple[List[List[str]], List[int]]:
    if not table_rows:
        return [], []

    num_cols = max(len(row) for row in table_rows)
    empty_cols = set()

    for col_idx in range(num_cols):
        is_empty = True
        for row in table_rows:
            if col_idx < len(row):
                if normalize_cell(row[col_idx]):
                    is_empty = False
                    break
        if is_empty:
            empty_cols.add(col_idx)

    cleaned_rows: List[List[str]] = []
    for row in table_rows:
        cleaned_row: List[str] = []
        for idx in range(num_cols):
            if idx in empty_cols:
                continue
            cell_value = normalize_cell(row[idx]) if idx < len(row) else ""
            cleaned_row.append(cell_value)
        while cleaned_row and cleaned_row[-1] == "":
            cleaned_row.pop()
        if any(cell for cell in cleaned_row):
            cleaned_rows.append(cleaned_row)

    # Fill merged cells after cleaning
    if len(cleaned_rows) > 1:
        cleaned_rows = fill_merged_cells(cleaned_rows)

    return cleaned_rows, sorted(empty_cols)


def header_stats(row: List[str]) -> Dict[str, float]:
    non_empty = [cell for cell in row if cell]
    if not non_empty:
        return {"empty_ratio": 1.0, "numeric_ratio": 1.0, "unit_ratio": 0.0}
    empty_ratio = 1 - (len(non_empty) / len(row))
    numeric_ratio = sum(1 for cell in non_empty if is_numeric_cell(cell)) / len(non_empty)
    unit_ratio = sum(1 for cell in non_empty if is_unit_like(cell)) / len(non_empty)
    return {
        "empty_ratio": empty_ratio,
        "numeric_ratio": numeric_ratio,
        "unit_ratio": unit_ratio,
    }


def is_index_row(row: List[str]) -> bool:
    non_empty = [cell for cell in row if cell]
    if len(non_empty) < 2:
        return False
    allowed_tokens = {"no", "no.", "no:", "sl", "sl.", "s.no", "s.no.", "s.no:"}
    for cell in non_empty:
        token = cell.strip().lower()
        if token in allowed_tokens:
            continue
        if not token.isdigit():
            return False
        if len(token) > 2:
            return False
    return True


def should_merge_header(first: List[str], second: List[str]) -> bool:
    stats_first = header_stats(first)
    stats_second = header_stats(second)

    if stats_first["numeric_ratio"] >= 0.6:
        return False
    if stats_second["numeric_ratio"] >= 0.6 and not is_index_row(second):
        return False
    if is_index_row(second):
        return True

    max_cols = max(len(first), len(second))
    overlap = 0
    for i in range(max_cols):
        cell1 = first[i] if i < len(first) else ""
        cell2 = second[i] if i < len(second) else ""
        if cell1 and cell2:
            overlap += 1

    if stats_second["unit_ratio"] >= 0.5:
        return True

    if (stats_first["empty_ratio"] + stats_second["empty_ratio"]) >= 0.3:
        return overlap <= max(1, int(max_cols * 0.2))

    return False


def merge_multi_line_header(rows: List[List[str]]) -> Tuple[List[List[str]], bool]:
    if len(rows) < 2:
        return rows, False

    first_row = rows[0]
    second_row = rows[1]
    
    # Check if first row starts with a number (like "1. No") - this is likely a data row
    if first_row and len(first_row) > 0:
        first_cell = normalize_cell(first_row[0])
        if first_cell and first_cell.strip()[0].isdigit() and "." in first_cell:
            # Check if it's a pattern like "1.", "2.", etc. (data row)
            try:
                num_part = first_cell.split(".")[0].strip()
                if num_part.isdigit() and int(num_part) <= 10:
                    # This is likely a data row, not a header
                    return rows, False
            except (ValueError, IndexError):
                pass
    
    # Check if second row is actually a data row (has many numeric values)
    second_row_numeric_ratio = header_stats(second_row)["numeric_ratio"]
    if second_row_numeric_ratio > 0.5:
        # Likely a data row, don't merge
        return rows, False
    
    if not should_merge_header(first_row, second_row):
        return rows, False

    max_cols = max(len(first_row), len(second_row))
    merged_header: List[str] = []
    for i in range(max_cols):
        cell1 = first_row[i] if i < len(first_row) else ""
        cell2 = second_row[i] if i < len(second_row) else ""
        
        merged_header.append(f"{cell1} {cell2}".strip())

    return [merged_header] + rows[2:], True


def format_header_with_units(header_text: str) -> str:
    """Format header text to put units in parentheses when appropriate."""
    if not header_text:
        return header_text
    
    # Pattern: "FW1.P % Filled" -> "FW1.P (% Filled)"
    # Pattern: "Weight (MT)" -> keep as is
    # Pattern: "Draft Aft (m)" -> keep as is
    
    # Check if it already has parentheses
    if "(" in header_text and ")" in header_text:
        return header_text
    
    # Check for "% Filled" pattern
    if "% Filled" in header_text:
        parts = header_text.split("% Filled")
        if len(parts) == 2 and parts[0].strip():
            return f"{parts[0].strip()} (% Filled)"
    
    # Check for standalone unit-like text at the end
    words = header_text.split()
    if len(words) >= 2:
        last_word = words[-1]
        if is_unit_like(last_word) and last_word not in ["Filled", "filled"]:
            # Check if it's a unit that should be in parentheses
            if last_word in ["MT", "m", "deg", "kN", "cm", "cm2", "cm3", "m-deg", "kN/cm^2"]:
                base = " ".join(words[:-1])
                return f"{base} ({last_word})"
    
    return header_text


def compact_header_and_units(rows: List[List[str]]) -> Tuple[List[List[str]], bool]:
    if len(rows) < 2:
        return rows, False

    num_cols = max(len(row) for row in rows)
    padded = [row + [""] * (num_cols - len(row)) for row in rows]
    header = padded[0]
    data_rows = padded[1:]

    data_has_value = [any(row[col] for row in data_rows) for col in range(num_cols)]
    header_has_value = [bool(header[col]) for col in range(num_cols)]

    changed = False
    for col in range(num_cols - 1):
        if data_has_value[col] and not header_has_value[col]:
            if header_has_value[col + 1] and not data_has_value[col + 1]:
                header[col] = header[col + 1]
                header[col + 1] = ""
                header_has_value[col] = True
                header_has_value[col + 1] = False
                changed = True

    for col in range(num_cols):
        if header_has_value[col] and not data_has_value[col]:
            target = None
            for left in range(col - 1, -1, -1):
                if data_has_value[left]:
                    target = left
                    break
            if target is None:
                for right in range(col + 1, num_cols):
                    if data_has_value[right]:
                        target = right
                        break
            if target is not None:
                if is_unit_like(header[target]) and not is_unit_like(header[col]):
                    header[target] = f"{header[col]} {header[target]}".strip()
                else:
                    header[target] = f"{header[target]} {header[col]}".strip()
                header[col] = ""
                changed = True

    keep_cols = [i for i in range(num_cols) if data_has_value[i] or header[i]]
    compacted_rows: List[List[str]] = []
    for row in [header] + data_rows:
        compacted_row = [row[i] for i in keep_cols]
        while compacted_row and compacted_row[-1] == "":
            compacted_row.pop()
        if compacted_row:
            compacted_rows.append(compacted_row)
    
    # Format headers with units in parentheses
    if compacted_rows:
        formatted_header = [format_header_with_units(cell) for cell in compacted_rows[0]]
        compacted_rows[0] = formatted_header
        changed = True

    return compacted_rows, changed


def is_layout_table(rows: List[List[str]]) -> bool:
    if len(rows) < 2:
        return True
    max_cols = max(len(row) for row in rows)
    if max_cols < 2:
        return True
    total_cells = sum(len(row) for row in rows)
    if total_cells == 0:
        return True
    non_empty_cells = sum(1 for row in rows for cell in row if cell)
    density = non_empty_cells / total_cells
    if density < 0.15:
        return True
    rows_with_two = sum(1 for row in rows if sum(1 for cell in row if cell) >= 2)
    return rows_with_two < 2


def is_complex_table_structure(rows: List[List[str]]) -> bool:
    """Detect if table has complex 2x2 or multi-section structure."""
    if len(rows) < 3:
        return False
    
    # Check for patterns like "Allowable X | Value | Allowable Y | Value"
    header = rows[0]
    if len(header) == 4:
        # Check if it's a 2x2 structure
        second_row = rows[1] if len(rows) > 1 else []
        if len(second_row) >= 4:
            # Pattern: Label1 | Value1 | Label2 | Value2
            non_empty = sum(1 for cell in header if cell)
            if non_empty == 4:
                # Check if data rows follow the pattern
                data_pattern_match = True
                for row in rows[1:3]:  # Check first 2 data rows
                    if len(row) >= 4:
                        # Should have values in columns 1 and 3
                        if not (row[1] and row[3]):
                            data_pattern_match = False
                            break
                if data_pattern_match:
                    return True
    return False


def split_complex_table(rows: List[List[str]]) -> List[List[List[str]]]:
    """Split complex 2x2 table into two simpler tables."""
    if not is_complex_table_structure(rows):
        return [rows]
    
    header = rows[0]
    if len(header) < 4:
        return [rows]
    
    # Split into two tables: columns 0-1 and columns 2-3
    table1_rows = []
    table2_rows = []
    
    for row in rows:
        table1_row = [row[0] if len(row) > 0 else "", row[1] if len(row) > 1 else ""]
        table2_row = [row[2] if len(row) > 2 else "", row[3] if len(row) > 3 else ""]
        table1_rows.append(table1_row)
        table2_rows.append(table2_row)
    
    return [table1_rows, table2_rows]


def table_to_markdown(table_rows: List[List[str]], table_id: str) -> str:
    if not table_rows:
        return f"### {table_id}\n\n*[No valid data]*\n\n"

    # Check if table should be split
    split_tables = split_complex_table(table_rows)
    
    if len(split_tables) > 1:
        # Multiple tables - format each separately
        md_parts = []
        for idx, sub_table in enumerate(split_tables):
            sub_id = f"{table_id}-{idx+1}" if len(split_tables) > 1 else table_id
            md_parts.append(table_to_markdown_single(sub_table, sub_id))
        return "\n".join(md_parts)
    else:
        return table_to_markdown_single(table_rows, table_id)


def table_to_markdown_single(table_rows: List[List[str]], table_id: str) -> str:
    """Convert a single table to markdown format."""
    if not table_rows:
        return f"### {table_id}\n\n*[No valid data]*\n\n"

    header = table_rows[0]
    md_lines = [f"### {table_id}\n"]
    md_lines.append("| " + " | ".join(header) + " |")
    md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")

    for row in table_rows[1:]:
        padded_row = row + [""] * (len(header) - len(row))
        md_lines.append("| " + " | ".join(padded_row[: len(header)]) + " |")

    md_lines.append("")
    return "\n".join(md_lines)


def extract_pdf_to_markdown_improved(
    pdf_path: Path, keep_layout: bool = False, keep_layout_tables: bool = False
) -> Dict[str, Any]:
    meta: Dict[str, Any] = {
        "source": str(pdf_path),
        "type": "pdf",
        "pages": 0,
        "parsed_at": utc_now_iso(),
        "ocr": {"used": False, "engine": "none", "lang": None},
    }

    text_parts: List[str] = []
    tables: List[Dict[str, Any]] = []
    warnings: List[str] = []
    metrics = {
        "tables_found": 0,
        "tables_kept": 0,
        "tables_skipped": 0,
        "header_merged": 0,
        "header_compacted": 0,
    }

    with pdfplumber.open(pdf_path) as pdf:
        meta["pages"] = len(pdf.pages)

        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(
                x_tolerance=1.5 if keep_layout else 3.0,
                y_tolerance=1.5 if keep_layout else 3.0,
            )
            if text:
                text_parts.append(f"## Page {page_index}\n\n{text}")
            else:
                text_parts.append(f"## Page {page_index}\n\n*[No text content]*")

            try:
                page_tables = page.extract_tables() or []
            except Exception as exc:
                warnings.append(f"Table extraction failed on page {page_index}: {exc}")
                continue

            for table_index, table in enumerate(page_tables):
                metrics["tables_found"] += 1
                if not table:
                    metrics["tables_skipped"] += 1
                    continue

                cleaned_rows, empty_cols = clean_table_rows(table)
                cleaned_rows, merged = merge_multi_line_header(cleaned_rows)
                if merged:
                    metrics["header_merged"] += 1
                cleaned_rows, compacted = compact_header_and_units(cleaned_rows)
                if compacted:
                    metrics["header_compacted"] += 1
                
                # Fix S.No. column positioning
                if len(cleaned_rows) > 1:
                    cleaned_rows = fix_sno_column(cleaned_rows)
                
                # Clean data rows: remove label text and stage indices
                if len(cleaned_rows) > 1:
                    header = cleaned_rows[0]
                    cleaned_data = [header]
                    for data_row in cleaned_rows[1:]:
                        cleaned_data_row = clean_data_row_labels(data_row, header)
                        
                        # Ensure row length matches header (pad if needed)
                        while len(cleaned_data_row) < len(header):
                            cleaned_data_row.append("")
                        
                        # Remove empty cells at the end (but keep structure)
                        while len(cleaned_data_row) > len(header) and cleaned_data_row[-1] == "":
                            cleaned_data_row.pop()
                        
                        # Keep row if it has any non-empty cells
                        if any(cell for cell in cleaned_data_row):
                            cleaned_data.append(cleaned_data_row[:len(header)])
                    cleaned_rows = cleaned_data
                
                # Additional pass: fill remaining merged cell values
                # This handles cases where Allowable values need to be repeated
                if len(cleaned_rows) > 2:
                    cleaned_rows = fill_merged_cells_final_pass(cleaned_rows)

                if cleaned_rows and not keep_layout_tables and is_layout_table(cleaned_rows):
                    metrics["tables_skipped"] += 1
                    warnings.append(f"Layout-like table skipped on page {page_index} index {table_index}")
                    continue

                if cleaned_rows:
                    metrics["tables_kept"] += 1
                    original_cols = len(table[0]) if table and table[0] else 0
                    cleaned_cols = len(cleaned_rows[0]) if cleaned_rows else 0
                    tables.append(
                        {
                            "page": page_index,
                            "index": table_index,
                            "rows": cleaned_rows,
                            "original_cols": original_cols,
                            "cleaned_cols": cleaned_cols,
                            "removed_empty_cols": empty_cols,
                            "header_merged": merged,
                        }
                    )
                    table_id = f"Table {page_index}-{table_index}"
                    text_parts.append(table_to_markdown(cleaned_rows, table_id))
                else:
                    metrics["tables_skipped"] += 1

    markdown_content = "\n\n".join(text_parts).strip()

    metadata_header = (
        f"# {pdf_path.stem}\n\n"
        f"**Source:** `{pdf_path.name}`  \n"
        f"**Type:** PDF  \n"
        f"**Pages:** {meta['pages']}  \n"
        f"**Parsed At:** {meta['parsed_at']}  \n"
        f"**Tables Found:** {metrics['tables_found']}  \n\n"
        "---\n\n"
    )

    full_markdown = metadata_header + markdown_content

    return {
        "meta": meta,
        "text": "\n\n".join(text_parts).replace("## ", "").replace("### ", "").strip(),
        "markdown": full_markdown,
        "tables": tables,
        "warnings": warnings,
        "metrics": metrics,
    }


def update_run_report(report_path: Path, new_run: Dict[str, Any]) -> None:
    existing: Any = None
    if report_path.exists():
        try:
            existing = json.loads(report_path.read_text(encoding="utf-8"))
        except Exception:
            existing = None

    history: List[Dict[str, Any]] = []
    if isinstance(existing, dict):
        prior = {k: v for k, v in existing.items() if k != "history"}
        if prior:
            history.append(prior)
        if isinstance(existing.get("history"), list):
            history = existing["history"]
    elif isinstance(existing, list):
        history = existing

    history.append(new_run)
    updated = dict(new_run)
    updated["history"] = history
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(updated, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert PDF to Markdown with improved table extraction."
    )
    parser.add_argument("pdf_file", help="Input PDF file path")
    parser.add_argument(
        "--out",
        dest="out_path",
        default=None,
        help="Output Markdown file path (default: out/<stem>_improved.md)",
    )
    parser.add_argument("--json", dest="write_json", action="store_true", help="Write JSON output")
    parser.add_argument(
        "--report",
        dest="report_path",
        default="out/_run_report.json",
        help="Run report path",
    )
    parser.add_argument(
        "--keep-layout",
        dest="keep_layout",
        action="store_true",
        help="Keep text layout during extraction",
    )
    parser.add_argument(
        "--keep-layout-tables",
        dest="keep_layout_tables",
        action="store_true",
        help="Do not filter layout-like tables",
    )
    args = parser.parse_args()

    pdf_path = Path(args.pdf_file)
    if not pdf_path.exists():
        print(f"Error: File not found: {pdf_path}")
        return 1
    if pdf_path.suffix.lower() != ".pdf":
        print(f"Error: Not a PDF file: {pdf_path}")
        return 1

    if args.out_path:
        output_path = Path(args.out_path)
    else:
        output_path = Path("out") / f"{pdf_path.stem}_improved.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    report_path = Path(args.report_path)
    start = time.time()
    status = "success"
    error_message = None
    result: Dict[str, Any] = {}

    print(f"Converting {pdf_path.name} to Markdown (improved table extraction)...")
    try:
        result = extract_pdf_to_markdown_improved(
            pdf_path, keep_layout=args.keep_layout, keep_layout_tables=args.keep_layout_tables
        )
        output_path.write_text(result["markdown"], encoding="utf-8")
        print(f"[OK] Markdown saved to: {output_path}")

        if args.write_json:
            json_path = output_path.with_suffix(".json")
            json_path.write_text(
                json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            print(f"[OK] JSON saved to: {json_path}")
    except Exception as exc:
        status = "failed"
        error_message = str(exc)
        print(f"Error during conversion: {exc}")

    elapsed = time.time() - start
    run_report = {
        "run_at": utc_now_iso(),
        "task": "pdf_to_md_improved",
        "status": status,
        "input": str(pdf_path),
        "output": str(output_path),
        "report": str(report_path),
        "elapsed_sec": round(elapsed, 3),
        "tables_found": result.get("metrics", {}).get("tables_found", 0),
        "tables_kept": result.get("metrics", {}).get("tables_kept", 0),
        "tables_skipped": result.get("metrics", {}).get("tables_skipped", 0),
        "header_merged": result.get("metrics", {}).get("header_merged", 0),
        "plugin_keys": ["tables.pdfplumber"],
        "warnings": result.get("warnings", []),
        "error": error_message,
    }
    update_run_report(report_path, run_report)

    if status != "success":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
