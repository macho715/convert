#!/usr/bin/env python3
"""
Repair malformed HVDC JSON exports.
Targets:
- hvdc logistics status*.json
- HVDC SATUS.JSON

The input files look like JSON but contain:
- doubled quotes ("")
- outer quotes around the whole array
- keys/values split across lines
- stray quote-only lines

This script normalizes the text, reconstructs objects, and rewrites JSON.
Original files are backed up before overwrite.
"""

from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def strip_outer_quotes(text: str) -> str:
    text = text.lstrip("\ufeff")
    # Remove a single outer quote wrapper if present.
    m = re.search(r"\S", text)
    if m:
        i = m.start()
        if text[i] == '"' and i + 1 < len(text) and text[i + 1] in "[{":
            text = text[:i] + text[i + 1 :]
    m = re.search(r"\S(?=\s*$)", text)
    if m:
        j = m.start()
        if text[j] == '"' and j - 1 >= 0 and text[j - 1] in "]}":
            text = text[:j] + text[j + 1 :]
    return text


def remove_control_chars(text: str) -> str:
    # Keep whitespace that is valid in JSON.
    return "".join(ch for ch in text if ch >= " " or ch in "\n\r\t")


def parse_string_token(s: str, start: int) -> Tuple[Optional[str], int]:
    if start >= len(s) or s[start] != '"':
        return None, start
    i = start + 1
    escape = False
    while i < len(s):
        c = s[i]
        if escape:
            escape = False
        else:
            if c == "\\":
                escape = True
            elif c == '"':
                literal = s[start : i + 1]
                try:
                    return json.loads(literal), i + 1
                except json.JSONDecodeError:
                    return None, start
        i += 1
    return None, start


def parse_line(line: str) -> Optional[Tuple[str, str, Optional[str], bool]]:
    s = line.strip()
    if not s:
        return None
    has_comma = s.endswith(",")
    if has_comma:
        s = s[:-1].rstrip()
    if not s:
        return None
    if s == '"' and has_comma:
        return ("string", "", None, True)
    if s == '"':
        return None
    if not s.startswith('"'):
        return None

    key, idx = parse_string_token(s, 0)
    if key is None:
        return None
    while idx < len(s) and s[idx].isspace():
        idx += 1
    if idx < len(s) and s[idx] == ":":
        idx += 1
        while idx < len(s) and s[idx].isspace():
            idx += 1
        if idx < len(s) and s[idx] == '"':
            value, _ = parse_string_token(s, idx)
            if value is None:
                return ("key", key, "", has_comma)
            return ("kv", key, value, has_comma)
        return ("key", key, "", has_comma)
    return ("string", key, None, has_comma)


def normalize_lines(lines: List[str], warnings: List[Dict[str, Any]]) -> List[str]:
    fixed = []
    for idx, line in enumerate(lines, start=1):
        if line.strip() == '"':
            warnings.append(
                {"line": idx, "issue": "quote_only_line", "content": line}
            )
            continue
        # Remove leading quote-space before a key: '" "KEY":' -> '"KEY":'
        line = re.sub(r'^(\s*)"\s+"(?=[^"]*":)', r'\1"', line)
        # Remove stray extra quote at end of a value.
        line = re.sub(r'(":\s*"[^"]*)""(\s*,?\s*)$', r'\1"\2', line)
        fixed.append(line)
    return fixed


def extract_broken_key_pairs(text: str) -> List[Tuple[str, str]]:
    # Detect keys split at a broken "\n" escape, e.g.
    # "JDN\" + newline + "nWaterfront": -> "JDN\nWaterfront"
    pattern = r'"([^"]*)\\\"\s*\r?\n\s*\"n([^"]*)\"\s*:'
    pairs = []
    for match in re.finditer(pattern, text):
        prefix = match.group(1)
        suffix = match.group(2)
        pairs.append((prefix, suffix))
    # Deduplicate while keeping order.
    seen = set()
    unique = []
    for prefix, suffix in pairs:
        key = (prefix, suffix)
        if key in seen:
            continue
        seen.add(key)
        unique.append(key)
    return unique


def apply_broken_key_fixes(
    data: List[Dict[str, Any]],
    pairs: List[Tuple[str, str]],
    warnings: List[Dict[str, Any]],
) -> None:
    if not pairs:
        return
    for obj in data:
        for prefix, suffix in pairs:
            full_key = f"{prefix}\n{suffix}"
            broken_key = f"n{suffix}"
            if full_key in obj or broken_key not in obj:
                continue
            obj[full_key] = obj.pop(broken_key)
            warnings.append(
                {
                    "issue": "fixed_split_newline_key",
                    "content": broken_key,
                    "fixed": full_key,
                }
            )


def build_objects(lines: List[str], warnings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    data: List[Dict[str, Any]] = []
    obj: Optional[Dict[str, Any]] = None
    pending_key_prefix: Optional[str] = None
    pending_kv: Optional[Tuple[str, str]] = None

    for idx, raw in enumerate(lines, start=1):
        s = raw.strip()
        if not s:
            continue
        if s.startswith("[") or s in ["]", "],"]:
            continue
        if s.startswith("{"):
            if obj is not None:
                data.append(obj)
                warnings.append(
                    {"line": idx, "issue": "nested_object_start", "content": raw}
                )
            obj = {}
            continue
        if s.startswith("}"):
            if obj is None:
                obj = {}
            if pending_kv is not None:
                key, value = pending_kv
                obj[key] = value
                pending_kv = None
            if pending_key_prefix is not None:
                warnings.append(
                    {
                        "line": idx,
                        "issue": "dangling_key_prefix",
                        "content": pending_key_prefix,
                    }
                )
                pending_key_prefix = None
            data.append(obj)
            obj = None
            continue

        if obj is None:
            obj = {}

        parsed = parse_line(raw)
        if parsed is None:
            warnings.append({"line": idx, "issue": "unparsed_line", "content": raw})
            continue

        kind, key, value, has_comma = parsed

        if kind == "kv":
            if pending_key_prefix is not None:
                key = pending_key_prefix + key
                pending_key_prefix = None
            if pending_kv is not None:
                prev_key, prev_value = pending_kv
                obj[prev_key] = prev_value
                warnings.append(
                    {
                        "line": idx,
                        "issue": "flush_pending_value",
                        "content": prev_key,
                    }
                )
                pending_kv = None
            if has_comma:
                obj[key] = value if value is not None else ""
            else:
                pending_kv = (key, value if value is not None else "")
        elif kind == "string":
            if pending_kv is not None:
                prev_key, prev_value = pending_kv
                prev_value += key
                if has_comma:
                    obj[prev_key] = prev_value
                    pending_kv = None
                else:
                    pending_kv = (prev_key, prev_value)
            elif pending_key_prefix is not None:
                obj[pending_key_prefix] = key
                pending_key_prefix = None
                if not has_comma:
                    pending_kv = (key, "")
            else:
                pending_key_prefix = key
        else:  # kind == "key"
            if pending_key_prefix is not None:
                key = pending_key_prefix + key
                pending_key_prefix = None
            pending_kv = (key, "")
            if has_comma:
                obj[key] = ""
                pending_kv = None

    if obj is not None:
        if pending_kv is not None:
            key, value = pending_kv
            obj[key] = value
        data.append(obj)
    return data


def repair_content(text: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], bool]:
    has_non_ascii = any(ord(ch) > 127 for ch in text)
    text = strip_outer_quotes(text)
    text = text.replace('""', '"')
    broken_key_pairs = extract_broken_key_pairs(text)
    # Fix broken \n escape sequences split by a newline.
    text = re.sub(r'\\\"\\s*\\r?\\n\\s*\"n', r'\\n', text)
    text = remove_control_chars(text)
    lines = text.splitlines()
    warnings: List[Dict[str, Any]] = []
    lines = normalize_lines(lines, warnings)
    data = build_objects(lines, warnings)
    apply_broken_key_fixes(data, broken_key_pairs, warnings)
    return data, warnings, has_non_ascii


def next_backup_path(path: Path) -> Path:
    base = path.with_name(path.name + ".backup")
    if not base.exists():
        return base
    for i in range(1, 1000):
        candidate = path.with_name(f"{path.name}.backup.{i}")
        if not candidate.exists():
            return candidate
    return base


def find_targets(base_dir: Path) -> List[Path]:
    targets: List[Path] = []
    for path in base_dir.glob("hvdc logistics status*.json"):
        name_lower = path.name.lower()
        if any(
            token in name_lower
            for token in [".backup", ".debug", "_fixed_preview", "_fixed"]
        ):
            continue
        targets.append(path)
    hvdc_status = base_dir / "HVDC SATUS.JSON"
    if hvdc_status.exists():
        targets.append(hvdc_status)
    return sorted(targets)


def write_json(path: Path, data: Any, ascii_only: bool) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=ascii_only, indent=2)


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent / "AGI TR 1-6 Transportation Master Gantt Chart"
    out_dir = Path(__file__).resolve().parent.parent / "out"
    out_dir.mkdir(exist_ok=True)
    report_path = out_dir / "json_repair_report.json"
    run_report_path = out_dir / "_run_report.json"

    started = time.time()
    targets = find_targets(base_dir)

    results: List[Dict[str, Any]] = []
    for path in targets:
        entry: Dict[str, Any] = {"file": str(path)}
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
            data, warnings, has_non_ascii = repair_content(text)
            backup_path = next_backup_path(path)
            backup_path.write_text(text, encoding="utf-8")
            write_json(path, data, ascii_only=not has_non_ascii)
            # Verify JSON
            with open(path, "r", encoding="utf-8") as f:
                json.load(f)
            entry.update(
                {
                    "status": "ok",
                    "items": len(data),
                    "backup": str(backup_path),
                    "warnings": warnings[:20],
                    "warnings_count": len(warnings),
                }
            )
        except Exception as exc:
            entry.update({"status": "failed", "error": str(exc)})
        results.append(entry)

    elapsed = time.time() - started
    report = {
        "task": "repair_hvdc_json",
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "base_dir": str(base_dir),
        "targets": [str(p) for p in targets],
        "results": results,
    }
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    run_report = {
        "task": "repair_hvdc_json",
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(started)),
        "elapsed_seconds": round(elapsed, 3),
        "files_total": len(targets),
        "files_ok": [r["file"] for r in results if r.get("status") == "ok"],
        "files_failed": [
            {"file": r["file"], "error": r.get("error")}
            for r in results
            if r.get("status") != "ok"
        ],
        "warnings_count": sum(r.get("warnings_count", 0) for r in results),
    }
    run_report_path.write_text(json.dumps(run_report, indent=2), encoding="utf-8")

    print(f"Report: {report_path}")
    print(f"Run report: {run_report_path}")


if __name__ == "__main__":
    main()
