#!/usr/bin/env python3
import argparse
import json
import os
import re
from datetime import datetime

ENTRYPOINT_HINTS = (
    "__main__.py",
    "main.py",
    "app.py",
)

CONFIG_HINTS = (
    "pyproject.toml",
    "requirements.txt",
    "environment.yml",
    "Pipfile",
    "setup.cfg",
)

README_HINTS = ("README.md", "readme.md")

def is_probable_entrypoint(filename: str) -> bool:
    base = os.path.basename(filename)
    if base in ENTRYPOINT_HINTS:
        return True
    if base.endswith("_cli.py"):
        return True
    return False

def scan(root: str):
    modules = []
    for dirpath, dirnames, filenames in os.walk(root):
        # skip common noise
        parts = set(dirpath.split(os.sep))
        if any(p in parts for p in (".git", ".venv", "node_modules", "dist", "build")):
            continue

        hits = {
            "readme": [],
            "configs": [],
            "entrypoints": [],
            "excel": [],
        }

        for fn in filenames:
            if fn in README_HINTS:
                hits["readme"].append(os.path.join(dirpath, fn))
            if fn in CONFIG_HINTS:
                hits["configs"].append(os.path.join(dirpath, fn))
            if fn.lower().endswith((".xlsx", ".xlsm")):
                hits["excel"].append(os.path.join(dirpath, fn))
            if fn.lower().endswith(".py") and is_probable_entrypoint(fn):
                hits["entrypoints"].append(os.path.join(dirpath, fn))

        if any(hits.values()):
            modules.append({
                "path": dirpath,
                "readme": sorted(hits["readme"]),
                "configs": sorted(hits["configs"]),
                "entrypoints": sorted(hits["entrypoints"]),
                "excel": sorted(hits["excel"])[:50],  # cap
            })

    return modules

def main():
    ap = argparse.ArgumentParser(description="CONVERT folder inventory (entrypoints/configs/readmes/excel).")
    ap.add_argument("--root", default=".", help="Root directory to scan.")
    ap.add_argument("--out", default="", help="Write JSON output to file path.")
    args = ap.parse_args()

    payload = {
        "generated_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "root": os.path.abspath(args.root),
        "modules": scan(args.root),
    }

    data = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(data)
    else:
        print(data)

if __name__ == "__main__":
    main()
