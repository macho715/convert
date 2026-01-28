#!/usr/bin/env python3
import argparse
import os
import re
import sys

NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")

def read_frontmatter_name(skill_md_path: str) -> str:
    with open(skill_md_path, "r", encoding="utf-8") as f:
        txt = f.read()
    if not txt.startswith("---"):
        return ""
    # naive YAML frontmatter parse: find name: line before second '---'
    fm_end = txt.find("\n---", 3)
    if fm_end == -1:
        return ""
    fm = txt[3:fm_end]
    for line in fm.splitlines():
        if line.strip().startswith("name:"):
            return line.split(":", 1)[1].strip()
    return ""

def validate_skills(root: str):
    problems = []
    skill_roots = [
        os.path.join(root, ".cursor", "skills"),
        os.path.join(root, ".codex", "skills"),
    ]
    for sr in skill_roots:
        if not os.path.isdir(sr):
            continue
        for name in os.listdir(sr):
            skill_dir = os.path.join(sr, name)
            if not os.path.isdir(skill_dir):
                continue
            if not NAME_RE.match(name):
                problems.append(f"[SKILL] invalid folder name: {skill_dir}")
            skill_md = os.path.join(skill_dir, "SKILL.md")
            if not os.path.exists(skill_md):
                problems.append(f"[SKILL] missing SKILL.md: {skill_dir}")
                continue
            fm_name = read_frontmatter_name(skill_md)
            if fm_name and fm_name != name:
                problems.append(f"[SKILL] name mismatch folder({name}) != frontmatter({fm_name}) in {skill_md}")
    return problems

def validate_subagents(root: str):
    problems = []
    agent_dirs = [
        os.path.join(root, ".cursor", "agents"),
        os.path.join(root, ".codex", "agents"),
    ]
    for ad in agent_dirs:
        if not os.path.isdir(ad):
            continue
        for fn in os.listdir(ad):
            if not fn.endswith(".md"):
                continue
            path = os.path.join(ad, fn)
            with open(path, "r", encoding="utf-8") as f:
                head = f.read(200)
            if not head.startswith("---"):
                problems.append(f"[AGENT] missing YAML frontmatter: {path}")
    return problems

def main():
    ap = argparse.ArgumentParser(description="Validate Cursor/Codex agent-skill assets.")
    ap.add_argument("--root", default=".", help="Repo root.")
    args = ap.parse_args()
    root = os.path.abspath(args.root)

    problems = []
    problems += validate_skills(root)
    problems += validate_subagents(root)

    if problems:
        print("VERDICT: FAIL")
        for p in problems:
            print(p)
        sys.exit(1)

    print("VERDICT: PASS")

if __name__ == "__main__":
    main()
