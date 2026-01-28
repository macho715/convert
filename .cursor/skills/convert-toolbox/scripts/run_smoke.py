#!/usr/bin/env python3
import argparse
import os
import subprocess
import sys

def run(cmd, cwd):
    p = subprocess.run(cmd, cwd=cwd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    return p.returncode, p.stdout

def has_pytest(root: str) -> bool:
    # heuristic: tests/ or pytest.ini or pyproject has [tool.pytest]
    if os.path.isdir(os.path.join(root, "tests")):
        return True
    for fn in ("pytest.ini", "pyproject.toml"):
        if os.path.exists(os.path.join(root, fn)):
            return True
    return False

def main():
    ap = argparse.ArgumentParser(description="Conservative smoke runner: compileall + optional pytest.")
    ap.add_argument("--root", default=".", help="Project root.")
    args = ap.parse_args()

    root = os.path.abspath(args.root)

    checks = []

    rc, out = run([sys.executable, "-m", "compileall", "-q", "."], cwd=root)
    checks.append(("compileall", rc, f"{sys.executable} -m compileall -q .", out[-2000:]))

    if has_pytest(root):
        rc2, out2 = run([sys.executable, "-m", "pytest", "-q"], cwd=root)
        checks.append(("pytest", rc2, f"{sys.executable} -m pytest -q", out2[-2000:]))

    verdict = "PASS" if all(rc == 0 for _, rc, _, _ in checks) else "FAIL"
    print(f"VERDICT: {verdict}")
    print("| Check | Result | Command | Notes |")
    print("| --- | --- | --- | --- |")
    for name, rcx, cmd, notes in checks:
        res = "PASS" if rcx == 0 else f"FAIL({rcx})"
        safe_notes = notes.replace("\n", " ")[:300]
        print(f"| {name} | {res} | `{cmd}` | {safe_notes} |")

    if verdict != "PASS":
        sys.exit(1)

if __name__ == "__main__":
    main()
