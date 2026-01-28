from __future__ import annotations
import argparse, subprocess, sys, json

def run(cmd: str) -> None:
    print(f"$ {cmd}")
    subprocess.check_call(cmd, shell=True)

def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--apply-precommit", action="store_true")
    p.add_argument("--apply-ci", action="store_true")
    p.add_argument("--python", default="3.13")
    args = p.parse_args()

    try:
        run("git init -b main")
    except Exception:
        pass
    if args.apply_precommit:
        run("pip install pre-commit")
        run("pre-commit install")
        run("pre-commit install --hook-type commit-msg")
    if args.apply-ci:
        print("CI ready: .github/workflows/ci.yml")
    print(json.dumps({"ok": True}, ensure_ascii=False))

if __name__ == "__main__":
    main()
