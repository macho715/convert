
from __future__ import annotations
from pathlib import Path
from .errors import IoError

def read_text(path: Path, encodings: list[str]) -> str:
    last_err = None
    for enc in encodings:
        try:
            return path.read_text(encoding=enc, errors="strict")
        except Exception as e:
            last_err = e
    raise IoError(f"failed to read {path}") from last_err

def ensure_dir(path: Path) -> None:
    try:
        path.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise IoError(f"failed to create dir {path}") from e
