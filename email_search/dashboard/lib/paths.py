# -*- coding: utf-8 -*-

from pathlib import Path


def resolve_data_root(raw_path: str) -> Path:
    """
    Resolve data_root into an absolute path using the project root.
    """
    if not raw_path:
        return Path()
    p = Path(raw_path).expanduser()
    if p.is_absolute():
        return p
    dashboard_dir = Path(__file__).resolve().parents[1]
    return (dashboard_dir / p).resolve()


def resolve_thread_paths(root: Path) -> tuple[Path, Path]:
    """
    Resolve threads.json / edges.csv in either direct or nested layout.
    """
    direct_threads = root / "threads.json"
    direct_edges = root / "edges.csv"
    if direct_threads.exists():
        return direct_threads, direct_edges
    nested_threads = root / "threads" / "threads.json"
    nested_edges = root / "threads" / "edges.csv"
    return nested_threads, nested_edges


def resolve_search_path(root: Path) -> Path:
    """
    Resolve search_result.csv in either direct or nested layout.
    """
    direct = root / "search_result.csv"
    if direct.exists():
        return direct
    return root / "searches" / "search_result.csv"
