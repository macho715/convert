# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import List

import pandas as pd


def assert_threads_contract(threads: List[dict]) -> None:
    if not isinstance(threads, list):
        raise ValueError("threads must be a list")
    if not threads:
        return
    sample = threads[0]
    if not isinstance(sample, dict):
        raise ValueError("threads elements must be dict")
    if "thread_id" not in sample or "members" not in sample:
        raise ValueError("threads must include thread_id and members")


def assert_edges_contract(edges: pd.DataFrame) -> None:
    required = {"thread_id", "parent_row", "child_row"}
    missing = [c for c in required if c not in edges.columns]
    if missing:
        raise ValueError(f"edges missing columns: {missing}")


def assert_search_contract(search_df: pd.DataFrame) -> None:
    if not isinstance(search_df, pd.DataFrame):
        raise ValueError("search must be a DataFrame")
