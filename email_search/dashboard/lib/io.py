# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from pathlib import Path
from typing import List

import pandas as pd


def load_threads(path: Path) -> List[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return data if isinstance(data, list) else []


def load_edges(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)


def load_search(path: Path) -> pd.DataFrame:
    return pd.read_csv(path)
