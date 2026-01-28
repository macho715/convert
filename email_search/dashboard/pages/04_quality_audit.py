# -*- coding: utf-8 -*-

import io
from pathlib import Path
import pandas as pd
import streamlit as st

from lib.io import load_threads, load_edges
from lib.contracts import assert_threads_contract, assert_edges_contract
from lib.paths import resolve_data_root, resolve_thread_paths

try:
    import plotly.express as px
except Exception:
    px = None

try:
    import networkx as nx
except Exception:
    nx = None


@st.cache_data
def _load_emails_excel(excel_path: str, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    df = df.reset_index(drop=True)
    return df


def _ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    must = [
        "Subject", "SenderName", "SenderEmail", "RecipientTo",
        "DeliveryTime", "PlainTextBody",
        "case_numbers", "hvdc_cases", "primary_case",
        "sites", "primary_site", "site",
        "lpo", "lpo_numbers",
    ]
    for c in must:
        if c not in df.columns:
            df[c] = ""
    return df


def _threads_df(threads: list) -> pd.DataFrame:
    rows = []
    for t in threads:
        members = t.get("members", []) or []
        rows.append({
            "thread_id": t.get("thread_id", ""),
            "size": int(t.get("size", len(members)) if t.get("size") is not None else len(members)),
            "confidence": float(t.get("confidence", 0.0)),
            "subject_norm": t.get("subject_norm", ""),
            "start": t.get("start_dt_local", t.get("start_dt", "")),
            "end": t.get("end_dt_local", t.get("end_dt", "")),
            "cases": ", ".join(t.get("cases", []) or []),
            "sites": ", ".join(t.get("sites", []) or []),
            "lpos": ", ".join(t.get("lpos", []) or []),
            "members": t.get("members", []) or [],
        })
    return pd.DataFrame(rows)


def _download_df(label: str, df: pd.DataFrame, file_name: str):
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(label, data=csv_bytes, file_name=file_name, mime="text/csv")


def _split_csvish(s: str):
    s = (s or "").strip()
    if not s:
        return set()
    return {x.strip().upper() for x in s.split(",") if x.strip()}


def _suspect_false_merge(thread_msgs: pd.DataFrame) -> bool:
    subj = thread_msgs["Subject"].fillna("").astype(str).str.replace(r"^(re|fw|fwd)\s*:\s*", "", regex=True).str.upper()
    subj = subj.str.replace(r"^\[.*?\]\s*", "", regex=True)
    uniq_subj = set([x.strip() for x in subj.tolist() if x.strip()])
    if len(uniq_subj) < 2:
        return False

    cases = []
    sites = []
    lpos = []
    for _, r in thread_msgs.iterrows():
        cases.append(_split_csvish(f"{r.get('case_numbers','')},{r.get('hvdc_cases','')},{r.get('primary_case','')}"))
        sites.append(_split_csvish(f"{r.get('sites','')},{r.get('primary_site','')},{r.get('site','')}"))
        lpos.append(_split_csvish(f"{r.get('lpo','')},{r.get('lpo_numbers','')}"))

    inter_cases = set.intersection(*cases) if cases and all(isinstance(x, set) for x in cases) else set()
    inter_sites = set.intersection(*sites) if sites and all(isinstance(x, set) for x in sites) else set()
    inter_lpos = set.intersection(*lpos) if lpos and all(isinstance(x, set) for x in lpos) else set()

    return (len(inter_cases) == 0) and (len(inter_sites) == 0) and (len(inter_lpos) == 0)


def _has_cycle_fallback(edges_df: pd.DataFrame) -> bool:
    if edges_df.empty:
        return False
    nodes = set(edges_df["parent_row"].tolist()) | set(edges_df["child_row"].tolist())
    indeg = {n: 0 for n in nodes}
    adj = {n: [] for n in nodes}
    for _, e in edges_df.iterrows():
        u = int(e["parent_row"])
        v = int(e["child_row"])
        adj[u].append(v)
        indeg[v] += 1

    q = [n for n in nodes if indeg[n] == 0]
    seen = 0
    while q:
        n = q.pop()
        seen += 1
        for v in adj[n]:
            indeg[v] -= 1
            if indeg[v] == 0:
                q.append(v)
    return seen != len(nodes)


st.title("Quality / Audit")

if "data_root" not in st.session_state:
    st.session_state.data_root = "../outputs/threads_full"
if "selected_thread_id" not in st.session_state:
    st.session_state.selected_thread_id = ""

root = resolve_data_root(st.session_state.data_root)
threads_path, edges_path = resolve_thread_paths(root)

if not threads_path.exists() or not edges_path.exists():
    st.error("threads.json or edges.csv not found. Please run CLI export first.")
    st.stop()

threads = load_threads(threads_path)
edges = load_edges(edges_path)
assert_threads_contract(threads)
assert_edges_contract(edges)

tdf = _threads_df(threads)
total_threads = len(tdf)

default_excel = str((Path(__file__).resolve().parents[2] / "data" / "OUTLOOK_HVDC_ALL_rev.xlsx").as_posix())
if "excel_path" not in st.session_state:
    st.session_state.excel_path = default_excel
if "excel_sheet" not in st.session_state:
    st.session_state.excel_sheet = "전체_데이터"

with st.expander("⚙️ Settings", expanded=False):
    st.session_state.excel_path = st.text_input("Excel Path", value=st.session_state.excel_path)
    st.session_state.excel_sheet = st.text_input("Sheet Name", value=st.session_state.excel_sheet)

use_excel = Path(st.session_state.excel_path).exists()
emails = None
if use_excel:
    emails = _load_emails_excel(st.session_state.excel_path, st.session_state.excel_sheet)
    emails = _ensure_cols(emails)

st.markdown("### Low-confidence Queue")
th = st.slider("Confidence Threshold", 0.0, 1.0, 0.60, 0.01)
low = tdf[tdf["confidence"] < float(th)].copy()
low = low.sort_values(["confidence", "size"], ascending=[True, False])

c1, c2, c3 = st.columns(3)
c1.metric("Total Threads", f"{total_threads:,}")
c2.metric("Low-confidence", f"{len(low):,}")
c3.metric("Ratio", f"{(len(low) / max(total_threads, 1)) * 100:.2f}%")

if px is not None and not tdf.empty:
    st.markdown("### Distributions")
    cc1, cc2 = st.columns(2)
    with cc1:
        fig1 = px.histogram(tdf, x="size", nbins=30, title="Thread Size Distribution", color_discrete_sequence=["#2563eb"])
        fig1.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(t=30, l=10, r=10, b=10))
        st.plotly_chart(fig1, use_container_width=True)
    with cc2:
        fig2 = px.histogram(tdf, x="confidence", nbins=30, title="Confidence Distribution", color_discrete_sequence=["#10b981"])
        fig2.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(t=30, l=10, r=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

st.dataframe(low[["thread_id", "size", "confidence", "start", "end", "subject_norm", "cases", "sites", "lpos"]].head(300), use_container_width=True)
_download_df("Download Low-confidence CSV", low, "low_confidence_threads.csv")

st.divider()

st.markdown("### Edge Cycle Analysis")
st.caption("Detects circular dependencies in thread edges (requires networkx for samples).")

cycle_rows = []
sample_cycles = []

for tid in tdf["thread_id"].tolist():
    te = edges[edges["thread_id"] == tid].copy()
    if te.empty:
        continue

    has_cycle = False
    cycle_edge_sample = None

    if nx is not None:
        G = nx.DiGraph()
        for _, e in te.iterrows():
            u = int(e["parent_row"])
            v = int(e["child_row"])
            G.add_edge(u, v, confidence=float(e.get("confidence", 0.0)))
        try:
            cyc = nx.find_cycle(G, orientation="original")
            has_cycle = True
            cycle_edge_sample = [(a, b) for a, b, *_ in cyc][:10]
        except Exception:
            has_cycle = False
    else:
        has_cycle = _has_cycle_fallback(te)

    if has_cycle:
        cycle_rows.append({
            "thread_id": tid,
            "edges": int(len(te)),
            "size": int(tdf.loc[tdf["thread_id"] == tid, "size"].iloc[0]),
            "confidence": float(tdf.loc[tdf["thread_id"] == tid, "confidence"].iloc[0]),
        })
        if cycle_edge_sample:
            sample_cycles.append({"thread_id": tid, "cycle_edges_sample": cycle_edge_sample})

cycle_df = pd.DataFrame(cycle_rows).sort_values(["edges", "size"], ascending=[False, False]) if cycle_rows else pd.DataFrame()
st.metric("Cycle Threads", f"{len(cycle_df):,}")

if not cycle_df.empty:
    st.dataframe(cycle_df.head(200), use_container_width=True)
    _download_df("Download Cycle Threads CSV", cycle_df, "cycle_threads.csv")

    if sample_cycles:
        st.markdown("### Cycle Samples (Max 10)")
        st.json(sample_cycles[:10], expanded=False)
else:
    st.info("No cycles detected.")

st.divider()

st.markdown("### False-merge Suspects (Split Candidates)")
st.caption("Conservative rule: Diverse subjects + No intersection in Case/Site/LPO")

if not use_excel or emails is None:
    st.warning("Excel file required for false-merge detection.")
    st.stop()

sus_rows = []
target = low.head(500) if not low.empty else tdf.head(200)

for _, r in target.iterrows():
    tid = r["thread_id"]
    members = r["members"] if isinstance(r["members"], list) else []
    if not members:
        continue
    try:
        msgs = emails.loc[[int(x) for x in members]].copy()
    except Exception:
        continue

    if _suspect_false_merge(msgs):
        subj_set = set(
            msgs["Subject"].fillna("").astype(str)
            .str.replace(r"^(re|fw|fwd)\s*:\s*", "", regex=True)
            .str.upper()
            .str.replace(r"^\[.*?\]\s*", "", regex=True)
            .str.strip()
            .tolist()
        )
        subj_set = {s for s in subj_set if s}
        sus_rows.append({
            "thread_id": tid,
            "size": int(r["size"]),
            "confidence": float(r["confidence"]),
            "distinct_subjects": int(len(subj_set)),
            "subject_samples": " | ".join(list(subj_set)[:3]),
            "cases": r["cases"],
            "sites": r["sites"],
            "lpos": r["lpos"],
            "recommendation": "SPLIT Candidate",
        })

sus_df = pd.DataFrame(sus_rows).sort_values(["confidence", "size"], ascending=[True, False]) if sus_rows else pd.DataFrame()
st.metric("Split Candidates", f"{len(sus_df):,}")

if not sus_df.empty:
    st.dataframe(sus_df.head(200), use_container_width=True)
    _download_df("Download Split Candidates CSV", sus_df, "suspected_false_merge_threads.csv")

    pick = st.selectbox("Drilldown to Thread Explorer", options=sus_df["thread_id"].tolist()[:200])
    if st.button("Open Selected Thread"):
        st.session_state.selected_thread_id = pick
        try:
            st.switch_page("pages/03_thread_explorer.py")
        except Exception:
            st.info("Please click Thread Explorer in the sidebar.")
else:
    st.info("No split candidates found.")
