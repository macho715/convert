# -*- coding: utf-8 -*-

import re
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

# Add scripts to path for OutlookAqsSearcher
scripts_dir = Path(__file__).parents[2] / "scripts"
if str(scripts_dir) not in sys.path:
    sys.path.append(str(scripts_dir))

try:
    from outlook_aqs_searcher import OutlookAqsSearcher
except ImportError:
    st.error("Cannot import OutlookAqsSearcher. Please check the scripts directory.")
    st.stop()

from lib.io import load_threads, load_search
from lib.contracts import assert_threads_contract, assert_search_contract
from lib.formatters import normalize_body_text
from lib.paths import resolve_data_root, resolve_thread_paths, resolve_search_path

try:
    import plotly.express as px  # optional
except Exception:
    px = None


def _to_local_iso(series_dt: pd.Series, tz: str = "Asia/Dubai") -> pd.Series:
    dt = pd.to_datetime(series_dt, errors="coerce", utc=True)
    try:
        return dt.dt.tz_convert(tz).dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return dt.dt.strftime("%Y-%m-%d %H:%M:%S")


def _build_thread_maps(threads: list) -> tuple[dict, dict]:
    thread_members = {}
    row_to_thread = {}
    for t in threads:
        tid = t.get("thread_id", "")
        members = set(t.get("members", []) or [])
        thread_members[tid] = members
        for r in members:
            if r not in row_to_thread:
                row_to_thread[r] = tid
    return thread_members, row_to_thread


def _render_result_card(row, idx, query_terms=None, is_context=False):
    subject = str(row.get("Subject", "(No Subject)"))
    sender = f"{row.get('SenderName', '')} <{row.get('SenderEmail', '')}>"
    date = str(row.get("_delivery_local", ""))
    body = normalize_body_text(row.get("PlainTextBody", ""))
    
    # Snippet generation
    snippet = body[:350] + "..." if len(body) > 350 else body
    snippet = snippet.replace("\n", " ")
    
    # Highlight terms in snippet
    if query_terms:
        for term in query_terms:
            if len(term) > 2:
                # Use a yellow highlight background defined in CSS
                pattern = re.compile(re.escape(term), re.IGNORECASE)
                snippet = pattern.sub(lambda m: f"<strong>{m.group(0)}</strong>", snippet)
                subject = pattern.sub(lambda m: f"<span style='background-color: #fef3c7; padding: 0 2px; border-radius: 2px;'>{m.group(0)}</span>", subject)

    # Tags
    tags_html = ""
    if row.get("case_numbers"):
        tags_html += f"<span class='card-tag tag-case'>{row['case_numbers']}</span>"
    if row.get("sites"):
        tags_html += f"<span class='card-tag tag-site'>{row['sites']}</span>"
    if row.get("lpo"):
        tags_html += f"<span class='card-tag tag-lpo'>{row['lpo']}</span>"
    
    score = row.get("_score", 0)
    if score > 0:
        tags_html += f"<span class='card-tag tag-score'>Score: {score:.1f}</span>"
        # Intent tag example (heuristic)
        if score >= 30.0: # Arbitrary high score threshold for "Top Match"
            tags_html += f"<span class='card-tag tag-intent'>Top Match</span>"

    card_class = "search-card context-card" if is_context else "search-card direct-card"
    
    # Use st.container to group card and button
    with st.container():
        html = f"""
        <div class="{card_class}">
            <div class="card-header">
                <div class="card-subject">{subject}</div>
                <div class="card-meta">{date}</div>
            </div>
            <div class="card-meta" style="margin-bottom: 0.75rem;">
                <span style="font-weight: 500; color: #334155;">From:</span> {sender}
            </div>
            <div class="card-snippet">
                {snippet}
            </div>
            <div class="card-tags">
                {tags_html}
            </div>
        </div>
        """
        st.markdown(html, unsafe_allow_html=True)
        
        # Drilldown button - placed right after the card content
        # We use columns to align it to the right or make it full width
        c1, c2 = st.columns([6, 1])
        with c2:
            if st.button(f"Details", key=f"btn_{idx}_{'ctx' if is_context else 'dir'}", help="View full email details"):
                st.session_state["_selected_row_idx"] = idx
                st.rerun()


st.title("Search")

if "data_root" not in st.session_state:
    st.session_state.data_root = "../outputs/threads_full"
if "selected_thread_id" not in st.session_state:
    st.session_state.selected_thread_id = ""
if "query" not in st.session_state:
    st.session_state.query = ""

root = resolve_data_root(st.session_state.data_root)
threads_path, _edges_path = resolve_thread_paths(root)
search_path = resolve_search_path(root)

if not threads_path.exists():
    st.error("threads.json이 없습니다. 먼저 CLI로 threads.json / edges.csv를 생성하세요.")
    st.stop()

threads = load_threads(threads_path)
assert_threads_contract(threads)
thread_members, row_to_thread = _build_thread_maps(threads)

default_excel = str((Path(__file__).resolve().parents[2] / "data" / "OUTLOOK_HVDC_ALL_rev.xlsx").as_posix())
if "excel_path" not in st.session_state:
    st.session_state.excel_path = default_excel
if "excel_sheet" not in st.session_state:
    st.session_state.excel_sheet = "전체_데이터"

with st.expander("⚙️ Data Source Settings", expanded=False):
    st.session_state.excel_path = st.text_input("Excel Path", value=st.session_state.excel_path)
    st.session_state.excel_sheet = st.text_input("Sheet Name", value=st.session_state.excel_sheet)
    tz = st.text_input("Display Timezone", value="Asia/Dubai")

use_excel = Path(st.session_state.excel_path).exists()

# Search Bar Area
st.markdown("""
<style>
div[data-testid="stTextInput"] input {
    font-size: 1.1rem;
    padding: 0.75rem;
}
</style>
""", unsafe_allow_html=True)

with st.container():
    c1, c2 = st.columns([5, 1])
    with c1:
        q = st.text_input("Search Query", value=st.session_state.query, placeholder="e.g., subject:TR site:AGI", label_visibility="collapsed")
    with c2:
        submitted = st.button("Search", type="primary", use_container_width=True)

    # Filters
    with st.expander("Advanced Filters", expanded=True):
        c3, c4, c5 = st.columns(3)
        with c3:
            context_expand = st.checkbox("Expand Context (Threads)", value=True, help="Include other emails from the same thread")
        with c4:
            fuzzy_enabled = st.checkbox("Fuzzy Matching", value=False, help="Enable typo tolerance (slower)")
        with c5:
            limit = st.number_input("Max Results", min_value=10, max_value=1000, value=50, step=10)

if submitted:
    st.session_state.query = q
    st.session_state["_selected_row_idx"] = None  # Reset selection

    if use_excel:
        with st.spinner("Searching..."):
            searcher = OutlookAqsSearcher(st.session_state.excel_path, st.session_state.excel_sheet, auto_normalize=True)
            results = searcher.search(q, max_results=int(limit), fuzzy=fuzzy_enabled)
            
            results["_delivery_local"] = _to_local_iso(results["DeliveryTime"], tz=tz)
            direct_idx = set(results.index.astype(int).tolist())
            
            # Store results in session for persistence
            st.session_state["_last_results"] = results
            st.session_state["_search_direct_idx"] = sorted(list(direct_idx))
    else:
        st.error("Excel file not found.")
        st.stop()

    # Context Expansion
    expanded_idx = set(direct_idx)
    expanded_threads = set()

    if context_expand and direct_idx:
        for r in list(direct_idx):
            tid = row_to_thread.get(r, "")
            if tid:
                expanded_threads.add(tid)
                expanded_idx |= set(thread_members.get(tid, set()))

    st.session_state["_search_expanded_idx"] = sorted(list(expanded_idx))
    st.session_state["_search_expanded_threads"] = sorted(list(expanded_threads))


# Display Results
if st.session_state.get("_last_results") is not None:
    results = st.session_state["_last_results"]
    direct_idx = st.session_state.get("_search_direct_idx", [])
    expanded_idx = st.session_state.get("_search_expanded_idx", [])
    expanded_threads = st.session_state.get("_search_expanded_threads", [])
    
    # Metrics
    st.markdown("### Results Overview")
    m1, m2, m3 = st.columns(3)
    m1.metric("Direct Matches", f"{len(direct_idx):,}")
    m2.metric("With Context", f"{len(expanded_idx):,}")
    m3.metric("Related Threads", f"{len(expanded_threads):,}")
    
    st.divider()

    # Separate Direct vs Context
    tab1, tab2 = st.tabs(["Direct Matches (Ranked)", "Context / Threads"])
    
    with tab1:
        if results.empty:
            st.info("No results found.")
        else:
            query_terms = [t.lower() for t in st.session_state.query.split() if ":" not in t]
            for idx, row in results.iterrows():
                _render_result_card(row, idx, query_terms, is_context=False)

    with tab2:
        if not context_expand:
            st.caption("Context expansion is disabled.")
        else:
            # Load context rows
            context_ids = sorted(list(set(expanded_idx) - set(direct_idx)))
            if not context_ids:
                st.info("No additional context found.")
            else:
                if use_excel:
                    @st.cache_data
                    def load_full_df(path, sheet):
                        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
                    
                    full_df = load_full_df(st.session_state.excel_path, st.session_state.excel_sheet)
                    context_rows = full_df.loc[context_ids].copy()
                    context_rows["_delivery_local"] = _to_local_iso(context_rows["DeliveryTime"], tz=tz)
                    
                    # Render context as cards too, but with different style
                    query_terms = [t.lower() for t in st.session_state.query.split() if ":" not in t]
                    for idx, row in context_rows.iterrows():
                        _render_result_card(row, idx, query_terms, is_context=True)
                else:
                    st.warning("Excel file missing, cannot load context.")

# Detail View (Modal-like or below)
if st.session_state.get("_selected_row_idx") is not None:
    sel_idx = st.session_state["_selected_row_idx"]
    if use_excel:
        full_df = pd.read_excel(st.session_state.excel_path, sheet_name=st.session_state.excel_sheet, engine="openpyxl")
        row = full_df.loc[sel_idx]
        
        with st.sidebar:
            st.markdown("### 📧 Email Details")
            st.markdown(f"**Subject:** {row.get('Subject')}")
            st.markdown(f"**From:** {row.get('SenderName')} <{row.get('SenderEmail')}>")
            st.markdown(f"**Date:** {row.get('DeliveryTime')}")
            st.divider()
            st.text_area("Body", normalize_body_text(row.get("PlainTextBody", "")), height=500)
            
            tid = row_to_thread.get(sel_idx)
            if tid:
                if st.button("Explore Thread"):
                    st.session_state.selected_thread_id = tid
                    st.switch_page("pages/03_thread_explorer.py")
