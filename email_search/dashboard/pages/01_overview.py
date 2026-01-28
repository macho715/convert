# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd

from lib.io import load_threads, load_edges
from lib.contracts import assert_threads_contract, assert_edges_contract
from lib.paths import resolve_data_root, resolve_thread_paths

try:
    import plotly.express as px
except Exception:
    px = None


st.title("Overview")

if "data_root" not in st.session_state:
    st.session_state.data_root = "../outputs/threads_full"

root = resolve_data_root(st.session_state.data_root)
threads_path, edges_path = resolve_thread_paths(root)

with st.expander("⚙️ Data Source Settings", expanded=False):
    data_root_options = {
        "threads_full (Full Run)": "../outputs/threads_full",
        "threads_test_v2 (Sample)": "../outputs/threads_test_v2",
        "threads_test (Sample)": "../outputs/threads_test",
        "threads (Default)": "../outputs/threads",
    }
    selected_root = st.selectbox(
        "Select Data Source",
        options=list(data_root_options.keys()),
        index=0,
    )
    if st.button("Apply"):
        st.session_state.data_root = data_root_options[selected_root]
        st.rerun()

    st.caption(f"Current Path: `{st.session_state.data_root}`")

if threads_path.exists() and edges_path.exists():
    try:
        threads = load_threads(threads_path)
        edges = load_edges(edges_path)
        assert_threads_contract(threads)
        assert_edges_contract(edges)

        total_threads = len(threads)
        total_edges = len(edges)
        total_messages = sum(len(t.get("members", [])) for t in threads)
        avg_thread_size = total_messages / max(total_threads, 1)
        confidences = [t.get("confidence", 0.0) for t in threads]
        avg_confidence = sum(confidences) / max(total_threads, 1) if confidences else 0.0

        # Metrics Row
        st.markdown("### Key Metrics")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Threads", f"{total_threads:,}")
        col2.metric("Total Messages", f"{total_messages:,}")
        col3.metric("Avg Thread Size", f"{avg_thread_size:.1f}")
        col4.metric("Avg Confidence", f"{avg_confidence:.2f}")

        st.divider()

        if px is not None and threads:
            st.markdown("### Distributions")
            thread_sizes = [len(t.get("members", [])) for t in threads]
            col1, col2 = st.columns(2)
            with col1:
                fig1 = px.histogram(
                    x=thread_sizes,
                    nbins=30,
                    title="Thread Size Distribution",
                    labels={"x": "Thread Size", "y": "Count"},
                    color_discrete_sequence=["#2563eb"]
                )
                fig1.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(t=30, l=10, r=10, b=10))
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                fig2 = px.histogram(
                    x=confidences,
                    nbins=30,
                    title="Confidence Distribution",
                    labels={"x": "Confidence", "y": "Count"},
                    color_discrete_sequence=["#10b981"]
                )
                fig2.update_layout(plot_bgcolor="white", paper_bgcolor="white", margin=dict(t=30, l=10, r=10, b=10))
                st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### Top Threads (by size)")
        thread_data = []
        for t in threads:
            members = t.get("members", []) or []
            thread_data.append(
                {
                    "Thread ID": t.get("thread_id", ""),
                    "Size": len(members),
                    "Confidence": f"{t.get('confidence', 0.0):.2f}",
                    "Subject": (t.get("subject_norm", "") or "")[:60],
                    "Entities": ", ".join((t.get("cases", []) or [])[:3]),
                }
            )

        thread_df = pd.DataFrame(thread_data).sort_values("Size", ascending=False)
        st.dataframe(
            thread_df.head(20), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Thread ID": st.column_config.TextColumn("Thread ID", width="medium"),
                "Subject": st.column_config.TextColumn("Subject", width="large"),
                "Size": st.column_config.NumberColumn("Size", format="%d"),
            }
        )

    except Exception as exc:
        st.error(f"Failed to load data: {exc}")
        st.code(f"Path: {threads_path}\n{edges_path}")
else:
    st.warning("Data files not found.")
    st.code(
        f"""
Current Path: {st.session_state.data_root}
Expected:
  - {threads_path}
  - {edges_path}

Solution:
1) Run CLI export:
   python email_search/scripts/run_full_export.py --excel ... --out ...
2) Select correct path in settings.
"""
    )
