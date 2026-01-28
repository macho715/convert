# -*- coding: utf-8 -*-

import io
from pathlib import Path
import pandas as pd
import streamlit as st

from lib.io import load_threads, load_edges
from lib.contracts import assert_threads_contract, assert_edges_contract
from lib.formatters import normalize_body_text
from lib.paths import resolve_data_root, resolve_thread_paths

try:
    import plotly.express as px
    import plotly.graph_objects as go
except Exception:
    px = None
    go = None

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
        "no", "Subject", "SenderName", "SenderEmail", "RecipientTo",
        "DeliveryTime", "PlainTextBody",
        "case_numbers", "hvdc_cases", "primary_case",
        "sites", "primary_site", "site",
        "lpo", "lpo_numbers",
    ]
    for c in must:
        if c not in df.columns:
            df[c] = ""
    return df


def _to_local(series_dt: pd.Series, tz: str = "Asia/Dubai") -> pd.Series:
    dt = pd.to_datetime(series_dt, errors="coerce", utc=True)
    try:
        return dt.dt.tz_convert(tz)
    except Exception:
        return dt


def _download_df(label: str, df: pd.DataFrame, file_name: str):
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(label, data=csv_bytes, file_name=file_name, mime="text/csv")


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
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["size", "confidence"], ascending=[False, False])
    return df


st.title("Thread Explorer")

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

default_excel = str((Path(__file__).resolve().parents[2] / "data" / "OUTLOOK_HVDC_ALL_rev.xlsx").as_posix())
if "excel_path" not in st.session_state:
    st.session_state.excel_path = default_excel
if "excel_sheet" not in st.session_state:
    st.session_state.excel_sheet = "전체_데이터"
if "tz" not in st.session_state:
    st.session_state.tz = "Asia/Dubai"

with st.expander("⚙️ Settings", expanded=False):
    st.session_state.excel_path = st.text_input("Excel Path", value=st.session_state.excel_path)
    st.session_state.excel_sheet = st.text_input("Sheet Name", value=st.session_state.excel_sheet)
    st.session_state.tz = st.text_input("Timezone", value=st.session_state.tz)

use_excel = Path(st.session_state.excel_path).exists()

# Filter Bar
st.markdown("### Thread List")
c1, c2, c3, c4 = st.columns([1, 1, 2, 2])
with c1:
    min_conf = st.slider("Min Confidence", 0.0, 1.0, 0.0, 0.01)
with c2:
    min_size = st.number_input("Min Size", min_value=1, max_value=100000, value=1, step=1)
with c3:
    q = st.text_input("Filter (Subject/Case/Site/LPO)", value="")
with c4:
    sort_mode = st.selectbox("Sort By", options=["Latest (start)", "Size", "Confidence"], index=1)

view = tdf.copy()
view = view[(view["confidence"] >= float(min_conf)) & (view["size"] >= int(min_size))]

if q.strip():
    qq = q.strip().lower()
    view = view[
        view["subject_norm"].fillna("").astype(str).str.lower().str.contains(qq, na=False)
        | view["cases"].fillna("").astype(str).str.lower().str.contains(qq, na=False)
        | view["sites"].fillna("").astype(str).str.lower().str.contains(qq, na=False)
        | view["lpos"].fillna("").astype(str).str.lower().str.contains(qq, na=False)
    ]

if sort_mode == "Confidence":
    view = view.sort_values(["confidence", "size"], ascending=[False, False])
elif sort_mode == "Size":
    view = view.sort_values(["size", "confidence"], ascending=[False, False])
else:
    view = view.sort_values(["start"], ascending=[False])

selected_tid = st.session_state.selected_thread_id
selected_rows = []
try:
    event = st.dataframe(
        view[["thread_id", "size", "confidence", "start", "end", "subject_norm", "cases", "sites", "lpos"]],
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
    )
    selected_rows = (event.selection.rows or []) if event is not None else []
except TypeError:
    selected_rows = []

if selected_rows:
    pos = selected_rows[0]
    selected_tid = view.iloc[pos]["thread_id"]
    st.session_state.selected_thread_id = selected_tid

if not selected_tid:
    options = view["thread_id"].tolist()[:500]
    selected_tid = st.selectbox("Select Thread (Fallback)", options=options, index=0 if options else None)
    st.session_state.selected_thread_id = selected_tid

st.divider()

if not selected_tid:
    st.info("No thread selected.")
    st.stop()

meta = next((t for t in threads if t.get("thread_id") == selected_tid), None)
if not meta:
    st.warning("Thread metadata not found.")
    st.stop()

st.markdown(f"### Thread Details: `{selected_tid}`")
mc1, mc2, mc3, mc4 = st.columns(4)
mc1.metric("Size", f"{int(float(meta.get('size', 0))):,}")
mc2.metric("Confidence", f"{float(meta.get('confidence', 0.0)):.2f}")
mc3.metric("Start", meta.get("start_dt_local", meta.get("start_dt", "")) or "-")
mc4.metric("End", meta.get("end_dt_local", meta.get("end_dt", "")) or "-")

st.caption("Entities")
st.write({
    "cases": meta.get("cases", []),
    "sites": meta.get("sites", []),
    "lpos": meta.get("lpos", []),
})

members = meta.get("members", []) or []
members = [int(x) for x in members]

if use_excel:
    emails = _load_emails_excel(st.session_state.excel_path, st.session_state.excel_sheet)
    emails = _ensure_cols(emails)
    msgs = emails.loc[members].copy() if len(members) else pd.DataFrame()
else:
    msgs = pd.DataFrame()

if msgs.empty:
    st.warning("Excel file missing. Cannot display message details.")
else:
    tz = st.session_state.tz
    msgs["_dt_utc"] = pd.to_datetime(msgs["DeliveryTime"], errors="coerce", utc=True)
    msgs["_dt_local"] = _to_local(msgs["DeliveryTime"], tz=tz)
    msgs["_dt_local_str"] = msgs["_dt_local"].dt.strftime("%Y-%m-%d %H:%M:%S%z")
    msgs["SenderEmail_raw"] = msgs["SenderEmail"].fillna("").astype(str)
    msgs["body_preview_200"] = (
        msgs["PlainTextBody"]
        .fillna("")
        .astype(str)
        .str.replace("_x000D_", " ", regex=False)
        .str.replace("\n", " ")
        .str.slice(0, 200)
    )

    msgs = msgs.sort_values("_dt_utc", ascending=True)

    st.markdown("### Messages (Chronological)")
    show_cols = [
        "_dt_local_str", "SenderEmail_raw", "SenderName", "RecipientTo",
        "Subject", "case_numbers", "hvdc_cases", "primary_site", "sites", "lpo", "lpo_numbers",
        "body_preview_200",
    ]
    msgs_view = msgs[show_cols].copy()

    msg_selected_rows = []
    try:
        event2 = st.dataframe(
            msgs_view,
            use_container_width=True,
            hide_index=False,
            on_select="rerun",
            selection_mode="single-row",
        )
        msg_selected_rows = (event2.selection.rows or []) if event2 is not None else []
    except TypeError:
        msg_selected_rows = []

    if not msg_selected_rows:
        opts = list(msgs_view.index[: min(300, len(msgs_view))])
        pick = st.selectbox(
            "Select Message (Fallback)",
            options=opts,
            format_func=lambda i: f"[{i}] {msgs_view.loc[i,'_dt_local_str']} | {msgs_view.loc[i,'SenderEmail_raw']} | {str(msgs_view.loc[i,'Subject'])[:60]}",
        )
        msg_selected_rows = [opts.index(pick)] if pick in opts else []

    if msg_selected_rows:
        sel_pos = msg_selected_rows[0]
        sel_idx = msgs_view.index[sel_pos]
        st.markdown("### Message Body")
        st.write(f"- Row Index: `{sel_idx}` / no: `{msgs.loc[sel_idx, 'no']}`")
        st.write(f"- Subject: {msgs.loc[sel_idx, 'Subject']}")
        body_text = normalize_body_text(msgs.loc[sel_idx, "PlainTextBody"])
        body_is_long = len(body_text) > 2000
        collapse_long = st.checkbox(
            "Auto-collapse long body",
            value=body_is_long,
            key=f"auto_collapse_body_{sel_idx}",
        )
        with st.expander("Full Body", expanded=not collapse_long):
            st.text_area("PlainTextBody", value=body_text, height=400, disabled=True)

    cdl1, cdl2 = st.columns(2)
    with cdl1:
        _download_df("Download Messages CSV", msgs_view.reset_index().rename(columns={"index": "row_index"}), f"{selected_tid}_messages.csv")
    with cdl2:
        buf = io.StringIO()
        msgs_view.reset_index().rename(columns={"index": "row_index"}).to_json(buf, orient="records", force_ascii=False)
        st.download_button("Download Messages JSON", data=buf.getvalue().encode("utf-8"), file_name=f"{selected_tid}_messages.json", mime="application/json")

st.divider()
st.markdown("### Edges (Parent → Child)")
t_edges = edges[edges["thread_id"] == selected_tid].copy()
if not t_edges.empty:
    if "child_delivery_time" in t_edges.columns:
        t_edges = t_edges.sort_values("child_delivery_time", ascending=True)
    st.dataframe(t_edges, use_container_width=True, height=320)
    _download_df("Download Edges CSV", t_edges, f"{selected_tid}_edges.csv")
else:
    st.info("No edges found for this thread.")

st.divider()
st.markdown("### Visualizations")

tab_timeline, tab_network = st.tabs(["Timeline", "Network Graph"])

with tab_timeline:
    if px is None:
        st.warning("Plotly not installed.")
    else:
        if use_excel and not msgs.empty:
            tl = msgs[["_dt_local", "SenderName", "SenderEmail_raw", "Subject"]].copy()
            tl = tl.reset_index().rename(columns={"index": "row_index"})
            tl = tl.sort_values("_dt_local", ascending=True)
            tl["start"] = tl["_dt_local"]
            tl["end"] = tl["start"].shift(-1)
            tl["end"] = tl["end"].fillna(tl["start"] + pd.Timedelta(minutes=5))
            tl["end"] = tl[["start", "end"]].apply(
                lambda r: min(r["end"], r["start"] + pd.Timedelta(hours=12)), axis=1
            )
            tl["label"] = tl["SenderEmail_raw"].astype(str) + " | " + tl["Subject"].astype(str).str.slice(0, 50)

            unique_senders = tl["SenderName"].unique()
            colors = px.colors.qualitative.Pastel
            color_map = {sender: colors[i % len(colors)] for i, sender in enumerate(unique_senders)}

            fig = px.timeline(
                tl,
                x_start="start",
                x_end="end",
                y="label",
                color="SenderName",
                color_discrete_map=color_map,
                hover_data={"row_index": True, "SenderName": True, "SenderEmail_raw": True, "Subject": True},
                height=max(400, len(tl) * 30)
            )
            fig.update_yaxes(autorange="reversed", title="")
            fig.update_xaxes(title="Time")
            fig.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(family="Inter, sans-serif", size=12, color="#1e293b"),
                margin=dict(l=10, r=10, t=30, b=30),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Excel missing, cannot generate timeline.")

with tab_network:
    if nx is None or go is None:
        st.warning("NetworkX or Plotly not installed.")
    elif t_edges.empty:
        st.info("No edges to visualize.")
    else:
        G = nx.DiGraph()
        # Add nodes
        for m in members:
            G.add_node(m)
        
        # Add edges
        for _, e in t_edges.iterrows():
            G.add_edge(int(e["parent_row"]), int(e["child_row"]), confidence=e.get("confidence", 0.0))

        # Layout
        pos = nx.spring_layout(G, seed=42)
        
        edge_x = []
        edge_y = []
        for edge in G.edges():
            x0, y0 = pos[edge[0]]
            x1, y1 = pos[edge[1]]
            edge_x.append(x0)
            edge_x.append(x1)
            edge_x.append(None)
            edge_y.append(y0)
            edge_y.append(y1)
            edge_y.append(None)

        edge_trace = go.Scatter(
            x=edge_x, y=edge_y,
            line=dict(width=1, color='#888'),
            hoverinfo='none',
            mode='lines')

        node_x = []
        node_y = []
        node_text = []
        for node in G.nodes():
            x, y = pos[node]
            node_x.append(x)
            node_y.append(y)
            # Try to get sender info if available
            if not msgs.empty and node in msgs.index:
                sender = msgs.loc[node, "SenderName"]
                subj = str(msgs.loc[node, "Subject"])[:30]
                node_text.append(f"Row {node}<br>{sender}<br>{subj}")
            else:
                node_text.append(f"Row {node}")

        node_trace = go.Scatter(
            x=node_x, y=node_y,
            mode='markers',
            hoverinfo='text',
            text=node_text,
            marker=dict(
                showscale=False,
                colorscale='YlGnBu',
                reversescale=True,
                color='#2563eb',
                size=15,
                line_width=2))

        fig_net = go.Figure(data=[edge_trace, node_trace],
                     layout=go.Layout(
                        showlegend=False,
                        hovermode='closest',
                        margin=dict(b=20,l=5,r=5,t=40),
                        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        height=500
                        ))
        
        st.plotly_chart(fig_net, use_container_width=True)
