# -*- coding: utf-8 -*-
import json
import random
from pathlib import Path
import pandas as pd
import streamlit as st

# Reuse logic from other pages if possible, or reimplement simply
def _load_emails_excel(excel_path: str, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    # Ensure we have an index that matches the row numbers used in threads.json
    # threads.json uses 0-based index of the dataframe.
    # We must ensure the dataframe is loaded exactly the same way.
    df = df.reset_index(drop=True)
    return df

def _load_existing_labels(path: Path) -> list:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except:
            return []
    return []

def _save_labels(path: Path, labels: list):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(labels, indent=2, ensure_ascii=False), encoding="utf-8")

st.set_page_config(page_title="Labeling Tool", layout="wide")
st.title("Thread Labeling Tool")

# --- Sidebar / Config ---
if "data_root" not in st.session_state:
    st.session_state.data_root = "../outputs/threads_full"
if "excel_path" not in st.session_state:
    # Try to find the default path
    default_excel = Path(__file__).resolve().parents[2] / "data" / "OUTLOOK_HVDC_ALL_rev.xlsx"
    st.session_state.excel_path = str(default_excel) if default_excel.exists() else ""
if "excel_sheet" not in st.session_state:
    st.session_state.excel_sheet = "전체_데이터"

with st.sidebar:
    st.header("Settings")
    excel_path = st.text_input("Excel Path", st.session_state.excel_path)
    sheet_name = st.text_input("Sheet Name", st.session_state.excel_sheet)
    
    if st.button("Load Data"):
        st.session_state.excel_path = excel_path
        st.session_state.excel_sheet = sheet_name
        st.rerun()

    st.divider()
    st.info("Labels are saved to `data/evaluation_set.json`")

# --- Main Logic ---
LABEL_FILE = Path(__file__).resolve().parents[2] / "data" / "evaluation_set.json"

if "labels" not in st.session_state:
    st.session_state.labels = _load_existing_labels(LABEL_FILE)

if "current_pair" not in st.session_state:
    st.session_state.current_pair = None

# Load Data
if not st.session_state.excel_path or not Path(st.session_state.excel_path).exists():
    st.warning("Please configure and load the Excel file in the sidebar.")
    st.stop()

@st.cache_data
def get_data(path, sheet):
    return _load_emails_excel(path, sheet)

try:
    df = get_data(st.session_state.excel_path, st.session_state.excel_sheet)
except Exception as e:
    st.error(f"Failed to load Excel: {e}")
    st.stop()

# Load Threads (for sampling)
THREADS_FILE = Path(st.session_state.data_root) / "threads.json"
if THREADS_FILE.exists():
    threads_data = json.loads(THREADS_FILE.read_text(encoding="utf-8"))
else:
    st.warning(f"threads.json not found at {THREADS_FILE}. Sampling will be random.")
    threads_data = []

# --- Sampling Logic ---
def get_next_pair():
    # Strategy: 50% Same Thread (Verify), 50% Random (Negative)
    # But for now, let's focus on verifying the *computed* threads.
    
    if not threads_data:
        # Fallback: Random pair
        idx1 = random.randint(0, len(df) - 1)
        idx2 = random.randint(0, len(df) - 1)
        return (idx1, idx2, "random")

    # Pick a thread with > 1 members
    candidates = [t for t in threads_data if len(t["members"]) > 1]
    if not candidates:
        return (0, 0, "error")
    
    thread = random.choice(candidates)
    members = thread["members"]
    
    # 70% chance to pick from same thread (Positive sample candidate)
    if random.random() < 0.7:
        i1, i2 = random.sample(members, 2)
        return (i1, i2, "same_thread_candidate")
    else:
        # Negative sample: One from this thread, one from random other thread
        i1 = random.choice(members)
        other_thread = random.choice(threads_data)
        if other_thread == thread:
            other_thread = random.choice(threads_data) # Retry once
        
        if other_thread["members"]:
            i2 = random.choice(other_thread["members"])
            return (i1, i2, "diff_thread_candidate")
        else:
            return (i1, i1, "error")

if st.session_state.current_pair is None:
    st.session_state.current_pair = get_next_pair()

idx1, idx2, pair_type = st.session_state.current_pair

# --- Display ---
row1 = df.iloc[idx1]
row2 = df.iloc[idx2]

st.subheader(f"Compare Pair ({pair_type})")

c1, c2 = st.columns(2)

with c1:
    st.markdown("### Email A")
    st.text(f"Row: {idx1}")
    st.markdown(f"**Subject:** {row1.get('Subject', '')}")
    st.markdown(f"**Sender:** {row1.get('SenderName', '')} ({row1.get('SenderEmail', '')})")
    st.markdown(f"**Date:** {row1.get('DeliveryTime', '')}")
    st.text_area("Body", str(row1.get('PlainTextBody', ''))[:1000], height=300, key="body1", disabled=True)

with c2:
    st.markdown("### Email B")
    st.text(f"Row: {idx2}")
    st.markdown(f"**Subject:** {row2.get('Subject', '')}")
    st.markdown(f"**Sender:** {row2.get('SenderName', '')} ({row2.get('SenderEmail', '')})")
    st.markdown(f"**Date:** {row2.get('DeliveryTime', '')}")
    st.text_area("Body", str(row2.get('PlainTextBody', ''))[:1000], height=300, key="body2", disabled=True)

# --- Actions ---
st.divider()
col_yes, col_no, col_skip = st.columns([1, 1, 2])

def save_and_next(match: bool):
    record = {
        "id_a": int(idx1),
        "id_b": int(idx2),
        "match": match,
        "reason": pair_type,
        "timestamp": pd.Timestamp.now().isoformat()
    }
    st.session_state.labels.append(record)
    _save_labels(LABEL_FILE, st.session_state.labels)
    st.session_state.current_pair = get_next_pair()
    # st.rerun() # Callback handles rerun automatically usually, but let's be safe if needed

with col_yes:
    if st.button("MATCH (Same Thread)", type="primary", use_container_width=True):
        save_and_next(True)
        st.rerun()

with col_no:
    if st.button("NO MATCH (Different)", type="secondary", use_container_width=True):
        save_and_next(False)
        st.rerun()

with col_skip:
    if st.button("Skip", use_container_width=True):
        st.session_state.current_pair = get_next_pair()
        st.rerun()

# --- Stats ---
st.divider()
st.metric("Total Labeled Pairs", len(st.session_state.labels))
if st.session_state.labels:
    st.dataframe(pd.DataFrame(st.session_state.labels).sort_values("timestamp", ascending=False).head(10))
