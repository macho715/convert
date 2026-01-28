# Walkthrough - Email Search System Improvements

I have successfully implemented the **Thread Score Model**, created an **Evaluation Suite**, and completely redesigned the **Dashboard** with a premium UI/UX.

## 1. Premium Dashboard Redesign
The dashboard has been overhauled with a modern, clean interface using custom CSS and improved layout.

### Global Styling
- **File**: `dashboard/style.css`
- **Font**: Inter & JetBrains Mono
- **Theme**: Clean white/gray with blue accents (`#2563eb`)
- **Components**: Rounded cards, soft shadows, and glassmorphism effects.

### Search Page
- **File**: `dashboard/pages/02_search.py`
- **Integrated Filters**: Search bar and filters are now in a unified, clean container.
- **Result Cards**: Emails are displayed as interactive cards with:
    - **Snippet Preview**: Highlighted search terms in the body snippet.
    - **Smart Tags**: `Case`, `Site`, `LPO`, and `Score` tags.
    - **Context Mode**: Clear visual distinction between "Direct Match" (Blue border) and "Thread Context" (Gray border).

### Thread Explorer
- **File**: `dashboard/pages/03_thread_explorer.py`
- **Timeline Visualization**: Interactive Plotly timeline showing message flow colored by sender.
- **Network Graph**: Interactive node-link diagram showing thread structure (requires `networkx`).
- **Message Detail**: Collapsible full-body view with auto-collapse for very long emails.

### Overview
- **File**: `dashboard/pages/01_overview.py`
- **Key Metrics**: High-level stats (Total Threads, Avg Confidence, etc.) in a clean grid.
- **Charts**: Distribution histograms for Thread Size and Confidence.

### Quality Audit
- **File**: `dashboard/pages/04_quality_audit.py`
- **Consistent Theme**: Updated to match the global premium style.
- **Advanced Diagnostics**: Cycle detection and false-merge suspect analysis.

## 2. Evaluation & Labeling
I have added tools to verify and improve the threading logic.

### Labeling Tool
- **File**: `dashboard/pages/05_labeling.py`
- **Purpose**: Manually verify if two emails belong to the same thread.
- **Features**:
    - Side-by-side email comparison.
    - "Match" / "No Match" buttons.
    - Saves results to `data/evaluation_set.json`.

### Evaluation Script
- **File**: `scripts/evaluate_threads.py`
- **Purpose**: Calculate accuracy metrics (Precision, Recall, F1) against the labeled set.
- **Usage**:
  ```bash
  python scripts/evaluate_threads.py outputs/threads_full/threads.json data/evaluation_set.json
  ```

## 3. Thread Score Model
The threading logic in `scripts/outlook_thread_tracker_v3.py` now includes:
- **Subject Similarity**: Levenshtein distance for fuzzy subject matching.
- **Entity Overlap**: Checks for shared Case/Site/LPO tags.
- **Time Decay**: Penalizes messages that are too far apart (>14 days).
- **Reply Hint**: Boosts score if `RE:`/`FW:` or quote patterns are found.

## Next Steps
1.  **Run the Dashboard**: `streamlit run dashboard/app.py`
2.  **Label Data**: Use the "Labeling Tool" page to create a ground truth set.
3.  **Evaluate**: Run the evaluation script to benchmark performance.
