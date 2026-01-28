# Implementation Plan - Email Search System Improvement

The `email_search` project seems to have most of the "Future Plan" (Option A) implemented, including the thread tracker v3 and Streamlit dashboard. This plan focuses on verifying current functionality, improving search accuracy, and overhauling the dashboard.

## User Review Required

> [!IMPORTANT]
> Please confirm if we should proceed with **Option B: Graph Integration (SSOT)** or focus on **Option C: Operational Features** (Delta queries, event logs). (Currently focusing on search quality and UI improvements)

## Proposed Changes

### 1. Search Improvements (Accuracy & Ranking)
- [x] **Relevance Scoring**: Algorithm prioritizing `Subject` > `Sender` > `Body`.
- [x] **Fuzzy Matching**: Option to handle typos.
- [x] **Context Separation**: Distinct UI for "Direct Match" vs "Context/Thread Match".
- [x] **Refined Tokenization**: Better handling of multi-word terms.

#### 1.1 Matching Rate Enhancement (Ontology-Based)
> Strategies derived from `Logi ontol core doc` analysis.
- [x] **Ontology-Driven Synonyms**:
    - Leverage relationships defined in [CONSOLIDATED-08-communication.md](file:///c:/Users/SAMSUNG/Downloads/CONVERT/email_search/Logi%20ontol%20core%20doc/CONSOLIDATED-08-communication.md) and `CORE` docs.
    - Example: `LCT` -> `Barge`, `Landing Craft`, `Vessel`. `Transformer` -> `TR`, `Trafo`.
    - Expand [config/synonyms.json](file:///c:/Users/SAMSUNG/Downloads/CONVERT/email_search/config/synonyms.json).
- [x] **Intent & Tag Boosting**:
    - Boost emails with `[URGENT]`, `[ACTION]` tags if query contains "Urgent", "Action", "Request".
- [x] **Entity Pattern Matching**:
    - Regex for `BL No`, `Container ID`, `Project Tag`.
- [x] **Semantic Expansion**:
    - "Customs" -> `BOE`, `Clearance`, `Inspection`, `Duty`.
- [x] **Thread Context Boosting**:
    - Show other messages in a thread as "related context" even if only one matches.

#### 1.2 Thread Score Model (Option A)
> Advanced scoring model for Excel-only environment.
- [ ] **Subject Similarity**: Levenshtein distance (0.0 ~ 1.0).
- [ ] **Entity Overlap**: Intersection of Case, Site, LPO.
- [ ] **Time Decay**: Decay over 14-day window.
- [ ] **Reply Hint**: Bonus for `RE:`, `FW:` and quote patterns.

#### 1.3 Evaluation Set
- [ ] **Evaluation Script**: Create `scripts/evaluate_threads.py`.
- [ ] **Labeling Tool**: Add simple labeling interface to Streamlit dashboard.

### 2. Dashboard Redesign (Premium UI/UX)
- [ ] **Global Styling**: Custom CSS for modern look.
- [ ] **Search Page Revamp**: Integrated filter bar, snippet previews, clear distinction of results.
- [ ] **Thread Explorer**: Interactive timeline or tree view.
- [ ] **Overview**: High-level metrics with trend charts.

### 3. Existing Function Verification
- [x] `scripts/outlook_aqs_searcher.py` verified.
- [x] `scripts/export_email_threads_cli.py` verified.
- [x] `dashboard/app.py` verified.

### 4. Documentation Update
- [ ] Integrate `dashboard/README.md` into main README.
- [ ] Ensure all CLI args are documented.

### 5. Graph Integration Prep (Option B)
- [ ] Create `sql/email_ssot.sql` for Supabase.
- [ ] Research Microsoft Graph API requirements.

## Verification Plan

### Automated Tests
```bash
# 1. Search (Synonyms & Patterns)
python scripts/outlook_aqs_searcher.py data/OUTLOOK_HVDC_ALL_rev.xlsx -q "subject:tr" --config config/aqs_column_aliases.json

# 2. Thread Export
python scripts/export_email_threads_cli.py --excel data/OUTLOOK_HVDC_ALL_rev.xlsx --out outputs/test_run
```

### Manual Verification
- Run dashboard: `streamlit run dashboard/app.py`
- Check Overview, Search (Ranking, Synonyms), and Thread Visualization.
