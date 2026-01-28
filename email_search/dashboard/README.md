# Email Search Dashboard

Quick start

1) Install dependencies

```bash
pip install -r email_search/dashboard/requirements.txt
```

2) Generate outputs (threads.json, edges.csv, search_result.csv)

```bash
python email_search/scripts/run_full_export.py \
  --excel email_search/data/OUTLOOK_HVDC_ALL_rev.xlsx \
  --sheet "전체_데이터" \
  --out email_search/outputs/threads_full \
  --query "LPO-1599"
```

3) Run the dashboard

```bash
streamlit run email_search/dashboard/app.py
```

Data root

- Default: `../outputs/threads_full`
- If your outputs live elsewhere, update the data root from the Overview page.

Notes

- If `plotly` is not installed, charts and timelines are hidden.
- If `networkx` is not installed, cycle samples in Quality/Audit fall back to a basic check.
