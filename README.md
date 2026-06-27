# DataPolish — Data Cleaning & Reporting Automation

A two-part deliverable for the **Data Cleaning & Reporting Automation** assignment:

1. **Web app** (TanStack Start + React) — upload CSV/Excel, auto-clean, interactive dashboard, multi-format export.
2. **Python script** ([`/python/clean_report.py`](./python/clean_report.py)) — same logic on the command line, emits Excel + PDF.

## Key features

- Handles **missing values, duplicates, inconsistent casing, whitespace, date formats, and numeric strings** (`"$1,250.00"` → `1250`).
- Auto-detects column types and applies type-aware filling (median / mode).
- **Interactive dashboard**: bar / pie / line charts, column-stats table, live data preview.
- **Automated reports**: download a multi-sheet Excel workbook, a printable PDF, or just the cleaned CSV.
- All cleaning rules are toggleable in the UI and re-run instantly.

## Run the web app

```bash
bun install
bun run dev
```

Then open the preview and click **Try sample dataset** (or upload your own).

## Run the Python automation

```bash
cd python
pip install pandas numpy openpyxl matplotlib
python clean_report.py --sample
```

See [`python/README.md`](./python/README.md) for full options.

## Push to GitHub

This project is wired for Lovable's GitHub integration:

1. In the Lovable editor open the **+** menu → **GitHub** → **Connect project**
2. Authorize the Lovable GitHub App and pick an account/org
3. Click **Create Repository** — Lovable pushes the full codebase and keeps it in two-way sync from then on

Once connected, anything you change here (or on GitHub) syncs automatically.

## Project structure

```
src/
  lib/data-cleaner.ts        # pure TS cleaning engine (mirrors clean_report.py)
  routes/index.tsx           # upload, dashboard, export UI
  styles.css                 # design tokens (oklch)
python/
  clean_report.py            # CLI: clean + export Excel + PDF with charts
  sample_messy_data.csv      # generated on first run
```
