# DataPolish — Python automation

Mirror of the web app's cleaning + reporting logic, runnable from the command line.

## Install

```bash
pip install pandas numpy openpyxl matplotlib
```

## Run

```bash
# Use the bundled messy sample dataset
python clean_report.py --sample

# Clean your own file
python clean_report.py path/to/data.csv
python clean_report.py path/to/data.xlsx --out-dir ./reports
```

## What it does

1. Detects column types (numeric / date / string / empty)
2. Normalizes missing-value tokens (`""`, `NA`, `N/A`, `null`, `-`, …) to `NaN`
3. Trims whitespace and title-cases short categorical strings
4. Coerces numeric columns (handles `$` and `,`)
5. Parses dates and rewrites them as ISO `YYYY-MM-DD`
6. Fills missing numerics with the median, missing strings with the mode
7. Drops exact duplicates after normalization
8. Emits:
   - `<name>_cleaned.xlsx` — three sheets: Cleaned Data, Summary, Column Stats
   - `<name>_report.pdf` — printable report with metrics, column stats, histograms, and category counts

## Files

| File | Purpose |
|------|---------|
| `clean_report.py` | Main automation script |
| `sample_messy_data.csv` | Built-in messy sales dataset (220+ rows, with duplicates / nulls / format drift) |

The same logic powers the web UI in `src/lib/data-cleaner.ts` and `src/routes/index.tsx`.
