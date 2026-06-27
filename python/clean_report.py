"""
DataPolish — Automated data cleaning & reporting (Python edition)
-----------------------------------------------------------------
Mirrors the web app logic: detects column types, fills missing values,
drops duplicates, normalizes formats, and emits a multi-sheet Excel
workbook plus a PDF report with charts.

Usage
-----
    python clean_report.py input.csv
    python clean_report.py input.xlsx --out-dir ./reports
    python clean_report.py --sample            # use bundled sample data

Outputs (next to the input by default):
    <name>_cleaned.xlsx   - cleaned data + summary + column stats
    <name>_report.pdf     - printable report with charts

Requirements
------------
    pip install pandas numpy openpyxl matplotlib
"""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

MISSING_TOKENS = {"", "na", "n/a", "null", "none", "nan", "-", "--", "?"}


@dataclass
class CleanReport:
    original_rows: int = 0
    cleaned_rows: int = 0
    duplicates_removed: int = 0
    missing_filled: int = 0
    whitespace_trimmed: int = 0
    case_normalized: int = 0
    dates_normalized: int = 0
    numbers_coerced: int = 0
    notes: list[str] = field(default_factory=list)
    column_stats: pd.DataFrame = field(default_factory=pd.DataFrame)


def _is_missing(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    if isinstance(v, str) and v.strip().lower() in MISSING_TOKENS:
        return True
    return False


def _try_number(v: Any) -> float | None:
    if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v)):
        return float(v)
    if not isinstance(v, str):
        return None
    cleaned = v.replace(",", "").replace("$", "").strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _detect_type(series: pd.Series) -> str:
    non_missing = [v for v in series if not _is_missing(v)]
    if not non_missing:
        return "empty"
    nums = sum(1 for v in non_missing if _try_number(v) is not None)
    if nums / len(non_missing) >= 0.8:
        return "number"
    try:
        parsed = pd.to_datetime(pd.Series(non_missing), errors="coerce", dayfirst=True)
        if parsed.notna().mean() >= 0.8:
            return "date"
    except Exception:
        pass
    return "string"


def clean_dataframe(
    df: pd.DataFrame,
    *,
    drop_duplicates: bool = True,
    trim_whitespace: bool = True,
    normalize_case: bool = True,
    normalize_dates: bool = True,
    drop_empty_columns: bool = True,
    fill_numeric_with: str = "median",  # mean | median | zero | none
    fill_string_with: str = "mode",      # mode | unknown | none
) -> tuple[pd.DataFrame, CleanReport]:
    report = CleanReport(original_rows=len(df))
    df = df.copy()

    types = {c: _detect_type(df[c]) for c in df.columns}

    if drop_empty_columns:
        empties = [c for c, t in types.items() if t == "empty"]
        if empties:
            df = df.drop(columns=empties)
            report.notes.append(f"Dropped {len(empties)} empty column(s): {', '.join(empties)}")
            for c in empties:
                types.pop(c)

    if trim_whitespace:
        for c in df.select_dtypes(include="object").columns:
            before = df[c].copy()
            df[c] = df[c].map(lambda x: " ".join(x.split()) if isinstance(x, str) else x)
            report.whitespace_trimmed += int((before != df[c]).sum())

    # Replace missing tokens with NaN
    df = df.applymap(lambda v: np.nan if _is_missing(v) else v)

    for col, t in types.items():
        if t == "number":
            coerced = df[col].map(_try_number)
            report.numbers_coerced += int((coerced.notna() & (coerced != df[col])).sum())
            df[col] = pd.to_numeric(coerced, errors="coerce")
            missing = df[col].isna().sum()
            if missing and fill_numeric_with != "none":
                fill = {"mean": df[col].mean(), "median": df[col].median(), "zero": 0}[fill_numeric_with]
                df[col] = df[col].fillna(fill)
                report.missing_filled += int(missing)

        elif t == "date" and normalize_dates:
            parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            report.dates_normalized += int((parsed.notna() & (parsed.dt.strftime("%Y-%m-%d") != df[col])).sum())
            df[col] = parsed.dt.strftime("%Y-%m-%d")

        elif t == "string":
            if normalize_case:
                def _tc(x):
                    if isinstance(x, str) and len(x) <= 40 and all(c.isalpha() or c.isspace() or c in "-'/" for c in x):
                        return x.title()
                    return x
                before = df[col].copy()
                df[col] = df[col].map(_tc)
                report.case_normalized += int((before != df[col]).sum())
            missing = df[col].isna().sum()
            if missing and fill_string_with != "none":
                if fill_string_with == "mode":
                    mode = df[col].mode(dropna=True)
                    fill = mode.iloc[0] if not mode.empty else "Unknown"
                else:
                    fill = "Unknown"
                df[col] = df[col].fillna(fill)
                report.missing_filled += int(missing)

    if drop_duplicates:
        before = len(df)
        df = df.drop_duplicates().reset_index(drop=True)
        report.duplicates_removed = before - len(df)

    # Column stats
    stats = []
    for col in df.columns:
        s = df[col]
        row = {
            "Column": col, "Type": types.get(col, "string"),
            "Filled": int(s.notna().sum()), "Missing": int(s.isna().sum()),
            "Unique": int(s.nunique(dropna=True)),
        }
        if types.get(col) == "number":
            row.update({"Min": float(s.min()), "Max": float(s.max()), "Mean": float(s.mean())})
        else:
            row.update({"Min": "", "Max": "", "Mean": ""})
        stats.append(row)
    report.column_stats = pd.DataFrame(stats)
    report.cleaned_rows = len(df)
    return df, report


# ---------- Reporting ---------------------------------------------------------

def export_excel(df: pd.DataFrame, report: CleanReport, path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="Cleaned Data", index=False)
        summary = pd.DataFrame({
            "Metric": [
                "Original rows", "Cleaned rows", "Duplicates removed",
                "Missing values filled", "Whitespace trimmed", "Case normalized",
                "Dates normalized", "Numbers coerced", "Generated at",
            ],
            "Value": [
                report.original_rows, report.cleaned_rows, report.duplicates_removed,
                report.missing_filled, report.whitespace_trimmed, report.case_normalized,
                report.dates_normalized, report.numbers_coerced,
                datetime.now().isoformat(timespec="seconds"),
            ],
        })
        summary.to_excel(xl, sheet_name="Summary", index=False)
        report.column_stats.to_excel(xl, sheet_name="Column Stats", index=False)


def export_pdf(df: pd.DataFrame, report: CleanReport, path: Path) -> None:
    with PdfPages(path) as pdf:
        # Cover page with summary metrics
        fig, ax = plt.subplots(figsize=(8.5, 11))
        ax.axis("off")
        ax.set_title("Data Cleaning Report", fontsize=22, weight="bold", loc="left", pad=20)
        ax.text(0, 0.95, f"Generated {datetime.now():%Y-%m-%d %H:%M}", fontsize=10, color="gray")

        metrics = [
            ("Original rows", report.original_rows),
            ("Cleaned rows", report.cleaned_rows),
            ("Duplicates removed", report.duplicates_removed),
            ("Missing values filled", report.missing_filled),
            ("Whitespace trimmed", report.whitespace_trimmed),
            ("Case normalized", report.case_normalized),
            ("Dates normalized", report.dates_normalized),
            ("Numbers coerced", report.numbers_coerced),
        ]
        table = ax.table(
            cellText=[[k, f"{v:,}"] for k, v in metrics],
            colLabels=["Metric", "Value"],
            loc="upper left", cellLoc="left", colWidths=[0.5, 0.3],
            bbox=[0, 0.45, 0.8, 0.45],
        )
        table.auto_set_font_size(False); table.set_fontsize(10); table.scale(1, 1.4)
        pdf.savefig(fig); plt.close(fig)

        # Column stats page
        fig, ax = plt.subplots(figsize=(11, 8.5))
        ax.axis("off")
        ax.set_title("Column statistics", fontsize=16, weight="bold", loc="left")
        tbl = ax.table(
            cellText=report.column_stats.values.tolist(),
            colLabels=list(report.column_stats.columns),
            loc="center", cellLoc="left",
        )
        tbl.auto_set_font_size(False); tbl.set_fontsize(8); tbl.scale(1, 1.3)
        pdf.savefig(fig); plt.close(fig)

        # Charts: numeric distributions
        num_cols = df.select_dtypes(include=np.number).columns.tolist()[:4]
        if num_cols:
            fig, axes = plt.subplots(2, 2, figsize=(11, 8.5))
            axes = axes.flatten()
            for i, col in enumerate(num_cols):
                axes[i].hist(df[col].dropna(), bins=20, color="#5252DC", edgecolor="white")
                axes[i].set_title(f"Distribution: {col}", fontsize=11)
                axes[i].grid(alpha=0.3)
            for i in range(len(num_cols), 4):
                axes[i].axis("off")
            fig.suptitle("Numeric distributions", fontsize=14, weight="bold")
            fig.tight_layout()
            pdf.savefig(fig); plt.close(fig)

        # Category counts
        cat_cols = [c for c in df.select_dtypes(include="object").columns if 1 < df[c].nunique() <= 12][:2]
        if cat_cols:
            fig, axes = plt.subplots(1, len(cat_cols), figsize=(11, 5))
            if len(cat_cols) == 1: axes = [axes]
            for ax, col in zip(axes, cat_cols):
                counts = df[col].value_counts().head(8)
                ax.bar(counts.index, counts.values, color="#5252DC")
                ax.set_title(f"Counts by {col}")
                ax.tick_params(axis="x", rotation=30)
            fig.tight_layout(); pdf.savefig(fig); plt.close(fig)


def generate_sample(path: Path, n: int = 220) -> None:
    rng = np.random.default_rng(7)
    regions = ["north", "South", "EAST", " west ", "North", "south"]
    products = ["Widget A", "widget a", "Widget B", "WIDGET B", "Gadget", "gadget ", "Sprocket"]
    channels = ["Online", "online", "Retail", "RETAIL", "Partner"]
    start = datetime(2024, 1, 1)
    rows = []
    for i in range(n):
        date = start + pd.Timedelta(days=int(rng.integers(0, 330)))
        fmt = [date.strftime("%Y-%m-%d"),
               date.strftime("%-d/%-m/%Y") if os.name != "nt" else date.strftime("%d/%m/%Y"),
               date.strftime("%m-%d-%Y")][i % 3]
        qty = None if rng.random() < 0.08 else int(rng.integers(1, 50))
        price = "" if rng.random() < 0.1 else f"{rng.uniform(10, 210):.2f}"
        revenue = f"${(qty * float(price)):.2f}" if qty and price else "N/A"
        rows.append({
            "OrderID": f"ORD-{1000+i}",
            "Date": fmt,
            "Region": "" if rng.random() < 0.05 else regions[i % len(regions)],
            "Product": products[i % len(products)],
            "Channel": channels[i % len(channels)],
            "Quantity": qty,
            "UnitPrice": price,
            "Revenue": revenue,
            "Customer": "  " if rng.random() < 0.07 else f"Customer {int(rng.integers(0,60))}",
        })
        if i % 25 == 0 and i > 0:
            rows.append(rows[-1].copy())
    pd.DataFrame(rows).to_csv(path, index=False)


def main() -> int:
    parser = argparse.ArgumentParser(description="Clean a dataset and emit Excel + PDF reports.")
    parser.add_argument("input", nargs="?", help="CSV or Excel file to clean")
    parser.add_argument("--out-dir", type=Path, default=None)
    parser.add_argument("--sample", action="store_true", help="Use bundled messy sample data")
    args = parser.parse_args()

    if args.sample or not args.input:
        sample_path = Path(__file__).parent / "sample_messy_data.csv"
        if not sample_path.exists():
            generate_sample(sample_path)
        args.input = str(sample_path)

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"File not found: {in_path}", file=sys.stderr); return 1

    if in_path.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(in_path)
    else:
        df = pd.read_csv(in_path)

    cleaned, report = clean_dataframe(df)

    out_dir = args.out_dir or in_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = in_path.stem
    xlsx_path = out_dir / f"{stem}_cleaned.xlsx"
    pdf_path = out_dir / f"{stem}_report.pdf"

    export_excel(cleaned, report, xlsx_path)
    export_pdf(cleaned, report, pdf_path)

    print(f"✓ Rows: {report.original_rows} → {report.cleaned_rows} "
          f"(removed {report.duplicates_removed} duplicates, filled {report.missing_filled} missing)")
    print(f"✓ Excel: {xlsx_path}")
    print(f"✓ PDF:   {pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
