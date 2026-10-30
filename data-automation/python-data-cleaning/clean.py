"""Clean a messy customer CSV/Excel export.

Usage:
    python clean.py sample/customers_raw.csv --out output/
    python clean.py big_export.xlsx --out output/ --excel

Outputs (in --out):
    customers_clean.csv     cleaned, de-duplicated rows (+ an `issues` column)
    rejected_rows.csv       rows that cannot be used, with the reason
    duplicates.csv          every removed duplicate and the id that was kept
    report.md / report.json summary of what changed
    customers_clean.xlsx    (with --excel) the same clean data as a formatted workbook
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd

from cleaner.pipeline import clean_frame, to_markdown

EXCEL_MAX_ROWS = 1_048_575


def load(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str, keep_default_na=False)
    return pd.read_csv(path, dtype=str, keep_default_na=False, encoding_errors="replace")


def write_excel(df: pd.DataFrame, path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Customers")
        ws = xw.sheets["Customers"]
        ws.freeze_panes = "A2"
        for cell in ws["G"][1:]:
            cell.number_format = "yyyy-mm-dd"
        for cell in ws["H"][1:]:
            cell.number_format = "#,##0.00"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in col[:500]) + 2
            ws.column_dimensions[col[0].column_letter].width = min(width, 40)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path)
    ap.add_argument("--out", type=Path, default=Path("output"))
    ap.add_argument("--excel", action="store_true", help="also write a formatted .xlsx")
    a = ap.parse_args()

    raw = load(a.input)
    res = clean_frame(raw)
    a.out.mkdir(parents=True, exist_ok=True)

    clean = res.clean.copy()
    clean["signup_date"] = clean["signup_date"].dt.strftime("%Y-%m-%d")
    clean["total_spend"] = clean["total_spend"].map(lambda v: "" if pd.isna(v) else f"{v:.2f}")
    clean.to_csv(a.out / "customers_clean.csv", index=False)
    res.rejected.to_csv(a.out / "rejected_rows.csv", index=False)
    res.duplicates.to_csv(a.out / "duplicates.csv", index=False)
    (a.out / "report.json").write_text(json.dumps(res.stats, indent=2))
    (a.out / "report.md").write_text(to_markdown(res.stats, raw, clean))
    if a.excel:
        if len(clean) > EXCEL_MAX_ROWS:
            print("too many rows for one Excel sheet - skipped .xlsx")
        else:
            write_excel(res.clean, a.out / "customers_clean.xlsx")  # real dates and numbers

    s = res.stats
    print(f"{s['rows_in']:,} rows in -> {s['rows_clean']:,} clean | "
          f"{s['duplicates_removed']:,} duplicates | {s['rows_rejected']:,} rejected | {s['seconds']} s")


if __name__ == "__main__":
    main()
