

import argparse
from pathlib import Path
import sys
import pandas as pd


def load_sheet(path: Path, sheet, key_cols):
    """Read an Excel sheet into a DataFrame, set the key as the index, and
    coerce everything else to string so compare() works cleanly."""
    df = pd.read_excel(path, sheet_name=sheet)
    if key_cols:
        try:
            df = df.set_index(key_cols)
        except KeyError as e:
            sys.exit(f"[error] Key column not found in {path}: {e}")
    df = df.sort_index()
    return df.astype(str)


def main():
    ap = argparse.ArgumentParser(description="Compare two Excel files.")
    ap.add_argument("file_a", type=Path, help="Older / baseline workbook")
    ap.add_argument("file_b", type=Path, help="Newer workbook to compare to")
    ap.add_argument("--sheet", default=0,
                    help="Sheet name or 0‑based index (default: first sheet)")
    ap.add_argument("--key", nargs="+", metavar="COL",
                    help="Column(s) that uniquely identify a row")
    ap.add_argument("--out", type=Path,
                    help="Write an Excel file with three tabs (added, deleted, modified)")
    args = ap.parse_args()

    df_a = load_sheet(args.file_a, args.sheet, args.key)
    df_b = load_sheet(args.file_b, args.sheet, args.key)

    only_in_a = df_a.loc[~df_a.index.isin(df_b.index)]
    only_in_b = df_b.loc[~df_b.index.isin(df_a.index)]

    common_a = df_a.loc[df_a.index.intersection(df_b.index)]
    common_b = df_b.loc[common_a.index]
    modified = common_a.compare(common_b, keep_equal=False)  

    print(f"Rows only in {args.file_a.name}: {len(only_in_a)}")
    print(f"Rows only in {args.file_b.name}: {len(only_in_b)}")
    print(f"Rows with modified data:     {modified.index.nlevels and modified.index.get_level_values(0).nunique() or 0}")

    if args.out:
        with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
            only_in_a.to_excel(xl, sheet_name="Deleted_rows")
            only_in_b.to_excel(xl, sheet_name="Added_rows")
            modified.to_excel(xl, sheet_name="Modified_cells")
        print(f"Full detail written to {args.out.resolve()}")


if __name__ == "__main__":
    main()
