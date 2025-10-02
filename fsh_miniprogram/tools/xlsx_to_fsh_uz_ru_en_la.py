#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_to_fsh_uz_ru_en_la.py

Convert a large Excel/CSV/TSV file (10k+ rows) to FSH with the format:

* #<code> "<uz>"
  * ^designation[0].language = #ru
  * ^designation[=].value = "<russian>"
  * ^designation[+].language = #en
  * ^designation[=].value = "<english>"
  * ^designation[+].language = #la
  * ^designation[=].value = "<latin>"
  (+ extra languages appended at the end)

Usage:
  python xlsx_to_fsh_uz_ru_en_la.py input.xlsx -o output.fsh --sheet Sheet1 \
      --code CodeCol --uz UzbekCol --ru RussianCol --en EnglishCol --la LatinCol

Also supports CSV/TSV:
  python xlsx_to_fsh_uz_ru_en_la.py input.csv -o output.fsh --code code --uz uz --ru ru --en en --la la

Notes:
- Only "code" and "uz" (display) are strictly required. ru/en/la are optional but recommended.
- Extra languages: columns named like 'lang:xx' (e.g., lang:tr, lang:kk) will be appended automatically.
- Properly escapes FSH strings (quotes, backslashes, newlines).

"""

import argparse
import pandas as pd
from pathlib import Path
import sys
import re
from typing import Dict, List, Optional

def fsh_escape(value: str) -> str:
    """Escape a string for FSH double-quoted literals."""
    if value is None:
        return ""
    s = str(value)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    return s

def detect_extra_langs(columns: List[str]) -> Dict[str, str]:
    """Return mapping iso2 -> column name for any column named 'lang:xx'."""
    extra = {}
    for c in columns:
        cl = str(c).lower().strip()
        if cl.startswith("lang:") and len(cl) == 7 and cl[5:7].isalpha():
            iso = cl[5:7]
            extra[iso] = c
    return extra

def build_designation_block(row, lang_code: str, value_col: Optional[str], is_first: bool=False) -> List[str]:
    lines: List[str] = []
    if value_col and value_col in row and pd.notna(row[value_col]) and str(row[value_col]).strip() != "":
        value = fsh_escape(str(row[value_col]).strip())
        if is_first:
            lines.append(f'  * ^designation[0].language = #{lang_code}')
        else:
            lines.append(f'  * ^designation[+].language = #{lang_code}')
        lines.append(f'  * ^designation[=].value = "{value}"')
    return lines

def generate_fsh(df: pd.DataFrame,
                 code_col: str,
                 uz_col: str,
                 ru_col: Optional[str] = None,
                 en_col: Optional[str] = None,
                 la_col: Optional[str] = None) -> str:
    """Generate FSH text for the whole DataFrame with required order."""
    out_lines: List[str] = []
    # Auto-detect extra languages (lang:xx)
    extra_langs = detect_extra_langs(list(df.columns))

    # Iterate rows
    for idx, row in df.iterrows():
        code_val = row.get(code_col, None)
        uz_val = row.get(uz_col, None)

        if pd.isna(code_val) or str(code_val).strip() == "":
            continue
        if pd.isna(uz_val) or str(uz_val).strip() == "":
            continue

        code_str = str(code_val).strip()
        uz_str = fsh_escape(str(uz_val).strip())

        # Header line (display = uz)
        out_lines.append(f'* #{code_str} "{uz_str}"')

        # designation order: ru -> en -> la
        out_lines.extend(build_designation_block(row, "ru", ru_col, is_first=True))
        out_lines.extend(build_designation_block(row, "en", en_col))
        out_lines.extend(build_designation_block(row, "la", la_col))

        # Append extra languages at the end, excluding ru/en/la
        for iso, colname in extra_langs.items():
            if iso in {"ru", "en", "la"}:
                continue
            out_lines.extend(build_designation_block(row, iso, colname))

        out_lines.append("")

    return "\n".join(out_lines).rstrip() + "\n"

def read_table(path: Path, sheet: Optional[str]) -> pd.DataFrame:
    """Read Excel/CSV/TSV into a DataFrame; uses openpyxl for .xlsx."""
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xlsm", ".xls"]:
        # For big sheets, pandas with openpyxl is fine for 10k+ rows
        return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    elif suffix in [".csv"]:
        return pd.read_csv(path)
    elif suffix in [".tsv", ".tab"]:
        return pd.read_csv(path, sep="\t")
    else:
        raise ValueError(f"Unsupported file type: {suffix}")

def main():
    ap = argparse.ArgumentParser(description="Convert Excel/CSV/TSV to FSH with Uzbek display and RU→EN→LA designations.")
    ap.add_argument("input", help="Path to input file (.xlsx, .csv, .tsv)")
    ap.add_argument("-o", "--output", default=None, help="Path to output .fsh (default: <input-name>.fsh)")
    ap.add_argument("--sheet", default=None, help="Excel sheet name (if reading .xlsx)")

    # Column mappings
    ap.add_argument("--code", required=True, help="Column name for concept code")
    ap.add_argument("--uz", required=True, help="Column name for Uzbek display (main)")
    ap.add_argument("--ru", default=None, help="Column name for Russian designation")
    ap.add_argument("--en", default=None, help="Column name for English designation")
    ap.add_argument("--la", default=None, help="Column name for Latin designation")

    args = ap.parse_args()
    in_path = Path(args.input)
    if not in_path.exists():
        print(f"ERROR: input file not found: {in_path}", file=sys.stderr)
        sys.exit(1)

    try:
        df = read_table(in_path, args.sheet)
    except Exception as e:
        print(f"ERROR reading input: {e}", file=sys.stderr)
        sys.exit(1)

    # If Excel returns a dict of sheets, pick provided sheet or first
    if isinstance(df, dict):
        if args.sheet and args.sheet in df:
            df = df[args.sheet]
        else:
            first = next(iter(df.values()), None)
            if first is None:
                print("ERROR: No sheets found.", file=sys.stderr)
                sys.exit(1)
            df = first

    # Normalize/strip column names
    df.columns = [str(c).strip() for c in df.columns]

    # Validate columns
    needed = [args.code, args.uz]
    for col in needed:
        if col not in df.columns:
            print(f'ERROR: Required column "{col}" not found in data.', file=sys.stderr)
            sys.exit(1)

    # Generate FSH
    try:
        text = generate_fsh(df, code_col=args.code, uz_col=args.uz, ru_col=args.ru, en_col=args.en, la_col=args.la)
    except Exception as e:
        print(f"ERROR generating FSH: {e}", file=sys.stderr)
        sys.exit(1)

    out_path = Path(args.output) if args.output else in_path.with_suffix(".fsh")
    try:
        out_path.write_text(text, encoding="utf-8")
    except Exception as e:
        print(f"ERROR writing output: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Done. Wrote {out_path}")

if __name__ == "__main__":
    main()
