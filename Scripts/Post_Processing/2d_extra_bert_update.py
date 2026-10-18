#!/usr/bin/env python3
"""
Update BERT_Base_NPs when BERT_Output contains extra useful text.

Input:
  2c_withBaseNPs_post_jun15_np.xlsx

Output:
  A new workbook with all input rows. Rows updated by this rule get Update=Yes.
"""

from __future__ import annotations
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402

import re
from pathlib import Path
from typing import Any

import pandas as pd


INFILE = Path(
    work("2c_withBaseNPs_post_np.xlsx")
)
OUTFILE = Path(
    work("2d_bert_np_update_extra_bert.xlsx")
)

SHEET_NAME = None  # use None to process every sheet with the required columns

PMID_COL = "pmid"
SENT_INDEX_COL = "sent_index"
SENT_TEXT_COL = "sent_text"
BERT_OUTPUT_COL = "BERT_Output"
BASE_NP_COL = "BERT_Base_NPs"
UPDATE_COL = "Update"

REQUIRED_COLUMNS = {
    PMID_COL,
    SENT_INDEX_COL,
    SENT_TEXT_COL,
    BERT_OUTPUT_COL,
    BASE_NP_COL,
}

# If True, keep only rows where BERT_Base_NPs appears inside BERT_Output.
# If False, keep every row where len(BERT_Output) > len(BERT_Base_NPs).
REQUIRE_BASE_INSIDE_OUTPUT = False


def is_empty_cell(value: Any) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def clean_cell(value: Any) -> str:
    return "" if is_empty_cell(value) else str(value).strip()


def normalize_for_contains(value: str) -> str:
    return re.sub(r"\s+", " ", value.casefold()).strip()


def base_inside_output(bert_output: str, base_np: str) -> bool:
    output_norm = normalize_for_contains(bert_output)
    base_norm = normalize_for_contains(base_np)
    return bool(output_norm and base_norm and base_norm in output_norm)


def find_casefold_span(text: str, needle: str) -> tuple[int, int] | None:
    if not text or not needle:
        return None

    start = text.casefold().find(needle.casefold())
    if start < 0:
        return None
    return start, start + len(needle)


def should_update_base_np(bert_output: str, base_np: str) -> tuple[bool, str]:
    span = find_casefold_span(bert_output, base_np)
    if span is None:
        return False, ""

    start, end = span
    left_extra = bert_output[:start]
    right_extra = bert_output[end:]

    if left_extra.strip():
        return True, "extra_left_of_base_np"
    if re.fullmatch(r"\s+[A-Za-z0-9]", right_extra):
        return True, "single_letter_or_digit_right_of_base_np"

    return False, ""


def update_sheet(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if not REQUIRED_COLUMNS.issubset(df.columns):
        return df, 0

    df = df.copy()
    if UPDATE_COL not in df.columns:
        df[UPDATE_COL] = ""

    update_count = 0
    for row_index, row in df.iterrows():
        bert_output = clean_cell(row.get(BERT_OUTPUT_COL))
        base_np = clean_cell(row.get(BASE_NP_COL))
        if not bert_output or not base_np:
            continue

        bert_output_length = len(bert_output)
        base_np_length = len(base_np)
        if bert_output_length <= base_np_length:
            continue

        contains_base = base_inside_output(bert_output, base_np)
        if REQUIRE_BASE_INSIDE_OUTPUT and not contains_base:
            continue

        should_update, _update_reason = should_update_base_np(bert_output, base_np)

        if should_update:
            df.at[row_index, BASE_NP_COL] = bert_output
            df.at[row_index, UPDATE_COL] = "Yes"
            update_count += 1

    return df, update_count


def process_workbook(input_path: Path, output_path: Path) -> dict[str, int]:
    xls = pd.ExcelFile(input_path)
    total_rows = 0
    total_updated = 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(input_path, sheet_name=sheet_name)

            if SHEET_NAME is None or sheet_name == SHEET_NAME:
                df, update_count = update_sheet(df)
                total_updated += update_count

            total_rows += len(df)
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return {"rows_written": total_rows, "rows_updated": total_updated}


def main() -> None:
    if not INFILE.exists():
        raise FileNotFoundError(f"Not found: {INFILE}")

    summary = process_workbook(INFILE, OUTFILE)

    print(f"Saved: {OUTFILE}")
    print(f"Rows written: {summary['rows_written']}")
    print(f"Rows updated: {summary['rows_updated']}")


if __name__ == "__main__":
    main()
