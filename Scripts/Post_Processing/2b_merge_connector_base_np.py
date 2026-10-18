#!/usr/bin/env python3
"""
Merge connector-linked BERT_Output fragments into one BERT_Base_NPs row.

Input:
  The output workbook from 2_remove_ending.py.

Output:
  A new workbook where rows like:
    globo | blank BaseNP
    ganglio-series | ganglio-series neutral glycolipids
  become one row:
    globo- and ganglio-series neutral glycolipids
"""

from __future__ import annotations
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402

import re
from pathlib import Path
from typing import Any

import pandas as pd


INFILE = work("withBaseNPs_post.xlsx")
OUTFILE = work("post_connector_merged.xlsx")

SHEET_NAME = None  # use None to process every sheet with the required columns
PMID_COL = "pmid"
SENT_INDEX_COL = "sent_index"
SENT_TEXT_COL = "sent_text"
TERM_COL = "BERT_Output"
BASE_COL = "BERT_Base_NPs"
START_COL = "start_offset"
END_COL = "end_offset"
DISCARD_COL = "Discard"
UPDATE_COL = "Update"

CONNECTOR_GAP_RE = re.compile(
    r"^(?:\s|[,;:/&()+\-]|\band/or\b|\band\b|\bor\b|\bplus\b)+$",
    re.IGNORECASE,
)


def is_empty_cell(value: Any) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def clean_cell_text(value: Any) -> str:
    if is_empty_cell(value):
        return ""
    return str(value).strip()


def is_yes(value: Any) -> bool:
    return str(value).strip().casefold() == "yes"


def to_int(value: Any) -> int | None:
    if is_empty_cell(value):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def is_connector_gap(value: str) -> bool:
    return bool(value and CONNECTOR_GAP_RE.fullmatch(value))


def matching_span(text: str, needle: str, min_start: int = 0) -> tuple[int, int] | None:
    if not text or not needle:
        return None

    start = text.find(needle, min_start)
    if start < 0:
        start = text.casefold().find(needle.casefold(), min_start)
    if start < 0:
        return None
    return start, start + len(needle)


def matching_spans(text: str, needle: str) -> list[tuple[int, int]]:
    if not text or not needle:
        return []

    spans: list[tuple[int, int]] = []
    search_text = text.casefold()
    search_needle = needle.casefold()
    start = 0
    while True:
        found = search_text.find(search_needle, start)
        if found < 0:
            break
        spans.append((found, found + len(needle)))
        start = found + 1
    return spans


def offset_span(row: pd.Series, text: str, value: str) -> tuple[int, int] | None:
    if START_COL not in row.index or END_COL not in row.index:
        return None

    start = to_int(row.get(START_COL))
    end = to_int(row.get(END_COL))
    if start is None or end is None or start < 0 or end <= start or end > len(text):
        return None
    if text[start:end].casefold() != value.casefold():
        return None
    return start, end


def row_term_span(
    row: pd.Series,
    text: str,
    min_start: int = 0,
) -> tuple[int, int] | None:
    term = clean_cell_text(row.get(TERM_COL))
    if not term:
        return None

    span = offset_span(row, text, term)
    if span is not None and span[0] >= min_start:
        return span
    return matching_span(text, term, min_start)


def row_base_span(
    row: pd.Series,
    text: str,
    term_span: tuple[int, int] | None,
) -> tuple[int, int] | None:
    base_np = clean_cell_text(row.get(BASE_COL))
    if not base_np:
        return None

    spans = matching_spans(text, base_np)
    if term_span is not None:
        term_start, term_end = term_span
        containing_spans = [
            span for span in spans if span[0] <= term_start and span[1] >= term_end
        ]
        if containing_spans:
            return containing_spans[0]

    return spans[0] if spans else None


def span_inside(
    span: tuple[int, int] | None,
    outer_start: int,
    outer_end: int,
) -> bool:
    return span is not None and span[0] >= outer_start and span[1] <= outer_end


def row_has_term_without_base(row: pd.Series) -> bool:
    return not is_empty_cell(row.get(TERM_COL)) and is_empty_cell(row.get(BASE_COL))


def row_has_term_and_base(row: pd.Series) -> bool:
    return not is_empty_cell(row.get(TERM_COL)) and not is_empty_cell(row.get(BASE_COL))


def row_span_inside_merged_phrase(
    row: pd.Series,
    text: str,
    merged_start: int,
    merged_end: int,
) -> bool:
    term = clean_cell_text(row.get(TERM_COL))
    base_np = clean_cell_text(row.get(BASE_COL))

    for value in [term, base_np]:
        if not value:
            continue

        span = offset_span(row, text, value)
        if span_inside(span, merged_start, merged_end):
            return True

        for local_start, local_end in matching_spans(text[merged_start:merged_end], value):
            shifted_span = (merged_start + local_start, merged_start + local_end)
            if span_inside(shifted_span, merged_start, merged_end):
                return True

    return False


def group_columns(df: pd.DataFrame) -> list[str]:
    columns = [PMID_COL, SENT_TEXT_COL]
    if SENT_INDEX_COL in df.columns:
        columns.insert(1, SENT_INDEX_COL)
    return columns


def merge_connector_base_np_rows(df: pd.DataFrame) -> pd.DataFrame:
    required = {PMID_COL, SENT_TEXT_COL, BASE_COL, TERM_COL}
    if not required.issubset(df.columns):
        return df

    if UPDATE_COL not in df.columns:
        df[UPDATE_COL] = ""

    consumed_indexes: set[Any] = set()

    for _, group in df.groupby(group_columns(df), sort=False, dropna=False):
        indexes = list(group.index)
        pos = 0

        while pos < len(indexes):
            first_index = indexes[pos]
            first_row = df.loc[first_index]
            if first_index in consumed_indexes or not row_has_term_without_base(first_row):
                pos += 1
                continue

            text = clean_cell_text(first_row.get(SENT_TEXT_COL))
            if not text:
                pos += 1
                continue

            chain_indexes: list[Any] = []
            chain_spans: list[tuple[int, int]] = []
            scan_pos = pos
            last_end = 0

            while scan_pos < len(indexes):
                current_index = indexes[scan_pos]
                current_row = df.loc[current_index]
                if (
                    current_index in consumed_indexes
                    or not row_has_term_without_base(current_row)
                    or is_yes(current_row.get(DISCARD_COL))
                ):
                    break

                min_start = last_end if chain_spans else 0
                term_span = row_term_span(current_row, text, min_start)
                if term_span is None:
                    break

                if chain_spans and not is_connector_gap(text[last_end : term_span[0]]):
                    break

                chain_indexes.append(current_index)
                chain_spans.append(term_span)
                last_end = term_span[1]
                scan_pos += 1

            if not chain_indexes or scan_pos >= len(indexes):
                pos += 1
                continue

            target_index = indexes[scan_pos]
            target_row = df.loc[target_index]
            if target_index in consumed_indexes or not row_has_term_and_base(target_row):
                pos += 1
                continue

            target_term_span = row_term_span(target_row, text, last_end)
            target_base_span = row_base_span(target_row, text, target_term_span)
            if target_base_span is None:
                pos += 1
                continue

            target_start = (
                target_term_span[0] if target_term_span is not None else target_base_span[0]
            )
            if target_start < last_end or not is_connector_gap(text[last_end:target_start]):
                pos += 1
                continue

            merged_start = chain_spans[0][0]
            merged_end = target_base_span[1]
            raw_merged_text = text[merged_start:merged_end]
            left_trimmed = raw_merged_text.lstrip()
            merged_start += len(raw_merged_text) - len(left_trimmed)
            merged_text = left_trimmed.rstrip()
            merged_end = merged_start + len(merged_text)
            if not merged_text:
                pos += 1
                continue

            df.at[first_index, TERM_COL] = merged_text
            df.at[first_index, BASE_COL] = merged_text
            df.at[first_index, UPDATE_COL] = "Yes"
            if START_COL in df.columns:
                df.at[first_index, START_COL] = merged_start
            if END_COL in df.columns:
                df.at[first_index, END_COL] = merged_end

            consumed_indexes.update(chain_indexes[1:])
            consumed_indexes.add(target_index)

            for other_index in indexes:
                if other_index == first_index or other_index in consumed_indexes:
                    continue
                other_row = df.loc[other_index]
                if row_span_inside_merged_phrase(
                    other_row,
                    text,
                    merged_start,
                    merged_end,
                ):
                    consumed_indexes.add(other_index)

            pos = scan_pos + 1

    if not consumed_indexes:
        return df
    return df.drop(index=list(consumed_indexes))


def process_workbook(infile: Path, outfile: Path) -> None:
    if not infile.exists():
        raise FileNotFoundError(f"Not found: {infile}")

    xls = pd.ExcelFile(infile)
    outfile.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(infile, sheet_name=sheet_name)
            should_process = SHEET_NAME is None or sheet_name == SHEET_NAME
            if should_process:
                df = merge_connector_base_np_rows(df)
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def main() -> None:
    process_workbook(Path(INFILE), Path(OUTFILE))
    print(f"Wrote connector-merged workbook to: {OUTFILE}")


if __name__ == "__main__":
    main()
