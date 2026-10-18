#!/usr/bin/env python3
"""
Collapse nested BERT_Base_NPs rows into one longest-BaseNP row.

For each pmid and sent_index:
1. Find BERT_Base_NPs spans that are contained inside a longer BaseNP span.
2. Keep one row for the longest BaseNP and remove its nested BaseNP rows.
3. Rebuild BERT_Output from the first detected output item through the last
   detected output item, following their order in sent_text.
4. Allow punctuation and at most one intervening word in total between the
   detected output items.
5. Update start_offset/end_offset and mark Nested="Yes" on changed rows.

Offsets are 0-based and end_offset is exclusive.
"""

from __future__ import annotations
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


INFILE = Path(
    work("2d_bert_np_update_extra_bert.xlsx")
)
OUTFILE = Path(
    work("2e_nested_base_np_updated.xlsx")
)

SHEET_NAME = None  # use None to process every sheet with the required columns

PMID_COL = "pmid"
SENT_INDEX_COL = "sent_index"
SENT_TEXT_COL = "sent_text"
BERT_OUTPUT_COL = "BERT_Output"
START_COL = "start_offset"
END_COL = "end_offset"
BASE_NP_COL = "BERT_Base_NPs"
NESTED_COL = "Nested"

REQUIRED_COLUMNS = {
    PMID_COL,
    SENT_INDEX_COL,
    SENT_TEXT_COL,
    BERT_OUTPUT_COL,
    START_COL,
    END_COL,
    BASE_NP_COL,
}

# Punctuation does not count. Across all gaps from the first output item to
# the last output item, allow zero or one intervening alphanumeric word.
MAX_TOTAL_GAP_WORDS = 1
GAP_WORD_RE = re.compile(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*")


@dataclass
class Candidate:
    row_index: Any
    row: pd.Series
    base_np: str
    bert_output: str
    base_start: int
    base_end: int
    output_start: int | None
    output_end: int | None

    @property
    def base_span(self) -> tuple[int, int]:
        return self.base_start, self.base_end

    @property
    def base_length(self) -> int:
        return self.base_end - self.base_start


def is_empty_cell(value: Any) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def clean_cell(value: Any) -> str:
    return "" if is_empty_cell(value) else str(value).strip()


def to_int(value: Any) -> int | None:
    if is_empty_cell(value):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def spans_for_text(sent_text: str, needle: str) -> list[tuple[int, int]]:
    if not sent_text or not needle:
        return []

    spans: list[tuple[int, int]] = []
    haystack = sent_text.casefold()
    query = needle.casefold()
    search_start = 0

    while True:
        found = haystack.find(query, search_start)
        if found < 0:
            break
        spans.append((found, found + len(needle)))
        search_start = found + 1

    return spans


def bert_output_span(row: pd.Series, sent_text: str) -> tuple[int, int] | None:
    bert_output = clean_cell(row.get(BERT_OUTPUT_COL))
    if not bert_output:
        return None

    start = to_int(row.get(START_COL))
    end = to_int(row.get(END_COL))
    if start is not None and end is not None and 0 <= start < end <= len(sent_text):
        if sent_text[start:end].casefold() == bert_output.casefold():
            return start, end

    spans = spans_for_text(sent_text, bert_output)
    return spans[0] if spans else None


def overlap_length(first: tuple[int, int], second: tuple[int, int]) -> int:
    return max(0, min(first[1], second[1]) - max(first[0], second[0]))


def base_np_span(
    row: pd.Series,
    sent_text: str,
    output_span: tuple[int, int] | None,
) -> tuple[int, int] | None:
    base_np = clean_cell(row.get(BASE_NP_COL))
    if not base_np:
        return None

    base_spans = spans_for_text(sent_text, base_np)
    if not base_spans:
        return None

    if output_span is not None:
        best_span = max(base_spans, key=lambda span: overlap_length(span, output_span))
        if overlap_length(best_span, output_span) > 0:
            return best_span

    return base_spans[0]


def make_candidate(
    row_index: Any,
    row: pd.Series,
    sent_text: str,
) -> Candidate | None:
    base_np = clean_cell(row.get(BASE_NP_COL))
    bert_output = clean_cell(row.get(BERT_OUTPUT_COL))
    if not base_np or not bert_output:
        return None

    output_span = bert_output_span(row, sent_text)
    base_span = base_np_span(row, sent_text, output_span)
    if base_span is None:
        return None

    output_start = output_span[0] if output_span is not None else None
    output_end = output_span[1] if output_span is not None else None

    return Candidate(
        row_index=row_index,
        row=row,
        base_np=base_np,
        bert_output=bert_output,
        base_start=base_span[0],
        base_end=base_span[1],
        output_start=output_start,
        output_end=output_end,
    )


def first_sent_text(group: pd.DataFrame) -> str:
    for value in group[SENT_TEXT_COL]:
        sent_text = clean_cell(value)
        if sent_text:
            return sent_text
    return ""


def strictly_contains(
    parent_span: tuple[int, int],
    child_span: tuple[int, int],
) -> bool:
    return (
        parent_span[0] <= child_span[0]
        and parent_span[1] >= child_span[1]
        and parent_span != child_span
    )


def trim_span(sent_text: str, start: int, end: int) -> tuple[int, int] | None:
    while start < end and sent_text[start].isspace():
        start += 1
    while end > start and sent_text[end - 1].isspace():
        end -= 1

    if start >= end or not re.search(r"[A-Za-z0-9]", sent_text[start:end]):
        return None
    return start, end


def output_item_span(
    candidate: Candidate,
    parent_span: tuple[int, int],
    sent_text: str,
) -> tuple[int, int] | None:
    parent_start, parent_end = parent_span

    if candidate.output_start is not None and candidate.output_end is not None:
        clipped_start = max(candidate.output_start, parent_start)
        clipped_end = min(candidate.output_end, parent_end)
        if clipped_start < clipped_end:
            clipped = trim_span(sent_text, clipped_start, clipped_end)
            if clipped is not None:
                return clipped

    parent_text = sent_text[parent_start:parent_end]
    local_spans = spans_for_text(parent_text, candidate.bert_output)
    if not local_spans:
        return None

    local_start, local_end = local_spans[0]
    return trim_span(
        sent_text,
        parent_start + local_start,
        parent_start + local_end,
    )


def merge_overlapping_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not spans:
        return []

    merged: list[list[int]] = []
    for start, end in sorted(set(spans)):
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)

    return [(start, end) for start, end in merged]


def total_gap_words(
    spans: list[tuple[int, int]],
    sent_text: str,
) -> int:
    total = 0
    for (_start, previous_end), (next_start, _end) in zip(spans, spans[1:]):
        gap = sent_text[previous_end:next_start]
        total += len(GAP_WORD_RE.findall(gap))
    return total


def choose_keep_candidate(
    parent_candidates: list[Candidate],
    parent_span: tuple[int, int],
    sent_text: str,
) -> Candidate:
    def sort_key(candidate: Candidate) -> tuple[int, int]:
        item_span = output_item_span(candidate, parent_span, sent_text)
        item_start = item_span[0] if item_span is not None else 10**12
        return item_start, int(candidate.row_index)

    return min(parent_candidates, key=sort_key)


def process_sentence_group(
    df: pd.DataFrame,
    group: pd.DataFrame,
    remove_indexes: set[Any],
) -> tuple[int, int]:
    sent_text = first_sent_text(group)
    if not sent_text:
        return 0, 0

    candidates = []
    for row_index, row in group.iterrows():
        candidate = make_candidate(row_index, row, sent_text)
        if candidate is not None:
            candidates.append(candidate)

    if len(candidates) < 2:
        return 0, 0

    span_groups: dict[tuple[int, int], list[Candidate]] = {}
    for candidate in candidates:
        span_groups.setdefault(candidate.base_span, []).append(candidate)

    if len(span_groups) < 2:
        return 0, 0

    parent_spans = [
        span
        for span in span_groups
        if not any(
            strictly_contains(other_span, span)
            for other_span in span_groups
            if other_span != span
        )
    ]

    nested_groups = 0
    rows_removed = 0

    for parent_span in sorted(parent_spans, key=lambda span: (span[0], -span[1])):
        child_spans = [
            span
            for span in span_groups
            if strictly_contains(parent_span, span)
        ]
        if not child_spans:
            continue

        involved_spans = [parent_span] + child_spans
        involved_candidates = [
            candidate
            for span in involved_spans
            for candidate in span_groups[span]
        ]

        item_spans_by_base: dict[tuple[int, int], list[tuple[int, int]]] = {}
        for span in involved_spans:
            item_spans = []
            for candidate in span_groups[span]:
                item_span = output_item_span(candidate, parent_span, sent_text)
                if item_span is not None:
                    item_spans.append(item_span)
            if not item_spans:
                break
            item_spans_by_base[span] = item_spans
        else:
            ordered_items = merge_overlapping_spans(
                [
                    item_span
                    for item_spans in item_spans_by_base.values()
                    for item_span in item_spans
                ]
            )
            if not ordered_items:
                continue

            if total_gap_words(ordered_items, sent_text) > MAX_TOTAL_GAP_WORDS:
                continue

            new_start = ordered_items[0][0]
            new_end = ordered_items[-1][1]
            trimmed = trim_span(sent_text, new_start, new_end)
            if trimmed is None:
                continue
            new_start, new_end = trimmed
            new_output = sent_text[new_start:new_end]

            parent_candidates = span_groups[parent_span]
            keep_candidate = choose_keep_candidate(
                parent_candidates,
                parent_span,
                sent_text,
            )

            parent_base_np = max(
                parent_candidates,
                key=lambda candidate: len(candidate.base_np),
            ).base_np

            df.at[keep_candidate.row_index, BERT_OUTPUT_COL] = new_output
            df.at[keep_candidate.row_index, BASE_NP_COL] = parent_base_np
            df.at[keep_candidate.row_index, START_COL] = new_start
            df.at[keep_candidate.row_index, END_COL] = new_end
            df.at[keep_candidate.row_index, NESTED_COL] = "Yes"

            for candidate in involved_candidates:
                if candidate.row_index == keep_candidate.row_index:
                    continue
                if candidate.row_index not in remove_indexes:
                    remove_indexes.add(candidate.row_index)
                    rows_removed += 1

            nested_groups += 1

    return nested_groups, rows_removed


def process_sheet(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    if not REQUIRED_COLUMNS.issubset(df.columns):
        return df, {
            "original_rows": len(df),
            "output_rows": len(df),
            "nested_groups": 0,
            "rows_removed": 0,
        }

    df = df.copy()
    if NESTED_COL not in df.columns:
        df[NESTED_COL] = ""

    remove_indexes: set[Any] = set()
    nested_groups = 0
    rows_removed = 0

    for (_pmid, _sent_index), group in df.groupby(
        [PMID_COL, SENT_INDEX_COL],
        sort=False,
        dropna=False,
    ):
        group_nested, group_removed = process_sentence_group(
            df,
            group,
            remove_indexes,
        )
        nested_groups += group_nested
        rows_removed += group_removed

    output_df = df.drop(index=list(remove_indexes)).sort_index().reset_index(drop=True)
    return output_df, {
        "original_rows": len(df),
        "output_rows": len(output_df),
        "nested_groups": nested_groups,
        "rows_removed": rows_removed,
    }


def process_workbook(input_path: Path, output_path: Path) -> dict[str, int]:
    if not input_path.exists():
        raise FileNotFoundError(f"Not found: {input_path}")

    xls = pd.ExcelFile(input_path)
    totals = {
        "original_rows": 0,
        "output_rows": 0,
        "nested_groups": 0,
        "rows_removed": 0,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(input_path, sheet_name=sheet_name)

            if SHEET_NAME is None or sheet_name == SHEET_NAME:
                df, summary = process_sheet(df)
            else:
                summary = {
                    "original_rows": len(df),
                    "output_rows": len(df),
                    "nested_groups": 0,
                    "rows_removed": 0,
                }

            for key, value in summary.items():
                totals[key] += value

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return totals


def main() -> None:
    summary = process_workbook(INFILE, OUTFILE)

    print(f"Saved: {OUTFILE}")
    print(f"Original data rows: {summary['original_rows']}")
    print(f"Output data rows: {summary['output_rows']}")
    print(f"Nested groups updated: {summary['nested_groups']}")
    print(f"Rows removed: {summary['rows_removed']}")


if __name__ == "__main__":
    main()
