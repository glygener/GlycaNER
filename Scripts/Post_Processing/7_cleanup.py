#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
"""
Clean the conjunction/parentheses-resolved output.

Rules:
1. Use the previous generated output Excel as input.
2. Find rows where BERT_Output is only one word and that word ends with:
      osyl, lyl, or ated
   Examples removed/blanked:
      fucosyl
      galactosyl
      hexaosyl
      acetyl
      sulfated
3. If another row exists with the same sent_index and pmid, remove the entire
   incomplete row.
4. If no other row exists with the same sent_index and pmid, keep the row but
   blank only BERT_Output.
5. Rename BERT_Output to Final_GST_String at the end.
"""

import re
import pandas as pd


# -------------------------
# PATHS
# -------------------------
# Change these paths as needed.
INPUT_XLSX = work("6b_conjunction_resolved.xlsx")
OUTPUT_XLSX = work("7_final_gst_string.xlsx")

BERT_COL = "BERT_Output"
FINAL_COL = "Final_GST_String"
PMID_COL = "pmid"
SENT_INDEX_COL = "sent_index"

# A single token only, ending with osyl, lyl, or ated.
# Allows letters, digits, apostrophe, underscore, and hyphen inside the token,
# but no spaces.
REMOVE_SINGLE_WORD_RE = re.compile(
    r"""(?ix)
    ^
    (?:
        [A-Za-z0-9-]+(?:osyl|lyl|ated)
        |sulfates?
        |phosphates?
    )
    $
    """
)


def normalize_text(value) -> str:
    """Convert Excel cell value to a safe stripped string."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_incomplete_one_word(value) -> bool:
    """True when value is exactly one word ending with osyl, lyl, or ated."""
    text = normalize_text(value)
    if not text:
        return False

    # Only one whitespace-free token is allowed.
    if re.search(r"\s", text):
        return False

    return REMOVE_SINGLE_WORD_RE.fullmatch(text) is not None


def main():
    df = pd.read_excel(INPUT_XLSX)

    required_cols = [BERT_COL, PMID_COL, SENT_INDEX_COL]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {missing}")

    df = df.copy()

    # Mark rows where BERT_Output is a single incomplete word.
    incomplete_mask = df[BERT_COL].apply(is_incomplete_one_word)

    # Count rows per same pmid and sent_index group.
    group_sizes = df.groupby([PMID_COL, SENT_INDEX_COL], dropna=False)[BERT_COL].transform("size")

    # If the same sent_index and pmid is present in multiple rows,
    # remove the entire incomplete row.
    drop_mask = incomplete_mask & (group_sizes > 1)

    # If this is the only row for that pmid/sent_index, blank only BERT_Output.
    blank_mask = incomplete_mask & (group_sizes == 1)

    df.loc[blank_mask, BERT_COL] = ""
    df = df.loc[~drop_mask].copy()

    # Rename BERT_Output to Final_GST_String at the very end.
    df = df.rename(columns={BERT_COL: FINAL_COL})

    df.to_excel(OUTPUT_XLSX, index=False)

    print(f"Input rows: {len(incomplete_mask)}")
    print(f"Rows removed: {int(drop_mask.sum())}")
    print(f"Rows blanked: {int(blank_mask.sum())}")
    print(f"Done. Output saved to: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
