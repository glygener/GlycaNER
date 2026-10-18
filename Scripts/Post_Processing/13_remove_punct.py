#!/usr/bin/env python3
import pandas as pd
import re
from pathlib import Path
import os

STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
STAGE2_DIR = os.environ.get("GLYCANER_STAGE2_DIR", "/data/output/post_processing")
os.makedirs(STAGE2_DIR, exist_ok=True)

# ------------------------------
# CONFIG
# ------------------------------
INPUT_FILE  = os.path.join(STAGE2_DIR, "gst_test_merged.xlsx")
OUTPUT_FILE = os.path.join(STAGE2_DIR, "gst_punct_removed_cleaned.xlsx")

OLD_TERM_COL = "BERT_Output"       # input column name
NEW_TERM_COL = "BERT_Output"
START_COL = "start_offset"
END_COL   = "end_offset"
PUNCT_COL = "Punctuation"

# Target punctuation to remove (from END only)
REMOVE_END_CHARS = set([":", ",", ";", ".", "/", "-", "+", ")", "]"])


# ------------------------------
# CLEAN ONE TERM
# ------------------------------
def clean_term(term):
    if not isinstance(term, str):
        return "", ""

    t = term.strip()
    removed = []

    while len(t) > 0 and t[-1] in REMOVE_END_CHARS:
        ch = t[-1]

        # Only remove ")" if it has no opening "("
        if ch == ")":
            if "(" not in t:
                removed.append(ch)
                t = t[:-1]
                continue
            break

        # Keep same logic for unmatched closing bracket
        if ch == "]":
            if "[" not in t:
                removed.append(ch)
                t = t[:-1]
                continue
            break

        # General punctuation, including "+"
        removed.append(ch)
        t = t[:-1]

    return t, "".join(removed)


# ------------------------------
# MAIN
# ------------------------------
def main():
    if not Path(INPUT_FILE).exists():
        raise FileNotFoundError(INPUT_FILE)

    df = pd.read_excel(INPUT_FILE, dtype=str)

    # ensure no NaN anywhere (convert to empty)
    df = df.fillna("")

    if OLD_TERM_COL not in df.columns:
        raise ValueError(f"Column '{OLD_TERM_COL}' not found in sheet.")

    # Save the original column order
    cols = list(df.columns)

    cleaned_terms = []
    removed_puncts = []
    new_end_offsets = []

    for term, end_o in zip(df[OLD_TERM_COL], df[END_COL]):
        # enforce correct types
        term = "" if not isinstance(term, str) else term
        try:
            end_o = int(end_o)
        except:
            end_o = 0

        cleaned, removed = clean_term(term)
        new_end = end_o - len(removed)

        cleaned_terms.append(cleaned)
        removed_puncts.append(removed)
        new_end_offsets.append(new_end)

    # Replace "term" column with "BERT_Output" IN THE SAME POSITION
    term_index = cols.index(OLD_TERM_COL)
    df.drop(columns=[OLD_TERM_COL], inplace=True)

    # Insert new BERT_Output column in exact same position
    df.insert(term_index, NEW_TERM_COL, cleaned_terms)

    # Insert Punctuation column immediately after BERT_Output
    df.insert(term_index + 1, PUNCT_COL, removed_puncts)

    # Update end_offset
    df[END_COL] = new_end_offsets

    # Export
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✓ Completed successfully.")
    print(f"→ Output written to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
