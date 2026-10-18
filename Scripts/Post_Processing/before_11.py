#!/usr/bin/env python3
import pandas as pd
from pathlib import Path
import re
import os

STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
STAGE2_DIR = os.environ.get("GLYCANER_STAGE2_DIR", "/data/output/post_processing")
os.makedirs(STAGE2_DIR, exist_ok=True)

# -----------------------------------------------------------
# Sequential matching (correct behavior)
# -----------------------------------------------------------

def find_offsets_sequential(sentence, terms):
    """
    Sequential, non-overlapping, full-string matching.
    Ensures terms are matched exactly, not as substrings.
    Case-insensitive.
    """
    offsets = []
    s_low = sentence.lower()
    pos = 0

    for term in terms:
        t_low = term.lower()

        # EXACT TERM MATCH -- no substring match allowed
        pattern = (
            r"(?<![A-Za-z0-9])"     # no letter/number before term
            + re.escape(t_low) +
            r"(?![A-Za-z0-9])"      # no letter/number after term
        )

        regex = re.compile(pattern, re.IGNORECASE)
        match = regex.search(s_low, pos)

        if not match:
            offsets.append(-1)
        else:
            start = match.start()
            offsets.append(start)
            pos = start + len(t_low)

    return offsets



# -----------------------------------------------------------
# Main processing
# -----------------------------------------------------------

def process_file(
    infile=os.path.join(STAGE1_DIR, "final_gst1.xlsx"),
    sheet_name="Sheet1",
    pmid_col="pmid",
    sent_col="sent_text",
    bert_col="BERT_Output",
    index_col="sent_index",
    outfile=os.path.join(STAGE2_DIR, "gst_test_modified.xlsx")
):

    if not Path(infile).exists():
        raise FileNotFoundError(f"File not found: {infile}")

    print(f"Reading: {infile}")
    df = pd.read_excel(infile, sheet_name=sheet_name)

    # Drop only unwanted columns
    for col in ["charStart", "charEnd"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    expanded_rows = []

    for _, row in df.iterrows():
        sentence = str(row.get(sent_col, ""))
        pmid = row.get(pmid_col, None)
        sent_idx = row.get(index_col, None)
        bert_raw = row.get(bert_col)

        # -----------------------------------------------------------
        # Case 1: empty BERT_Output → keep one row
        # -----------------------------------------------------------
        if pd.isna(bert_raw) or not isinstance(bert_raw, str) or not bert_raw.strip():
            expanded_rows.append({
                pmid_col: pmid,
                index_col: sent_idx,
                sent_col: sentence,
                "term": "",
                "start_offset": -1,
                "end_offset": -1
            })
            continue

        # -----------------------------------------------------------
        # Case 2: normal processing
        # -----------------------------------------------------------
        terms = [t.strip() for t in bert_raw.split("|") if t.strip()]

        starts = find_offsets_sequential(sentence, terms)

        for term, start in zip(terms, starts):
            if start == -1:
                end = -1
            else:
                end = start + len(term)

            expanded_rows.append({
                pmid_col: pmid,
                index_col: sent_idx,
                sent_col: sentence,
                "term": term,
                "start_offset": start,
                "end_offset": end
            })

    out_df = pd.DataFrame(expanded_rows)

    print(f"Writing: {outfile}")
    out_df.to_excel(outfile, index=False)

    print("✓ Done. Added start_offset and end_offset correctly.")


if __name__ == "__main__":
    process_file()
