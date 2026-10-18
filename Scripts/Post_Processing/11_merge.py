#!/usr/bin/env python3
import pandas as pd
from pathlib import Path
import os

STAGE1_DIR = os.environ.get("GLYCANER_STAGE1_DIR", "/data/output/bert_detection")
STAGE2_DIR = os.environ.get("GLYCANER_STAGE2_DIR", "/data/output/post_processing")
os.makedirs(STAGE2_DIR, exist_ok=True)

def merge_continuous_terms(
    infile=os.path.join(STAGE2_DIR, "gst_test_modified.xlsx"),
    outfile=os.path.join(STAGE2_DIR, "gst_test_merged.xlsx"),
    pmid_col="pmid",
    index_col="sent_index",
    sentent_col="sent_text",
    term_col="term",
    start_col="start_offset",
    end_col="end_offset"
):

    if not Path(infile).exists():
        raise FileNotFoundError(f"Input file not found: {infile}")

    print(f"Reading: {infile}")
    df = pd.read_excel(infile)

    # Sort by pmid, sent_index, start_offset
    df = df.sort_values([pmid_col, index_col, start_col])

    merged_rows = []

    for (pmid, sent_idx), group in df.groupby([pmid_col, index_col]):
        group = group.sort_values(start_col).reset_index(drop=True)

        i = 0
        while i < len(group):
            row = group.loc[i]

            # clean current term
            cur_term = row[term_col]
            if not isinstance(cur_term, str):
                cur_term = ""

            cur_start = row[start_col]
            cur_end = row[end_col]
            merged_flag = 0

            j = i + 1
            while j < len(group):

                next_row = group.loc[j]

                # Are they adjacent?
                adjacent = next_row[start_col] == cur_end + 1

                # BLOCK merge if cur_term ends with comma or semicolon
                boundary_block = cur_term.endswith((",", ";"))

                if adjacent and not boundary_block:

                    next_term = next_row[term_col]
                    if not isinstance(next_term, str):
                        next_term = ""

                    # merge
                    merged_flag = 1
                    cur_term = f"{cur_term} {next_term}".strip()
                    cur_end = next_row[end_col]

                    j += 1
                else:
                    break

            merged_rows.append({
                pmid_col: pmid,
                index_col: sent_idx,
                sentent_col: row[sentent_col],
                "BERT_Output": cur_term.strip(),
                start_col: cur_start,
                end_col: cur_end,
                "Merge": merged_flag
            })

            i = j

    out_df = pd.DataFrame(merged_rows)

    print(f"Writing merged output to: {outfile}")
    out_df.to_excel(outfile, index=False)

    print("✓ Merging completed with comma/semicolon boundary rule.")

if __name__ == "__main__":
    merge_continuous_terms()
