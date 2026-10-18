#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import re
from pathlib import Path

import pandas as pd


INFILE = work("colon_slash_split.xlsx")
OUTFILE = work("withBaseNPs_post.xlsx")

SHEET_NAME = None  # use None to process every sheet with the required columns
PMID_COL = "pmid"
SENT_INDEX_COL = "sent_index"
BASE_COL = "BERT_Base_NPs"
TERM_COL = "BERT_Output"
MODIFY_COL = "Modify"
DISCARD_COL = "Discard"

# punctuation to ignore at the end of the BaseNP / BERT_Output
TRAIL_PUNCT_RE = re.compile(r"[,.;:!?/\\|()\[\]{}\s\-–—]+$")

# last-word endings to flag as Discard=Yes for BERT_Base_NPs
DISCARD_LASTWORD_RE = re.compile(
    r"(GlyCAM-1|proteins?|genes?|transporters?|peptides?|lectins?|cytes?|substrates?|phenotypes?|sites?|domains?|repeats?|receptors?|esters?|adhesins?|mRNAs?|RNA|DNA|enzymes?|antibody|antibodies|family|families|cells?|mouse|mice|celllines?|cell-lines?|ases?|fragments?|osis|(?<![A-Za-z0-9])lipids?|thesis|oses)(?:[\s-]+(?:\d+|[IVXLCDM]+|[A-Za-z]{1,5}))?$",
    re.IGNORECASE,
)


def is_empty_cell(x) -> bool:
    # True if NaN/None or empty/whitespace string
    if pd.isna(x):
        return True
    return str(x).strip() == ""


def clean_trailing_punct(value) -> str:
    if is_empty_cell(value):
        return ""
    return TRAIL_PUNCT_RE.sub("", str(value).strip())


def last_word(value) -> str:
    s = clean_trailing_punct(value)
    if not s:
        return ""
    parts = s.split()
    return parts[-1] if parts else ""


def bert_output_ends_with_upper_t(term) -> bool:
    word = last_word(term)
    if not word:
        return False
    return bool(re.search(r"[A-Za-z0-9][A-Za-z0-9_./-]*T$", word))


def is_one_word_anti_term(term) -> bool:
    if is_empty_cell(term):
        return False
    s = str(term).strip()
    return bool(re.fullmatch(r"(?i)anti-\S+", s))


def base_np_ends_with_nt(base_np) -> bool:
    word = last_word(base_np)
    return bool(re.search(r"nTs?$", word))


def base_np_discard_flag(base_np) -> bool:
    if is_empty_cell(base_np):
        return False

    s = clean_trailing_punct(base_np)
    return bool(DISCARD_LASTWORD_RE.search(s) or base_np_ends_with_nt(s))


def is_yes(value) -> bool:
    return str(value).strip().casefold() == "yes"


def set_discard_yes(df: pd.DataFrame, mask) -> None:
    if DISCARD_COL not in df.columns:
        df[DISCARD_COL] = ""
    df.loc[mask, DISCARD_COL] = "Yes"


def split_colon_base_np_for_consecutive_modify_rows(df: pd.DataFrame) -> None:
    required = {PMID_COL, SENT_INDEX_COL, BASE_COL}
    if not required.issubset(df.columns):
        return

    modify_col = MODIFY_COL
    if modify_col not in df.columns:
        if "Modified" in df.columns:
            modify_col = "Modified"
        else:
            return

    i = 0
    while i < len(df) - 1:
        row1 = df.iloc[i]
        row2 = df.iloc[i + 1]

        same_sentence = (
            row1[PMID_COL] == row2[PMID_COL]
            and row1[SENT_INDEX_COL] == row2[SENT_INDEX_COL]
        )
        both_modified = is_yes(row1[modify_col]) and is_yes(row2[modify_col])

        if same_sentence and both_modified:
            base1 = "" if is_empty_cell(row1[BASE_COL]) else str(row1[BASE_COL])
            base2 = "" if is_empty_cell(row2[BASE_COL]) else str(row2[BASE_COL])
            base_with_colon = base1 if ":" in base1 else base2 if ":" in base2 else ""

            if base_with_colon:
                left, right = base_with_colon.split(":", 1)
                df.at[df.index[i], BASE_COL] = left.strip()
                df.at[df.index[i + 1], BASE_COL] = right.strip()
                i += 2
                continue

        i += 1


def main():
    infile = Path(INFILE)
    if not infile.exists():
        raise FileNotFoundError(f"Not found: {INFILE}")

    xls = pd.ExcelFile(INFILE)

    with pd.ExcelWriter(OUTFILE, engine="openpyxl") as writer:
        for sh in xls.sheet_names:
            df = pd.read_excel(INFILE, sheet_name=sh)

            should_process_sheet = SHEET_NAME is None or sh == SHEET_NAME
            if should_process_sheet and BASE_COL in df.columns and TERM_COL in df.columns:
                df[DISCARD_COL] = ""

                # 1) Before other rules, discard BERT_Output ending with uppercase T
                #    and one-word BERT_Output values starting with anti-.
                term_discard_mask = df[TERM_COL].apply(
                    lambda term: bert_output_ends_with_upper_t(term)
                    or is_one_word_anti_term(term)
                )
                set_discard_yes(df, term_discard_mask)

                # 2) If two consecutive rows for the same PMID/sentence have Modify=Yes,
                #    split BERT_Base_NPs at ":" and assign left/right pieces.
                split_colon_base_np_for_consecutive_modify_rows(df)

                # 3) If BaseNP exists but BERT_Output is empty/missing, remove BaseNP.
                mask_remove = df.apply(
                    lambda r: (not is_empty_cell(r[BASE_COL]))
                    and is_empty_cell(r[TERM_COL]),
                    axis=1,
                )
                df.loc[mask_remove, BASE_COL] = ""

                # 4) Add Discard=Yes based on BaseNP last-word ending rules,
                #    including BaseNP ending with nT.
                base_discard_mask = df[BASE_COL].apply(base_np_discard_flag)
                set_discard_yes(df, base_discard_mask)

            df.to_excel(writer, sheet_name=sh, index=False)

    print(f"✓ Wrote: {OUTFILE}")


if __name__ == "__main__":
    main()
