#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import re
from pathlib import Path
import pandas as pd
import spacy
import benepar
import nltk
from transformers import AutoTokenizer

# ---------------- spaCy + benepar ---------------- #
nlp = spacy.load("en_core_web_sm")
nlp.add_pipe("benepar", config={"model": "benepar_en3"})

tok = AutoTokenizer.from_pretrained("t5-small")   # benepar tokenizer

# ---------------- helpers ---------------- #
PUNCT_SKIP = {";", ":"}
PTB2CHAR = {
    "-LRB-": "(", "-RRB-": ")",
    "-LSB-": "[", "-RSB-": "]",
    "-LCB-": "{", "-RCB-": "}",
    "`": '"',  "''": '"',
}

def _cleanup(s: str) -> str:
    for k, v in PTB2CHAR.items():
        s = s.replace(k, v)
    return re.sub(r"\s+", " ", s).strip()

def strip_angle_markers(t):
    if not isinstance(t, str):
        return ""
    return t

def get_tree(sent):
    if hasattr(sent._, "constituent"):
        return sent._.constituent.convert()
    if hasattr(sent._, "parse_string"):
        return nltk.Tree.fromstring(sent._.parse_string)
    raise RuntimeError("No parse tree found.")

def _leaf_texts(tree):
    return [PTB2CHAR.get(x, x) for x in tree.leaves() if x not in PUNCT_SKIP]

def _is_base_np(node):
    return (
        isinstance(node, nltk.Tree)
        and node.label() == "NP"
        and all(not (isinstance(c, nltk.Tree) and c.label() == "NP") for c in node)
    )

def _pp_has_in_and_np(pp_node):
    """Check if PP has an IN and an NP as (direct) children."""
    if not isinstance(pp_node, nltk.Tree) or pp_node.label() != "PP":
        return False
    has_in = any(isinstance(ch, nltk.Tree) and ch.label() == "IN" for ch in pp_node)
    has_np = any(isinstance(ch, nltk.Tree) and ch.label() == "NP" for ch in pp_node)
    return has_in and has_np

def _align_leaves_to_sent(sent, leaf_tokens, full_sentence):
    """
    Align a sequence of leaf tokens to spaCy tokens in this sent.
    Return (start_char, end_char, text) or (None, None, "") if fail.
    """
    if not leaf_tokens:
        return None, None, ""

    for i in range(len(sent)):
        if sent[i].text == leaf_tokens[0]:
            ok = True
            for j in range(1, len(leaf_tokens)):
                if i + j >= len(sent) or sent[i + j].text != leaf_tokens[j]:
                    ok = False
                    break
            if ok:
                start_char = sent[i].idx
                end_tok = sent[i + len(leaf_tokens) - 1]
                end_char = end_tok.idx + len(end_tok.text)
                text = _cleanup(full_sentence[start_char:end_char])
                return start_char, end_char, text

    return None, None, ""

# ---------------- Extract base NPs + "group NPs" ---------------- #
def extract_base_nps_with_groups(sentence):
    """
    Returns list of:
    (base_text, base_start, base_end,
     group_text, group_start, group_end)

    group_text/group_start/group_end correspond to:
      - parent NP (if parent is NP), or
      - grandparent NP (if parent is PP with IN+NP and grandparent is NP),
      - "" / -1 / -1 otherwise.
    """

    sub_tokens = tok.tokenize(sentence)
    if len(sub_tokens) > 480:
        return []

    doc = nlp(sentence)
    output = []

    for sent in doc.sents:
        tree = get_tree(sent)

        def walk(node, parent, grandparent):
            if not isinstance(node, nltk.Tree):
                return

            if _is_base_np(node):
                leaf_tokens = _leaf_texts(node)
                b_start, b_end, base_text = _align_leaves_to_sent(
                    sent, leaf_tokens, sentence
                )
                if b_start is not None:
                    group_text = ""
                    g_start = -1
                    g_end = -1
                    candidate_np = None

                    if isinstance(parent, nltk.Tree) and parent.label() == "NP":
                        candidate_np = parent
                    elif (
                        isinstance(parent, nltk.Tree)
                        and parent.label() == "PP"
                        and isinstance(grandparent, nltk.Tree)
                        and grandparent.label() == "NP"
                        and _pp_has_in_and_np(parent)
                    ):
                        candidate_np = grandparent

                    if candidate_np is not None:
                        p_leaf_tokens = _leaf_texts(candidate_np)
                        g_start, g_end, group_text = _align_leaves_to_sent(
                            sent, p_leaf_tokens, sentence
                        )

                    output.append(
                        (base_text, b_start, b_end, group_text, g_start, g_end)
                    )

            for child in node:
                walk(child, node, parent)

        walk(tree, None, None)

    return output

# ---------------- Find best base NP by offsets ---------------- #
def find_base_np(np_structs, t_start, t_end):
    """
    np_structs is list of:
      (base_text, base_start, base_end, group_text, group_start, group_end)
    We pick the smallest base NP overlapping the term span.
    """
    candidates = []
    for entry in np_structs:
        b_text, b_s, b_e, g_text, g_s, g_e = entry
        if not (t_end <= b_s or t_start >= b_e):
            candidates.append(entry)

    if not candidates:
        return None

    candidates.sort(key=lambda x: len(x[0].split()))
    return candidates[0]

# ---------------- MAIN ---------------- #
def main(
    infile=work("2g_pattern_list_merged.xlsx"),
    outfile=work("3_withMergedBaseNPs.xlsx"),
    sent_col="sent_text",
    term_col="BERT_Output",
    start_col="start_offset",
    end_col="end_offset",
    base_col="BERT_Base_NPs",
    merged_col="Merged_Base_NPs",
    discard_col="Discard",
):
    if not Path(infile).exists():
        raise FileNotFoundError(f"Input file not found: {infile}")

    df = pd.read_excel(infile)
    np_cache = {}

    if base_col not in df.columns:
        raise ValueError(f"Missing column: {base_col}")

    group_np_list = []

    # -------- First pass: compute only NP_Group, do not touch BERT_Base_NPs -------- #
    for _, row in df.iterrows():
        discard_val = str(row.get(discard_col, "")).strip().lower()

        if discard_val == "yes":
            group_np_list.append("")
            continue

        sent = strip_angle_markers(row[sent_col])

        try:
            s_o = int(row[start_col])
            e_o = int(row[end_col])
        except Exception:
            s_o = row[start_col]
            e_o = row[end_col]

        if sent not in np_cache:
            np_cache[sent] = extract_base_nps_with_groups(sent)

        match = find_base_np(np_cache[sent], s_o, e_o)

        if match is None:
            group_np_list.append("")
        else:
            base_text, b_s, b_e, group_text, g_s, g_e = match
            group_np_list.append(group_text)

    df["NP_Group"] = group_np_list

    # -------- Second pass: merge only if 2+ non-discard rows share same NP_Group -------- #
    merged_values = []
    grouping = {}

    for idx, row in df.iterrows():
        discard_val = str(row.get(discard_col, "")).strip().lower()
        if discard_val == "yes":
            continue

        key = (row[sent_col], row["NP_Group"])
        if row["NP_Group"]:
            grouping.setdefault(key, []).append(idx)

    for idx, row in df.iterrows():
        discard_val = str(row.get(discard_col, "")).strip().lower()

        if discard_val == "yes":
            merged_values.append("")
            continue

        key = (row[sent_col], row["NP_Group"])
        rows_in_same_group = grouping.get(key, [])

        if row["NP_Group"] and len(rows_in_same_group) >= 2:
            merged_values.append(row["NP_Group"])
        else:
            merged_values.append("")

    df[merged_col] = merged_values

    if "NP_Group" in df.columns:
        df.drop(columns=["NP_Group"], inplace=True)

    df.to_excel(outfile, index=False)
    print(f"✓ Done. Added only {merged_col} and wrote output to {outfile}")

if __name__ == "__main__":
    main()