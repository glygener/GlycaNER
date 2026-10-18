#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import re
from pathlib import Path
import pandas as pd
import spacy
import benepar
import nltk

# ---------------- spaCy + benepar ---------------- #
nlp = spacy.load("en_core_web_sm")
nlp.add_pipe("benepar", config={"model": "benepar_en3"})

# ---------------- helper utilities ---------------- #
PUNCT_SKIP = {";", ":"}
PRONOUNS = {"i", "we", "you", "he", "she", "it", "they"}

PTB2CHAR = {
    "-LRB-": "(", "-RRB-": ")",
    "-LSB-": "[", "-RSB-": "]",
    "-LCB-": "{", "-RCB-": "}",
    "`": '"', "''": '"',
}

COORD_WORDS = {"and", "or"}
SPLIT_PUNCT = {","}

LONG_SENTENCE_COL = "Long_Sentence"

def extract_all_nps_with_char_offsets(sentence: str):
    doc = nlp(sentence)
    out = []
    seen = set()

    for sent in doc.sents:
        if not hasattr(sent._, "constituents"):
            continue

        for span in sent._.constituents:
            labels = getattr(span._, "labels", ())
            if "NP" not in labels:
                continue

            s = span.start_char
            e = span.end_char
            key = (s, e)
            if key in seen:
                continue
            seen.add(key)

            txt = sentence[s:e]
            txt = _cleanup_ptb(txt)
            out.append((txt, s, e))

    return out

def split_coordinated_np_into_segments(np_tree: nltk.Tree):
    items = [(normalize_leaf(w), t) for (w, t) in np_tree.pos()]

    split_idx = []
    for i, (w, t) in enumerate(items):
        if (t == "CC" and w.lower() in COORD_WORDS) or (w in SPLIT_PUNCT) or (t == ","):
            split_idx.append(i)

    if not split_idx:
        return []

    segments = []
    start = 0
    for k in split_idx + [len(items)]:
        seg = items[start:k]
        start = k + 1

        seg_words = [w for (w, _) in seg if w]
        if not seg_words:
            continue

        has_noun = any(tag.startswith("NN") for (_, tag) in seg)
        if has_noun:
            segments.append(seg_words)

    return segments if len(segments) >= 2 else []

def fix_offsets_if_needed(sent: str, term: str, s: int, e: int):
    s = int(s)
    e = int(e)

    if sent[s:e] == term:
        return s, e

    if e + 1 <= len(sent) and sent[s:e + 1] == term:
        return s, e + 1
    if s - 1 >= 0 and sent[s - 1:e] == term:
        return s - 1, e
    if s - 1 >= 0 and e + 1 <= len(sent) and sent[s - 1:e + 1] == term:
        return s - 1, e + 1

    hits = [(m.start(), m.end()) for m in re.finditer(re.escape(term), sent)]
    if not hits:
        return s, e

    return min(hits, key=lambda x: abs(x[0] - s) + abs(x[1] - e))

def strip_angle_markers_preserve_offsets(text: str) -> str:
    return text

def normalize_for_cache_key(text: str) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()

def _get_nltk_tree(sent):
    if hasattr(sent._, "constituent"):
        return sent._.constituent.convert()
    if hasattr(sent._, "parse_string"):
        return nltk.Tree.fromstring(sent._.parse_string)
    raise RuntimeError("No parse tree for this sentence.")

def _cleanup_ptb(phrase: str) -> str:
    for ptb, real in PTB2CHAR.items():
        phrase = phrase.replace(ptb, real)
    return re.sub(r"\s+", " ", phrase).strip()

def normalize_leaf(tok: str) -> str:
    tok = str(tok)
    for ptb, real in PTB2CHAR.items():
        tok = tok.replace(ptb, real)
        tok = tok.replace(" -", "-")
    tok = re.sub(r"\[\-", "[", tok)
    tok = re.sub(r"\-\]", "]", tok)
    tok = re.sub(r"\s+", " ", tok).strip()
    return tok

def _is_base_np(tree):
    return tree.label() == "NP" and all(
        not (isinstance(child, nltk.Tree) and child.label() == "NP")
        for child in tree
    )

def extract_base_nps_with_char_offsets(sentence):
    doc = nlp(sentence)
    base_np_list = []
    seen_spans = set()

    for sent in doc.sents:
        tree = _get_nltk_tree(sent)

        def add_span_by_token_sequence(token_texts):
            L = len(token_texts)
            if L == 0:
                return

            for i in range(0, len(sent) - L + 1):
                ok = True
                for j in range(L):
                    if sent[i + j].text != token_texts[j]:
                        ok = False
                        break
                if not ok:
                    continue

                start_tok = sent[i]
                end_tok = sent[i + L - 1]
                start_char = start_tok.idx
                end_char = end_tok.idx + len(end_tok.text)

                span_key = (start_char, end_char)
                if span_key in seen_spans:
                    continue
                seen_spans.add(span_key)

                np_text = sentence[start_char:end_char]
                np_text = _cleanup_ptb(np_text)
                base_np_list.append((np_text, start_char, end_char))

        def walk(t):
            if not isinstance(t, nltk.Tree):
                return

            if _is_base_np(t):
                segments = split_coordinated_np_into_segments(t)

                if segments:
                    for seg_words in segments:
                        add_span_by_token_sequence(seg_words)
                else:
                    leaves = [normalize_leaf(x) for x in t.leaves()]
                    add_span_by_token_sequence(leaves)

            for child in t:
                walk(child)

        walk(tree)

    return base_np_list

def find_np_by_offset(np_list, term_start, term_end):
    term_start = int(term_start)
    term_end = int(term_end)

    overlaps = []
    contains = []

    for np_text, np_s, np_e in np_list:
        if not (term_end <= np_s or term_start >= np_e):
            overlaps.append((np_text, np_s, np_e))
        if term_start >= np_s and term_end <= np_e:
            contains.append((np_text, np_s, np_e))

    candidates = contains if contains else overlaps

    if not candidates:
        return ""

    candidates.sort(key=lambda x: (x[2] - x[1], len(x[0])))
    return candidates[0][0]

def is_tokenizer_length_error(err: Exception) -> bool:
    msg = str(err)

    if isinstance(err, AssertionError):
        return True

    return (
        "exceeds the maximum supported length" in msg
        or "Token indices sequence length is longer" in msg
        or "indexing errors" in msg
        or "offset_mapping_iter" in msg
        or "retokenize" in msg
    )

def main(
    infile=work("bert_output_full_updated_modified.xlsx"),
    outfile=work("withBaseNPs_full.xlsx"),
    sheet_name="Sheet1",
    sent_col="sent_text",
    bert_col="BERT_Output",
    start_col="start_offset",
    end_col="end_offset",
    target_col="BERT_Base_NPs",
):
    if not Path(infile).exists():
        raise FileNotFoundError(f"{infile} not found.")

    xls = pd.ExcelFile(infile)
    df = pd.read_excel(infile, sheet_name=sheet_name)

    base_np_results = []
    long_sentence_flags = []
    np_cache = {}

    for _, row in df.iterrows():
        sent = strip_angle_markers_preserve_offsets(row[sent_col])
        term = str(row[bert_col]).strip()

        t_start = int(row[start_col])
        t_end = int(row[end_col])

        t_start, t_end = fix_offsets_if_needed(sent, term, t_start, t_end)

        cache_key = normalize_for_cache_key(sent)

        if cache_key not in np_cache:
            try:
                np_cache[cache_key] = {
                    "base": extract_base_nps_with_char_offsets(sent),
                    "all": extract_all_nps_with_char_offsets(sent),
                    "long": False,
                }
            except Exception as e:
                if (
                    isinstance(e, (AssertionError, StopIteration, ValueError))
                    or is_tokenizer_length_error(e)
                ):
                    np_cache[cache_key] = {
                        "base": [],
                        "all": [],
                        "long": True,
                    }
                else:
                    raise

        if np_cache[cache_key]["long"]:
            base_np_results.append("")
            long_sentence_flags.append("Yes")
            continue

        base_list = np_cache[cache_key]["base"]
        all_list = np_cache[cache_key]["all"]

        base_np = find_np_by_offset(base_list, t_start, t_end)

        if (not base_np) or (len(base_np) < len(term)):
            base_np = find_np_by_offset(all_list, t_start, t_end)

        base_np_results.append(base_np)
        long_sentence_flags.append("")

    df[target_col] = base_np_results
    df[LONG_SENTENCE_COL] = long_sentence_flags

    # ===== CLEAN BERT_Base_NPs VALUES ONLY, DO NOT REMOVE ROWS =====
    before_rows = len(df)

    df[bert_col] = df[bert_col].fillna("").astype(str).str.strip()
    df[target_col] = df[target_col].fillna("").astype(str).str.strip()

    # Case 1: BERT_Output is empty but BERT_Base_NPs has a value
    mask_empty_bert = (df[bert_col] == "") & (df[target_col] != "")

    # Case 2: BERT_Base_NPs is only A, B, or H
    bad_values = {"A", "B", "H"}
    mask_bad_base_np = df[target_col].isin(bad_values)

    # Clear only BERT_Base_NPs
    df.loc[mask_empty_bert, target_col] = ""
    df.loc[mask_bad_base_np, target_col] = ""

    print("\n===== DEBUG: BERT_Base_NPs VALUE CLEANING =====")
    print(f"Total rows: {before_rows}")
    print(f"BERT_Base_NPs cleared because BERT_Output was empty: {mask_empty_bert.sum()}")
    print(f"BERT_Base_NPs cleared because value was A/B/H: {mask_bad_base_np.sum()}")
    print(f"Total BERT_Base_NPs values cleared: {(mask_empty_bert | mask_bad_base_np).sum()}")
    print(f"Total rows after cleaning: {len(df)}")

    print("\nSample rows cleared because BERT_Output was empty:")
    print(df.loc[mask_empty_bert, [bert_col, target_col]].head(5))

    print("\nSample rows cleared because BERT_Base_NPs was A/B/H:")
    print(df.loc[mask_bad_base_np, [bert_col, target_col]].head(5))

    with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
        for s in xls.sheet_names:
            if s == sheet_name:
                df.to_excel(writer, sheet_name=s, index=False)
            else:
                pd.read_excel(infile, sheet_name=s).to_excel(writer, sheet_name=s, index=False)

    print(f"✓ BaseNPs written to: {outfile}")

if __name__ == "__main__":
    main()