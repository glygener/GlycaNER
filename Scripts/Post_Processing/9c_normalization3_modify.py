#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import os
import re
import shutil
import tempfile

import pandas as pd


# === 1. Paths ===
input_path = (
    work("9b_matched_entities_gold.xlsx")
)

reference_gold_path = (
    resource("GSD_Entries/output_with_gold.xlsx")
)

GSD_DICT_FILE = (
    resource("GSD_Entries/GWU_Dict.xlsx")
)

output_path = (
    work("9c_normalized_form.xlsx")
)


# === 2. Columns / sheet rules ===
GOLD_COL = "Gold"
MATCHED_COL = "Matched_parts_in_Final_GST_String"

COMBINED_COLUMNS = [
    "pmid",
    "sent_index",
    "sent_text",
    "Final_String",
    "Type",
    "GSD_ID",
    "GSD_Name",
    "Match",
]

SHEET_RULES = {
    "Sheet1": {
        "final_cols": ["Final_GST_String"],
        "type": "GST",
        "output_sheet": "GST",
    },
    "GST": {
        "final_cols": ["Final_GST_String"],
        "type": "GST",
        "output_sheet": "GST",
    },
    "GCT": {
        "final_cols": ["Final_GST_String"],
        "type": "GCT",
        "output_sheet": "GCT",
    },
    "Composition": {
        "final_cols": ["Composition", "Final_GST_String"],
        "type": "Composition",
        "output_sheet": "Composition",
    },
    "Sequence": {
        "final_cols": ["Final_GST_String", "Sequence"],
        "type": "Sequence",
        "output_sheet": "Sequence",
    },
    "Sugar": {
        "final_cols": ["Final_GST_String", "Sugar"],
        "type": "Sugar",
        "output_sheet": "Sugar",
    },
    "Generic": {
        "final_cols": ["Final_GST_String", "Generic", "Composition", "Sequence", "Sugar"],
        "type": "Generic",
        "output_sheet": "Generic",
    },
    "Oxford_glycan_notation": {
        "final_cols": ["Oxford_glycan_notation", "Final_GST_String"],
        "type": "GST",
        "output_sheet": "Oxford_glycan_notation",
    },
}


# === 3. Normalize / matching helpers ===
def clean_cell(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def normalize_term(term):
    s = str(term).strip().lower()
    s = re.sub(r"(?<=\d)['′`]", "", s)
    s = re.sub(r"-", "", s)
    return s


ALLOWED_EXTRAS = {
    "antigen", "antigens",
    "chain", "chains",
    "glycolipid", "glycolipids",
    "ceramide", "ceramides",
    "ganglioside", "gangliosides",
    "gangliosidosis", "gangliosidoses",
    "glycosphingolipid", "glycosphingolipids",
    "N-glycan", "N-glycans",
    "O-glycan", "O-glycans",
    "epitope", "epitopes",
    "structure", "structures",
    "glycan", "glycans",
    "determinant", "determinants",
    "glycoprotein", "glycoproteins",
    "glycoside", "glycosides",
    "carbohydrate", "carbohydrates",
}
ALLOWED_EXTRA_NORMS = {normalize_term(x) for x in ALLOWED_EXTRAS}


def split_gold_parts(gold_value):
    if pd.isna(gold_value) or str(gold_value).strip() == "":
        return []

    gold_raw_items = [x.strip() for x in str(gold_value).split(";") if x.strip()]
    return [part.strip() for x in gold_raw_items for part in x.split("+") if part.strip()]


def unique_preserve_order(items):
    seen = set()
    out = []

    for x in items:
        key = normalize_term(x)
        if key not in seen:
            out.append(x)
            seen.add(key)

    return out


# === 4. Build exact match map and detailed GST-Gold map ===
df2 = pd.read_excel(reference_gold_path)

exact_map = {}
gold_term_map = []

for _, row in df2.iterrows():
    if pd.isna(row["Gold"]) or pd.isna(row["GST"]):
        continue

    gst = str(row["GST"]).strip()
    acc = "" if pd.isna(row["GST_Accession"]) else str(row["GST_Accession"]).strip()

    gold_parts = split_gold_parts(row["Gold"])
    if not gold_parts:
        continue

    norm_items = [normalize_term(x) for x in gold_parts]
    gold_set = frozenset(norm_items)

    exact_map.setdefault(gold_set, set())
    if acc:
        exact_map[gold_set].add(acc)

    gold_term_map.append(
        {
            "GST": gst,
            "Gold_parts_raw": gold_parts,
            "Gold_norm": gold_set,
        }
    )


def match_gold_value(gold_value):
    df_gold_parts = split_gold_parts(gold_value)

    if not df_gold_parts:
        return "", "", ""

    df_gold_norm_list = [normalize_term(x) for x in df_gold_parts]
    df_gold_norm = frozenset(df_gold_norm_list)

    all_matched_norms = set()

    for entry in gold_term_map:
        overlap = df_gold_norm & entry["Gold_norm"]
        if overlap:
            all_matched_norms.update(overlap)

    matched_raw_parts = []
    for raw, norm in zip(df_gold_parts, df_gold_norm_list):
        if norm in all_matched_norms:
            matched_raw_parts.append(raw)

    matched_raw_parts = unique_preserve_order(matched_raw_parts)
    matched_parts = " + ".join(matched_raw_parts)

    if df_gold_norm in exact_map:
        matched_accessions = sorted(set(exact_map[df_gold_norm]))
        return "; ".join(matched_accessions), "", matched_parts

    partial_list = []

    for entry in gold_term_map:
        overlap = df_gold_norm & entry["Gold_norm"]
        if not overlap:
            continue

        missing = sorted(entry["Gold_norm"] - df_gold_norm)
        raw_extra = df_gold_norm - entry["Gold_norm"]
        extra = sorted([x for x in raw_extra if x not in ALLOWED_EXTRA_NORMS])

        if not missing and not extra:
            matched_accessions = sorted(set(exact_map.get(entry["Gold_norm"], [])))
            return "; ".join(matched_accessions), "", matched_parts

        parts = []
        if missing:
            parts.append(f"missing: {'; '.join(missing)}")
        if extra:
            parts.append(f"extra: {'; '.join(extra)}")

        diff_str = f"{entry['GST']} ({'; '.join(parts)})" if parts else entry["GST"]
        partial_list.append(diff_str)

    return "", "; ".join(partial_list), matched_parts


# === 5. GSD ID to GSD name map ===
gsd_df = pd.read_excel(GSD_DICT_FILE)

gsd_map = {}

for _, r in gsd_df.iterrows():
    gsd_id = clean_cell(r.get("GST_Accession", ""))
    gsd_name = clean_cell(r.get("GST", ""))

    if gsd_id and gsd_id not in gsd_map:
        gsd_map[gsd_id] = gsd_name


def get_gsd_names(gsd_ids):
    ids = [x.strip() for x in str(gsd_ids).split(";") if x.strip()]
    names = []

    for gsd_id in ids:
        name = gsd_map.get(gsd_id, "")
        if name:
            names.append(name)

    return "; ".join(names)


# === 6. Sheet processing helpers ===
def get_sheet_rule(sheet_name, df):
    rule = SHEET_RULES.get(sheet_name, {})

    final_cols = rule.get("final_cols", [])
    final_col = next((col for col in final_cols if col in df.columns), None)

    if final_col is None:
        if "Final_GST_String" in df.columns:
            final_col = "Final_GST_String"
        elif sheet_name in df.columns:
            final_col = sheet_name

    type_value = rule.get("type", sheet_name)
    output_sheet = rule.get("output_sheet", sheet_name)

    return final_col, type_value, output_sheet


def process_match_columns(df, final_col):
    df = df.copy()

    if GOLD_COL not in df.columns:
        return df

    exact_matches = []
    partial_matches = []
    matched_parts_values = []

    for _, row in df.iterrows():
        exact, partial, matched_parts = match_gold_value(row[GOLD_COL])
        exact_matches.append(exact)
        partial_matches.append(partial)
        matched_parts_values.append(matched_parts)

    if MATCHED_COL in df.columns:
        df.drop(columns=[MATCHED_COL], inplace=True)

    if final_col in df.columns:
        insert_at = list(df.columns).index(final_col) + 1
        df.insert(insert_at, MATCHED_COL, matched_parts_values)
    else:
        df[MATCHED_COL] = matched_parts_values

    df["Exact"] = exact_matches
    df["Partial"] = partial_matches

    return df


def term_position(sent_text, final_string):
    sent = clean_cell(sent_text).lower()
    term = clean_cell(final_string).lower()

    if not sent or not term:
        return 10**9

    candidates = [term]
    for delimiter in [" | ", ";", "+"]:
        candidates.extend([x.strip().lower() for x in term.split(delimiter) if x.strip()])

    positions = []
    for candidate in candidates:
        pos = sent.find(candidate)
        if pos != -1:
            positions.append(pos)

    return min(positions) if positions else 10**9


def numeric_sort_value(value):
    text = clean_cell(value)
    if text == "":
        return float("inf")

    try:
        return float(text)
    except ValueError:
        return float("inf")


def build_combined_rows(df, sheet_name, final_col, type_value, source_start):
    rows = []

    if "sent_index" not in df.columns:
        return rows

    for offset, (_, row) in enumerate(df.iterrows()):
        final_string = clean_cell(row.get(final_col, "")) if final_col else ""

        gsd_id = clean_cell(row.get("Exact", ""))
        gsd_name = get_gsd_names(gsd_id) if gsd_id else ""
        partial = clean_cell(row.get("Partial", ""))

        if gsd_id:
            match = "Exact"
        elif partial:
            match = "Partial"
        elif GOLD_COL in df.columns:
            match = "None"
        else:
            match = ""

        rows.append(
            {
                "pmid": row.get("pmid", ""),
                "sent_index": row.get("sent_index", ""),
                "sent_text": row.get("sent_text", ""),
                "Final_String": final_string,
                "Type": type_value if final_string else "",
                "GSD_ID": gsd_id,
                "GSD_Name": gsd_name,
                "Match": match,
                "_pmid_num": numeric_sort_value(row.get("pmid", "")),
                "_pmid_text": clean_cell(row.get("pmid", "")),
                "_sent_index_num": numeric_sort_value(row.get("sent_index", "")),
                "_sent_index_text": clean_cell(row.get("sent_index", "")),
                "_term_pos": term_position(row.get("sent_text", ""), final_string),
                "_source_order": source_start + offset,
            }
        )

    return rows


def make_unique_sheet_name(name, used_names):
    base = str(name)[:31]

    if base not in used_names:
        used_names.add(base)
        return base

    counter = 2
    while True:
        suffix = f"_{counter}"
        candidate = base[: 31 - len(suffix)] + suffix

        if candidate not in used_names:
            used_names.add(candidate)
            return candidate

        counter += 1


# === 7. Process workbook ===
all_sheets = pd.read_excel(input_path, sheet_name=None)

output_sheets = []
combined_rows = []
source_counter = 0

for sheet_name, df in all_sheets.items():
    if sheet_name == "Combined":
        continue

    final_col, type_value, output_sheet = get_sheet_rule(sheet_name, df)

    processed_df = process_match_columns(df, final_col)

    combined_rows.extend(
        build_combined_rows(
            processed_df,
            sheet_name,
            final_col,
            type_value,
            source_counter,
        )
    )
    source_counter += len(processed_df)

    output_sheets.append((output_sheet, processed_df))


if combined_rows:
    df_combined = pd.DataFrame(combined_rows)
    df_combined = df_combined.sort_values(
        by=[
            "_pmid_num",
            "_pmid_text",
            "_sent_index_num",
            "_sent_index_text",
            "_term_pos",
            "_source_order",
        ],
        ascending=[True, True, True, True, True, True],
        kind="stable",
    )
    df_combined = df_combined[COMBINED_COLUMNS]
else:
    df_combined = pd.DataFrame(columns=COMBINED_COLUMNS)


# === 8. Save output ===
output_dir = os.path.dirname(output_path)
os.makedirs(output_dir, exist_ok=True)

tmp_dir = tempfile.mkdtemp(prefix="gst_normalized_form_")
tmp_output_path = os.path.join(tmp_dir, os.path.basename(output_path))

try:
    written_sheet_names = set()

    with pd.ExcelWriter(tmp_output_path, engine="xlsxwriter") as writer:
        df_combined.to_excel(writer, sheet_name="Combined", index=False)
        written_sheet_names.add("Combined")

        for sheet_name, df in output_sheets:
            safe_sheet_name = make_unique_sheet_name(sheet_name, written_sheet_names)
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

    shutil.move(tmp_output_path, output_path)

finally:
    shutil.rmtree(tmp_dir, ignore_errors=True)


print(f"Matching complete. Output saved to:\n{output_path}")