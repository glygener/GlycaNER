#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, dictionary_files, resource, work  # noqa: E402
import os
import re
import shutil
import tempfile

import pandas as pd


# ----------------------------
# Paths
# ----------------------------
input_path = (
    work("8_final_gst_string_classified.xlsx")
)

output_path = (
    work("9a_normalization.xlsx")
)

dict_dir = resource("Dict_Normalization")

sheets_to_process = {"GCT", "Sheet1"}
required_columns = ["pmid", "sent_index", "sent_text", "Final_GST_String"]

# Dictionary groups that act as modifiers
modifier_dicts = {"digit_modifier", "bi_di_tri", "globo"}


# ----------------------------
# Load all dictionary files
# ----------------------------
dictionaries = {}

for file in dictionary_files(dict_dir):
    if file.endswith(".txt"):
        dict_name = file.replace(".txt", "")
        with open(os.path.join(dict_dir, file), "r", encoding="utf-8") as f:
            dictionaries[dict_name] = [line.strip() for line in f if line.strip()]


def process_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing required column(s): {', '.join(missing)}"
        )

    # Keep only required input columns from GCT / Sheet1. Extra columns are ignored.
    df = df[required_columns].copy()

    # Add output columns only for non-modifier dictionaries.
    for dict_name in dictionaries:
        if dict_name not in modifier_dicts:
            df[dict_name] = None

    df["Unmatched"] = ""

    for idx, gst_term in df["Final_GST_String"].items():
        gst_str = "" if pd.isna(gst_term) else str(gst_term)

        if gst_str.strip() == "":
            df.at[idx, "Unmatched"] = ""
            continue

        global_matches = []

        for dict_name, patterns in dictionaries.items():
            for pattern in patterns:
                try:
                    for m in re.finditer(pattern, gst_str):
                        global_matches.append(
                            {
                                "match": m.group(),
                                "start": m.start(),
                                "end": m.end(),
                                "length": len(m.group()),
                                "dict": dict_name,
                            }
                        )
                except re.error as e:
                    print(f"Regex error in {dict_name!r} with pattern {pattern!r}: {e}")

        if not global_matches:
            df.at[idx, "Unmatched"] = gst_str
            continue

        global_matches.sort(key=lambda x: (x["start"], -x["length"]))

        matched_positions = [0] * len(gst_str)
        accepted_spans = []

        matched_terms_per_dict = {
            dict_name: []
            for dict_name in dictionaries
            if dict_name not in modifier_dicts
        }
        pending_modifiers = []

        for match_info in global_matches:
            start = match_info["start"]
            end = match_info["end"]
            match_text = match_info["match"]
            dict_name = match_info["dict"]

            if any(matched_positions[start:end]):
                continue

            for k in range(start, end):
                matched_positions[k] = 1

            accepted_spans.append((start, end))

            if dict_name in modifier_dicts:
                pending_modifiers.append(match_text)
            else:
                if pending_modifiers:
                    match_text = " + ".join(pending_modifiers) + " + " + match_text
                    pending_modifiers = []

                matched_terms_per_dict[dict_name].append(match_text)

        for dict_name, terms in matched_terms_per_dict.items():
            if terms:
                df.at[idx, dict_name] = " | ".join(terms)

        accepted_spans.sort()

        merged = []
        for s, e in accepted_spans:
            if not merged or s > merged[-1][1]:
                merged.append([s, e])
            else:
                merged[-1][1] = max(merged[-1][1], e)

        unmatched_parts = []
        cursor = 0

        for s, e in merged:
            if cursor < s:
                frag = gst_str[cursor:s].strip()
                if frag:
                    unmatched_parts.append(frag)
            cursor = e

        if cursor < len(gst_str):
            frag = gst_str[cursor:].strip()
            if frag:
                unmatched_parts.append(frag)

        df.at[idx, "Unmatched"] = " | ".join(unmatched_parts) if unmatched_parts else ""

    return df


# ----------------------------
# Load workbook, process only GCT and Sheet1, keep other sheets unchanged
# ----------------------------
all_sheets = pd.read_excel(input_path, sheet_name=None)

output_dir = os.path.dirname(output_path)
os.makedirs(output_dir, exist_ok=True)

tmp_dir = tempfile.mkdtemp(prefix="gst_normalization1_")
tmp_output_path = os.path.join(tmp_dir, os.path.basename(output_path))

try:
    with pd.ExcelWriter(tmp_output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in all_sheets.items():
            if sheet_name in sheets_to_process:
                df = process_sheet(df, sheet_name)

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    shutil.move(tmp_output_path, output_path)

finally:
    shutil.rmtree(tmp_dir, ignore_errors=True)

print(f"Done. Written to: {output_path}")