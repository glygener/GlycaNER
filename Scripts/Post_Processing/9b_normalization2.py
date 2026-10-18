#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import os
import re
import shutil
import tempfile

import pandas as pd


# ----------------------------
# Paths
# ----------------------------
file_path = (
    work("9a_normalization.xlsx")
)
output_path = (
    work("9b_matched_entities_gold.xlsx")
)

mapping_file = (
    resource("Mapping_File/mapping_file.txt")
)

sheets_to_process = {"GCT", "Sheet1"}


# ----------------------------
# Load grouped replacement rules
# ----------------------------
replacements = []

with open(mapping_file, "r", encoding="utf-8") as f:
    for line in f:
        if "->" not in line:
            continue

        lhs, rhs = map(str.strip, line.split("->", 1))
        variants = [v.strip() for v in lhs.split("|") if v.strip()]

        for var in variants:
            # Add word boundaries only for pure alphanumeric tokens.
            if re.match(r"^\w+$", var):
                pattern = r"\b" + re.escape(var) + r"\b"
            else:
                pattern = re.escape(var)

            replacements.append((re.compile(pattern, flags=re.IGNORECASE), rhs))


def apply_replacements(text: str) -> str:
    for pattern, repl in replacements:
        text = pattern.sub(repl, text)
    return text


def process_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Exclude all columns up to H inclusive, plus the last column.
    # A = 1, H = 8
    exclude_upto_col = 8
    exclude_columns = set(df.columns[:exclude_upto_col])

    if len(df.columns) > 0:
        exclude_columns.add(df.columns[-1])

    gold_values = []

    for _, row in df.iterrows():
        gold_parts = []

        for col in df.columns:
            if col in exclude_columns or pd.isna(row[col]):
                continue

            parts = str(row[col]).split(" | ")

            for part in parts:
                part = part.strip()
                if not part:
                    continue

                normalized = apply_replacements(part)
                gold_parts.append(normalized)

        gold_values.append("; ".join(gold_parts))

    df["Gold"] = gold_values
    return df


# ----------------------------
# Process only GCT and Sheet1, keep other sheets unchanged
# ----------------------------
all_sheets = pd.read_excel(file_path, sheet_name=None)


# ----------------------------
# Save output using temp file first to avoid openpyxl timeout
# ----------------------------
output_dir = os.path.dirname(output_path)
os.makedirs(output_dir, exist_ok=True)

tmp_dir = tempfile.mkdtemp(prefix="gst_matched_gold_")
tmp_output_path = os.path.join(tmp_dir, os.path.basename(output_path))

try:
    with pd.ExcelWriter(tmp_output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in all_sheets.items():
            if sheet_name in sheets_to_process:
                df = process_sheet(df)

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    shutil.move(tmp_output_path, output_path)

finally:
    shutil.rmtree(tmp_dir, ignore_errors=True)


print(f"Gold column created for GCT and Sheet1. Written to: {output_path}")