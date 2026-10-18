#!/usr/bin/env python3
"""
Create final_gst.xlsx from output_predictions1.csv without removing GST terms.

Usage:
  python create_final_gst_no_filter.py
  python create_final_gst_no_filter.py --input /path/to/output_predictions1.csv --output /path/to/final_gst.xlsx

By default, the script reads output_predictions1.csv from the same folder as this
script and writes final_gst.xlsx to that same folder.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
from uuid import uuid4

import pandas as pd


DEFAULT_COLUMNS_TO_DROP = [
    "predicted_labels",
    "GST_BERT",
    "I-GST_Start",
    "Removed_Terms",
]


def resolve_path(path_text: str, base_dir: Path) -> Path:
    """Resolve user-provided paths, keeping relative defaults local to this script."""
    path = Path(path_text).expanduser()
    if not path.is_absolute():
        path = base_dir / path
    return path


def create_final_gst(input_file: Path, output_file: Path) -> None:
    data = pd.read_csv(input_file)

    if "GST_BERT" not in data.columns:
        raise ValueError("Input CSV must contain a GST_BERT column.")

    # Preserve GST_BERT exactly as the BERT output column, except NaN values become blank.
    data["BERT_Output"] = data["GST_BERT"].fillna("").astype(str)

    data = data.drop(columns=DEFAULT_COLUMNS_TO_DROP, errors="ignore")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    temp_output = output_file.with_name(f".{output_file.stem}.{uuid4().hex}.tmp.xlsx")

    try:
        data.to_excel(temp_output, index=False)
        os.replace(temp_output, output_file)
    except PermissionError as exc:
        if temp_output.exists():
            message = (
                f"Permission denied while replacing {output_file}.\n"
                f"A complete temporary Excel file was written here instead:\n"
                f"  {temp_output}\n\n"
                "Close the existing Excel file if it is open, or fix ownership/permissions "
                "on the output file or directory."
            )
        else:
            message = (
                f"Permission denied while writing inside {output_file.parent}.\n"
                "Fix ownership/permissions on the output directory or choose a different "
                "--output path."
            )
        raise PermissionError(message) from exc


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(
        description="Create final_gst.xlsx without blacklist or alpha/beta removal heuristics."
    )
    parser.add_argument(
        "-i",
        "--input",
        default="output_predictions1.csv",
        help="Input CSV path. Default: output_predictions1.csv beside this script.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="final_gst.xlsx",
        help="Output XLSX path. Default: final_gst.xlsx beside this script.",
    )

    args = parser.parse_args()
    args.input = resolve_path(args.input, script_dir)
    args.output = resolve_path(args.output, script_dir)
    return args


def main() -> None:
    args = parse_args()
    create_final_gst(args.input, args.output)
    print(f"The data has been processed and saved successfully at {args.output}.")


if __name__ == "__main__":
    main()
