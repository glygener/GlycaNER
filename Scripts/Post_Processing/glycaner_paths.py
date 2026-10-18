"""
Shared locations for the stage 3 post-processing scripts.

Every script in this folder reads and writes through these helpers, so the
pipeline can run anywhere by setting environment variables instead of editing
hardcoded paths.

  GLYCANER_STAGE3_INPUT   workbook produced by stage 2 (gst_punct_removed_cleaned.xlsx)
  GLYCANER_STAGE3_DIR     folder for every intermediate and final stage 3 workbook
  GLYCANER_RESOURCES      folder holding Dict_Normalization/, Mapping_File/, GSD_Entries/
"""

import os

_HERE = os.path.dirname(os.path.abspath(__file__))

RESOURCES_DIR = os.environ.get("GLYCANER_RESOURCES", os.path.join(_HERE, "resources"))
WORK_DIR = os.environ.get("GLYCANER_STAGE3_DIR", os.path.join(_HERE, "outputs"))
STAGE3_INPUT = os.environ.get(
    "GLYCANER_STAGE3_INPUT", os.path.join(WORK_DIR, "gst_punct_removed_cleaned.xlsx")
)

os.makedirs(WORK_DIR, exist_ok=True)


def work(name: str) -> str:
    """Path of an intermediate/output workbook inside the stage 3 work folder."""
    return os.path.join(WORK_DIR, name)


def resource(name: str) -> str:
    """Path of a bundled dictionary, mapping file or GSD workbook."""
    return os.path.join(RESOURCES_DIR, name)


# When a term matches several dictionaries, the first one read wins, so the read
# order changes results. os.listdir() order differs between filesystems (macOS
# APFS vs. Linux), so it is pinned here to the order the published results were
# produced with. Files not in this list are read afterwards in alphabetical order.
DICT_ORDER = [
    "type_no", "ceramide", "chem_moieties", "unknown", "gags", "globo",
    "named_antigens", "prefix", "lewis", "adj_sugar", "blood_group",
    "digit_modifier", "sugar", "antenna", "structural", "sialo_sulfo",
    "chemical", "adj_aminyl", "bi_di_tri", "sulfated", "named_gags",
    "headwords", "branch",
]


def dictionary_files(folder) -> list:
    """File names of the *.txt dictionaries in `folder`, in DICT_ORDER."""
    names = [f for f in os.listdir(folder) if f.endswith(".txt")]
    rank = {name: i for i, name in enumerate(DICT_ORDER)}
    return sorted(names, key=lambda f: (rank.get(f[:-4], len(rank)), f))
