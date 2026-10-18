from glycaner_paths import STAGE3_INPUT, resource, work  # noqa: E402
import pandas as pd
import re

# =========================
# Input / Output
# =========================
# This script is intended to run after your conjunction / parenthesis cleanup step,
# where BERT_Output has already been renamed to Final_GST_String.
in_path = work("7_final_gst_string.xlsx")
out_path = work("8_final_gst_string_classified.xlsx")
source_sheet = "Sheet1"

df = pd.read_excel(in_path, sheet_name=source_sheet)

target_col = "Final_GST_String"
sent_col = "sent_text"
pmid_col = "pmid"
sent_index_col = "sent_index"

composition_col = "Composition"
sequence_col = "Sequence"
sugar_col = "Sugar"
gct_col = "GCT"
generic_col = "Generic"
ignored_col = "Ignored"
oxford_col = "Oxford_glycan_notation"

for c in [target_col, sent_col, pmid_col, sent_index_col]:
    if c not in df.columns:
        raise ValueError(f"Expected a '{c}' column in the Excel file.")

# =========================
# Shared helpers
# =========================

SPACE_RE = re.compile(r"\s+")


def normalize_text(s: str) -> str:
    s = "" if pd.isna(s) else str(s).strip()
    s = SPACE_RE.sub(" ", s)
    return s


def join_matches(x):
    if not x:
        return ""
    return " | ".join(x)


# =========================
# Composition detector
# =========================

COMP_TOKEN_ABBR = [
    "GlcNAc", "GalNAc", "ManNAc", "HexNAc",
    "Neu5Ac", "Neu5Gc", "NeuGc", "NeuAc", "NeuNAc",
    "Glc", "GlcA", "Gal", "GalA", "IdoA", "HexA",
    "GlcN", "GalN", "Fuc", "Man", "Hex"
]

COMP_TOKEN_FULL = [
    r"N-acetylglucosamine",
    r"N-acetylgalactosamine",
    r"N-acetylmannosamine",
    r"N-acetylhexosamine",
    r"N-acetylneuraminic acid",
    r"N-glycolylneuraminic acid",
    r"neuraminic acid",
    r"glucose",
    r"glucuronic acid",
    r"galactose",
    r"galacturonic acid",
    r"iduronic acid",
    r"hexuronic acid",
    r"glucosamine",
    r"galactosamine",
    r"fucose",
    r"mannose",
    r"hexose",
]

COMP_TOKEN_RE = r"(?:%s|%s)" % (
    "|".join(COMP_TOKEN_ABBR),
    "|".join(COMP_TOKEN_FULL),
)

COMPOSITION_RE = re.compile(
    rf"""(?ix)
    (?:
        (?:\b{COMP_TOKEN_RE}\b|\({COMP_TOKEN_RE}\))
        \s*
        (?:\(\d+\)|\d+)
    )
    (?:
        \s*
        (?:
            (?:\b{COMP_TOKEN_RE}\b|\({COMP_TOKEN_RE}\))
            \s*
            (?:\(\d+\)|\d+)
        )
    )+
    """
)


def find_composition_list(text: str):
    if not isinstance(text, str) or not text.strip():
        return []

    matches = []
    seen = set()

    for m in COMPOSITION_RE.finditer(text):
        val = re.sub(r"\s+", "", m.group(0))
        key = val.lower()
        if key not in seen:
            matches.append(val)
            seen.add(key)

    return matches


# =========================
# Sequence detector
# =========================

H1_RE = re.compile(
    r"""^
    (?!.*\s)
    (?=(?:[^-]*-){2,})
    .*
    (?:
        -
        (?:\d+(?:-\d+)?(?:')?)?
        (?:cers?|sers?|ceramides?)
      |
        d\s*(?:,|-)\s*d
        (?:cers?|sers?|ceramides?)
    )
    $
    """,
    re.IGNORECASE | re.VERBOSE
)

ARROW_RE = re.compile(r"-+>")
LINKAGE_RE = re.compile(r"\d+\s*(?:,|-+)\s*\d+")

GLYCAN_TERM_RE = re.compile(
    r"(GlcNAc|GalNAc|ManNAc|HexNAc|Neu5Ac|Neu5Gc|NeuGc|NeuAc|NeuNAc|Glc|GlcA|Gal|GalA|IdoA|HexA|GlcN|GalN|Fuc|Cer|Man|Hex|Ser|alpha|beta)",
    re.IGNORECASE
)

GLYCAN_SEQ_RE = re.compile(
    r"""(
        (GlcNAc|GalNAc|ManNAc|HexNAc|Neu5Ac|Neu5Gc|NeuGc|NeuAc|NeuNAc|Glc|GlcA|Gal|GalA|IdoA|HexA|GlcN|GalN|Fuc|Cer|Man|Hex|Ser|alpha|beta)
        .*?
        (\d+\s*(?:,|-+)\s*\d+|-+>)
      |
        (\d+\s*(?:,|-+)\s*\d+|-+>)
        .*?
        (GlcNAc|GalNAc|ManNAc|HexNAc|Neu5Ac|Neu5Gc|NeuGc|NeuAc|NeuNAc|Glc|GlcA|Gal|GalA|IdoA|HexA|GlcN|GalN|Fuc|Cer|Man|Hex|Ser|alpha|beta)
    )""",
    re.IGNORECASE | re.VERBOSE
)


def matches_heuristics(x: str) -> bool:
    if not isinstance(x, str) or not x.strip():
        return False

    s = x.strip()
    h1 = bool(H1_RE.search(s))
    h2 = bool(ARROW_RE.search(s))
    h3 = bool(GLYCAN_TERM_RE.search(s) and LINKAGE_RE.search(s))
    h4 = bool(GLYCAN_SEQ_RE.search(s))

    return h1 or h2 or h3 or h4


# =========================
# Oxford glycan notation from sent_text
# =========================

OXFORD_GLYCAN_RE = re.compile(
    r"""(?x)

    \bM[3-9]A[0-4]\b
    |
    \bG[0-2](?:F[0-6])(?:S[0-4])?\b
    |
    \bF?(?:[0-6])?A[1-4]
        (?=(?:\([0-6]\)|\[\d\]|(?:B|Bi)(?=[GSLF])|G[0-4]|L[0-6]|S|F))
        (?:\([0-6]\)|\[\d\])?
        (?:(?:B|Bi)(?=[GSLF]))?
        (?:G[0-4](?:\([0-6]\))?)?
        (?:L[0-6](?:\([0-6]\))?)?
        (?:S(?:[0-4](?:\([0-6]\))?)?)?
        (?:F(?:[0-6]|B)?)?
        (?=$|[^A-Za-z0-9])
    |
    \bN(?:G[0-4](?:\([0-6]\))?)?A[0-4](?:(?:B|Bi)(?=[GSLF]))?(?:F(?:[0-6]|B)?)?(?:S(?:[0-4](?:\([0-6]\))?)?)?(?=$|[^A-Za-z0-9])
    |
    \bA[0-4]A[0-4](?:F[0-6])?(?=$|[^A-Za-z0-9GSLFB(])
    |
    \bNa\d-\d\(AF\)(?:\(A\))?(?:\(F\))?F\d\b
    |
    \bNa\d-\d(?:\(A\))?(?:\(F\))?F\d\b

    """,
    flags=0
)

GLYC_WORD_RE = re.compile(r"\bglyc\w*\b", flags=re.IGNORECASE)
A_START_RE = re.compile(r"^A[1-4]\b")
WORD_RE = re.compile(r"\b\w+\b")


def has_nearby_glyc_word(text, start, end, window=3):
    words = list(WORD_RE.finditer(text))
    if not words:
        return False

    hit_idx = []
    for i, w in enumerate(words):
        if not (w.end() <= start or w.start() >= end):
            hit_idx.append(i)

    if not hit_idx:
        return False

    left = max(0, min(hit_idx) - window)
    right = min(len(words) - 1, max(hit_idx) + window)

    for i in range(left, right + 1):
        if GLYC_WORD_RE.fullmatch(words[i].group()):
            return True

    return False


def find_oxford_glycan_notation_list(text: str):
    if not isinstance(text, str) or not text.strip():
        return []

    matches = []
    seen = set()

    for m in OXFORD_GLYCAN_RE.finditer(text):
        val = m.group(0)

        if A_START_RE.match(val):
            if not has_nearby_glyc_word(text, m.start(), m.end(), window=3):
                continue

        key = val.lower()
        if key not in seen:
            matches.append(val)
            seen.add(key)

    return matches


# =========================
# Sugar detector
# =========================

SINGLE_OSE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9-]*ose$", re.IGNORECASE)
RIGHTMOST_OSE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9-]*ose$", re.IGNORECASE)
OSYL_RE = re.compile(r"^[A-Za-z0-9-]*osyl$", re.IGNORECASE)
SIALYL_RE = re.compile(r"^(sialyl|sialylated)$", re.IGNORECASE)

EXCLUDED_SACCHARIDE_RE = re.compile(
    r"""(?ix)
    \b(
        terminal |
        high[-\s]?mannose |
        core \s+ fucose |
        fucosylated |
        sialylated |
        glycosylated |
        aglycosylated |
        galactosylated |
        core \s+ fucosylated
    )\b
    """
)

UDP_GDP_START_RE = re.compile(r"^(?:UDP|GDP)\b", re.IGNORECASE)

RIGHT_OR_LEFT_MODIFIER_RE = re.compile(
    r"""(?ix)^
    (?:
        .*-(?:sulfo|sulfated|acetyl|acetylated|methyl|methylated|phosphoryl|phosphorylated)
        |
        (?:sulfo|sulfated|acetyl|acetylated|methyl|methylated|phosphoryl|phosphorylated|modified)
    )$
    """
)


def is_excluded_term(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    return bool(EXCLUDED_SACCHARIDE_RE.search(s))


def starts_with_udp_gdp(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    return bool(UDP_GDP_START_RE.match(s))


def matches_single_word_ose(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    words = s.split()
    if len(words) != 1:
        return False
    return bool(SINGLE_OSE_RE.fullmatch(words[0]))


def matches_single_word_osyl(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    words = s.split()
    if len(words) != 1:
        return False
    return bool(OSYL_RE.fullmatch(words[0]))


def matches_sialyl_word(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    words = s.split()
    if len(words) != 1:
        return False
    return bool(SIALYL_RE.fullmatch(words[0]))


MONO_AND_DISACCHARIDE_PATTERNS = [
    r"(L|D|O|N)?-?(linked)?[-\s]?(fucose|glucose|mannose|galactose|fructose|ribose|xylose|arabinose)",
    r"(gal|glc|gluco|man|fuc|xyl)",
    r"(glucuronic|iduronic|galacturonic|hexuronic|uronic)\sacid(s)?",
    r"(GlcA|IdoA|GalA|HexA)",
    r"(glucosamine|galactosamine|mannosamine)",
    r"(GlcN|GalN|ManN)",
    r"(L|D|O|N)?-?(linked)?[-\s]?(sucrose|lactose|maltose|cellobiose)",
    r"chitobiose",
    r"lac",
    r"LFuc",
]

MODIFIED_MONO_AND_DISACCHARIDE_PATTERNS = [
    r"[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?[Gg]lucosamine",
    r"[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?[Gg]alactosamine",
    r"[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?[Mm]annosamine",
    r"(GlcNAc|GlcNac|Glcnac)",
    r"(GalNAc|GalNac|Galnac)",
    r"(ManNAc|ManNac|Mannac)",
    r"Glc[-]?NAc",
    r"Gal[-]?NAc",
    r"Man[-]?NAc",
    r"[Ss]ialic\s[Aa]cid(s)?",
    r"neuraminic[-\s]acid",
    r"glycolylneuraminic[-\s]acid",
    r"acetylneuraminic[-\s]acid",
    r"N?-?acetylneuraminic\sacid(s)?",
    r"N?-?glycolylneuraminic\sacid(s)?",
    r"(Neu5Ac|Neu5Gc|NeuGc|NeuAc|NeuNAc)",
    r"S?-?Neu(5)?(A|G)c-?S?",
    r"S?-?Neu(5)?(A|G)c-?H?-?[12345678]?",
    r"GalNAc-?(β|beta)-?1[-,]4GlcNAc",
    r"[Nn]-[Aa]cetyl[-\s]?[Gg]alactosaminyl[-\s](β|beta)1[,-]4-[Nn]-[Aa]cetyl[-\s]?[Gg]lucosamine",
    r"(UDP|GDP)\S*",
    r"(Gal|Glc|Man)-?NAc-?H?-?[12345678]?",
]

OTHER_SACCHARIDE_PATTERNS = [
    r"polySia",
    r"(tetraose|heptaose|octaose|nonaose|decaose)",
    r"(triose|pentose|hexose)",
]


def compile_fullmatch_patterns(patterns):
    return [re.compile(pat, re.IGNORECASE) for pat in patterns]


MONO_REGEX = compile_fullmatch_patterns(MONO_AND_DISACCHARIDE_PATTERNS)
MOD_MONO_REGEX = compile_fullmatch_patterns(MODIFIED_MONO_AND_DISACCHARIDE_PATTERNS)
OTHER_SAC_REGEX = compile_fullmatch_patterns(OTHER_SACCHARIDE_PATTERNS)


def matches_any_full(s: str, regex_list) -> bool:
    if not isinstance(s, str):
        return False
    s = normalize_text(s)
    if not s:
        return False
    return any(rx.fullmatch(s) for rx in regex_list)


def contains_any_full(s: str, regex_list) -> bool:
    if not isinstance(s, str):
        return False
    s = normalize_text(s)
    if not s:
        return False
    return any(rx.fullmatch(s) for rx in regex_list)


def two_word_sugar_rule(x: str):
    if not isinstance(x, str):
        return ""

    s = normalize_text(x)
    if not s or is_excluded_term(s):
        return ""

    words = s.split()
    if len(words) < 2:
        return ""

    prev_word = words[-2]
    last_word = words[-1]

    if re.match(r"(?i)^core(?:-\S+)?$", prev_word):
        return ""

    last_is_mono = contains_any_full(last_word, MONO_REGEX)
    last_is_mod = contains_any_full(last_word, MOD_MONO_REGEX)

    if RIGHT_OR_LEFT_MODIFIER_RE.fullmatch(prev_word) and (last_is_mono or last_is_mod):
        return "sugar"

    if last_is_mono and not RIGHT_OR_LEFT_MODIFIER_RE.fullmatch(prev_word):
        return "sugar"

    return ""


def matches_monosaccharide(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s or is_excluded_term(s):
        return False

    if matches_any_full(s, MONO_REGEX):
        return True

    return two_word_sugar_rule(s) == "sugar"


def matches_modified_monosaccharide(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s or is_excluded_term(s):
        return False

    if matches_any_full(s, MOD_MONO_REGEX):
        return True

    return two_word_sugar_rule(s) == "sugar"


def matches_other_saccharide(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s or is_excluded_term(s):
        return False
    return matches_any_full(s, OTHER_SAC_REGEX)


def matches_unmarked_multiword_ose(x: str) -> bool:
    if not isinstance(x, str):
        return False

    s = normalize_text(x)
    if not s:
        return False

    words = s.split()
    if len(words) < 2:
        return False

    rightmost_word = words[-1]
    return bool(RIGHTMOST_OSE_RE.fullmatch(rightmost_word))


# =========================
# GCT detector
# =========================
# GCT is evaluated after Composition, Sequence, and Sugar detection.
# Uses search-style patterns so compound glycolipid names such as
# glucocerebroside, glucosylceramide, lactosylceramide, and globotriaosylceramide
# are captured without needing every possible prefix.

GCT_RE = re.compile(
    r"""(?ix)
    (?:
        # Globoside/oside terms
        \bgloboside\s(?:[1234]|I{1,3}|IV)\b
        |
        \b[A-Za-z0-9-]*(?:
            oside|osides|
            cerebroside|cerebrosides|
            cerbroside|cerbrosides|
            ganglioside|gangliosides|
            globoside|globosides
        )\b
        |
        \b(?:gangli|isoglob|lact|glyc|cerebr|fuc|mann|sial|hex|fruct|xyl|all|altr|gul|id|tal|arabin|rib|lyx|rhamn|digitox)osides?\b
        |
        \b(?:CTH|CMH|CDH|CTeH|CPH)\b
        |
        \bn?[Ll]c[-_\s]?\([345]\)\b
        |
        \bn?[Ll]c[-_\s]?[345]\b
        |
        \bG[Bb][-]?[123456789]\b
        |
        \b(?:DS|i)?G[Bb][-_]?[123456789]\b
        |
        \b(?:DS|i)?G[Bb]\([123456789]\)\b
        |
        \bG[Ll][-_]?[123456789]\b
        |
        \bG[Ll]\([123456789]\)\b
        |
        \bsulfatide\b
        |
        \bsulphatide\b

        # Ganglioside/glycolipid abbreviations
        |
        \bGM[-]?[123456789](?:a|b)?\b
        |
        \bGO[-]?[123456789](?:a|b)?\b
        |
        \bGP[-]?[123456789](?:a|b|c)?(?:a|b|c)?\b
        |
        \bGT[-]?[123456789](?:a|b|c)?\s?(?:a|b|c)?\s?(?:alpha)?\b
        |
        \bGQ[-]?[123456789](?:[Aa]|[Bb]|[Cc])?\s?(?:a|b|c)?\s?(?:alpha)?\b
        |
        \bGA[-]?[1234]\b
        |
        \bG[Dd][-]?[123456789][Bb]?-?L?\b
        |
        \bG[Dd][-]?[123456789](?:[ABCabc])?\b
        |
        \bG[Dd][-]?[123456789](?:alpha|Alpha)?\b
        |
        \bC[Dd]\d\d[Bb]?\b
        |
        \b(?:Lc)?G[Gg]O?\d\b
        |
        \bLD[-]?[12](?:a|b)?\b
        |
        \bLE[-]?[12](?:a|b)?\b
        |
        \bLQ[-]?1(?:a|b)?\b
        |
        \bTD[-]?[12](?:a|b)?\b
        |
        \bTE[-]?[12](?:a|b)?\b
        |
        \bTQ[-]?1(?:a|b)?\b
        |
        \bLM[-]?1(?:a|b)?\b

        # GCT headwords and compound endings
        |
        \b[A-Za-z0-9-]*(?:
            lipid|lipids|
            conjugate|conjugates|
            cer|cers|
            ceramide|ceramides|
            globoseries|
            glycosphingolipid|glycosphingolipids|
            glycopshingolipid|glycopshingolipids| 
            sphingolipid|sphingolipids|
            sphingoshine|sphingoshines
        )(?:-?\d+)?-?\b
    )
    """
)


def matches_gct(x: str) -> bool:
    if not isinstance(x, str):
        return False
    s = normalize_text(x)
    if not s:
        return False
    return bool(GCT_RE.search(s))


# =========================
# Generic detector
# =========================

GENERIC_TOKEN_RE = re.compile(
    r"""^(
        moiet(?:y|ies) |
        [A-Za-z0-9-]*saccharides? |
        cores? |
        carbohydrates? |
        structures? |
        epitopes? |
        determinants? |
        blood |
        groups? |
        globo[-\s]?series |
        ganglio[-\s]?series |
        lacto[-\s]?series |
        neolacto[-\s]?series |
        series |
        ganglio |
        globo |
        lacto |
        neolacto |
        N-glycans? |
        O-glycans? |
        neutral |
        antigens? |
        Ag |
        chains? |
        complex |
        sulfates?
    )$""",
    re.IGNORECASE | re.VERBOSE
)


def tokenize_generic_text(s: str):
    return re.findall(r"[A-Za-z0-9-]+", s)


def matches_generic(x: str) -> bool:
    if not isinstance(x, str) or not x.strip():
        return False

    s = x.strip()
    if re.search(r"[>|/]", s):
        return False

    s_norm = re.sub(r"[,:;()\[\]{}]+", " ", s)
    s_norm = re.sub(r"\s+", " ", s_norm).strip()

    tokens = tokenize_generic_text(s_norm)
    if not tokens:
        return False

    return all(GENERIC_TOKEN_RE.fullmatch(tok) for tok in tokens)


# =========================
# Apply detectors
# =========================

composition_hits = df[sent_col].apply(find_composition_list)
composition_mask = composition_hits.apply(lambda x: len(x) > 0)

sequence_mask = df[target_col].apply(matches_heuristics) & ~composition_mask
preclassified_mask = composition_mask | sequence_mask

excluded_mask = df[target_col].apply(is_excluded_term)

mono_mask = df[target_col].apply(matches_monosaccharide) & ~excluded_mask & ~preclassified_mask
mod_mono_mask = df[target_col].apply(matches_modified_monosaccharide) & ~excluded_mask & ~preclassified_mask
other_sac_mask = df[target_col].apply(matches_other_saccharide) & ~excluded_mask & ~preclassified_mask

single_ose_mask = df[target_col].apply(matches_single_word_ose)
single_osyl_mask = df[target_col].apply(matches_single_word_osyl)
single_sialyl_mask = df[target_col].apply(matches_sialyl_word)
udp_gdp_mask = df[target_col].apply(starts_with_udp_gdp)

sugar_mask = mono_mask | mod_mono_mask | other_sac_mask
unmarked_sugar_mask = ~(sugar_mask | preclassified_mask)

sugar_mask = sugar_mask | (single_ose_mask & unmarked_sugar_mask & ~excluded_mask)
unmarked_sugar_mask = ~(sugar_mask | preclassified_mask)

sugar_mask = sugar_mask | (single_sialyl_mask & unmarked_sugar_mask & ~excluded_mask)
unmarked_sugar_mask = ~(sugar_mask | preclassified_mask)

sugar_mask = sugar_mask | (udp_gdp_mask & ~excluded_mask & ~preclassified_mask)
unmarked_sugar_mask = ~(sugar_mask | preclassified_mask)

ignored_mask = single_osyl_mask & ~excluded_mask & ~(sugar_mask | preclassified_mask)
unmarked_sugar_mask = ~(sugar_mask | ignored_mask | preclassified_mask)

fallback_other_sac_mask = df[target_col].apply(matches_unmarked_multiword_ose) & unmarked_sugar_mask & ~excluded_mask
sugar_mask = sugar_mask | fallback_other_sac_mask

pre_gct_mask = composition_mask | sequence_mask | sugar_mask
gct_mask = df[target_col].apply(matches_gct) & ~pre_gct_mask

pre_generic_mask = pre_gct_mask | gct_mask | ignored_mask
generic_mask = df[target_col].apply(matches_generic) & ~pre_generic_mask

# Oxford notation from unique sent_text per pmid/sent_index
unique_sent = df[[pmid_col, sent_index_col, sent_col]].drop_duplicates(
    subset=[pmid_col, sent_index_col, sent_col]
).copy()

unique_sent[oxford_col + "_list"] = unique_sent[sent_col].apply(find_oxford_glycan_notation_list)
unique_sent[oxford_col] = unique_sent[oxford_col + "_list"].apply(join_matches)

df = df.merge(
    unique_sent[[pmid_col, sent_index_col, sent_col, oxford_col]],
    on=[pmid_col, sent_index_col, sent_col],
    how="left"
)

# =========================
# Add / update columns
# =========================

for col in [composition_col, sequence_col, sugar_col, gct_col, generic_col, ignored_col]:
    if col not in df.columns:
        df[col] = ""

df.loc[composition_mask, composition_col] = "Yes"
df.loc[sequence_mask, sequence_col] = "Yes"
df.loc[sugar_mask, sugar_col] = "Yes"
df.loc[gct_mask, gct_col] = "Yes"
df.loc[generic_mask, generic_col] = "Yes"
df.loc[ignored_mask, ignored_col] = "Yes"

# =========================
# Prepare per-type sheets
# =========================

upto_idx = df.columns.get_loc(target_col)
base_cols = list(df.columns[:upto_idx + 1])

composition_unique = df[[pmid_col, sent_index_col, sent_col]].copy()
composition_unique[composition_col + "_list"] = composition_hits
composition_rows = composition_unique.explode(composition_col + "_list").copy()
composition_rows = composition_rows[
    composition_rows[composition_col + "_list"].notna() &
    (composition_rows[composition_col + "_list"].astype(str).str.strip() != "")
][[pmid_col, sent_index_col, sent_col, composition_col + "_list"]].copy()
composition_rows = composition_rows.rename(columns={composition_col + "_list": composition_col})
composition_rows = composition_rows.drop_duplicates()

sequence_rows = df.loc[df[sequence_col] == "Yes", base_cols].copy()
sugar_rows = df.loc[df[sugar_col] == "Yes", base_cols].copy()
gct_rows = df.loc[df[gct_col] == "Yes", base_cols].copy()
generic_rows = df.loc[df[generic_col] == "Yes", base_cols].copy()
ignored_rows = df.loc[df[ignored_col] == "Yes", base_cols].copy()

oxford_rows = unique_sent.explode(oxford_col + "_list").copy()
oxford_rows = oxford_rows[
    oxford_rows[oxford_col + "_list"].notna() &
    (oxford_rows[oxford_col + "_list"].astype(str).str.strip() != "")
][[pmid_col, sent_index_col, sent_col, oxford_col + "_list"]].copy()
oxford_rows = oxford_rows.rename(columns={oxford_col + "_list": oxford_col})

# =========================
# Remove classified rows from main sheet
# =========================

classified_mask = (
    (df[composition_col] == "Yes") |
    (df[sequence_col] == "Yes") |
    (df[sugar_col] == "Yes") |
    (df[gct_col] == "Yes") |
    (df[generic_col] == "Yes") |
    (df[ignored_col] == "Yes")
)

df_remaining = df.loc[~classified_mask].copy()

cols_to_drop = [
    composition_col,
    sequence_col,
    sugar_col,
    gct_col,
    generic_col,
    ignored_col,
    oxford_col,
]

df_remaining = df_remaining.drop(columns=cols_to_drop, errors="ignore")

# =========================
# Add fully empty sent_index rows back to Sheet1
# =========================

sent_text_idx = df.columns.get_loc(sent_col)
after_sent_cols = list(df.columns[sent_text_idx + 1:])

empty_after_sent_mask = df[after_sent_cols].fillna("").astype(str).apply(
    lambda row: all(x.strip() == "" for x in row),
    axis=1
)

empty_sentence_rows = df.loc[
    empty_after_sent_mask,
    [pmid_col, sent_index_col, sent_col]
].drop_duplicates()

df_remaining = pd.concat(
    [df_remaining, empty_sentence_rows],
    ignore_index=True,
    sort=False
).drop_duplicates()

# =========================
# Write output
# =========================

with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
    df_remaining.to_excel(writer, sheet_name=source_sheet, index=False)

    composition_rows.to_excel(writer, sheet_name="Composition", index=False)
    sequence_rows.to_excel(writer, sheet_name="Sequence", index=False)
    sugar_rows.to_excel(writer, sheet_name="Sugar", index=False)
    gct_rows.to_excel(writer, sheet_name="GCT", index=False)
    generic_rows.to_excel(writer, sheet_name="Generic", index=False)
    ignored_rows.to_excel(writer, sheet_name="Ignored", index=False)
    oxford_rows.to_excel(writer, sheet_name="Oxford_glycan_notation", index=False)

print("✅ Done.")
print(f"Remaining rows in '{source_sheet}': {len(df_remaining)}")
print(f"Composition rows: {len(composition_rows)}")
print(f"Sequence rows: {len(sequence_rows)}")
print(f"Sugar rows: {len(sugar_rows)}")
print(f"GCT rows: {len(gct_rows)}")
print(f"Generic rows: {len(generic_rows)}")
print(f"Ignored rows: {len(ignored_rows)}")
print(f"Oxford glycan notation rows: {len(oxford_rows)}")
