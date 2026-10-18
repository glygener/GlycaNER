#!/usr/bin/env python3
from glycaner_paths import STAGE3_INPUT, dictionary_files, resource, work  # noqa: E402
"""
Resolve coordinated expressions in BERT_Output.

When a rule produces multiple resolved values, the source row is duplicated
once for each value. BERT_Output is replaced in each new row, and
Conj_Resolve records the rule that created it. Rows for which no rule applies
are retained once with their original BERT_Output.
"""

import re
from pathlib import Path

import pandas as pd


# -------------------------
# PATHS
# -------------------------
INPUT_XLSX = work("6a_percent_removed.xlsx")
REGEX_DIR = resource("Dict_Normalization")
OUTPUT_XLSX = work("6b_conjunction_resolved.xlsx")

COL = "BERT_Output"
# Resolved values replace BERT_Output in each duplicated output row.
NEWCOL = COL
OUTCOL = "Conj_Resolve"
DISCARD_COL = "Discard"
DIAG_COL = "Dict_Matches_Maximal"
PAREN_COL = "Parentheses_Resolve"

IGNORE_DICTS = {"digit_modifier"}
RULE3_EXCLUDE_TYPES = set()

CONJ_AND_OR_RE = re.compile(r"(?i)\b(?:and/or|and|or)\b")

LEFT_BREAK_TOKEN_RE = re.compile(
    r"^(?:"
    r"[A-Za-z]+(?:osyl|tyl|lyl)-?"
    r"|bi-?|tri-?|tetra-?|penta-?|mono-?|di-?"
    r"|oligo-?|poly-?"
    r"|[NO]-?"
    r"|\w+ed"
    r"|straight"
    r"|branched"
    r")$",
    flags=re.IGNORECASE,
)

RIGHT_BREAK_PREFIX_RE = re.compile(
    r"^(?:"
    r"[A-Za-z]+(?:osyl|tyl|lyl)"
    r"|bi|tri|tetra|penta|mono|di"
    r"|oligo|poly"
    r"|[NO]"
    r"|\w+ed"
    r"|straight"
    r"|branched"
    r")",
    flags=re.IGNORECASE,
)

CORE_DIGIT_TOKEN_RE = re.compile(
    r"\bcore\s+(?:\d+|[A-Za-z]\d+)\b",
    flags=re.IGNORECASE,
)

GM_LIKE_RE = re.compile(
    r"^[A-Za-z]{1,3}\d+[A-Za-z]?$",
    flags=re.IGNORECASE,
)


def load_regex_groups(regex_dir: str):
    groups = {}
    regex_dir = Path(regex_dir)

    for p in (regex_dir / name for name in dictionary_files(regex_dir)):
        if p.stem in IGNORE_DICTS:
            continue

        patterns = []
        for line in p.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                patterns.append(re.compile(line))
            except re.error:
                pass

        if patterns:
            groups[p.stem] = patterns

    return groups


regex_groups = load_regex_groups(REGEX_DIR)
if not regex_groups:
    raise RuntimeError(f"No regex patterns loaded. Check REGEX_DIR: {REGEX_DIR}")

AND_OR_RE = re.compile(r"\b(and|or)\b", flags=re.IGNORECASE)
CONJ_ANY_RE = re.compile(r"(?i)\b(and|or)\b")
WORDISH_RE = re.compile(r"[A-Za-z0-9'\-\(\)\[\]_/]+")

RESPECTIVELY_TAIL_RE = re.compile(
    r",\s+\b[A-Za-z]+ly\b\s*$",
    flags=re.IGNORECASE,
)

TYPE_NUM_TOKEN_RE = re.compile(
    r"\b(?:[ABOH]\s+)?type(?:\s*-\s*|\s+)\d+\b",
    flags=re.IGNORECASE,
)

TYPE_HEAD_RE = re.compile(
    r"^(?:(?P<abo>[ABOH])\s+)?type(?:\s*-\s*|\s+)(?P<num>\d+)$",
    flags=re.IGNORECASE,
)

TYPE_BARE_NUM_RE = re.compile(r"^\d+$")
OSIDE_WORD_RE = re.compile(r"[A-Za-z0-9'\-\(\)\[\]_/]*osides?\b", flags=re.IGNORECASE)
SINGLE_TOKEN_RE = re.compile(r"^[A-Za-z0-9'\-\(\)\[\]_/,]+$")

MULTIWORD_OSIDE_PHRASE_RE = re.compile(
    r"^[A-Za-z0-9'\-\(\)\[\]_/]+(?:\s+[A-Za-z0-9'\-\(\)\[\]_/]+)+\s+[A-Za-z0-9'\-\(\)\[\]_/]*osides?\b$",
    flags=re.IGNORECASE,
)

FORSSMAN_START_RE = re.compile(r"\bforssman\b", flags=re.IGNORECASE)
FORSSMAN_WORD_RE = re.compile(r"[A-Za-z0-9'\-\(\)\[\]_/]+")
MULTIWORD_FORSSMAN_PHRASE_RE = re.compile(
    r"^forssman(?:\s+[A-Za-z0-9'\-\(\)\[\]_/]+)+$",
    flags=re.IGNORECASE,
)

GLYCAN_LINK_RE = re.compile(
    r"\b(?:alpha|beta)\s+\d+\s*-+\s*\d+(?=\s|[A-Za-z(]|$)",
    flags=re.IGNORECASE,
)

PCT_TAIL_RE = re.compile(r"\s*\(\s*\d+(?:\.\d+)?\s*%\s*\)\s*$")
PCT_AFTER_RE = re.compile(r"\s*\(\s*\d+(?:\.\d+)?\s*%\s*\)\s*")

EXCL_HYPHEN_BEFORE_CONJ = re.compile(r"[-/]\s+(and|or)\b", flags=re.IGNORECASE)
EXCL_GLUED_HYPHEN_ANDOR = re.compile(r"\w+[-/]\s*(and|or)\b", flags=re.IGNORECASE)
Z_OK_RE = re.compile(r"(?i).*(tive|series)$")

INCOMPLETE_PATTERNS = [
    r"\b\w*osyl\b",
    r"\b\w*aminyl\b",
    r"\b\w*lyl\b",
    r"\b\w*sulfo\b",
    r"\btype[-\s]?[1-6](?:/[1-6])*\b",
    r"\banternnar(?:y|ies)\b",
    r"\bbi\b",
    r"\bdi\b",
    r"\bdim\b",
    r"\btri\b",
    r"\btria\b",
    r"\btetra\b",
    r"\bmono\b",
    r"\bpara\b",
    r"\bpenta\b",
    r"\bhexa\b",
    r"\bhepta\b",
    r"\bpoly\b",
    r"\boligo\b",
    r"\blyso\b",
    r"\bocta\b",
    r"\bnona\b",
    r"\bneo\b",
    r"\blacto\b",
    r"\bgluco\b",
    r"\bglobo\b",
    r"\bgalacto\b",
    r"\bganglio\b",
    r"\bcore[-\s]?[A-Za-z]?\d+\b",
    r"\bcore[-\s]?(?:[IVX]+)\b",
    r"\b\w+ed\b",
]

INCOMPLETE_END_RE = re.compile(
    r"(?:" + "|".join(f"(?:{p})" for p in INCOMPLETE_PATTERNS) + r")\s*$",
    flags=re.IGNORECASE,
)

COMPLETE_PATTERNS = [
    r"\b\w*osides?\b",
    r"\b\w*ose\b",
    r"\b\w*ceramides?\b",
    r"\b\w*lipids?\b",
    r"\b\w*tions?\b",
    r"\b\w*saccharides?\b",
    r"\b\w*cers?\b",
    r"\b\w*activit(?:y|ies)\b",
    r"\b\w*amine\b",
    r"\b\w*determinants?\b",
    r"\b\w*glycans?\b",
    r"\b\w*acids?\b",
    r"\b\w*forms?\b",
    r"\bblood groups?\b",
    r"\bCTH\b",
    r"\bCMH\b",
    r"\bCDH\b",
    r"\bn?[Ll]c[-_\s]?\([345]\)",
    r"\bn?[Ll]c[-_\s]?[345]",
    r"\bphosphate\b",
    r"\bsulfate\b",
    r"\bsulfatide\b",
    r"\bGM[-]?[123456789](a|b)?\b",
    r"\bGO[-]?[123456789](a|b)?\b",
    r"GM[-]?[123456789](a|b)?",
    r"\bGP[-]?[123456789](a|b|c)?(a|b|c)?\b",
    r"\bGT[-]?[123456789](a|b|c)?\s?(a|b|c)?\s?(alpha)?\b",
    r"\bGQ[-]?[123456789]([Aa]|[Bb]|[Cc])?\s?(a|b|c)?\s?(alpha)?\b",
    r"\bGA[-]?[1234]\b",
    r"\bG[Dd][-]?[123456789][Bb]?-?L?\b",
    r"G[Dd][-]?[123456789]([ABCabc])?",
    r"G[Dd][-]?[123456789](alpha|Alpha)?",
    r"\bG[Bb][-]?[123456789]\b",
    r"(DS|i)?G[Bb][-_]?[123456789]",
    r"(DS|i)?G[Bb]\([123456789]\)",
    r"G[Ll][-_]?[123456789]",
    r"G[Ll]\([123456789]\)",
    r"\bC[Dd]\d\d[Bb]?\b",
    r"(Lc)?G[Gg]O?\d",
    r"\bLD[-]?[12](a|b)?\b",
    r"\bLE[-]?[12](a|b)?\b",
    r"\bLQ[-]?1(a|b)?\b",
    r"\bTD[-]?[12](a|b)?\b",
    r"\bTE[-]?[12](a|b)?\b",
    r"\bTQ[-]?1(a|b)?\b",
    r"\bepitope(s)?\b",
    r"\b[Aa]ntigen(s)?\b",
    r"\brepeat(s)?\b",
    r"\bcarbohydrate(s)?\b",
    r"\bseries\b",
    r"\bhapten(s)?\b",
    r"\bstructure(s)?\b",
    r"\bchain(s)?\b",
    r"\bchain[-\s]?[ABOH]?\b",
    r"\bforssman(?:\s+\w+)*\b",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Le[\s^-]?\(?([ABXY])\)?",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Le[\s^-]?\(?([ABXY])\)?-([ABXY])?",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?-([ABXY])?",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?-antigen(s)?",
    r"\b(?:S|Sai|SD)?(lewis|le)\s?\(?([ABXY])-?(?:I{1,3}|IV)?\)?",
    r"\b(Lewis|Le)[\s^()-]*[ABXY]\b",
    r"\bSLX\b",
    r"\blewis[-\s]?(antigen|antigens)?\b",
    r"\b(GlcNAc|GlcNac|Glcnac)\b",
    r"\b[Ss]ialic\s[Aa]cid(s)?\b",
    r"\bneuraminic[-\s]acid\b",
    r"\bGlc[-]?NAc\b",
    r"\bGal[-]?NAc\b",
    r"\bMan[-]?NAc\b",
    r"\bS?-?Neu(5)?(A|G)c-?S?\b",
    r"\bS?-?Neu(5)?(A|G)c-?H?-?[12345678]?\b",
    r"\b[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?[Gg]lucosamine\b",
    r"\b[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?[Gg]alactosamin(e|yl)\b",
    r"\b(GalNAc|GalNac|Galnac)\b",
    r"\bHNK[-]?(1|2|3)?\b",
    r"\bHNK[-]?(1|2|3)?\s?[Aa]ntigen(s)?\b",
    r"\bSTN[-]?(1|2|3)?\s?[Aa]ntigen(s)?\b",
    r"\bCA19-9\b",
    r"\bCA-50\b",
    r"\b(LacNAc|LacNac|Lacnac)\b",
    r"\b(LacdiNAc|LacdiNac|Lacdinac)\b",
    r"\b(LN|LDN)\b",
    r"\b(lac|gal|glc)\b",
    r"(Gal|Glc|Lac|Man|Lacdi)-?NAc-?H?-?[12345678]?",
    r"\bCD57\b",
    r"\bcomplex[-\s](type)?(s)?\b",
    r"\bC1G\b",
    r"\b(N|O)[-\s]?linked\b",
    r"\bman[_]?[56]\b",
    r"\b(di|mono)meric\b",
    r"\bV3NeuAcIV6NeuAc[-]?\b",
    r"\bhigh[-\s]?mannose\b",
    r"\bhybrid\b",
    r"\b\(?SSEA[-\s]?[1234]\)?\b",
    r"\b\(?VIM[-\s]?[1234]\)?\b",
    r"\bsialic[-\s]?acid(s)?\b",
    r"\bsia\b",
    r"\bPSA\b",
]

COMPLETE_ANY_RE = re.compile(
    r"(?:" + "|".join(f"(?:{p})" for p in COMPLETE_PATTERNS) + r")",
    flags=re.IGNORECASE,
)


# Extra complete glycan feature targets that should be allowed to split at
# top-level commas, and/or.  These are intentionally independent from the
# external dictionary files so the conjunction resolver is stable even when a
# dictionary file does not contain one of these feature classes.
EXTRA_COMPLETE_FEATURE_PATTERNS = [
    # Lewis / Le / SSEA / CD antigen features
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Le[\s^-]?\(?([ABXY])\)?\b",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?\b",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Le[\s^-]?\(?([ABXY])\)?-([ABXY])?\b",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?-([ABXY])?\b",
    r"\b\(?((?:S|Sai|SD)-?)?\)?[-]?\(?((?:S|Sai|SD)-?)?\)?[-]?Lewis[\s^-]?\(?([ABXY])\)?-antigen(s)?\b",
    r"\b(?:S|Sai|SD)?(?:lewis|le)\s?\(?([ABXY])-?(?:I{1,3}|IV)?\)?\b",
    r"\b(?:Lewis|Le)[\s^()-]*[ABXY]\b",
    r"\bSLX\b",
    r"\blewis[-\s]?(?:antigen|antigens)?\b",
    r"\bCD15\b",
    r"\bCD174\b",
    r"\b\(?SSEA[-\s]?[1234]\)?\b",

    # Ganglio-series / glycolipid abbreviations
    r"\bGM[-]?[123456789](?:a|b)?\b",
    r"\bGO[-]?[123456789](?:a|b)?\b",
    r"\bGP[-]?[123456789](?:a|b|c)?(?:a|b|c)?\b",
    r"\bGT[-]?[123456789](?:a|b|c)?\s?(?:a|b|c)?\s?(?:alpha)?\b",
    r"\bGQ[-]?[123456789](?:[Aa]|[Bb]|[Cc])?\s?(?:a|b|c)?\s?(?:alpha)?\b",
    r"\bGA[-]?[1234]\b",
    r"\bG[Dd][-]?[123456789][Bb]?-?L?\b",
    r"\bG[Dd][-]?[123456789](?:[ABCabc])?\b",
    r"\bG[Dd][-]?[123456789](?:alpha|Alpha)?\b",
    r"\bC[Dd]\d\d[Bb]?\b",
    r"\b(?:Lc)?G[Gg]O?\d\b",
    r"\bLD[-]?[12](?:a|b)?\b",
    r"\bLE[-]?[12](?:a|b)?\b",
    r"\bLQ[-]?1(?:a|b)?\b",
    r"\bTD[-]?[12](?:a|b)?\b",
    r"\bTE[-]?[12](?:a|b)?\b",
    r"\bTQ[-]?1(?:a|b)?\b",
    r"\bLM[-]?1(?:a|b)?\b",

    # Globoside / oside / sulfatide classes
    r"\bgloboside\s(?:[1234]|I{1,3}|IV)\b",
    r"\boside(s)?\b",
    r"\b(?:gangli|isoglob|lact|glyc|cerebr|fuc|mann|sial|hex|fruct|xyl|all|altr|gul|id|tal|arabin|rib|lyx|rhamn|digitox)oside(s)?\b",
    r"\bCTH\b",
    r"\bCMH\b",
    r"\bCDH\b",
    r"\bCTeH\b",
    r"\bCPH\b",
    r"\bn?[Ll]c[-_\s]?\([345]\)\b",
    r"\bn?[Ll]c[-_\s]?[345]\b",
    r"\bG[Bb][-]?[123456789]\b",
    r"\b(?:DS|i)?G[Bb][-_]?[123456789]\b",
    r"\b(?:DS|i)?G[Bb]\([123456789]\)\b",
    r"\bG[Ll][-_]?[123456789]\b",
    r"\bG[Ll]\([123456789]\)\b",
    r"\bsulfatide\b",
    r"\bsulphatide\b",

    # Headwords that are complete features before top-level coordination
    r"\bcerebroside(s)?\b",
    r"\bganglioside(s)?\b",
    r"\bcer(s)?\b",
    r"\bceramide(s)?(?:-?\d+)?-?\b",

    # HexNAc / sialic-acid / nucleotide-sugar style features
    r"\b(?:glucosamine|galactosamine|mannosamine)\b",
    r"\b[Nn]-[Aa]cetyl[-\s]?[Dd]?[-\s]?(?:[Gg]lucosamine|[Gg]alactosamine|[Mm]annosamine|[Ff]ucosamine|[Ll]actosamine|[Gg]lycosamine|[Tt]etraosamine|[Hh]exosamine|[Hh]exaosamine|[Pp]entaosamine|[Hh]eptaosamin|[Hh]eptosamine|[Oo]ctaosamine)\b",
    r"\b[Nn]-[Aa]cetyl[-\s]?[Gg]alactosaminyl[-\\s](?:β|beta)1[,-]4-[Nn]-[Aa]cetyl[-\s]?[Gg]lucosamine\b",
    r"\b(?:LacNAc|LacNac|Lacnac)\b",
    r"\b(?:LacdiNAc|LacdiNac|Lacdinac)\b",
    r"\b(?:GalNAc|GalNac|Galnac)\b",
    r"\b(?:GlcNAc|GlcNac|Glcnac)\b",
    r"\b(?:ManNAc|ManNac|Mannac)\b",
    r"\b(?:HexNAc|HexNac|Hexnac)\b",
    r"\b(?:Gal|Glc|Hex|Lac|Man|Lacdi)-?NAc-?H?-?[12345678]?\b",
    r"\bLDN[-\s]?F\b",
    r"\b(?:Glc|Gal|Man|Lac|Lacdi|Hex)[-]?NAc\b",
    r"\b(?:LN|LDN|LDN-F)\b",
    r"\b[Nn]-[Aa]cetyl[-\s]?[Ll]actosamine\b",
    r"\bGalNAc-?(?:β|beta)-?1[-,]4GlcNAc\b",
    r"\b(?:UDP|GDP)\S*\b",
    r"\b(?:tetra|tri|tria|penta|pent|hex|hexa|hept|hepta|octa|nona|deca)ose\b",
    r"\b(?:UDP|GDP)\S*(?:fucose|glucose|mannose|galactose)\b",
    r"\bN?-?acetylneuraminic\sacid(s)?\b",
    r"\bN?-?glycolylneuraminic\sacid(s)?\b",
    r"\bS?-?Neu(5)?(?:A|G)c-?S?\b",
    r"\bS?-?Neu(5)?(?:A|G)c-?H?-?[12345678]?\b",
    r"\bsialic[-\s]acid(s)?\b",
    r"\bneuraminic[-\s]acid(s)?\b",
    r"\b(?:tri|tetra|penta|hexa|hepta|octa|poly)?[Ss]ia\b",
    r"\bV3NeuAcIV6NeuAc[-]?\b",
    r"\bLN-?[12345678]\b",
    r"\bLDN-?[12]\b",
    r"\bLDN-?F[12]\b",
    r"\bPSA\b",
    r"\b(?:keto)?-?(?:deoxy)?-?nonulosonic\sacid(s)?\b",
]

EXTRA_COMPLETE_FEATURE_RE = re.compile(
    r"(?:" + "|".join(f"(?:{p})" for p in EXTRA_COMPLETE_FEATURE_PATTERNS) + r")",
    flags=re.IGNORECASE,
)

COMPLETE_OR_EXTRA_ANY_RE = re.compile(
    r"(?:" + "|".join(f"(?:{p})" for p in (COMPLETE_PATTERNS + EXTRA_COMPLETE_FEATURE_PATTERNS)) + r")",
    flags=re.IGNORECASE,
)

CORE_UNIT_RE = re.compile(
    r"""(?ix)
    ^
    (?:(?P<mod>extended|internal|external)\s+)?
    (?:(?P<core>core)\s*-?\s*)?
    (?P<id>\d+|[A-Za-z]\d+)
    $
    """
)

CORE_WORD_UNIT_RE = re.compile(
    r"""(?ix)
    (?P<mod>\b(?:extended|internal|external)\s+)?
    (?P<core>\bcore\s*-?\s*)?
    (?P<id>\b(?:\d+|[A-Za-z]\d+)\b)
    """
)

TYPE_OR_ABO_FEATURE_RE = re.compile(
    r"""(?ix)
    ^(?:
        (?:[ABOH]\s+)?type\s*-?\s*\d+
        | Le\s*\(?[ABXY]\)?
        | Lewis\s*\(?[ABXY]\)?
        | [ABOH]
    )$
    """
)


def has_complete_or_extra_feature(s: str) -> bool:
    return COMPLETE_OR_EXTRA_ANY_RE.search((s or "").strip()) is not None


def try_complete_feature_list_rule(text: str):
    """
    Split coordinated complete features before top-level commas/and/or.

    Examples:
      Lea and H type 2 -> Lea | H type 2
      H type 1, H type 2, Le(a), Le(b), A, and B
        -> H type 1 | H type 2 | Le(a) | Le(b) | A | B
      GM1, GM2 and GD1a -> GM1 | GM2 | GD1a
    """
    if not has_top_level_coordination(text):
        return None

    segs = [normalize_spaces(s) for s in split_top_level_segments(text)]
    segs = [s for s in segs if s]
    if len(segs) <= 1:
        return None

    out = []
    for seg in segs:
        base = strip_pct_tail_only(seg)
        if (
            has_complete_or_extra_feature(base)
            or TYPE_OR_ABO_FEATURE_RE.fullmatch(base)
            or CORE_DIGIT_TOKEN_RE.fullmatch(base)
        ):
            out.append(seg)
        else:
            return None

    uniq = []
    seen = set()
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def _parse_core_segment(seg: str, require_core_word: bool = False):
    """Return prefix, normalized core token, and suffix for one coordinated core segment."""
    seg = seg or ""
    matches = []
    for m in CORE_WORD_UNIT_RE.finditer(seg):
        raw = m.group(0).strip()
        parsed = CORE_UNIT_RE.fullmatch(raw)
        if not parsed:
            continue
        if require_core_word and not parsed.group("core"):
            continue
        matches.append((m, parsed))

    if not matches:
        return None

    m, parsed = matches[-1] if require_core_word else matches[0]
    prefix = seg[:m.start()]
    suffix = seg[m.end():]
    mod = normalize_spaces(parsed.group("mod") or "")
    has_core = bool(parsed.group("core"))
    cid = parsed.group("id")
    token = normalize_spaces(" ".join(p for p in [mod, "core" if has_core else "", cid] if p))

    return {
        "prefix": normalize_spaces(prefix),
        "suffix": normalize_spaces(suffix),
        "token": token,
        "has_core": has_core,
        "id": cid,
        "mod": mod,
    }


def try_core_series_rule(text: str):
    """
    Resolve core lists while preserving useful left/right context.

    Examples:
      core 2 and core 4 -> core 2 | core 4
      O-glycan core 2, core 3 and extended core 1
        -> O-glycan core 2 | O-glycan core 3 | O-glycan extended core 1
      core M1, M2, and M3 glycans
        -> core M1 glycans | core M2 glycans | core M3 glycans
    """
    if not has_top_level_coordination(text):
        return None

    segs = [s for s in split_top_level_segments(text) if normalize_spaces(s)]
    if len(segs) <= 1:
        return None

    first = _parse_core_segment(segs[0], require_core_word=True)
    if first is None:
        return None

    parsed = [first]
    for seg in segs[1:]:
        item = _parse_core_segment(seg, require_core_word=False)
        if item is None:
            return None
        parsed.append(item)

    if not all(item["has_core"] or re.fullmatch(r"[A-Za-z]?\d+", item["id"]) for item in parsed):
        return None

    left_context = first["prefix"]
    right_context = parsed[-1]["suffix"]

    out = []
    for item in parsed:
        core_word = "core" if item["has_core"] or first["has_core"] else ""
        core_token = normalize_spaces(" ".join(p for p in [item["mod"], core_word, item["id"]] if p))
        out.append(normalize_spaces(" ".join(p for p in [left_context, core_token, right_context] if p)))

    uniq = []
    seen = set()
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def strip_pct_tail_only(s: str) -> str:
    return PCT_TAIL_RE.sub("", (s or "").strip())


def normalize_spaces(x: str) -> str:
    return re.sub(r"\s+", " ", x or "").strip()


def is_emptyish(x: str) -> bool:
    return re.fullmatch(r"[\s,;:]*", x or "") is not None


def is_dangling_hyphen_token(tok_base: str) -> bool:
    t = (tok_base or "").strip()
    return bool(t) and (t.startswith("-") or t.endswith("-"))


def remove_comma_ly_tail(text: str) -> str:
    text = "" if text is None else str(text)
    return RESPECTIVELY_TAIL_RE.sub("", text).strip()


def parse_type_head_token(tok: str):
    tok = (tok or "").strip()
    m = TYPE_HEAD_RE.fullmatch(tok)
    if not m:
        return None

    abo = (m.group("abo") or "").strip()
    num = m.group("num")
    canon = f"{abo} type {num}" if abo else f"type {num}"
    return abo, num, canon


def is_bare_type_number(tok: str) -> bool:
    return TYPE_BARE_NUM_RE.fullmatch((tok or "").strip()) is not None


def build_type_token(num: str, abo: str = "") -> str:
    num = str(num).strip()
    abo = (abo or "").strip()
    return f"{abo} type {num}" if abo else f"type {num}"


class SpanToken:
    def __init__(self, text: str, start: int, end: int):
        self._text = text
        self._start = start
        self._end = end

    def start(self):
        return self._start

    def end(self):
        return self._end

    def group(self, idx=0):
        return self._text[self._start:self._end]


def find_oside_phrase_spans(text: str):
    if not text:
        return []

    spans = []
    n = len(text)

    for m in OSIDE_WORD_RE.finditer(text):
        oside_start = m.start()
        oside_end = m.end()

        j = oside_end
        while j < n and text[j].isspace():
            j += 1

        followed_by_coord = False
        if j < n and text[j] == ",":
            followed_by_coord = True
        else:
            m_conj = AND_OR_RE.match(text, j)
            if m_conj:
                followed_by_coord = True

        if not followed_by_coord:
            continue

        left_comma = text.rfind(",", 0, oside_start)
        start = 0 if left_comma == -1 else left_comma + 1
        end = oside_end

        while start < end and text[start].isspace():
            start += 1
        while end > start and text[end - 1].isspace():
            end -= 1

        if start < end:
            spans.append((start, end))

    spans = sorted(set(spans), key=lambda x: (x[0], x[1]))
    kept = []
    for a, b in spans:
        if not any(A <= a and b <= B for A, B in kept):
            kept.append((a, b))

    return kept


def find_forssman_phrase_spans(text: str):
    if not text:
        return []

    spans = []
    n = len(text)

    for m in FORSSMAN_START_RE.finditer(text):
        start = m.start()
        i = m.end()
        word_count_after = 0
        end = m.end()

        while i < n:
            j = i
            while j < n and text[j].isspace():
                j += 1

            if j >= n:
                break
            if text[j] in ",/":
                break
            if AND_OR_RE.match(text, j):
                break

            w = FORSSMAN_WORD_RE.match(text, j)
            if not w:
                break

            end = w.end()
            word_count_after += 1
            i = w.end()

        if word_count_after >= 1:
            spans.append((start, end))

    spans = sorted(set(spans), key=lambda x: (x[0], x[1]))
    kept = []
    for a, b in spans:
        if not any(A <= a and b <= B for A, B in kept):
            kept.append((a, b))

    return kept


def iter_wordish_tokens(text: str):
    if not text:
        return []

    special = []
    special.extend(list(CORE_DIGIT_TOKEN_RE.finditer(text)))
    special.extend(list(TYPE_NUM_TOKEN_RE.finditer(text)))

    for a, b in find_oside_phrase_spans(text):
        special.append(SpanToken(text, a, b))

    for a, b in find_forssman_phrase_spans(text):
        special.append(SpanToken(text, a, b))

    taken = [(m.start(), m.end()) for m in special]

    def overlaps(a, b):
        return any(not (b <= x or a >= y) for x, y in taken)

    normal = []
    for m in WORDISH_RE.finditer(text):
        if not overlaps(m.start(), m.end()):
            normal.append(m)

    all_matches = special + normal
    all_matches.sort(key=lambda m: (m.start(), m.end()))
    return all_matches


def is_single_token(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return False

    if SINGLE_TOKEN_RE.fullmatch(s) is not None:
        return True
    if TYPE_NUM_TOKEN_RE.fullmatch(s) is not None:
        return True
    if MULTIWORD_OSIDE_PHRASE_RE.fullmatch(s):
        return True
    if MULTIWORD_FORSSMAN_PHRASE_RE.fullmatch(s):
        return True

    oside_spans = find_oside_phrase_spans(s)
    if len(oside_spans) == 1 and oside_spans[0] == (0, len(s)):
        return True

    forssman_spans = find_forssman_phrase_spans(s)
    if len(forssman_spans) == 1 and forssman_spans[0] == (0, len(s)):
        return True

    return False


def has_glycan_linkage_phrase(s: str) -> bool:
    return GLYCAN_LINK_RE.search((s or "").strip()) is not None


def is_rule1a_valid_segment(s: str) -> bool:
    s = (s or "").strip()
    if not s:
        return False
    if is_dangling_hyphen_token(s):
        return False
    if is_single_token(s):
        return True
    if has_glycan_linkage_phrase(s):
        return True
    return False


def has_top_level_coordination(text: str) -> bool:
    if not text:
        return False

    depth = 0
    i = 0
    n = len(text)

    while i < n:
        ch = text[i]

        if ch == "(":
            depth += 1
            i += 1
            continue

        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue

        if depth == 0:
            if ch == "/":
                return True
            if ch == "," and i + 1 < n and text[i + 1].isspace():
                return True
            if AND_OR_RE.match(text, i):
                return True

        i += 1

    return False


def split_top_level_segments(text: str):
    if not text:
        return []

    segs = []
    depth = 0
    start = 0
    i = 0
    n = len(text)

    def emit(a, b):
        raw = text[a:b]
        if raw.strip():
            segs.append(raw)

    while i < n:
        ch = text[i]

        if ch == "(":
            depth += 1
            i += 1
            continue

        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue

        if depth == 0 and ch == "/":
            emit(start, i)
            i += 1
            start = i
            continue

        if depth == 0 and ch == "," and i + 1 < n and text[i + 1].isspace():
            emit(start, i)
            i += 1
            start = i
            continue

        if depth == 0:
            m = AND_OR_RE.match(text, i)
            if m:
                emit(start, i)
                i = m.end()
                start = i
                continue

        i += 1

    emit(start, n)
    return segs


def exclude_rule2(text: str) -> bool:
    if not text:
        return True
    if EXCL_HYPHEN_BEFORE_CONJ.search(text):
        return True
    if EXCL_GLUED_HYPHEN_ANDOR.search(text):
        return True
    return False


def fullmatch_any(regex_list, s: str) -> bool:
    return any(rx.fullmatch(s) for rx in regex_list)


def types_covering_all_words(words):
    matched = []
    for type_name, regexes in regex_groups.items():
        ok = True
        for w in words:
            if not fullmatch_any(regexes, w):
                ok = False
                break
        if ok:
            matched.append(type_name)
    return matched


def last_wordish(seg: str):
    ms = iter_wordish_tokens(seg)
    if not ms:
        return None
    m = ms[-1]
    return seg[:m.start()], seg[m.start():m.end()], seg[m.end():]


def first_wordish(seg: str):
    ms = iter_wordish_tokens(seg)
    if not ms:
        return None
    m = ms[0]
    return seg[:m.start()], seg[m.start():m.end()], seg[m.end():]


def find_top_level_conj_span(text: str):
    depth = 0
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == "(":
            depth += 1
            i += 1
            continue

        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue

        if depth == 0:
            if text[i] == "/":
                return i, i + 1
            m = CONJ_ANY_RE.match(text, i)
            if m:
                return m.start(), m.end()

        i += 1

    return None


def dict_suffix_at_end(token: str):
    best = None

    for type_name, regexes in regex_groups.items():
        if type_name in RULE3_EXCLUDE_TYPES:
            continue

        for rx in regexes:
            for m in rx.finditer(token):
                if m.end() == len(token) and m.start() < m.end():
                    matched_text = token[m.start():m.end()]
                    length = m.end() - m.start()
                    cand = (length, -m.start(), type_name, matched_text)
                    if best is None or cand > best:
                        best = cand

    if best is None:
        return None

    _, _, type_name, matched_text = best
    return matched_text, type_name


def token_ends_with_dict_suffix(token: str) -> bool:
    return dict_suffix_at_end(token) is not None


def join_with_dict_suffix(left_tok: str, z: str) -> str:
    if left_tok.endswith("-"):
        return left_tok + z
    return left_tok + "-" + z


def decompose_suffix_chain(token: str, max_parts: int = 10):
    cur = token
    rev_parts = []

    for _ in range(max_parts):
        info = dict_suffix_at_end(cur)
        if not info:
            break

        z, z_type = info
        if len(cur) <= len(z):
            break

        base = cur[:-len(z)]
        if not base:
            break

        rev_parts.append((z, z_type))
        cur = base

    rev_parts.reverse()
    return cur, rev_parts


def segment_core_candidate(seg: str, position: str):
    seg = seg or ""

    if position == "first":
        lw = last_wordish(seg)
        if lw is None:
            return None

        before, tok, after = lw
        if not is_emptyish(after):
            return None

        return {"prefix": before, "token": tok.strip(), "suffix": after}

    if position == "last":
        fw = first_wordish(seg)
        if fw is None:
            return None

        before, tok, after = fw
        if not is_emptyish(before):
            return None

        return {"prefix": before, "token": tok.strip(), "suffix": after}

    seg_stripped = seg.strip()
    if not seg_stripped:
        return None

    if is_single_token(seg_stripped):
        return {"prefix": "", "token": seg_stripped, "suffix": ""}

    fw = first_wordish(seg)
    lw = last_wordish(seg)
    if fw is None or lw is None:
        return None

    b1, t1, _ = fw
    _, t2, a2 = lw

    if t1 != t2:
        return None
    if not is_emptyish(b1):
        return None
    if not is_emptyish(a2):
        return None

    return {"prefix": b1, "token": t1.strip(), "suffix": a2}


def should_keep_original_segments_due_to_complete_incomplete_guard(text: str):
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    segs_stripped = [normalize_spaces(s) for s in segs if normalize_spaces(s)]
    if len(segs_stripped) <= 1:
        return None

    def _segment_is_complete(seg_text: str) -> bool:
        seg_text = strip_pct_tail_only(normalize_spaces(seg_text))
        return bool(seg_text) and COMPLETE_ANY_RE.search(seg_text) is not None

    def _segment_has_leading_incomplete(seg_text: str) -> bool:
        seg_text = strip_pct_tail_only(normalize_spaces(seg_text))
        if not seg_text:
            return False

        fw = first_wordish(seg_text)
        if fw is None:
            return False

        first_tok = fw[1].strip()
        return INCOMPLETE_END_RE.search(first_tok) is not None

    all_previous_complete = _segment_is_complete(segs_stripped[0])

    for seg_text in segs_stripped[1:]:
        if all_previous_complete and _segment_has_leading_incomplete(seg_text):
            return segs_stripped
        all_previous_complete = all_previous_complete and _segment_is_complete(seg_text)

    return None


def find_top_level_conj_span_special(text: str):
    depth = 0
    i = 0

    while i < len(text):
        ch = text[i]

        if ch == "(":
            depth += 1
            i += 1
            continue

        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue

        if depth == 0:
            m = CONJ_AND_OR_RE.match(text, i)
            if m:
                return m.start(), m.end()

        i += 1

    return None


MODIFIER_END_RE = re.compile(
    r"^[A-Za-z0-9'\-\(\)\[\]_/]+(?:ated|osyl|tyl|lyl)-?$",
    flags=re.IGNORECASE,
)

HEADWORD_RE = re.compile(
    r"""(?ix)
    (?<![A-Za-z0-9])
    (
        -?alditols? |
        determinants? |
        cores? |
        chains? |
        epitopes? |
        antigens? |
        saccharides? |
        (?:N|O)?-?glycans? |
        carbohydrates? |
        cerebrosides? |
        gangliosides? |
        cers? |
        ceramides?(?:-?\d+)?-? |
        globosides? |
        globoseries |
        glycolipids? |
        glycosphingolipids? |
        sphingolipids? |
        sphingoshines? |
        types? |
        haptens? |
        structures? |
        Ag |
        derivatives?
    )
    (?![A-Za-z0-9])
    """,
)

PREFIX_INCOMPLETE_RE = re.compile(
    r"^(?:mono|bi|di|tri|tetra|penta|hexa|hepta|octa|nona|deca|oligo|poly)-?$",
    flags=re.IGNORECASE,
)

RIGHT_PREFIXED_INCOMPLETE_RE = re.compile(
    r"^(mono|bi|di|tri|tetra|penta|hexa|hepta|octa|nona|deca|oligo|poly)(.+)$",
    flags=re.IGNORECASE,
)

LETTER_TOKEN_RE = re.compile(r"^[A-Za-z]$")


def is_modifier_end_token(tok: str) -> bool:
    tok = (tok or "").strip()
    return bool(tok) and MODIFIER_END_RE.fullmatch(tok) is not None


def suffix_through_headword(text: str) -> str:
    text = text or ""
    m = HEADWORD_RE.search(text)
    if not m:
        return ""
    return normalize_spaces(text[:m.end()])


def try_before_other_rule1(text: str):
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    first = segs[0]
    last = segs[-1]
    mids = segs[1:-1]

    fw = last_wordish(first)
    if fw is None:
        return None

    first_before, first_tok, first_after = fw
    if not is_emptyish(first_after):
        return None

    first_tok = first_tok.strip()
    if not is_modifier_end_token(first_tok):
        return None

    modifier_tokens = [first_tok]

    for mid in mids:
        mid_core = segment_core_candidate(mid, "middle")
        if mid_core is None:
            return None

        mid_tok = mid_core["token"].strip()
        if not is_modifier_end_token(mid_tok):
            return None

        modifier_tokens.append(mid_tok)

    lw = first_wordish(last)
    if lw is None:
        return None

    last_before, last_tok, last_after = lw
    if not is_emptyish(last_before):
        return None

    last_tok = last_tok.strip()
    if not is_modifier_end_token(last_tok):
        return None

    modifier_tokens.append(last_tok)

    shared_suffix = suffix_through_headword(last_after)
    if not shared_suffix:
        return None

    prefix_text = normalize_spaces(first_before)

    out = []
    seen = set()

    for tok in modifier_tokens:
        val = normalize_spaces(
            " ".join([p for p in [prefix_text, tok, shared_suffix] if p])
        )
        if val and val not in seen:
            seen.add(val)
            out.append(val)

    return out if len(out) >= 2 else None


def is_prefix_incomplete_token(tok: str) -> bool:
    tok = (tok or "").strip()
    return bool(tok) and PREFIX_INCOMPLETE_RE.fullmatch(tok) is not None


def split_right_prefixed_incomplete(tok: str):
    tok = (tok or "").strip()
    m = RIGHT_PREFIXED_INCOMPLETE_RE.fullmatch(tok)
    if not m:
        return None

    prefix = m.group(1)
    suffix = m.group(2)

    if not suffix:
        return None

    return prefix, suffix


def complete_prefix_with_suffix(prefix_tok: str, suffix: str) -> str:
    prefix_tok = (prefix_tok or "").strip()
    suffix = (suffix or "").strip()

    if prefix_tok.endswith("-"):
        return prefix_tok + suffix

    return prefix_tok + "-" + suffix


def try_prefix_incomplete_rule(text: str):
    """
    prefix_incomplete_rule

    Examples:
      mono- or difucosyl type 2 chain A
      -> mono-fucosyl type 2 chain A
      -> difucosyl type 2 chain A

      mono-, di-, and trifucosyl chains
      -> mono-fucosyl chains
      -> di-fucosyl chains
      -> trifucosyl chains
    """
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    first = segs[0]
    last = segs[-1]
    mids = segs[1:-1]

    fw = last_wordish(first)
    if fw is None:
        return None

    first_before, first_tok, first_after = fw
    if not is_emptyish(first_after):
        return None

    first_tok = first_tok.strip()
    if not is_prefix_incomplete_token(first_tok):
        return None

    prefix_tokens = [first_tok]

    for mid in mids:
        mid_core = segment_core_candidate(mid, "middle")
        if mid_core is None:
            return None

        mid_tok = mid_core["token"].strip()
        if not is_prefix_incomplete_token(mid_tok):
            return None

        prefix_tokens.append(mid_tok)

    lw = first_wordish(last)
    if lw is None:
        return None

    last_before, last_tok, last_after = lw
    if not is_emptyish(last_before):
        return None

    last_tok = last_tok.strip()
    right_parts = split_right_prefixed_incomplete(last_tok)
    if right_parts is None:
        return None

    _, shared_suffix = right_parts

    prefix_text = normalize_spaces(first_before)
    suffix_text = normalize_spaces(last_after)

    out = []
    seen = set()

    for prefix_tok in prefix_tokens:
        completed = complete_prefix_with_suffix(prefix_tok, shared_suffix)
        val = normalize_spaces(
            " ".join([p for p in [prefix_text, completed, suffix_text] if p])
        )
        if val and val not in seen:
            seen.add(val)
            out.append(val)

    last_val = normalize_spaces(
        " ".join([p for p in [prefix_text, last_tok, suffix_text] if p])
    )
    if last_val and last_val not in seen:
        seen.add(last_val)
        out.append(last_val)

    return out if len(out) >= 2 else None


def is_letter_token(tok: str) -> bool:
    tok = (tok or "").strip()
    return bool(tok) and LETTER_TOKEN_RE.fullmatch(tok) is not None


def try_letter_conj_rule(text: str):
    """
    letter_conj_rule

    Examples:
      histo-blood group B and O individuals
      -> histo-blood group B individuals
      -> histo-blood group O individuals
    """
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    first = segs[0]
    last = segs[-1]
    mids = segs[1:-1]

    fw = last_wordish(first)
    if fw is None:
        return None

    first_before, first_tok, first_after = fw
    if not is_emptyish(first_after):
        return None

    first_tok = first_tok.strip()
    if not is_letter_token(first_tok):
        return None

    letter_tokens = [first_tok]

    for mid in mids:
        mid_core = segment_core_candidate(mid, "middle")
        if mid_core is None:
            return None

        mid_tok = mid_core["token"].strip()
        if not is_letter_token(mid_tok):
            return None

        letter_tokens.append(mid_tok)

    lw = first_wordish(last)
    if lw is None:
        return None

    last_before, last_tok, last_after = lw
    if not is_emptyish(last_before):
        return None

    last_tok = last_tok.strip()
    if not is_letter_token(last_tok):
        return None

    letter_tokens.append(last_tok)

    prefix_text = normalize_spaces(first_before)
    suffix_text = normalize_spaces(last_after)

    out = []
    seen = set()

    for tok in letter_tokens:
        val = normalize_spaces(
            " ".join([p for p in [prefix_text, tok, suffix_text] if p])
        )
        if val and val not in seen:
            seen.add(val)
            out.append(val)

    return out if len(out) >= 2 else None


def has_complete_feature_before_separator(text: str) -> bool:
    if not text:
        return False

    depth = 0
    segment_start = 0
    i = 0
    n = len(text)

    def segment_has_complete(seg: str) -> bool:
        seg = normalize_spaces(strip_pct_tail_only(seg))
        if not seg:
            return False

        tokens = iter_wordish_tokens(seg)

        if len(tokens) <= 1:
            return False

        return COMPLETE_ANY_RE.search(seg) is not None

    while i < n:
        ch = text[i]

        if ch == "(":
            depth += 1
            i += 1
            continue

        if ch == ")":
            depth = max(0, depth - 1)
            i += 1
            continue

        if depth == 0:
            if ch == "," and i + 1 < n and text[i + 1].isspace():
                if segment_has_complete(text[segment_start:i]):
                    return True
                i += 1
                segment_start = i
                continue

            m = CONJ_AND_OR_RE.match(text, i)
            if m:
                if segment_has_complete(text[segment_start:m.start()]):
                    return True
                i = m.end()
                segment_start = i
                continue

        i += 1

    return False



OSYL_HEADWORD_SUFFIX_RE = re.compile(
    r"^(?P<right_prefix>[A-Za-z0-9'()\[\]_/+-]*osyl)(?P<head>(?:ceramides?|cerebrosides?|gangliosides?|glycolipids?|glycosphingolipids?|sphingolipids?|saccharides?|glycans?|antigens?))$",
    flags=re.IGNORECASE,
)


def try_type_number_shared_context_rule(text: str):
    """
    Resolve only type-number coordinations where the first item has the word
    type and the later item is a bare number.

    Example:
      type 1 and 2 lacto series glycolipid antigens
      -> type 1 lacto series glycolipid antigens
      -> type 2 lacto series glycolipid antigens
    """
    if not text:
        return None

    conj_span = find_top_level_conj_span_special(text)
    if not conj_span:
        return None

    a, b = conj_span
    left = text[:a]
    right = text[b:]

    lw = last_wordish(left)
    rw = first_wordish(right)
    if lw is None or rw is None:
        return None

    left_before, left_tok, left_after = lw
    right_before, right_tok, right_after = rw

    if not is_emptyish(left_after) or not is_emptyish(right_before):
        return None

    head_info = parse_type_head_token(left_tok.strip())
    if not head_info:
        return None

    abo, first_num, _ = head_info

    # Allow both forms:
    #   type 1 and 2 ...
    #   type 1 and type 2 ...
    right_info = parse_type_head_token(right_tok.strip())
    if is_bare_type_number(right_tok.strip()):
        right_num = right_tok.strip()
    elif right_info is not None:
        right_abo, right_num, _ = right_info
        if (right_abo or "").lower() != (abo or "").lower():
            return None
    else:
        return None

    prefix_text = normalize_spaces(left_before)
    suffix_text = normalize_spaces(right_after)

    out = []
    for num in [first_num, right_num]:
        tok = build_type_token(num, abo)
        out.append(normalize_spaces(" ".join(p for p in [prefix_text, tok, suffix_text] if p)))

    return out if len(set(out)) >= 2 else None


def try_osyl_headword_coordination_rule(text: str):
    """
    Resolve coordinated incomplete modifier prefixes where the right item
    carries the shared suffix/headword.

    Examples:
      Ley-active hexaosyl- or octaosylceramides
      -> Ley-active hexaosylceramides
      -> Ley-active octaosylceramides

      Forssman penta- and tetraglycosylceramide are xenoantigens
      -> Forssman pentaglycosylceramide are xenoantigens
      -> Forssman tetraglycosylceramide are xenoantigens
    """
    if not text:
        return None

    conj_span = find_top_level_conj_span_special(text)
    if not conj_span:
        return None

    a, b = conj_span
    left = text[:a]
    right = text[b:]

    lw = last_wordish(left)
    rw = first_wordish(right)
    if lw is None or rw is None:
        return None

    left_before, left_tok, left_after = lw
    right_before, right_tok, right_after = rw

    if not is_emptyish(left_after) or not is_emptyish(right_before):
        return None

    left_tok = left_tok.strip()
    right_tok = right_tok.strip()

    prefix_alt_re = re.compile(
        r"^(mono|di|tri|tetra|penta|hexa|hepta|octa|nona|deca|oligo|poly|[A-Za-z0-9'()\[\]_/+-]*osyl)-?$",
        flags=re.IGNORECASE,
    )
    right_prefixed_re = re.compile(
        r"^(mono|di|tri|tetra|penta|hexa|hepta|octa|nona|deca|oligo|poly|[A-Za-z0-9'()\[\]_/+-]*osyl)(.+)$",
        flags=re.IGNORECASE,
    )

    lm = prefix_alt_re.fullmatch(left_tok)
    rm = right_prefixed_re.fullmatch(right_tok)
    if not lm or not rm:
        return None

    left_prefix = lm.group(1)
    right_prefix = rm.group(1)
    shared_suffix = rm.group(2)
    if not shared_suffix:
        return None

    # Keep this rule focused on glycan-like suffixes/headwords so it does not
    # over-expand ordinary English prefix coordinations.
    if not re.search(
        r"(?i)(osyl|glycosyl|ceramides?$|cerebrosides?$|gangliosides?$|glycolipids?$|glycosphingolipids?$|sphingolipids?$|saccharides?$|glycans?$|antigens?$)",
        shared_suffix,
    ):
        return None

    prefix_text = normalize_spaces(left_before)
    suffix_text = normalize_spaces(right_after)

    left_completed = left_prefix + shared_suffix
    right_completed = right_prefix + shared_suffix

    out1 = normalize_spaces(" ".join(p for p in [prefix_text, left_completed, suffix_text] if p))
    out2 = normalize_spaces(" ".join(p for p in [prefix_text, right_completed, suffix_text] if p))

    uniq = []
    seen = set()
    for x in [out1, out2]:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None

def try_special_breaking_conj_rule(text: str):
    if not text:
        return None

    conj_span = find_top_level_conj_span_special(text)
    if not conj_span:
        return None

    a, b = conj_span
    left = text[:a]
    right = text[b:]

    lw = last_wordish(left)
    rw = first_wordish(right)

    if lw is None or rw is None:
        return None

    left_before, left_tok, left_after = lw
    right_before, right_tok, right_after = rw

    if not is_emptyish(left_after):
        return None
    if not is_emptyish(right_before):
        return None

    left_before = normalize_spaces(left_before)
    right_after = normalize_spaces(right_after)

    left_tok = left_tok.strip()
    right_tok = right_tok.strip()

    if GM_LIKE_RE.fullmatch(left_tok) or GM_LIKE_RE.fullmatch(right_tok):
        return None

    if CORE_DIGIT_TOKEN_RE.fullmatch(left_tok) and CORE_DIGIT_TOKEN_RE.fullmatch(right_tok):
        out1 = normalize_spaces(" ".join([p for p in [left_before, left_tok, right_after] if p]))
        out2 = normalize_spaces(" ".join([p for p in [left_before, right_tok, right_after] if p]))
        return [out1, out2] if out1 != out2 else None

    head_info = parse_type_head_token(left_tok)
    if head_info and is_bare_type_number(right_tok):
        abo, first_num, _ = head_info
        out1_tok = build_type_token(first_num, abo)
        out2_tok = build_type_token(right_tok, abo)

        out1 = normalize_spaces(" ".join([p for p in [left_before, out1_tok, right_after] if p]))
        out2 = normalize_spaces(" ".join([p for p in [left_before, out2_tok, right_after] if p]))

        return [out1, out2] if out1 != out2 else None

    if not LEFT_BREAK_TOKEN_RE.fullmatch(left_tok):
        return None

    rm = RIGHT_BREAK_PREFIX_RE.match(right_tok)
    if not rm:
        return None

    right_prefix = rm.group(0)
    shared_suffix = right_tok[len(right_prefix):]

    if not shared_suffix and right_after:
        out1 = normalize_spaces(" ".join([p for p in [left_before, left_tok.rstrip("-"), right_after] if p]))
        out2 = normalize_spaces(" ".join([p for p in [left_before, right_tok, right_after] if p]))
        return [out1, out2] if out1 != out2 else None

    if not shared_suffix:
        return None

    left_completed = left_tok + shared_suffix
    left_completed = re.sub(r"-{2,}", "-", left_completed)

    out1 = normalize_spaces(" ".join([p for p in [left_before, left_completed, right_after] if p]))
    out2 = normalize_spaces(" ".join([p for p in [left_before, right_tok, right_after] if p]))

    uniq = []
    seen = set()
    for x in [out1, out2]:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def try_rule_1a(text: str):
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    parts = [s.strip() for s in segs]
    bases = [strip_pct_tail_only(p) for p in parts]

    for b in bases:
        if not is_rule1a_valid_segment(b):
            return None

    return parts


def candidates_in_segment(seg_raw: str):
    raw = seg_raw
    out = []

    for type_name, regexes in regex_groups.items():
        for rx in regexes:
            for m in rx.finditer(raw):
                a, b = m.start(), m.end()
                bb = b

                pct = PCT_AFTER_RE.match(raw[b:])
                if pct:
                    bb = b + pct.end()

                token_with_pct = raw[a:bb].strip()
                token_base = strip_pct_tail_only(token_with_pct)

                if is_dangling_hyphen_token(token_base):
                    continue

                out.append({
                    "type": type_name,
                    "a": a,
                    "b": bb,
                    "token_with_pct": token_with_pct,
                    "token_base": token_base,
                    "prefix": raw[:a],
                    "suffix": raw[bb:],
                })

    return out


def pick_candidate_for_position(cands, position):
    if position == "first":
        valid = [c for c in cands if is_emptyish(c["suffix"])]
        if not valid:
            return None
        valid.sort(key=lambda c: c["a"], reverse=True)
        return valid[0]

    if position == "last":
        valid = [c for c in cands if is_emptyish(c["prefix"])]
        if not valid:
            return None
        valid.sort(key=lambda c: c["a"])
        return valid[0]

    valid = [c for c in cands if is_emptyish(c["prefix"]) and is_emptyish(c["suffix"])]
    if not valid:
        return None

    valid.sort(key=lambda c: (c["b"] - c["a"]), reverse=True)
    return valid[0]


def try_rule_1b(text: str):
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    seg_cands = [candidates_in_segment(seg) for seg in segs]
    seg_types = [sorted({c["type"] for c in cands}) for cands in seg_cands]

    best = None
    n = len(segs)

    for i in range(n):
        if not seg_types[i]:
            continue

        shared = set(seg_types[i])

        for j in range(i + 1, n + 1):
            if j > i + 1:
                if not seg_types[j - 1]:
                    break
                shared &= set(seg_types[j - 1])

            if j - i >= 2 and shared:
                chosen_type = sorted(shared)[0]
                if best is None or (j - i) > (best[1] - best[0]):
                    best = (i, j, chosen_type)

    if best is None:
        return None

    i, j, chosen_type = best

    chosen = []
    for idx in range(i, j):
        cands = [c for c in seg_cands[idx] if c["type"] == chosen_type]
        if not cands:
            return None

        if idx == i:
            pos = "first"
        elif idx == j - 1:
            pos = "last"
        else:
            pos = "middle"

        pick = pick_candidate_for_position(cands, pos)
        if pick is None:
            return None

        chosen.append(pick)

    bases = [c["token_base"] for c in chosen]
    if not types_covering_all_words(bases):
        return None

    prefix_text = normalize_spaces(chosen[0]["prefix"])
    suffix_text = normalize_spaces(chosen[-1]["suffix"])

    expanded = []
    for c in chosen:
        tok = normalize_spaces(c["token_with_pct"])
        pieces = []
        if prefix_text:
            pieces.append(prefix_text)
        pieces.append(tok)
        if suffix_text:
            pieces.append(suffix_text)
        expanded.append(normalize_spaces(" ".join(pieces)))

    uniq = []
    seen = set()
    for x in expanded:
        if x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def try_rule_2(text: str):
    if not has_top_level_coordination(text):
        return None
    if exclude_rule2(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    first = segs[0]
    last = segs[-1]
    mids = segs[1:-1]

    fw = last_wordish(first)
    if fw is None:
        return None

    first_before, first_tok, first_after = fw

    lw = first_wordish(last)
    if lw is None:
        return None

    last_before, last_tok, last_after = lw

    mid_toks = []
    for m in mids:
        mw = first_wordish(m)
        if mw is None:
            return None
        _, mt, _ = mw
        rem = m.replace(mt, " ", 1)
        if WORDISH_RE.search(rem):
            return None
        mid_toks.append(mt.strip())

    last_wordish_tokens = [m.group(0) for m in iter_wordish_tokens(last)]
    suffix_text = ""
    x_last = last.strip()

    if len(last_wordish_tokens) >= 2:
        _, tok0, after0 = first_wordish(last)
        suffix_text = normalize_spaces(after0)
        x_last = tok0.strip()
    else:
        suffix_text = normalize_spaces(last_after)
        x_last = last_tok.strip()

    prefix_text = normalize_spaces(first_before)
    x_first = first_tok.strip()

    head_info = parse_type_head_token(x_first)
    if head_info:
        abo, first_num, _ = head_info
        all_rest_are_bare_nums = all(is_bare_type_number(t) for t in (mid_toks + [x_last]))

        if all_rest_are_bare_nums and is_emptyish(first_after) and is_emptyish(last_before):
            nums = [first_num] + mid_toks + [x_last]
            out = []

            for num in nums:
                type_tok = build_type_token(num, abo)
                out.append(normalize_spaces(" ".join([p for p in [prefix_text, type_tok, suffix_text] if p])))

            out_uniq = []
            seen = set()
            for x in out:
                if x and x not in seen:
                    seen.add(x)
                    out_uniq.append(x)

            return out_uniq if len(out_uniq) >= 2 else None

    if head_info:
        abo, first_num, _ = head_info
        later_type_infos = [parse_type_head_token(t) for t in (mid_toks + [x_last])]

        if all(info is not None for info in later_type_infos) and is_emptyish(first_after) and is_emptyish(last_before):
            nums = [first_num] + [info[1] for info in later_type_infos]
            out = []

            for num in nums:
                tok = build_type_token(num, abo)
                out.append(normalize_spaces(" ".join([p for p in [prefix_text, tok, suffix_text] if p])))

            out_uniq = []
            seen = set()
            for x in out:
                if x and x not in seen:
                    seen.add(x)
                    out_uniq.append(x)

            return out_uniq if len(out_uniq) >= 2 else None

    base_hy_num = re.fullmatch(r"([A-Za-z]+)-(\d+)", x_first)
    hy_num_mid = all(re.fullmatch(r"-(\d+)", t) for t in mid_toks + [x_last])

    if base_hy_num and hy_num_mid:
        base = base_hy_num.group(1)
        prefix2 = normalize_spaces((prefix_text + " " + base).strip()) if prefix_text else base
        suffix2 = suffix_text

        nums = [base_hy_num.group(2)]
        nums += [re.fullmatch(r"-(\d+)", t).group(1) for t in mid_toks]
        nums += [re.fullmatch(r"-(\d+)", x_last).group(1)]

        out = []
        for num in nums:
            out.append(normalize_spaces(" ".join([prefix2, f"{base}-{num}", suffix2]).strip()))

        out_uniq = []
        seen = set()
        for x in out:
            if x and x not in seen:
                seen.add(x)
                out_uniq.append(x)

        return out_uniq if len(out_uniq) >= 2 else None

    xs = [x_first] + mid_toks + [x_last]

    if not is_emptyish(first_after):
        return None
    if not is_emptyish(last_before):
        return None

    segs_stripped = [s.strip() for s in segs]

    def _segment_is_complete(seg_text: str) -> bool:
        seg_text = strip_pct_tail_only(normalize_spaces(seg_text))
        return bool(seg_text) and COMPLETE_ANY_RE.search(seg_text) is not None

    def _segment_has_leading_incomplete(seg_text: str) -> bool:
        seg_text = strip_pct_tail_only(normalize_spaces(seg_text))
        if not seg_text:
            return False

        fw2 = first_wordish(seg_text)
        if fw2 is None:
            return False

        tok2 = fw2[1].strip()
        return INCOMPLETE_END_RE.search(tok2) is not None

    all_previous_complete = _segment_is_complete(segs_stripped[0])
    for seg_text in segs_stripped[1:]:
        if all_previous_complete and _segment_has_leading_incomplete(seg_text):
            return segs_stripped
        all_previous_complete = all_previous_complete and _segment_is_complete(seg_text)

    out = []
    for x in xs:
        out.append(normalize_spaces(" ".join([p for p in [prefix_text, x, suffix_text] if p])))

    out_uniq = []
    seen = set()
    for x in out:
        if x and x not in seen:
            seen.add(x)
            out_uniq.append(x)

    return out_uniq if len(out_uniq) >= 2 else None


def split_x_and_z(token: str):
    if not token or "-" not in token:
        return None

    i = token.rfind("-")
    x = token[:i]
    z = token[i + 1:]

    if not x or not z:
        return None
    if not Z_OK_RE.fullmatch(z):
        return None

    return x, z


def try_rule_3(text: str):
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    first = segs[0]
    last = segs[-1]
    mids = segs[1:-1]

    fw = last_wordish(first)
    if fw is None:
        return None

    first_before, first_tok, first_after = fw
    if not first_tok.endswith("-"):
        return None
    if not is_emptyish(first_after):
        return None

    mid_toks = []
    for m in mids:
        mw = first_wordish(m)
        if mw is None:
            return None

        _, mt, _ = mw
        if not mt.endswith("-"):
            return None

        rem = m.replace(mt, " ", 1)
        if WORDISH_RE.search(rem):
            return None

        mid_toks.append(mt.strip())

    lw = first_wordish(last)
    if lw is None:
        return None

    last_before, last_tok, last_after = lw
    if not is_emptyish(last_before):
        return None

    xz = split_x_and_z(last_tok.strip())
    if xz is None:
        return None

    _, z = xz

    prefix_text = normalize_spaces(first_before)
    suffix_text = normalize_spaces(last_after)

    x_bases = [first_tok.strip()[:-1]] + [t[:-1] for t in mid_toks]
    x_tokens_final = [f"{xb}-{z}" for xb in x_bases] + [last_tok.strip()]

    expanded = []
    for xt in x_tokens_final:
        expanded.append(normalize_spaces(" ".join([p for p in [prefix_text, xt, suffix_text] if p])))

    uniq = []
    seen = set()
    for x in expanded:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def try_rule_4a(text: str):
    if not text:
        return None

    conj_span = find_top_level_conj_span(text)
    if not conj_span:
        return None

    a, b = conj_span
    left = text[:a]
    right = text[b:]

    left = re.sub(r"\s*[-/]\s*$", lambda m: m.group(0).replace(" ", ""), left)

    lw = last_wordish(left)
    rw = first_wordish(right)
    if lw is None or rw is None:
        return None

    left_before, left_tok, left_after = lw
    right_before, right_tok, right_after = rw

    if not is_emptyish(left_after):
        return None
    if not is_emptyish(right_before):
        return None

    left_tok = left_tok.strip()
    right_tok = right_tok.strip()

    _, suffix_chain = decompose_suffix_chain(right_tok)
    if len(suffix_chain) >= 2:
        return None

    suffix_info = dict_suffix_at_end(right_tok)
    if not suffix_info:
        return None

    z, _ = suffix_info

    if token_ends_with_dict_suffix(left_tok):
        return None
    if len(right_tok) <= len(z):
        return None

    prefix_text = normalize_spaces(left_before)
    suffix_text = normalize_spaces(right_after)

    completed_left = join_with_dict_suffix(left_tok, z)

    out1 = normalize_spaces(" ".join([p for p in [prefix_text, completed_left, suffix_text] if p]))
    out2 = normalize_spaces(" ".join([p for p in [prefix_text, right_tok, suffix_text] if p]))

    if out1 == out2:
        return None

    uniq = []
    seen = set()
    for x in [out1, out2]:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def try_rule_4b(text: str):
    if not text:
        return None
    if not has_top_level_coordination(text):
        return None

    segs = split_top_level_segments(text)
    if len(segs) <= 1:
        return None

    n = len(segs)
    cands = []

    for idx, seg in enumerate(segs):
        if idx == 0:
            pos = "first"
        elif idx == n - 1:
            pos = "last"
        else:
            pos = "middle"

        cand = segment_core_candidate(seg, pos)
        if cand is None:
            return None

        cands.append(cand)

    prefix_text = normalize_spaces(cands[0]["prefix"])
    suffix_text = normalize_spaces(cands[-1]["suffix"])

    left_tokens = [c["token"] for c in cands[:-1]]
    last_token = cands[-1]["token"]

    if not token_ends_with_dict_suffix(last_token):
        return None

    _, suffix_chain = decompose_suffix_chain(last_token)
    if not suffix_chain:
        return None

    shared_suffix = "".join(z for z, _ in suffix_chain)
    if not shared_suffix:
        return None

    outputs = []

    for tok in left_tokens:
        if token_ends_with_dict_suffix(tok):
            return None

        built = join_with_dict_suffix(tok, shared_suffix)
        row_text = normalize_spaces(" ".join([p for p in [prefix_text, built, suffix_text] if p]))

        if row_text:
            outputs.append(row_text)

    last_row = normalize_spaces(" ".join([p for p in [prefix_text, last_token, suffix_text] if p]))
    if last_row:
        outputs.append(last_row)

    uniq = []
    seen = set()
    for x in outputs:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    return uniq if len(uniq) >= 2 else None


def get_maximal_dict_matches(text: str, max_hits: int = 200) -> str:
    if not text:
        return ""

    hits = []

    for type_name, regexes in regex_groups.items():
        for rx in regexes:
            for m in rx.finditer(text):
                a, b = m.start(), m.end()
                if a == b:
                    continue
                hits.append((a, b, text[a:b], type_name))

    if not hits:
        return ""

    span_best_type = {}
    span_text = {}

    for a, b, s, t in hits:
        key = (a, b)
        span_text[key] = s
        if key not in span_best_type or t < span_best_type[key]:
            span_best_type[key] = t

    uniq = [
        (a, b, span_text[(a, b)], span_best_type[(a, b)])
        for (a, b) in span_best_type.keys()
    ]
    uniq.sort(key=lambda x: (-(x[1] - x[0]), x[0], x[3], x[2]))

    kept = []
    kept_spans = []

    for a, b, s, t in uniq:
        length = b - a
        contained = False

        for A, B, LL in kept_spans:
            if A <= a and b <= B and LL > length:
                contained = True
                break

        if not contained:
            kept.append((a, b, s, t))
            kept_spans.append((a, b, length))

        if len(kept) >= max_hits:
            break

    kept.sort(key=lambda x: (x[0], x[1]))
    return "; ".join([f"{s} ({t})" for _, _, s, t in kept])



PAREN_SPAN_RE = re.compile(r"\(([^()]*)\)")
SINGLE_LETTER_PAREN_RE = re.compile(r"^[A-Za-z]$")


def find_resolvable_parenthetical(text: str):
    """Find a parenthetical span preceded by a space and not containing one letter only."""
    text = "" if text is None else str(text)

    for m in PAREN_SPAN_RE.finditer(text):
        if m.start() == 0 or not text[m.start() - 1].isspace():
            continue

        inner = normalize_spaces(m.group(1))
        if not inner:
            continue
        if SINGLE_LETTER_PAREN_RE.fullmatch(inner):
            continue

        return m, inner

    return None


def resolve_parenthetical_text(text: str):
    """
    Split one BERT_Output value when it contains a resolvable parenthetical term.

    Rules:
      - Parentheses must be preceded by a space.
      - Parenthetical content that is a single letter, such as (A), is ignored.
      - If the parenthetical is final: keep the outside text and make the
        parenthetical content a separate entry.
      - If text follows the parenthetical: keep the outside text as one entry,
        and make parenthetical content + right outside text another entry.
    """
    text = normalize_spaces(text)
    found = find_resolvable_parenthetical(text)
    if not found:
        return None

    m, inner = found
    left = normalize_spaces(text[:m.start()])
    right = normalize_spaces(text[m.end():])

    without_parentheses = normalize_spaces(" ".join(p for p in [left, right] if p))

    if right:
        parenthetical_entry = normalize_spaces(" ".join(p for p in [inner, right] if p))
    else:
        parenthetical_entry = inner

    out = []
    seen = set()
    for val in [without_parentheses, parenthetical_entry]:
        if val and val not in seen:
            seen.add(val)
            out.append(val)

    return out if len(out) >= 2 else None


def apply_parentheses_resolution(out_df: pd.DataFrame) -> pd.DataFrame:
    if PAREN_COL not in out_df.columns:
        out_df[PAREN_COL] = ""

    new_rows = []

    for _, row in out_df.iterrows():
        text = row.get(NEWCOL, "")
        text = "" if pd.isna(text) else str(text)
        parts = resolve_parenthetical_text(text)

        if parts:
            for p in parts:
                row2 = row.copy()
                row2[NEWCOL] = p
                row2[PAREN_COL] = "Yes"
                row2[DIAG_COL] = ""
                new_rows.append(row2)
        else:
            if pd.isna(row.get(PAREN_COL, "")):
                row[PAREN_COL] = ""
            new_rows.append(row)

    return pd.DataFrame(new_rows)


df = pd.read_excel(INPUT_XLSX)

if COL not in df.columns:
    raise ValueError(f"Missing required column: {COL}")
if OUTCOL not in df.columns:
    df[OUTCOL] = ""
if DIAG_COL not in df.columns:
    df[DIAG_COL] = ""
if PAREN_COL not in df.columns:
    df[PAREN_COL] = ""

rows = []
df = df.copy()
df["_row_id_tmp"] = range(len(df))

for _, row in df.iterrows():
    text = row.get(COL, "")
    text = "" if pd.isna(text) else str(text)
    text = remove_comma_ly_tail(text)

    discard_val = row.get(DISCARD_COL, "")
    discard_val = "" if pd.isna(discard_val) else str(discard_val).strip().lower()

    if discard_val == "yes":
        row2 = row.copy()
        row2[NEWCOL] = text

        if pd.isna(row2.get(OUTCOL, "")):
            row2[OUTCOL] = ""

        if pd.isna(row2.get(DIAG_COL, "")):
            row2[DIAG_COL] = ""

        rows.append(row2)
        continue

    parts_letter_conj_rule = try_letter_conj_rule(text)
    if parts_letter_conj_rule:
        for p in parts_letter_conj_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "letter_conj_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue


    parts_type_number_shared_context_rule = try_type_number_shared_context_rule(text)
    if parts_type_number_shared_context_rule:
        for p in parts_type_number_shared_context_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "type_number_shared_context_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_osyl_headword_coordination_rule = try_osyl_headword_coordination_rule(text)
    if parts_osyl_headword_coordination_rule:
        for p in parts_osyl_headword_coordination_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "osyl_headword_coordination_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_prefix_incomplete_rule = try_prefix_incomplete_rule(text)
    if parts_prefix_incomplete_rule:
        for p in parts_prefix_incomplete_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "prefix_incomplete_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_core_series_rule = try_core_series_rule(text)
    if parts_core_series_rule:
        for p in parts_core_series_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "core_series_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_complete_feature_list_rule = try_complete_feature_list_rule(text)
    if parts_complete_feature_list_rule:
        for p in parts_complete_feature_list_rule:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "complete_feature_list_rule"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    if has_complete_feature_before_separator(text):
        row2 = row.copy()
        row2[NEWCOL] = text
        row2[DIAG_COL] = get_maximal_dict_matches(text)

        if pd.isna(row2.get(OUTCOL, "")):
            row2[OUTCOL] = ""

        rows.append(row2)
        continue

    parts_before_other_rule1 = try_before_other_rule1(text)
    if parts_before_other_rule1:
        for p in parts_before_other_rule1:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "before_other_rule1"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_special = try_special_breaking_conj_rule(text)
    if parts_special:
        for p in parts_special:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "YesSpecial"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_1a = try_rule_1a(text)
    if parts_1a:
        for p in parts_1a:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes1a"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    guard_parts = should_keep_original_segments_due_to_complete_incomplete_guard(text)
    if guard_parts:
        for p in guard_parts:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes2a"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_1b = try_rule_1b(text)
    if parts_1b:
        for p in parts_1b:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes1b"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_2 = try_rule_2(text)
    if parts_2:
        for p in parts_2:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes2a"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_3 = try_rule_3(text)
    if parts_3:
        for p in parts_3:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes2b"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_4b = try_rule_4b(text)
    if parts_4b:
        for p in parts_4b:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes3b"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    parts_4a = try_rule_4a(text)
    if parts_4a:
        for p in parts_4a:
            new_row = row.copy()
            new_row[NEWCOL] = p
            new_row[OUTCOL] = "Yes3a"
            new_row[DIAG_COL] = ""
            rows.append(new_row)
        continue

    row2 = row.copy()
    row2[NEWCOL] = text
    row2[DIAG_COL] = get_maximal_dict_matches(text)

    if pd.isna(row2.get(OUTCOL, "")):
        row2[OUTCOL] = ""

    rows.append(row2)

out = pd.DataFrame(rows)
out = apply_parentheses_resolution(out)

if "_row_id_tmp" in out.columns:
    out = out.drop_duplicates(subset=["_row_id_tmp", NEWCOL, OUTCOL, DIAG_COL, PAREN_COL])
    out = out.drop(columns=["_row_id_tmp"])

out.to_excel(OUTPUT_XLSX, index=False)

print(f"Done. Output saved to: {OUTPUT_XLSX}")
