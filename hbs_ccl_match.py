"""
HBS ↔ CCL matching logic recovered from cloud-agent transcript
bc-5eee4a56-c93e-443c-92a6-9397b12e2c67 (2026-08-07).

Primary source: assistant tool command with CANON_PATTERNS / extract_concepts / INCOMPAT
(msg index ~49). Ancillary maps STOP / REPL / GENERIC_ONLY / STRONG_PHRASES from an earlier
phrase-based attempt (msg ~45) are included for completeness.

Original upload paths (no longer present in this environment):
  HBS: /home/ubuntu/.cursor/projects/workspace/uploads/HBS___________1__fd77.xlsx
  CCL: /home/ubuntu/.cursor/projects/workspace/uploads/____________________________________________________________4_09_8e2b.xlsx
       (original name: Перечень_обслуживаемых_компонентов_Фастэйр_Технологии_Рев_№_4_09.xlsx)
  CCL sheet: 'Перечень обс. комп.'  (header row offset: iloc[2:], cols 2..9)
"""

from __future__ import annotations

import re
from collections import defaultdict
from difflib import SequenceMatcher
from typing import Any, Iterable

import pandas as pd

# ---------------------------------------------------------------------------
# P/N + name normalization (msg49: norm_pn / norm_name synonym map)
# ---------------------------------------------------------------------------

def normalize_pn(p: Any) -> str:
    """Strip whitespace and common separators; uppercase. Alias: norm_pn / npn."""
    return re.sub(r"[\s\-_/\\.]", "", str(p).upper().strip())


# Alias used in later scripts in the same transcript
norm_pn = normalize_pn
npn = normalize_pn

# Synonym / rewrite patterns applied inside norm_name (msg49 `reps`)
NAME_SYNONYM_REPS: dict[str, str] = {
    r"\bmachine assy\s*-\s*air cycle\b": "air cycle machine",
    r"\bair cycle machine\b": "air cycle machine",
    r"\brecirc(?:ulation)? fan\b": "recirculation fan",
    r"\bheat exchanger\b": "heat exchanger",
    r"\bexchanger assy\b": "heat exchanger",
    r"\bexchanger\b": "heat exchanger",
    r"\boverheat\b": "overtemperature",
    r"\bturbine inlet overtemperature\b": "overtemperature",
    r"\bcabin attendant.*handset\b": "attendant handset",
    r"\bflight deck.*handset\b": "attendant handset",
    r"\battendant handset\b": "attendant handset",
    r"\bhandset module\b": "attendant handset",
    r"\bbattery pack ass(?:y|embly)\b": "battery pack",
    r"\bbattery assy\b": "battery assembly",
    r"\bmain battery\b": "main battery",
    r"\bemergency battery\b": "emergency battery",
    r"\bemergency power supply\b": "emergency power supply",
    r"\bpower supply unit\b": "power supply unit",
    r"\bwater boiler\b": "water boiler",
    r"\bcoffee maker\b": "coffee maker",
    r"\bpropeller blade\b": "propeller blade",
    r"\bpropeller brake\b": "propeller brake",
    r"\batc control panel\b": "atc control",
    r"\bcta-?\d+[a-z/]* control unit\b": "atc control",
    r"\bpassenger address amplifier\b": "passenger address amplifier",
    r"\bdrain valve\b": "drain valve",
    r"\bfuel control unit\b": "fuel control unit",
    r"\bfuel control,?\s*mechanical\b": "fuel control unit",
    r"\bfuel control\b": "fuel control unit",
    r"\bdifferential pressure switch\b": "differential pressure switch",
    r"\bfuel differential pressure switch\b": "differential pressure switch",
    r"\bflap track(?: rail)?\b": "flap track",
    r"\bflap track assy\b": "flap track",
    r"\bengine fire extinguisher\b": "engine fire extinguisher",
    r"\bfire extinguisher bottle\b": "fire extinguisher bottle",
    r"\bportable extinguisher\b": "portable fire extinguisher",
    r"\bwater fire extinguisher\b": "portable fire extinguisher",
    r"\bballast unit\b": "ballast",
    r"\bballast\b": "ballast",
    r"\boven\b": "oven",
    r"\bstarter generator\b": "starter generator",
    r"\bdc starter generator\b": "starter generator",
    r"\bgenerator assy\b": "generator",
    r"\bintegrated drive generator\b": "integrated drive generator",
    r"\blanding light\b": "landing light",
    r"\btaxi light\b": "taxi light",
    r"\bnavigation light\b": "navigation light",
    r"\bstrobe light\b": "strobe light",
    r"\bsliding window\b": "window",
    r"\bwindow assy\b": "window",
}


def normalize_name(n: Any) -> str:
    """Lowercase, strip punctuation, apply aviation synonym rewrites. Alias: norm_name."""
    s = str(n).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9\s\-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    for pat, rep in NAME_SYNONYM_REPS.items():
        s = re.sub(pat, rep, s)
    return re.sub(r"\s+", " ", s).strip()


norm_name = normalize_name

# ---------------------------------------------------------------------------
# Canonical concept patterns + exclusion / incompatibility rules (msg49)
# ---------------------------------------------------------------------------

CANON_PATTERNS: list[tuple[str, str]] = [
    ("attendant handset", r"\battendant handset\b|\bhandset\b"),
    ("water boiler", r"\bwater boiler\b"),
    ("coffee maker", r"\bcoffee maker\b"),
    ("oven", r"\boven\b"),
    ("ballast", r"\bballast\b"),
    ("battery pack", r"\bbattery pack\b"),
    ("emergency battery", r"\bemergency battery\b"),
    ("emergency power supply", r"\bemergency power supply\b"),
    ("power supply unit", r"\bpower supply unit\b"),
    ("main battery", r"\bmain battery\b"),
    ("battery", r"\bbattery\b"),
    ("propeller blade", r"\bpropeller blade\b"),
    ("propeller brake", r"\bpropeller brake\b"),
    ("atc control", r"\batc control\b"),
    ("drain valve", r"\bdrain valve\b"),
    ("fuel control unit", r"\bfuel control unit\b"),
    ("differential pressure switch", r"\bdifferential pressure switch\b"),
    ("flap track", r"\bflap track\b"),
    ("air cycle machine", r"\bair cycle machine\b"),
    ("recirculation fan", r"\brecirculation fan\b"),
    ("heat exchanger", r"\bheat exchanger\b"),
    (
        "overtemperature switch",
        r"\bovertemperature\b.*\bswitch\b|\bswitch\b.*\bovertemperature\b|\bturbine inlet overtemperature\b",
    ),
    ("landing light", r"\blanding light\b"),
    ("taxi light", r"\btaxi light\b"),
    ("navigation light", r"\bnavigation light\b"),
    ("strobe light", r"\bstrobe light\b"),
    ("starter generator", r"\bstarter generator\b"),
    ("passenger address amplifier", r"\bpassenger address amplifier\b"),
    ("cabin pressure indicator", r"\bcabin press(?:ure)? indicator\b"),
    ("pressure control panel", r"\bpressure control panel\b"),
    ("window", r"\bwindow\b"),
    ("engine fire extinguisher", r"\bengine fire extinguisher\b|\bfire extinguisher bottle\b"),
    (
        "portable fire extinguisher",
        r"\bportable fire extinguisher\b|\bwater fire extinguisher\b|\bportable extinguisher\b",
    ),
]

# Concept pairs that must NOT cross-match
INCOMPAT: set[frozenset[str]] = {
    frozenset({"propeller blade", "propeller brake"}),
    frozenset({"engine fire extinguisher", "portable fire extinguisher"}),
    frozenset({"emergency battery", "battery pack"}),
    frozenset({"emergency battery", "main battery"}),
    frozenset({"emergency battery", "battery"}),
    frozenset({"emergency power supply", "power supply unit"}),
    frozenset({"main battery", "emergency battery"}),
}


def extract_concepts(nn: str) -> list[str]:
    """Extract canonical product concepts from a normalized name."""
    found: list[str] = []
    for label, pat in CANON_PATTERNS:
        if re.search(pat, nn):
            found.append(label)
    return found


# Alias used in transcript
extract_concept = extract_concepts  # singular form sometimes referenced in search


def compatible(a: str, b: str) -> bool:
    if a == b:
        return True
    if {a, b} <= {"battery", "battery pack", "main battery"}:
        if "emergency battery" in {a, b}:
            return False
        return True
    return False


# ---------------------------------------------------------------------------
# Ancillary maps from earlier phrase-based matcher (msg45) — optional
# ---------------------------------------------------------------------------

STOP = {
    "assy",
    "assembly",
    "assys",
    "module",
    "unit",
    "system",
    "panel",
    "switch",
    "indicator",
    "control",
    "controller",
    "sensor",
    "valve",
    "pump",
    "fan",
    "light",
    "lamp",
    "heater",
    "motor",
    "relay",
    "box",
    "computer",
    "transmitter",
    "receiver",
    "amplifier",
    "battery",
    "pack",
    "main",
    "cabin",
    "cockpit",
    "aircraft",
    "and",
    "the",
    "of",
    "for",
    "with",
    "type",
    "assy.",
    "ass",
    "equip",
    "equipment",
    "device",
    "set",
    "kit",
    "item",
    "part",
    "lh",
    "rh",
    "left",
    "right",
    "fwd",
    "aft",
    "upper",
    "lower",
    "no",
    "nr",
    "number",
    "pn",
}

REPL = [
    (r"\belt\b", "emergency locator transmitter"),
    (r"\bvhf\b", "vhf"),
    (r"\batc\b", "atc"),
    (r"\badf\b", "adf"),
    (r"\bdme\b", "dme"),
    (r"\bgps\b", "gps"),
    (r"\biru\b", "inertial reference"),
    (r"\badiru\b", "air data inertial reference"),
    (r"\bfadec\b", "fadec"),
    (r"\becu\b", "electronic control unit"),
    (r"\bpcu\b", "power control unit"),
    (r"\bbcu\b", "brake control unit"),
    (r"\bapu\b", "apu"),
    (r"\bidg\b", "integrated drive generator"),
    (r"\btr\b", "transformer rectifier"),
    (r"\bt/r\b", "transformer rectifier"),
    (r"\bacm\b", "air cycle machine"),
    (r"\bmachine assy\s*-\s*air cycle\b", "air cycle machine"),
    (r"\bair cycle\b", "air cycle machine"),
    (r"\brecirc(?:ulation)?\b", "recirculation"),
    (r"\boverheat\b", "overtemperature"),
    (r"\bovertemperature\b", "overtemperature"),
    (r"\bhandset\b", "handset"),
    (r"\battendant\b", "attendant"),
    (r"\bflight deck\b", "flight deck"),
    (r"\blanding light\b", "landing light"),
    (r"\bnavigation light\b", "navigation light"),
    (r"\bstrobe light\b", "strobe light"),
    (r"\btaxi light\b", "taxi light"),
    (r"\banti[- ]?collision\b", "anticollision"),
    (r"\bballast\b", "ballast"),
    (r"\bwater boiler\b", "water boiler"),
    (r"\bpropeller blade\b", "propeller blade"),
    (r"\bbattery pack\b", "battery pack"),
    (r"\bbattery assy\b", "battery assembly"),
    (r"\bgenerator assy\b", "generator assembly"),
    (r"\bstarter generator\b", "starter generator"),
    (r"\bemergency battery\b", "emergency battery"),
    (r"\bpressure control panel\b", "pressure control panel"),
    (r"\bcabin press(?:ure)? indicator\b", "cabin pressure indicator"),
    (r"\bexchanger\b", "heat exchanger"),
    (r"\bheat exchanger\b", "heat exchanger"),
]

GENERIC_ONLY = {
    "panel",
    "switch",
    "indicator",
    "module",
    "assembly",
    "assy",
    "unit",
    "box",
    "relay",
    "sensor",
    "valve",
    "pump",
    "fan",
    "light",
    "lamp",
    "motor",
    "battery",
    "transmitter",
    "receiver",
    "amplifier",
    "computer",
    "heater",
    "controller",
    "control",
}

STRONG_PHRASES = [
    "landing light",
    "taxi light",
    "navigation light",
    "strobe light",
    "anticollision light",
    "water boiler",
    "propeller blade",
    "attendant handset",
    "cabin attendant handset",
    "air cycle machine",
    "recirculation fan",
    "heat exchanger",
    "overtemperature switch",
    "starter generator",
    "integrated drive generator",
    "emergency locator transmitter",
    "passenger address amplifier",
    "cabin pressure indicator",
    "pressure control panel",
    "battery pack",
    "main battery",
    "emergency battery",
    "transformer rectifier",
    "vhf com control panel",
    "atc control panel",
    "ballast",
    "cooling turbine",
    "galley heating fan",
    "cockpit display fan",
    "lavatory galley exhaust fan",
    "aileron autopilot actuator",
    "thrust lever",
    "annunciator",
]

# ---------------------------------------------------------------------------
# Loaders matching the original Excel layouts
# ---------------------------------------------------------------------------

def load_ccl(path: str | Any, sheet_name: str = "Перечень обс. комп.") -> pd.DataFrame:
    ccl_raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    ccl = ccl_raw.iloc[2:, [2, 3, 4, 5, 6, 7, 8, 9]].copy()
    ccl.columns = ["num", "category", "ata", "pn", "name", "docs", "mfr", "level"]
    ccl = ccl[
        ccl["num"].apply(
            lambda x: bool(re.fullmatch(r"\d+", str(x).strip())) if pd.notna(x) else False
        )
    ].copy()
    for col in ["pn", "name", "ata", "category", "mfr", "level"]:
        ccl[col] = ccl[col].astype(str).str.strip()
    return ccl.reset_index(drop=True)


def load_hbs(path: str | Any, sheet_name: str = "Лист1") -> pd.DataFrame:
    hbs = pd.read_excel(path, sheet_name=sheet_name)
    hbs = hbs.rename(
        columns={
            "Part Number": "pn",
            "Part Description": "name",
            "TYPE": "ac_type",
            "ATA": "ata",
            "Категория GO/GO IF/NO GO": "go_cat",
        }
    )
    for col in ["pn", "name", "ata", "ac_type", "go_cat"]:
        if col in hbs.columns:
            hbs[col] = hbs[col].astype(str).str.strip()
    return hbs


# ---------------------------------------------------------------------------
# Core match: exact P/N + semantic concepts (msg49 algorithm)
# ---------------------------------------------------------------------------

def match_hbs_ccl(
    hbs: pd.DataFrame,
    ccl: pd.DataFrame,
    *,
    fuzzy_miss_threshold: float = 0.55,
    fuzzy_miss_keywords: Iterable[str] | None = None,
) -> dict[str, Any]:
    """
    Returns dict with:
      exact: list[(hbs_idx, ccl_row)]
      semantic: list[(hbs_idx, list[(concept, ccl_row)])]
      exact_idx: set of hbs indices with exact PN match
      hbs / ccl: frames with pn_n, nn, concepts columns
      fuzzy_misses: optional SequenceMatcher hints for HBS without concepts
    """
    hbs = hbs.copy()
    ccl = ccl.copy()
    hbs["pn_n"] = hbs["pn"].map(normalize_pn)
    ccl["pn_n"] = ccl["pn"].map(normalize_pn)
    hbs["nn"] = hbs["name"].map(normalize_name)
    ccl["nn"] = ccl["name"].map(normalize_name)
    hbs["concepts"] = hbs["nn"].map(extract_concepts)
    ccl["concepts"] = ccl["nn"].map(extract_concepts)

    ccl_by_pn: dict[str, list] = defaultdict(list)
    for _, r in ccl.iterrows():
        ccl_by_pn[r["pn_n"]].append(r)

    exact_idx: set = set()
    exact: list = []
    for i, r in hbs.iterrows():
        for m in ccl_by_pn.get(r["pn_n"], []):
            exact_idx.add(i)
            exact.append((i, m))

    ccl_by_concept: dict[str, list] = defaultdict(list)
    for _, r in ccl.iterrows():
        for c in r["concepts"]:
            ccl_by_concept[c].append(r)

    semantic: list = []
    for i, r in hbs.iterrows():
        if i in exact_idx:
            continue
        concepts = r["concepts"]
        if not concepts:
            continue
        matched_rows = []
        for concept in concepts:
            for m in ccl_by_concept.get(concept, []):
                if m["pn_n"] == r["pn_n"]:
                    continue
                ok = True
                for mc in m["concepts"]:
                    if frozenset({concept, mc}) in INCOMPAT and concept != mc:
                        ok = False
                if not ok:
                    continue
                if not any(compatible(concept, mc) or concept == mc for mc in m["concepts"]):
                    if concept not in m["concepts"]:
                        continue
                matched_rows.append((concept, m))
        seen: set = set()
        uniq = []
        for concept, m in matched_rows:
            if m["pn"] in seen:
                continue
            seen.add(m["pn"])
            uniq.append((concept, m))
        if uniq:
            semantic.append((i, uniq))

    keywords = list(
        fuzzy_miss_keywords
        or [
            "fan",
            "light",
            "generator",
            "exchanger",
            "valve",
            "pump",
            "heater",
            "battery",
            "oven",
            "boiler",
            "extinguisher",
            "handset",
            "ballast",
            "sensor",
            "actuator",
            "indicator",
            "amplifier",
            "transmitter",
            "panel",
            "window",
            "brake",
            "blade",
            "coffee",
            "fuel",
            "fire",
            "pressure",
            "temperature",
        ]
    )
    fuzzy_misses = []
    for i, r in hbs.iterrows():
        if i in exact_idx or r["concepts"]:
            continue
        nn = r["nn"]
        if not any(k in nn for k in keywords):
            continue
        best = []
        for _, m in ccl.iterrows():
            ratio = SequenceMatcher(None, nn, m["nn"]).ratio()
            if ratio >= fuzzy_miss_threshold:
                best.append((ratio, m))
        best.sort(reverse=True, key=lambda x: x[0])
        if best[:3]:
            fuzzy_misses.append((i, best[:3]))

    return {
        "exact": exact,
        "semantic": semantic,
        "exact_idx": exact_idx,
        "hbs": hbs,
        "ccl": ccl,
        "fuzzy_misses": fuzzy_misses,
    }


if __name__ == "__main__":
    import sys
    from pathlib import Path

    if len(sys.argv) < 3:
        print(
            "Usage: python hbs_ccl_match_recovered.py <hbs.xlsx> <ccl.xlsx>\n"
            "Recovered matching library; requires the original Excel files."
        )
        sys.exit(2)
    hbs_path, ccl_path = sys.argv[1], sys.argv[2]
    hbs = load_hbs(hbs_path)
    ccl = load_ccl(ccl_path)
    result = match_hbs_ccl(hbs, ccl)
    print("Exact HBS items:", len(result["exact_idx"]))
    print("Semantic HBS items:", len(result["semantic"]))
    print("CCL rows:", len(ccl), "HBS rows:", len(hbs))
