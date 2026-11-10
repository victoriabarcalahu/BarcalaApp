#!/usr/bin/env python3
"""
BarcalaReviewApp — a screening workbench for systematic literature reviews.

  python BarcalaReviewApp.py my_corpus.xlsx      name the corpus once
  python BarcalaReviewApp.py                     and it is remembered

Nothing about any particular subject is written into this program. Your
themes, inclusion and exclusion criteria, screening rules, the keyword
patterns that guess a country or a theme, your name and your email all live
in review_config.json beside it, which is created on first run.

Opens in your browser. Everything is saved as you work, next to the corpus,
in files named after it. The corpus itself is only ever read.

What is new in this version
---------------------------
EVERY ROW OF EVERY IMPORT IS KEPT.
    Importing used to throw away any record another database had already
    given you. That hid most of the duplicates — 228 of 259 on this corpus —
    so they never reached the Duplicates tab, could not be checked, and the
    better of the two versions could never win.

    Now every usable row is stored with the source it came from, and overlap
    is resolved in the corpus, where you can see it. The figure in the PRISMA
    diagram and the figure on the Duplicates tab are the same number, because
    they are now the same thing.

NOTHING IS THROWN AWAY WHEN A DUPLICATE IS RESOLVED.
    Choosing the most complete version is not the same as keeping the most
    complete record: Web of Science may have the abstract that OpenAlex
    lacks, while OpenAlex alone has the topic, subfield and field. The kept
    record now keeps everything it has and takes only what it is missing from
    the others. Borrowed values are marked on the Duplicates tab and named in
    the export's merged_fields column.

IMPORTING THE SAME FILE TWICE CHANGES NOTHING.
    Imported records are keyed by their own DOI or title rather than by their
    position in the import, so re-importing an export is recognised instead of
    stored again. Before, the same 126-row Web of Science file read twice was
    reported as 252 records identified. An _extra_records.csv written by an
    older version is migrated to the new keys on first load, and decisions
    follow their records automatically.

YOUR GOLD SET IS THE KEY-PAPER LIST.
    Import a list with a `status` column and the core/background marking
    comes with it. PRISMA counts those under other methods — added by hand,
    not found by a database — and the Stopping tab separates three questions
    that call for different action: how many key papers your SEARCH returns
    (widen the query), how many are in the corpus at all, and how many you
    have included (keep screening).

RULES THAT FEED PRISMA DIRECTLY, AND STAY OVERRIDABLE.
    A rule like "not in English" or "published 5+ years ago and never cited"
    excludes records automatically and those exclusions flow straight into
    the PRISMA diagram with their reason. Every one stays overridable on the
    record: open it, press Include, and your decision wins over the rule for
    good. The Rules tab shows what each rule WOULD remove before you switch
    it on, so you never enable one blind.

    A rule never silently overrides you. Manual beats automatic, always, and
    the Screen tab marks which is which.

DUPLICATES RESOLVED BY COMPLETENESS.
    Within each duplicate group the most complete record is kept
    automatically. A DOI counts for more than anything else, then an
    abstract, then authors, journal and year, with citation count as a
    tiebreak. You can switch which one is kept, and the app shows you what
    each version has that the others do not.

COUNTRY AND THEME GUESSED, THEN EDITED.
    Country is read from the title and abstract first, because that is the
    country STUDIED, and only falls back to author affiliation, which is
    merely where the authors work. Theme is matched against your three
    research themes. Both show a "guess" badge until you confirm or change
    them, so you always know which values you have actually checked.

A DASHBOARD YOU DRIVE.
    Choose what to group by and what to split by, apply any filter, and the
    chart redraws. Colours are the Okabe-Ito set, which stays readable for
    colour-blind readers.

Everything runs locally. No data leaves your machine except on the Ask tab,
which sends only the records it quotes, and only when you press Ask.

Standard library only, except openpyxl if your corpus is an .xlsx.
"""

import argparse
import csv
import datetime
import difflib
import hashlib
import http.server
import io
import json
import os
import re
import socketserver
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from collections import Counter
from pathlib import Path

PORT = 8766
SCRIPT_DIR = Path(__file__).resolve().parent
SPREADSHEET = (".xlsx", ".xlsm")

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_MODEL = "claude-sonnet-5-5"

# How much of the corpus a question gets to see. Both are adjustable in
# Settings for the session; change them here to alter the default.
#
# The whole corpus cannot be sent: 1,793 records come to roughly 645,000
# tokens against a 200,000-token context window. So a question is answered
# from the closest ASK_RECORDS records instead, which is also why a negative
# answer ("no study examines X") is never trustworthy — it means "not in
# these records", not "not in your corpus". Use the dashboard for counts.
#
# At 50 records and a 2,500-character cap this sends about 30,000 tokens per
# question. A 2,500-character cap passes nearly every abstract whole: the
# median in this corpus is 1,469 characters, and the old 1,200 cap was
# clipping 71% of them.
MAX_CONTEXT_RECORDS = 50
MAX_ABSTRACT_CHARS = 2500

THIS_YEAR = datetime.date.today().year

# Okabe-Ito: distinguishable with any common form of colour blindness.
PALETTE = ["#0072B2", "#E69F00", "#009E73", "#CC79A7",
           "#56B4E9", "#D55E00", "#F0E442", "#999999"]

DEFAULT_LABELS = [
    {"id": "t1", "name": "Theme 1", "colour": "#0072B2"},
    {"id": "t2", "name": "Theme 2", "colour": "#E69F00"},
    {"id": "t3", "name": "Theme 3", "colour": "#009E73"},
    {"id": "t4", "name": "Theme 4", "colour": "#CC79A7"},
]

DEFAULT_CRITERIA = {
    "include": [
        "Matches the population, intervention or phenomenon of interest",
        "Study design is eligible",
        "Addresses at least one of the research themes",
        "Peer-reviewed journal article, review, or book chapter",
        "Within the publication window",
        "Written in an accepted language",
    ],
    "exclude": [
        "Off topic",
        "No relevance to the research themes",
        "Wrong setting or population",
        "Wrong publication type (editorial, abstract, book review)",
        "Not in an accepted language",
        "Duplicate of another record",
        "Full text not retrievable",
        "Other (see notes)",
    ],
}

# Each rule excludes records automatically. `enabled` is the only thing most
# of them need; `params` holds anything adjustable. Rules that encode a
# methodological judgement start switched OFF, so you turn them on knowingly
# after seeing the count they would remove.
#
# `keywords` rules are yours to define: give a reason and a list of words,
# and any record whose title, abstract or keywords match none of them is
# excluded with that reason. That is how a review states its own scope
# without this file having to know anything about the subject.
DEFAULT_RULES = [
    {"id": "retracted", "name": "Retracted publications",
     "reason": "Retracted", "enabled": True, "params": {},
     "note": "Almost always right to exclude."},
    {"id": "lang", "name": "Not in an accepted language",
     "reason": "Not in an accepted language", "enabled": True,
     "params": {"allow": ["en"], "keep_unknown": True},
     "note": "Records whose language the source does not know are kept by "
             "default, because unknown is not the same as not-English."},
    {"id": "wrongtype", "name": "Wrong publication type",
     "reason": "Wrong publication type", "enabled": True,
     "params": {"block": ["paratext", "editorial", "erratum", "letter",
                          "note", "other"]},
     "note": "Editorials, errata and front matter."},
    {"id": "noabstract", "name": "No abstract available",
     "reason": "No abstract to screen", "enabled": False, "params": {},
     "note": "Off by default. These can only be screened on the title, but "
             "some are real papers whose publisher withholds abstracts — "
             "Elsevier especially. Read the titles before switching this on."},
    {"id": "old_uncited", "name": "Old and never cited",
     "reason": "5+ years old with no citations", "enabled": False,
     "params": {"years": 5},
     "note": "Off by default. After this long, zero citations is no longer "
             "explained by the paper being recent, so it is a real signal — "
             "but it still removes work that is simply in a quiet venue."},
    {"id": "region", "name": "Outside the region of interest",
     "reason": "Outside the region of interest", "enabled": False,
     "params": {"regions": [], "basis": "both", "keep_unknown": True},
     "note": "Off by default, and only does anything once you name regions "
             "in the config. `basis` is text, affiliation or both; "
             "keep_unknown keeps records whose affiliation is unrecorded, "
             "because unknown is not the same as elsewhere."},
]

DECISION_FIELDS = ["doi", "title", "ta_decision", "ta_reason", "ta_source",
                   "ft_decision", "ft_reason", "labels", "country",
                   "instrument_type", "notes", "abstract_manual",
                   # Corrections you type over what the source supplied. Kept
                   # apart from the record itself so the harvest and every
                   # import stay exactly as they arrived, and so a re-run can
                   # never silently drop a fix you made.
                   "doi_manual", "title_manual", "authors_manual",
                   "year_manual", "journal_manual",
                   "dup_status", "screener", "updated_at"]

# field on the record  ->  the decision column that overrides it
MANUAL_OVERRIDES = {"abstract": "abstract_manual", "doi": "doi_manual",
                    "title": "title_manual", "authors": "authors_manual",
                    "year": "year_manual", "journal": "journal_manual"}

STATE = {
    "records": [], "decisions": {}, "labels": [], "criteria": {},
    "rules": [], "dupes": [], "paths": {}, "screener": "", "source": "",
    "api_key": os.environ.get("ANTHROPIC_API_KEY", ""),
    "model": DEFAULT_MODEL,
    "ask_k": MAX_CONTEXT_RECORDS, "ask_chars": MAX_ABSTRACT_CHARS,
    # Which model orders the screening queue. One of ASREVIEW_MODELS, or
    # "builtin" to use the dependency-free ranker even when ASReview is here.
    "ranker": "nb",
    "stop": {}, "stop_cfg": {}, "extra": [], "imports": [],
}
LOCK = threading.Lock()


# ==========================================================================
#  reading files
# ==========================================================================

def sniff_delimiter(path):
    try:
        with open(path, encoding="utf-8-sig") as fh:
            header = fh.readline()
    except Exception:
        return ","
    counts = {d: header.count(d) for d in (";", "\t", ",")}
    best = max(counts, key=lambda d: counts[d])
    return best if counts[best] else ","


def read_header(path):
    p = Path(path)
    if p.suffix.lower() in SPREADSHEET:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(p, read_only=True)
            hdr = [("" if h is None else str(h))
                   for h in next(wb[wb.sheetnames[0]].iter_rows(values_only=True), [])]
            wb.close()
            return hdr
        except Exception:
            return []
    try:
        with open(p, newline="", encoding="utf-8-sig") as fh:
            return next(csv.reader(fh, delimiter=sniff_delimiter(p)), [])
    except Exception:
        return []


def read_table(path):
    p = Path(path)
    if p.suffix.lower() in SPREADSHEET:
        try:
            import openpyxl
        except ImportError:
            raise SystemExit(f"\n{p.name} is a spreadsheet; reading one needs:"
                             "\n    pip install openpyxl")
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        it = wb[wb.sheetnames[0]].iter_rows(values_only=True)
        hdr = [("" if h is None else str(h).strip()) for h in next(it, [])]
        out = []
        for row in it:
            if row is None or all(v in (None, "") for v in row):
                continue
            out.append({hdr[i]: ("" if i >= len(row) or row[i] is None
                                 else str(row[i]))
                        for i in range(len(hdr)) if hdr[i]})
        wb.close()
        return out
    with open(p, newline="", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh, delimiter=sniff_delimiter(p)))


def row_count(path):
    p = Path(path)
    if p.suffix.lower() in SPREADSHEET:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(p, read_only=True)
            n = wb[wb.sheetnames[0]].max_row or 1
            wb.close()
            return max(0, n - 1)
        except Exception:
            return 0
    try:
        with open(p, encoding="utf-8-sig") as fh:
            return max(0, sum(1 for _ in fh) - 1)
    except Exception:
        return 0


def looks_like_corpus(path):
    h = {str(c).strip().lower() for c in read_header(path)}
    if not h:
        return False
    if {"item type", "date added", "abstract note", "manual tags"} & h:
        return False          # a Zotero export
    if "ta_decision" in h or "dup_status" in h:
        return False          # a decisions file this app wrote
    if {"source_file", "added_at"} <= h:
        return False          # an _extra_records.csv this app wrote
    # A list you made by hand — a gold set, the references of a paper. It is
    # something to IMPORT, not a corpus to open. Its name can easily look like
    # a corpus: 04_goldset_from_review_complete.xlsx ends in _complete, same
    # as a step-2 output, and being the newest file it would otherwise win.
    if "status" in h and {"doi_source", "abstract_source", "match_note"} & h:
        return False
    if {"lens id", "ut (unique id)", "eid"} & h:
        return False          # a database export, also for importing
    if not {"abstract", "journal", "authors"} & h:
        return False
    return "title" in h and bool({"doi", "openalex_id", "id"} & h)


def work_done_against(path):
    """How many decisions beside this corpus look like screening you did.

    Decisions written many-at-once are not counted. A bulk action can put
    thousands of rows in a decisions file in one second, and counting those
    as screening makes the biggest batch win: a file holding 1,793 records
    marked in one go outranked the corpus with 814 judgements actually made
    on it, so the app opened the wrong one.

    A run of 25 or more decisions sharing a single second, none of them
    carrying a reason, is a batch. Everything else counts."""
    dec = Path(str(Path(path).with_suffix("")) + "_decisions.csv")
    if not dec.exists():
        return 0
    try:
        with open(dec, newline="", encoding="utf-8-sig") as fh:
            rows = [r for r in csv.DictReader(
                        fh, delimiter=sniff_delimiter(dec))
                    if (r.get("ta_decision") or "").strip()
                    or (r.get("ft_decision") or "").strip()]
    except Exception:
        return 0
    by_second = {}
    for r in rows:
        by_second.setdefault((r.get("updated_at") or "")[:19], []).append(r)
    real = 0
    for sec, group in by_second.items():
        if len(group) >= BULK_MIN and sec:
            unexplained = sum(1 for r in group
                              if not (r.get("ta_reason") or "").strip()
                              and not (r.get("ft_reason") or "").strip())
            if unexplained * 2 >= len(group):
                continue              # a batch, not screening
        real += len(group)
    return real


CONFIG_FILE = "review_config.json"


def load_config():
    """Everything about YOUR review, in one editable file beside the app.

    Themes, inclusion and exclusion criteria, the rules, the keyword patterns
    that guess a country or a theme, your name and your email. Nothing about
    any particular subject is written into this program: it is written here,
    by you, and a review of anything else needs only a different config.

    The file is created with neutral defaults the first time the app runs, so
    there is always something to edit rather than a blank page. Values you
    change in Settings are written back to it."""
    path = SCRIPT_DIR / CONFIG_FILE
    cfg = {}
    if path.exists():
        try:
            cfg = json.loads(path.read_text(encoding="utf-8-sig")) or {}
        except (OSError, ValueError) as e:
            print(f"  could not read {CONFIG_FILE} ({e}); using defaults",
                  file=sys.stderr)
            cfg = {}
    base = {
        "screener": "",
        "mailto": "",
        "themes": DEFAULT_LABELS,
        "criteria": DEFAULT_CRITERIA,
        "rules": DEFAULT_RULES,
        # name -> regex. Used to guess what a record is about. Yours to
        # replace: these decide the Country and Theme columns the app fills
        # in before you confirm them.
        "country_patterns": {},
        "theme_patterns": {},
        "stopping": {"consecutive": SAFE_CONSECUTIVE,
                     "min_percent": SAFE_MIN_PERCENT},
    }
    for k, v in base.items():
        cfg.setdefault(k, v)
    if not path.exists():
        try:
            save_config(cfg)
            print(f"  wrote {CONFIG_FILE} — your themes, criteria and rules "
                  f"live there")
        except OSError:
            pass
    return cfg


def save_config(cfg):
    path = SCRIPT_DIR / CONFIG_FILE
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def config_patterns(cfg):
    """User patterns if given, otherwise the built-in ones.

    Returns (country_list, theme_dict) in the shape the guessers expect."""
    cp = cfg.get("country_patterns") or {}
    tp = cfg.get("theme_patterns") or {}
    countries = ([(k, v) for k, v in cp.items()] if cp else COUNTRY_PATTERNS)
    themes = tp if tp else THEME_PATTERNS
    return countries, themes


ACTIVE_FILE = ".review_corpus.json"


def remember_corpus(path):
    """Note which corpus this is, so the next run opens the same one.

    Naming a corpus on the command line makes it the active one. Guessing
    from the folder is a fallback for the first run only — it has picked
    wrong twice, and a remembered choice cannot."""
    try:
        (SCRIPT_DIR / ACTIVE_FILE).write_text(json.dumps(
            {"corpus": str(path), "at": datetime.datetime.now().isoformat(
                timespec="seconds")}), encoding="utf-8")
    except OSError:
        pass


def remembered_corpus():
    p = SCRIPT_DIR / ACTIVE_FILE
    if not p.exists():
        return None
    try:
        q = Path(json.loads(p.read_text(encoding="utf-8")).get("corpus") or "")
    except (OSError, ValueError):
        return None
    return q if q.exists() else None


def candidate_tables(folder):
    out = []
    for pat in ("*.csv", "*.xlsx", "*.xlsm"):
        out += [p for p in Path(folder).glob(pat) if not p.name.startswith("~$")]
    return out


def find_corpus(explicit=None):
    me = Path(sys.argv[0]).name
    if explicit:
        p = Path(explicit).expanduser()
        if not p.is_absolute():
            beside = SCRIPT_DIR / p
            p = beside if beside.exists() else p.resolve()
        if not p.exists():
            raise SystemExit(f"\nCannot find {explicit}")
        remember_corpus(p)          # this is the corpus from now on
        return p
    # A corpus you named before is the corpus you meant, whatever else has
    # since appeared in the folder.
    kept = remembered_corpus()
    if kept:
        print(f"Using {kept.name}  (the corpus you opened last)")
        print(f"  To switch:  python \"{me}\" \"another_file.xlsx\"\n")
        return kept
    cands = [p for p in candidate_tables(SCRIPT_DIR) if looks_like_corpus(p)]
    if not cands:
        raise SystemExit(
            f"\nNo corpus found in:\n  {SCRIPT_DIR}\n\n"
            f"  Merge your database exports into one corpus first:\n"
            f"    python 10-merge_sources.py\n\n"
            f'  Or name a file:  python "{me}" "my_corpus.xlsx"')
    # Work already done decides it, then size. Sorting on the date alone is
    # what opened a 36-row gold set instead of a 1,082-record corpus with 446
    # records screened against it; ranking on the name alone then opened that
    # 1,080-record corpus in preference to a 1,887-record merged one holding
    # the same decisions, because its name happened to end in "_complete".
    # Among files you have screened equally, the one holding more of the
    # review is the one you meant.
    scored = [(work_done_against(p), row_count(p),
               p.stem.endswith("_complete"), p.stat().st_mtime, p)
              for p in cands]
    scored.sort(key=lambda t: t[:4], reverse=True)
    best = scored[0][4]
    if len(cands) > 1:
        print("Corpus files here:")
        for done, rows, _, mt, p in scored[:6]:
            when = datetime.datetime.fromtimestamp(mt)
            mark = "  <- using this" if p == best else ""
            work = f"{done:,} screened" if done else "nothing screened yet"
            print(f"   {p.name:46s} ~{rows:5,} rows  "
                  f"{when:%d %b %H:%M}  {work}{mark}")
        if scored[0][0]:
            print(f"\n  Opening the one holding most of your review.")
        else:
            print(f"\n  Nothing has been screened yet in any of these, so the "
                  f"largest was picked.")
        other = next((t[4] for t in scored if t[4] != best), best)
        print(f"  To open a different one, name it once and it will be "
              f"remembered:\n    python \"{me}\" \"{other.name}\"")
        print()
    else:
        print(f"Using {best.name}")
    remember_corpus(best)
    return best


# ==========================================================================
#  normalising
# ==========================================================================

# Sent to Crossref and friends so their polite pools know who is calling.
# Set it in the config file or in Settings; the placeholder is never used for
# a real request.
DEFAULT_MAILTO = ""
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def clean(v):
    """One line, no control characters, no runs of whitespace.

    Exports arrive with newlines inside abstracts and stray control bytes;
    everything downstream — the CSV, the .xlsx, the browser — is happier
    without them."""
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", _CTRL.sub(" ", s)).strip()


def strip_doi(s):
    """Drop the resolver prefix, keep the DOI's own capitalisation."""
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", str(s or "").strip(),
                  flags=re.I).strip()


def norm_doi(s):
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", str(s or "").strip(),
                  flags=re.I).strip().lower()


def norm_title(s):
    return re.sub(r"\s+", " ",
                  re.sub(r"[^a-z0-9 ]+", " ", str(s or "").lower())).strip()


def pick(r, *names):
    for n in names:
        if r.get(n):
            return str(r[n]).strip()
    return ""


def as_int(v, default=0):
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return default


def record_key(source, doi, title):
    """A stable id for an imported row, derived from what the row says.

    Ids used to be positions in the import list — WOS1, WOS2 — which meant
    that re-importing a file, or importing a second one before it, renumbered
    records that already carried screening decisions. A decision is stored
    against an id, so every decision silently moved to a different paper.

    Deriving the id from the DOI, or from the title when there is no DOI,
    fixes that: the same row always gets the same id, whatever order the
    imports happen in and however many times a file is imported."""
    basis = norm_doi(doi) or norm_title(title)
    if not basis:
        return ""
    h = hashlib.sha1(basis.encode("utf-8")).hexdigest()[:10]
    return f"{(source or 'x').upper().replace('-', '')}-{h}"


# ==========================================================================
#  guessing country and theme
# ==========================================================================

COUNTRY_PATTERNS = [
    ("Netherlands", r"netherlands|dutch|holland"), ("Germany", r"german(y|)"),
    ("France", r"\bfrance\b|french"), ("Italy", r"ital(y|ian)"),
    ("Spain", r"spain|spanish"), ("Portugal", r"portug(al|uese)"),
    ("Belgium", r"belgi(um|an)"), ("Austria", r"austria(n|)"),
    ("Switzerland", r"switzerland|swiss"), ("Sweden", r"swed(en|ish)"),
    ("Norway", r"norw(ay|egian)"), ("Denmark", r"denmark|danish"),
    ("Finland", r"finland|finnish"), ("Poland", r"poland|polish"),
    ("Czechia", r"czech"), ("Hungary", r"hungar(y|ian)"),
    ("Greece", r"greece|greek"), ("Ireland", r"ireland|irish"),
    ("Estonia", r"estonia"), ("Latvia", r"latvia"), ("Lithuania", r"lithuania"),
    ("Slovenia", r"slovenia"), ("Croatia", r"croatia"),
    ("Luxembourg", r"luxembourg"), ("Slovakia", r"slovak"),
    ("Bulgaria", r"bulgaria"), ("Romania", r"romania"),
    ("United Kingdom", r"united kingdom|\buk\b|britain|british|england|"
                       r"scotland|wales"),
    ("China", r"\bchina\b|chinese"), ("United States", r"united states|\busa\b|\bu\.s\.|american"),
    ("India", r"\bindia\b|indian"), ("Indonesia", r"indonesia"),
    ("Japan", r"\bjapan\b|japanese"), ("South Korea", r"south korea|korean"),
    ("Brazil", r"brazil"), ("Canada", r"canada|canadian"),
    ("Australia", r"australia"), ("Turkey", r"turkey|turkish"),
    ("Malaysia", r"malaysia"), ("Singapore", r"singapore"),
    ("Taiwan", r"taiwan"), ("Vietnam", r"vietnam"), ("Thailand", r"thailand"),
    ("Africa", r"\bafrica"), ("Latin America", r"latin america"),
    ("Europe (general)", r"\beurope|european union|\beurozone"),
]
ISO_TO_NAME = {
    "NL": "Netherlands", "DE": "Germany", "FR": "France", "IT": "Italy",
    "ES": "Spain", "PT": "Portugal", "BE": "Belgium", "AT": "Austria",
    "CH": "Switzerland", "SE": "Sweden", "NO": "Norway", "DK": "Denmark",
    "FI": "Finland", "PL": "Poland", "CZ": "Czechia", "HU": "Hungary",
    "GR": "Greece", "IE": "Ireland", "EE": "Estonia", "LV": "Latvia",
    "LT": "Lithuania", "SI": "Slovenia", "HR": "Croatia", "LU": "Luxembourg",
    "SK": "Slovakia", "BG": "Bulgaria", "RO": "Romania", "GB": "United Kingdom",
    "CN": "China", "US": "United States", "IN": "India", "ID": "Indonesia",
    "JP": "Japan", "KR": "South Korea", "BR": "Brazil", "CA": "Canada",
    "AU": "Australia", "TR": "Turkey", "MY": "Malaysia", "SG": "Singapore",
    "TW": "Taiwan", "VN": "Vietnam", "TH": "Thailand", "IS": "Iceland",
    "CY": "Cyprus", "MT": "Malta", "ZA": "South Africa", "IL": "Israel",
}
EU_ISO = {"NL", "DE", "FR", "IT", "ES", "BE", "AT", "SE", "NO", "DK", "FI",
          "PL", "PT", "IE", "GB", "CZ", "HU", "RO", "GR", "EE", "LV", "LT",
          "SI", "HR", "LU", "SK", "BG", "CH", "IS", "CY", "MT"}

THEME_PATTERNS = {
    "t1": r"investor motivation|investor behaviou?r|lender behaviou?r|warm glow|"
          r"prosocial|pro-social|altruis|human values|willingness to invest|"
          r"investor profile|investor confidence|investment decision|motive",
    "t2": r"signal|information asymmetr|quality signal|certification|"
          r"third.party|campaign success|funding success|rhetorical|"
          r"substantive|credit risk|risk of failure|overfund|goal attainment",
    "t3": r"impact measure|impact report|sustainability report|"
          r"non.financial report|impact metric|greenwash|disclos|\besg\b",
    "t4": r"european crowdfunding service|crowdfunding regulation|\becsp\b|"
          r"\becspr\b|\bsfdr\b|harmoniz|harmonis|regulat",
}


def guess_country(rec):
    """The country STUDIED, which the text reveals and affiliation does not.

    An author in Utrecht may be studying Chinese platforms, so text wins and
    affiliation is only a fallback. The basis is returned so you can see
    which it was."""
    hay = f"{rec['title']} {rec['abstract']}".lower()
    pats = STATE.get("country_patterns") or COUNTRY_PATTERNS
    hits = [name for name, pat in pats if re.search(pat, hay)]
    specific = [h for h in hits if h != "Europe (general)"]
    if specific:
        return " | ".join(specific[:3]), "text"
    if hits:
        return hits[0], "text"
    codes = [c.strip().upper() for c in (rec.get("countries") or "").split("|")
             if c.strip()]
    named = [ISO_TO_NAME.get(c, c) for c in dict.fromkeys(codes)]
    if named:
        return " | ".join(named[:3]), "affiliation"
    return "", ""


def guess_themes(rec):
    hay = f"{rec['title']} {rec['abstract']} {rec.get('keywords', '')}".lower()
    pats = STATE.get("theme_patterns") or THEME_PATTERNS
    out = []
    for tid, pat in pats.items():
        try:
            if re.search(pat, hay):
                out.append(tid)
        except re.error:
            continue        # a pattern you are still editing must not crash
    return out


# ==========================================================================
#  loading the corpus
# ==========================================================================

# ==========================================================================
#  Other sources: Web of Science, Lens, Scopus, and papers added by hand
#
#  Records from these live in <corpus>_extra_records.csv, never in the
#  harvest file. The harvest stays exactly as the search returned it, which
#  is what makes it reproducible; re-running step 1 therefore cannot lose an
#  imported or hand-added record. They are merged in memory at startup and
#  are records like any other from that point on: deduplicated against the
#  harvest, subject to the rules, screened, counted and exported.
#
#  Every record carries `source`. PRISMA 2020 separates studies found
#  through databases from studies found by other methods, and that column is
#  what tells the two apart.
# ==========================================================================

SOURCE_LABELS = {
    "openalex": "OpenAlex", "wos": "Web of Science", "lens": "Lens",
    "scopus": "Scopus", "other-db": "Other database",
    "citation": "Citation searching", "manual": "Added by hand",
    "reflist": "Added by hand (known papers)",
}
OTHER_METHODS = {"citation", "manual", "reflist"}   # PRISMA's second identification arm

# What each export calls things. First name that appears in the file wins.
SOURCE_PROFILES = {
    "wos": {
        "label": "Web of Science",
        "detect": ["UT (Unique ID)", "Times Cited, WoS Core"],
        "map": {"title": ["Article Title"], "authors": ["Authors"],
                "year": ["Publication Year"], "doi": ["DOI"],
                "abstract": ["Abstract"], "journal": ["Source Title"],
                "cited_by": ["Times Cited, WoS Core", "Times Cited, All Databases"],
                "type": ["Document Type"], "language": ["Language"],
                "keywords": ["Author Keywords", "Keywords Plus"],
                "issn": ["ISSN", "eISSN"]},
    },
    "lens": {
        "label": "Lens",
        "detect": ["Lens ID", "Citing Works Count"],
        "map": {"title": ["Title"], "authors": ["Author/s"],
                "year": ["Publication Year"], "doi": ["DOI"],
                "abstract": ["Abstract"], "journal": ["Source Title"],
                "cited_by": ["Citing Works Count"],
                "type": ["Publication Type"],
                "keywords": ["Keywords", "Fields of Study"],
                "issn": ["ISSNs"]},
    },
    "scopus": {
        "label": "Scopus",
        "detect": ["EID", "Source title"],
        "map": {"title": ["Title"], "authors": ["Authors", "Author full names"],
                "year": ["Year"], "doi": ["DOI"], "abstract": ["Abstract"],
                "journal": ["Source title"], "cited_by": ["Cited by"],
                "type": ["Document Type"], "language": ["Language of Original Document"],
                "keywords": ["Author Keywords", "Index Keywords"],
                "issn": ["ISSN"]},
    },
    "reflist": {
        # A list you made yourself — a gold set, or the references of your own
        # paper. Not a database hit, so PRISMA counts these under the second
        # identification arm, "other methods".
        #
        # `status` carries core/background through to the record, which is
        # what marks a key paper for the stopping rule. `country` is a value
        # you established yourself, so it arrives confirmed rather than as a
        # guess the app made.
        "label": "Added by hand (known papers)",
        "detect": ["status", "title", "authors"],
        "map": {"title": ["title", "Title"], "authors": ["authors", "Authors"],
                "year": ["year", "Year"], "doi": ["doi", "DOI"],
                "abstract": ["abstract", "Abstract"],
                "journal": ["journal", "Journal", "Source Title"],
                "cited_by": ["cited_by", "citations", "cited_by_count"],
                "type": ["type", "publication_type", "Type"],
                "language": ["language"],
                "keywords": ["keywords"], "issn": ["issn"],
                "key_paper": ["status", "key_paper", "gold"],
                "country": ["country", "Country"],
                "note": ["match_note", "note", "notes"]},
    },
    "other-db": {
        "label": "Other database",
        "detect": [],
        "map": {"title": ["title", "Title", "Article Title", "primary_title"],
                "authors": ["authors", "Authors", "Author/s", "author names"],
                "year": ["year", "Year", "Publication Year", "publication_year"],
                "doi": ["doi", "DOI"], "abstract": ["abstract", "Abstract"],
                "journal": ["journal", "Journal", "Source Title", "source title"],
                "cited_by": ["cited_by_count", "cited_by", "Cited by", "citations"],
                "type": ["type", "Document Type", "Publication Type"],
                "language": ["language", "Language"],
                "keywords": ["keywords", "Keywords", "Author Keywords"],
                "issn": ["issn", "ISSN", "ISSNs"]},
    },
}


def read_any_table(path):
    """Read .csv/.tsv/.xlsx — and .xls, which Web of Science still exports."""
    p = Path(path)
    if p.suffix.lower() == ".xls":
        try:
            import xlrd
        except ImportError:
            raise RuntimeError(
                f"{p.name} is an old-format Excel file, which Web of Science "
                "still exports. Reading one needs:\n    pip install xlrd\n"
                "Or in Web of Science, export as 'Tab delimited' or Excel "
                "(.xlsx) instead.")
        wb = xlrd.open_workbook(str(p))
        sh = wb.sheet_by_index(0)
        hdr = [str(sh.cell_value(0, c)).strip() for c in range(sh.ncols)]
        out = []
        for r in range(1, sh.nrows):
            row = {}
            for c in range(sh.ncols):
                if not hdr[c]:
                    continue
                v = sh.cell_value(r, c)
                # xlrd turns every number into a float, so 2021 arrives as
                # "2021.0" and a citation count as "48.0".
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                row[hdr[c]] = str(v).strip()
            out.append(row)
        return out
    return read_table(p)


def detect_source(rows):
    if not rows:
        return "other-db"
    cols = set(rows[0])
    for key, prof in SOURCE_PROFILES.items():
        if prof["detect"] and all(d in cols for d in prof["detect"]):
            return key
    return "other-db"


def map_source_rows(rows, source):
    """Turn an export into records shaped like the rest of the corpus."""
    prof = SOURCE_PROFILES.get(source, SOURCE_PROFILES["other-db"])
    cols = list(rows[0]) if rows else []
    lower = {c.strip().lower(): c for c in cols}

    def find(names):
        for n in names:
            if n in cols:
                return n
            if n.strip().lower() in lower:
                return lower[n.strip().lower()]
        return None

    picked = {f: find(names) for f, names in prof["map"].items()}
    out = []
    for r in rows:
        g = lambda f: clean(r.get(picked[f])) if picked.get(f) else ""
        title = g("title")
        if not title and not g("doi"):
            continue                      # nothing to identify it by
        yr = re.sub(r"\.0$", "", g("year"))
        yr = (re.search(r"(19|20)\d{2}", yr) or [""])[0] if yr else ""
        out.append({
            "doi": strip_doi(g("doi")), "title": title, "authors": g("authors"),
            "year": yr, "journal": g("journal"), "abstract": g("abstract"),
            "keywords": g("keywords"), "language": g("language"),
            "type": g("type"), "issn": g("issn"),
            "cited_by": as_int(re.sub(r"\.0$", "", g("cited_by"))),
            # Only a hand-made list declares these; for a database export
            # they come back empty, which is correct.
            "key_paper": g("key_paper"), "country": g("country"),
            "note": g("note"),
        })
    return out, {f: picked.get(f) for f in prof["map"]}


EXTRA_COLUMNS = ["id", "source", "source_file", "added_at", "doi", "title",
                 "authors", "year", "journal", "abstract", "keywords",
                 "language", "type", "issn", "cited_by", "note",
                 # core / background, from a hand-made list's status column.
                 # This is what marks a key paper for the stopping rule.
                 "key_paper",
                 # a country you established yourself, not one the app guessed
                 "country"]


def load_extra(path):
    if not os.path.exists(path):
        return []
    try:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = [dict(r) for r in csv.DictReader(
                        fh, delimiter=sniff_delimiter(path))]
    except Exception as e:
        print(f"  could not read {os.path.basename(path)}: {e}", file=sys.stderr)
        return []
    return migrate_extra(rows, path)


def migrate_extra(rows, path):
    """Bring a file written by an earlier version up to the current id scheme.

    Rows used to be numbered by position — WOS1, WOS2 — and are now keyed by
    their own DOI or title (see record_key). Without this, importing the same
    export again would not recognise a single row and would store all of them
    a second time: 126 Web of Science records becoming 252.

    Decisions follow the renamed records automatically, because
    reattach_decisions re-places them by DOI and title.

    Rows that turn out to be the same paper from the same source collapse
    into one; the fuller of the two wins, so nothing you had is lost."""
    changed = False
    best = {}
    order = []
    for r in rows:
        src = r.get("source") or "other-db"
        key = record_key(src, r.get("doi"), r.get("title"))
        if not key:
            key = r.get("id") or f"X{len(order)}"
        if key != (r.get("id") or ""):
            changed = True
        r["id"] = key
        if key in best:
            changed = True
            a, b = best[key], r
            # Keep whichever row carries more: abstract first, then the rest.
            score = lambda x: (bool(clean(x.get("abstract"))),
                               sum(1 for f in ("doi", "authors", "journal",
                                               "year", "keywords", "key_paper")
                                   if clean(x.get(f))))
            keep, drop = (a, b) if score(a) >= score(b) else (b, a)
            for f in EXTRA_COLUMNS:
                if not clean(keep.get(f)) and clean(drop.get(f)):
                    keep[f] = drop[f]
            best[key] = keep
        else:
            best[key] = r
            order.append(key)
    out = [best[k] for k in order]
    if changed:
        print(f"  updated {os.path.basename(path)} to the current id scheme"
              + (f" ({len(rows) - len(out)} repeated rows merged)"
                 if len(out) < len(rows) else ""))
        try:
            save_extra(path, out)
        except RuntimeError as e:
            print(f"  {e}", file=sys.stderr)
    return out


def save_extra(path, rows):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=EXTRA_COLUMNS, delimiter=";",
                           quoting=csv.QUOTE_ALL, lineterminator="\r\n",
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    try:
        os.replace(tmp, path)
    except PermissionError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise RuntimeError(
            f"Could not save: {os.path.basename(path)} is open in another "
            "program, most likely Excel. Close it and try again.")


def extra_to_record(e, idx):
    """An imported or hand-added row, shaped like a harvested record."""
    rec = {
        "id": e.get("id") or f"X{idx}",
        "idx": 100000 + idx,
        "title": clean(e.get("title")), "authors": clean(e.get("authors")),
        "year": clean(e.get("year")), "journal": clean(e.get("journal")),
        "doi": strip_doi(e.get("doi")), "abstract": clean(e.get("abstract")),
        "keywords": clean(e.get("keywords")), "language": clean(e.get("language")),
        "type": clean(e.get("type")), "cited_by": as_int(e.get("cited_by")),
        "retrieved": clean(e.get("added_at")), "institutions": "",
        "topic": "", "subfield": "", "field": "", "countries": "",
        "europe": "", "europe_basis": "", "oa_status": "", "url": "",
        "apa": "", "abstract_source": "", "retracted": False,
        "source": e.get("source") or "other-db",
        "source_file": clean(e.get("source_file")),
        "note": clean(e.get("note")),
        "key_paper": clean(e.get("key_paper")).lower(),
        "from_corpus": False,
    }
    rec["ndoi"] = norm_doi(rec["doi"])
    rec["ntitle"] = norm_title(rec["title"])
    rec["abstract_original"] = rec["abstract"]
    rec["abstract_added"] = False
    c, basis = guess_country(rec)
    given = clean(e.get("country"))
    if given:
        # You supplied this one, so it is not a guess and gets no guess badge.
        rec["country_guess"], rec["country_basis"] = given, "from your file"
    else:
        rec["country_guess"], rec["country_basis"] = c, basis
    rec["theme_guess"] = guess_themes(rec)
    rec["completeness"] = completeness(rec)
    return rec


def crossref_lookup(doi):
    """Fetch a record by DOI so adding a known paper is one field, not ten."""
    d = norm_doi(doi)
    if not d:
        raise RuntimeError("Give a DOI, for example 10.1016/j.jbusvent.2013.06.005")
    mail = STATE.get("mailto") or DEFAULT_MAILTO
    url = f"https://api.crossref.org/works/{urllib.parse.quote(d)}"
    if mail:
        url += f"?mailto={urllib.parse.quote(mail)}"
    req = urllib.request.Request(
        url, headers={"User-Agent": f"BarcalaReviewApp/1.0"
                                    + (f" (mailto:{mail})" if mail else "")})
    try:
        with urllib.request.urlopen(req, timeout=40) as r:
            msg = json.loads(r.read().decode("utf-8")).get("message") or {}
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise RuntimeError(f"Crossref has no record for {d}. Check the DOI, "
                               "or add the paper by typing the fields.")
        raise RuntimeError(f"Crossref returned {e.code}. Try again, or type "
                           "the fields instead.")
    except Exception as e:
        raise RuntimeError(f"Could not reach Crossref ({e}). Type the fields "
                           "instead.")
    au = []
    for a in (msg.get("author") or []):
        nm = " ".join(x for x in [a.get("given"), a.get("family")] if x) \
             or a.get("name") or ""
        if nm:
            au.append(clean(nm))
    dp = ((msg.get("issued") or {}).get("date-parts") or [[None]])[0]
    ct = msg.get("container-title") or []
    return {
        "doi": clean(msg.get("DOI")) or d,
        "title": clean((msg.get("title") or [""])[0]),
        "authors": " | ".join(au), "year": str(dp[0]) if dp and dp[0] else "",
        "journal": clean(ct[0]) if ct else "",
        "abstract": strip_jats_text(msg.get("abstract")),
        "type": clean(msg.get("type")), "issn": clean((msg.get("ISSN") or [""])[0]),
        "language": clean(msg.get("language"))[:2].lower(),
        "cited_by": as_int(msg.get("is-referenced-by-count")),
        "keywords": " | ".join(clean(s) for s in (msg.get("subject") or [])),
    }


def strip_jats_text(x):
    if not x:
        return ""
    s = re.sub(r"<jats:title>.*?</jats:title>", " ", str(x), flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = html_unescape(s)
    return clean(re.sub(r"^\s*abstract\s*[:.\-]?\s*", "", s, flags=re.I))


def html_unescape(s):
    import html as _html
    return _html.unescape(s)


def rebuild_records():
    """Harvest + everything imported or added by hand, as one list.

    Called at startup and after any import, so the duplicate groups and every
    count are recomputed against the whole set rather than the harvest alone.

    The base is whatever came out of the corpus file, flagged at load time.
    Everything else is rebuilt from _extra_records.csv, the file holding every
    row imported or added by hand after that."""
    base = [r for r in STATE["records"] if r.get("from_corpus")]
    extra = [extra_to_record(e, i) for i, e in enumerate(STATE.get("extra") or [])]
    # One id can only mean one record. A row whose id already belongs to the
    # harvest would otherwise shadow it and take its decisions with it.
    taken = {r["id"] for r in base}
    kept = []
    for r in extra:
        if not r["id"] or r["id"] in taken:
            continue
        taken.add(r["id"])
        kept.append(r)
    extra = kept
    STATE["records"] = base + extra
    STATE["dupes"] = find_duplicates(STATE["records"])
    apply_manual_fields()
    merge_duplicates()
    return len(base), len(extra)


def source_counts():
    """Records per source, as PRISMA needs them reported.

    Every figure is counted off the stored records. Nothing is read from the
    import log, which is why re-importing a file changes nothing: the same
    row keeps the same id (see record_key) and is simply recognised as
    already held.

    `identified` is therefore what each database contributed and what the
    diagram reports, `after_dedup` is what survived deduplication, and the
    two differ by records another source also returned. The earlier version
    reconstructed `identified` by adding the import log on top of the stored
    count, which double-counted every re-import — the same 126-row export
    read twice was reported as 252 records identified."""
    dups, _ = dup_state()
    held, kept = {}, {}
    for r in STATE["records"]:
        s = r.get("source") or "openalex"
        held[s] = held.get(s, 0) + 1
        if dups.get(r["id"]) != "duplicate":
            kept[s] = kept.get(s, 0) + 1

    # Rows a file contained but the corpus does not hold: the same paper
    # listed twice in one export, and rows with neither title nor DOI. They
    # were identified by the search, so the diagram reports them, but they
    # are not records and cannot be screened.
    not_stored = {}
    for imp in (STATE.get("imports") or []):
        s = imp.get("source") or "other-db"
        n = as_int(imp.get("dupes_within_file")) + as_int(imp.get("unusable"))
        not_stored[s] = not_stored.get(s, 0) + n

    keys = set(held) | set(not_stored)
    out = []
    for k in sorted(keys, key=lambda x: -(held.get(x, 0) + not_stored.get(x, 0))):
        out.append({"key": k, "label": SOURCE_LABELS.get(k, k),
                    "identified": held.get(k, 0) + not_stored.get(k, 0),
                    "stored": held.get(k, 0),
                    "not_stored": not_stored.get(k, 0),
                    "after_dedup": kept.get(k, 0),
                    "other_method": k in OTHER_METHODS})
    return out


def load_corpus(path):
    raw = read_table(path)
    if not raw:
        raise SystemExit(f"{Path(path).name} has no rows.")
    out = []
    for i, r in enumerate(raw):
        rid = pick(r, "openalex_id", "id") or f"row{i}"
        rec = {
            "id": rid, "idx": i,
            "title": pick(r, "title"),
            "authors": pick(r, "authors"),
            "year": pick(r, "year", "publication_year"),
            # "source" means the venue in an OpenAlex harvest and the
            # DATABASE in a merged corpus, so falling back to it blindly
            # labelled 40 records as having been published in "openalex".
            # Fall back only when the value is not a database name.
            "journal": (pick(r, "journal", "venue", "source_title")
                        or (lambda v: "" if v.lower() in SOURCE_LABELS else v)(
                            pick(r, "source"))),
            "doi": pick(r, "doi"),
            "abstract": pick(r, "abstract"),
            "keywords": pick(r, "keywords"),
            "language": pick(r, "language"),
            "type": pick(r, "type"),
            "cited_by": as_int(pick(r, "cited_by_count", "cited_by", "citations")),
            "retrieved": pick(r, "retrieved"),
            "institutions": pick(r, "institutions"),
            "topic": pick(r, "topic"),
            "subfield": pick(r, "subfield"),
            "field": pick(r, "field"),
            "countries": pick(r, "author_countries"),
            "europe": pick(r, "europe", "europe_any"),
            "europe_basis": pick(r, "europe_basis"),
            "oa_status": pick(r, "oa_status"),
            "url": pick(r, "url", "landing_page_url"),
            "apa": pick(r, "apa_citation"),
            "abstract_source": pick(r, "abstract_source"),
            "retracted": str(pick(r, "is_retracted", "retracted")).lower() == "true",
            # Everything in the harvest file came from the search itself.
            # A merged corpus carries these itself: one file holding every
            # database, each row saying where it came from. They used to be
            # blanked here, which forced every non-OpenAlex record to live in
            # a second file and be re-joined on startup — the arrangement
            # that let the wrong corpus be opened and the imports go missing.
            "source": pick(r, "source") or "openalex",
            "source_file": pick(r, "source_file"),
            "note": pick(r, "note"),
            "key_paper": pick(r, "key_paper").lower(),
        }
        # Loaded from the corpus file, as opposed to imported afterwards.
        # rebuild_records tells the two apart by this, not by source_file,
        # which a merged corpus legitimately fills in.
        rec["from_corpus"] = True
        rec["ndoi"] = norm_doi(rec["doi"])
        rec["ntitle"] = norm_title(rec["title"])
        # Kept so that removing a hand-typed abstract restores the original
        # rather than leaving the record blank.
        rec["abstract_original"] = rec["abstract"]
        rec["abstract_added"] = False
        if not rec["doi"]:
            rec["doi"] = ""
        c, basis = guess_country(rec)
        given = pick(r, "country")
        if given:
            # Supplied in the file, so it is not a guess and gets no badge.
            rec["country_guess"], rec["country_basis"] = given, "from your file"
        else:
            rec["country_guess"], rec["country_basis"] = c, basis
        rec["theme_guess"] = guess_themes(rec)
        rec["completeness"] = completeness(rec)
        out.append(rec)
    return out


def completeness(r):
    """How complete a record is, for choosing which duplicate to keep.

    A DOI outweighs everything else because it is what every later lookup
    depends on; an abstract comes next because you cannot screen without
    one. Citations only break ties."""
    s = 0
    if r.get("doi"):
        s += 1000
    if r.get("abstract"):
        s += 500
    if r.get("authors"):
        s += 100
    if r.get("journal"):
        s += 80
    if r.get("year"):
        s += 50
    if r.get("topic"):
        s += 20
    if r.get("institutions"):
        s += 20
    if r.get("language"):
        s += 10
    if r.get("url"):
        s += 5
    return s + min(as_int(r.get("cited_by")), 999) / 1000.0


def completeness_detail(r):
    """What this version has, in words, for the Duplicates tab."""
    have = []
    for field, label in (("doi", "DOI"), ("abstract", "abstract"),
                         ("authors", "authors"), ("journal", "journal"),
                         ("year", "year"), ("topic", "topic"),
                         ("institutions", "institutions")):
        if r.get(field):
            have.append(label)
    return have


# ==========================================================================
#  duplicates
# ==========================================================================

def find_duplicates(records):
    groups, seen = [], set()
    by_doi = {}
    for r in records:
        if r["ndoi"]:
            by_doi.setdefault(r["ndoi"], []).append(r["id"])
    for k, ids in by_doi.items():
        if len(ids) > 1:
            groups.append({"key": f"doi:{k}", "basis": "identical DOI", "ids": ids})
            seen.update(ids)

    by_title = {}
    for r in records:
        if r["ntitle"] and r["id"] not in seen:
            by_title.setdefault(r["ntitle"], []).append(r["id"])
    for k, ids in by_title.items():
        if len(ids) > 1:
            groups.append({"key": f"t:{k[:40]}", "basis": "identical title",
                           "ids": ids})
            seen.update(ids)

    blocks = {}
    for r in records:
        if r["ntitle"] and r["id"] not in seen:
            blocks.setdefault(r["ntitle"][:32], []).append(r)
    for _, grp in blocks.items():
        if len(grp) < 2:
            continue
        used = set()
        for i, a in enumerate(grp):
            if a["id"] in used:
                continue
            near = [a["id"]]
            for b in grp[i + 1:]:
                if b["id"] in used:
                    continue
                if difflib.SequenceMatcher(None, a["ntitle"],
                                           b["ntitle"]).ratio() > 0.93:
                    near.append(b["id"])
                    used.add(b["id"])
            if len(near) > 1:
                used.add(a["id"])
                groups.append({"key": f"n:{a['ntitle'][:32]}",
                               "basis": "near-identical title", "ids": near})
                seen.update(near)
    return groups


def auto_keep(group, by_id):
    """Which record in a duplicate group to keep: the most complete one."""
    ids = [i for i in group["ids"] if i in by_id]
    if not ids:
        return ""
    return max(ids, key=lambda i: (by_id[i]["completeness"], i))


# Fields worth taking from a duplicate. Citations are deliberately absent:
# every database counts them over a different set of citing works, so the
# highest figure is not a better figure, it is a different question.
MERGE_FIELDS = ["abstract", "doi", "authors", "journal", "year", "language",
                "type", "keywords", "institutions", "topic", "subfield",
                "field", "countries", "europe", "url", "apa", "oa_status",
                "key_paper"]


def merge_duplicates():
    """Fill gaps in each kept record from the duplicates it stands for.

    Choosing the most complete version is not the same as keeping the most
    complete record. Web of Science may carry the abstract that OpenAlex is
    missing, while OpenAlex is the only source with the topic, subfield and
    field — so whichever one wins, something is thrown away.

    Nothing has to be. The kept record keeps everything it has and takes only
    what it lacks from the others, best-first. It is never overwritten, so
    your own corrections and the source's own values both survive, and
    `merged_from` records where each borrowed value came from.

    Idempotent: anything borrowed on the previous run is cleared first, so
    changing which record you keep cannot leave a stale borrowed value
    behind. Run it after apply_manual_fields, never before, or a hand-typed
    abstract would be reset underneath it."""
    by_id = {r["id"]: r for r in STATE["records"]}
    for r in STATE["records"]:
        for f in (r.get("merged_from") or {}):
            r[f] = ""
        r["merged_from"] = {}

    dups, chosen = dup_state()
    filled = 0
    for g in STATE["dupes"]:
        keep = chosen.get(g["key"])
        rec = by_id.get(keep)
        if not rec:
            continue
        others = [by_id[i] for i in g["ids"]
                  if i in by_id and i != keep]
        others.sort(key=lambda x: -x["completeness"])
        for f in MERGE_FIELDS:
            if str(rec.get(f) or "").strip():
                continue
            for o in others:
                v = str(o.get(f) or "").strip()
                if v:
                    rec[f] = v
                    rec["merged_from"][f] = SOURCE_LABELS.get(
                        o.get("source"), o.get("source", "another record"))
                    filled += 1
                    break
        rec["ndoi"] = norm_doi(rec.get("doi"))
        rec["ntitle"] = norm_title(rec.get("title"))
        if rec["merged_from"]:
            # A borrowed abstract changes what the guesses read from.
            c, basis = guess_country(rec)
            rec["country_guess"], rec["country_basis"] = c, basis
            rec["theme_guess"] = guess_themes(rec)
            rec["completeness"] = completeness(rec)
    return filled


def dup_state():
    """{record id: 'keep' | 'duplicate'} from the rules, with your overrides.

    A manual dup_status on a record always wins; otherwise the most complete
    record in each group is kept and the rest marked duplicate."""
    by_id = {r["id"]: r for r in STATE["records"]}
    out, chosen = {}, {}
    dupe_rule = next((x for x in STATE["rules"] if x["id"] == "dup"), None)
    for g in STATE["dupes"]:
        ids = [i for i in g["ids"] if i in by_id]
        if len(ids) < 2:
            continue
        manual_keep = [i for i in ids
                       if STATE["decisions"].get(i, {}).get("dup_status") == "keep"]
        keep = manual_keep[0] if manual_keep else auto_keep(g, by_id)
        chosen[g["key"]] = keep
        for i in ids:
            manual = STATE["decisions"].get(i, {}).get("dup_status")
            if manual in ("keep", "duplicate", "not a duplicate"):
                out[i] = manual
            else:
                out[i] = "keep" if i == keep else "duplicate"
    return out, chosen


# ==========================================================================
#  the rules engine
# ==========================================================================

def rule_hits(rec, rule):
    """Does this rule exclude this record? Pure function, no state."""
    rid, p = rule["id"], rule.get("params") or {}
    if rid == "retracted":
        return bool(rec.get("retracted"))
    if rid == "lang":
        lang = (rec.get("language") or "").lower()
        allow = [x.lower() for x in (p.get("allow") or ["en"])]
        if not lang:
            return not p.get("keep_unknown", True)
        return lang not in allow
    if rid == "wrongtype":
        t = (rec.get("type") or "").lower()
        return bool(t) and t in [x.lower() for x in (p.get("block") or [])]
    if rid == "noabstract":
        return not rec.get("abstract")
    if rid == "old_uncited":
        years = as_int(p.get("years"), 5)
        y = as_int(rec.get("year"))
        return bool(y) and (THIS_YEAR - y) >= years and as_int(rec.get("cited_by")) == 0
    if rid in ("region", "noneurope"):
        # Named regions, from your config. With none named the rule does
        # nothing, which is what a review that has no geographic scope wants.
        regions = [str(x).strip().lower()
                   for x in (p.get("regions") or []) if str(x).strip()]
        if not regions:
            # The old European-only behaviour, for a corpus that carries the
            # europe column a harvest wrote.
            if rec.get("europe") or rec.get("europe_affiliation"):
                return (rec.get("europe") or "").lower() != "yes"
            return False
        basis = (p.get("basis") or "both").lower()
        text = ""
        if basis in ("text", "both"):
            text += f" {rec.get('title','')} {rec.get('abstract','')}"
        if basis in ("affiliation", "both"):
            text += f" {rec.get('institutions','')} {rec.get('countries','')}"
        t = text.lower()
        if not t.strip():
            return not p.get("keep_unknown", True)
        return not any(re.search(r"\b" + re.escape(w) + r"\b", t)
                       for w in regions)
    if rid.startswith("kw:"):
        # A keyword rule you defined yourself: exclude anything matching none
        # of these words. This is how a review states its own scope.
        words = [str(x).strip().lower()
                 for x in (p.get("keywords") or []) if str(x).strip()]
        if not words:
            return False
        t = (f"{rec.get('title','')} {rec.get('abstract','')} "
             f"{rec.get('keywords','')}").lower()
        hit = any(re.search(r"\b" + re.escape(w).replace(r"\ ", r"\s+") + r"\b", t)
                  for w in words)
        return (not hit) if p.get("mode", "require") == "require" else hit
    return False


def auto_reason(rec, rules=None):
    """The first enabled rule that excludes this record, or ('', '')."""
    for rule in (rules if rules is not None else STATE["rules"]):
        if rule.get("enabled") and rule_hits(rec, rule):
            return rule["reason"], rule["id"]
    return "", ""


def rule_preview():
    """How many records each rule actually removes, enabled or not.

    Three things make this number honest rather than merely suggestive.
    Duplicates are out, because they never reach screening. Records you have
    decided yourself are out, because a manual decision always wins and no
    rule can touch them. And a record already taken by an earlier enabled
    rule is not counted again, so the numbers sum to the exclusion total
    instead of double-counting. A disabled rule is counted as if it were
    last, which is exactly what switching it on would do."""
    recs = STATE["records"]
    dups, _ = dup_state()
    live = [r for r in recs
            if dups.get(r["id"]) != "duplicate"
            and not STATE["decisions"].get(r["id"], {}).get("ta_decision")]
    out, claimed = {}, set()
    for rule in STATE["rules"]:
        n = 0
        for r in live:
            if r["id"] in claimed:
                continue
            if rule_hits(r, rule):
                n += 1
                if rule.get("enabled"):
                    claimed.add(r["id"])
        out[rule["id"]] = n
    return out


def effective(rec):
    """What this record's title/abstract decision actually is, and why.

    Manual beats automatic. Always. Returns
    (decision, reason, source) with source 'manual', 'auto' or ''."""
    d = STATE["decisions"].get(rec["id"], {})
    if d.get("ta_decision"):
        # A label brought back from ASReview is still your decision, but it
        # was made elsewhere and the methods section should be able to say so.
        return (d["ta_decision"], d.get("ta_reason", ""),
                d.get("ta_source") or "manual")
    reason, _ = auto_reason(rec)
    if reason:
        return "exclude", reason, "auto"
    return "", "", ""


# ==========================================================================
#  persistence
# ==========================================================================

def apply_manual_fields():
    """Fold hand-typed abstracts into the in-memory records.

    Done once at load and again on every save, so that everything downstream
    — the no-abstract rule, the ranking, Ask, the dashboard, the export —
    sees the abstract you added without any of them needing to know it was
    added. The corpus file on disk is never touched: your text lives in the
    decisions file, which is also what carries it across a re-harvest."""
    n = 0
    for r in STATE["records"]:
        d = STATE["decisions"].get(r["id"], {})
        was = r.get("abstract")

        # Fields other than the abstract: a typed value replaces what the
        # source gave, and clearing it restores the original.
        edited = []
        for field, col in MANUAL_OVERRIDES.items():
            if field == "abstract":
                continue
            r.setdefault(field + "_original", r.get(field, ""))
            typed_v = (d.get(col) or "").strip()
            if typed_v:
                r[field] = typed_v
                edited.append(field)
            else:
                r[field] = r.get(field + "_original", "")
        r["edited_fields"] = edited
        # The DOI and title are what deduplication matches on, so their
        # normalised forms have to be rebuilt whenever either is corrected.
        r["ndoi"] = norm_doi(r.get("doi"))
        r["ntitle"] = norm_title(r.get("title"))

        typed = (d.get("abstract_manual") or "").strip()
        if typed:
            r["abstract"] = typed
            r["abstract_added"] = True
            n += 1
        elif r.get("abstract_added"):
            r["abstract"] = r.get("abstract_original", "")
            r["abstract_added"] = False
        else:
            r["abstract_added"] = False
        if r["abstract"] != was:
            # The country and theme guesses are read out of the abstract, so
            # a new one has to re-run them. Your own edits are kept separately
            # and still win; only the guess underneath is refreshed.
            c, basis = guess_country(r)
            r["country_guess"], r["country_basis"] = c, basis
            r["theme_guess"] = guess_themes(r)
    return n


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            v = json.load(fh)
            return v if v else default
    except Exception:
        return default


def save_json(path, obj):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, indent=2)
    os.replace(tmp, path)


def load_decisions(path):
    out = {}
    if not os.path.exists(path):
        return out
    try:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh, delimiter=sniff_delimiter(path)):
                rid = (row.get("id") or "").strip()
                if rid:
                    out[rid] = {k: (row.get(k) or "") for k in DECISION_FIELDS}
    except Exception as e:
        print(f"  could not read {os.path.basename(path)}: {e}", file=sys.stderr)
    return out


def save_decisions(path, decisions):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["id"] + DECISION_FIELDS,
                           delimiter=";", quoting=csv.QUOTE_ALL,
                           lineterminator="\r\n", extrasaction="ignore")
        w.writeheader()
        for rid, d in sorted(decisions.items()):
            row = {"id": rid}
            row.update({k: d.get(k, "") for k in DECISION_FIELDS})
            w.writerow(row)
    try:
        os.replace(tmp, path)
    except PermissionError:
        # Windows refuses to replace a file another program holds open, and
        # the usual culprit is the decisions CSV being open in Excel. Say so,
        # because the alternative is screening on while nothing is written.
        try:
            os.remove(tmp)      # don't leave a half-written twin behind
        except OSError:
            pass
        raise RuntimeError(
            f"Could not save: {os.path.basename(path)} is open in another "
            "program, most likely Excel. Close it there and make the decision "
            "again — nothing has been lost, but nothing is being saved until "
            "you do.")


def has_work(d):
    """Does this decision record an actual judgement, or is it an empty shell?"""
    return any((d.get(f) or "").strip() for f in
               ("ta_decision", "ft_decision", "labels", "notes", "dup_status",
                "country", "instrument_type", "abstract_manual", "doi_manual",
                "title_manual"))


def carry_over(records, own_path, already=None, quiet=False):
    """Reuse decisions made against another corpus file in the same folder.

    Re-running the search produces new OpenAlex ids for the same papers, so
    decisions are matched on id, then DOI, then normalised title. Without
    this, changing the query throws away every judgement you have made.

    This now runs EVERY time, not only when the corpus has no decisions of
    its own, and it only ever fills gaps — a record that already carries a
    judgement is never touched. That matters when the same review is spread
    over more than one harvest: opening a newer corpus used to show nothing,
    because the newer file had a decisions file of its own and the 446
    records screened against the older one were never looked for."""
    already = already or {}
    folder = Path(own_path).parent
    mine = Path(own_path).name
    by_doi, by_title, found = {}, {}, {}
    for r in records:
        if r["ndoi"]:
            by_doi[r["ndoi"]] = r["id"]
        if r["ntitle"]:
            by_title[r["ntitle"]] = r["id"]
    ids = {r["id"] for r in records}
    for p in sorted(folder.glob("*_decisions.csv")):
        if p.name == mine:
            continue
        try:
            with open(p, newline="", encoding="utf-8-sig") as fh:
                rows = list(csv.DictReader(fh, delimiter=sniff_delimiter(p)))
        except Exception:
            continue
        n = 0
        for row in rows:
            rid = (row.get("id") or "").strip()
            target = (rid if rid in ids else
                      by_doi.get(norm_doi(row.get("doi", ""))) or
                      by_title.get(norm_title(row.get("title", ""))))
            if not target or target in found:
                continue
            # Never overwrite a judgement this corpus already holds.
            if has_work(already.get(target, {})):
                continue
            if not has_work(row):
                continue
            found[target] = {k: (row.get(k) or "") for k in DECISION_FIELDS}
            n += 1
        if n and quiet:
            print(f"  {p.name} holds {n:,} decisions for records in this corpus "
                  f"that you have not screened here.\n"
                  f"    To bring them in:  python \"{Path(sys.argv[0]).name}\" "
                  f"\"{STATE.get('source') or Path(own_path).name}\" "
                  f"--merge-decisions\n"
                  f"    Check that file first — decisions made by a bulk "
                  f"operation look the same as screening you did yourself.")
        elif n:
            print(f"  brought {n:,} decisions across from {p.name}")
    return found


def backup_decisions(tag="backup"):
    """Copy the decisions file before anything rewrites it in bulk."""
    path = STATE["paths"]["decisions"]
    if not os.path.exists(path):
        return ""
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = f"{os.path.splitext(path)[0]}_{tag}_{stamp}.csv"
    import shutil
    shutil.copy2(path, dest)
    return os.path.basename(dest)


BULK_MIN = 25          # decisions in one second before it counts as a batch


def bulk_batches():
    """Decisions written all at once rather than one record at a time.

    Every decision carries the second it was made in. Twenty-five or more
    sharing a single second cannot have been made by reading twenty-five
    records: they came from a bulk action, an import or a script. That is
    not wrong in itself — applying a filter to a whole selection is ordinary
    screening, and faster than doing it by hand.

    What matters is what the batch WROTE. A batch that set a full-text
    decision is claiming those full texts were retrieved and read, and the
    PRISMA diagram reports it exactly that way; a batch that recorded no
    reason leaves the diagram with an exclusion it cannot explain. Both are
    questions a reviewer asks, so they are surfaced here rather than left to
    be discovered in the diagram.

    Only batches that are missing a reason are returned. Once every record in
    a batch has one, there is nothing left to answer for and it drops out of
    this list by itself."""
    by_second = {}
    for rid, d in STATE["decisions"].items():
        if not (d.get("ta_decision") or d.get("ft_decision")):
            continue
        sec = (d.get("updated_at") or "")[:19]
        if sec:
            by_second.setdefault(sec, []).append(rid)

    here = {r["id"]: r for r in STATE["records"]}
    out = []
    for sec, ids in sorted(by_second.items()):
        if len(ids) < BULK_MIN:
            continue
        ds = [STATE["decisions"][i] for i in ids]
        unexplained = sum(1 for d in ds if not (d.get("ta_reason") or "").strip()
                          and not (d.get("ft_reason") or "").strip())
        if unexplained * 2 < len(ds):
            continue                      # mostly explained; nothing to answer
        ta = Counter((d.get("ta_decision") or "—") for d in ds)
        ft = Counter((d.get("ft_decision") or "—") for d in ds)
        # What the records in this batch have in common, measured against the
        # rules, so a reason can be checked rather than assumed.
        recs = [here[i] for i in ids if i in here]
        checks = []
        for rule in STATE["rules"]:
            hits = sum(1 for r in recs if rule_hits(r, rule))
            if hits:
                checks.append({"id": rule["id"], "name": rule["name"],
                               "reason": rule.get("reason", ""),
                               "hits": hits, "of": len(recs)})
        checks.sort(key=lambda c: -c["hits"])
        out.append({
            "stamp": sec, "n": len(ids), "in_corpus": len(recs),
            "unexplained": unexplained,
            "ta": dict(ta), "ft": dict(ft),
            "claims_fulltext": ft.get("exclude", 0) + ft.get("include", 0),
            "checks": checks[:5],
            "sample": [{"title": here[i]["title"][:84], "year": here[i]["year"]}
                       for i in ids[:5] if i in here],
        })
    return out


def apply_bulk_fix(stamp, action, reason="", rule_id=""):
    """Re-file one batch. Returns what it changed.

    `exclude` moves the batch to a title/abstract exclusion with a reason and
    clears the full-text decision, which is what a filter applied to a
    selection actually was. `rule_id` guards it: a record the rule does not
    hit contradicts the reason, so it is handed back unscreened rather than
    given a reason that is untrue of it. `explain` keeps the batch as it
    stands and only writes the missing reason, for a batch that really was
    full-text work. `clear` returns the whole batch to unscreened."""
    ids = [rid for rid, d in STATE["decisions"].items()
           if (d.get("updated_at") or "")[:19] == stamp
           and (d.get("ta_decision") or d.get("ft_decision"))]
    if not ids:
        raise RuntimeError("That batch is no longer in the decisions file.")
    here = {r["id"]: r for r in STATE["records"]}
    rule = next((x for x in STATE["rules"] if x["id"] == rule_id), None)
    backup = backup_decisions("beforefix")

    changed = handed_back = 0
    for rid in ids:
        d = STATE["decisions"][rid]
        rec = here.get(rid)
        if action == "clear" or (
                action == "exclude" and rule is not None
                and rec is not None and not rule_hits(rec, rule)):
            # Nothing is asserted about this record that cannot be checked.
            d["ta_decision"] = d["ta_reason"] = ""
            d["ft_decision"] = d["ft_reason"] = ""
            handed_back += 1 if action == "exclude" else 0
            changed += 1 if action == "clear" else 0
        elif action == "exclude":
            d["ta_decision"] = "exclude"
            d["ta_reason"] = reason or "Reason not recorded"
            d["ft_decision"] = d["ft_reason"] = ""
            changed += 1
        elif action == "explain":
            if d.get("ft_decision"):
                d["ft_reason"] = reason or d.get("ft_reason") or ""
            else:
                d["ta_reason"] = reason or d.get("ta_reason") or ""
            changed += 1
        else:
            raise RuntimeError(f"Unknown action {action!r}")
        # The timestamp is left alone deliberately: it records when you made
        # the judgement, which is what the stopping rule reads as the order
        # you screened in. Only what the judgement SAYS has changed.
    save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
    return {"changed": changed, "handed_back": handed_back,
            "total": len(ids), "backup": backup}


def reattach_decisions():
    """Re-point decisions whose record id no longer exists.

    A decision is stored against a record id. Imported records used to be
    numbered by position — WOS1, WOS2 — so importing a second file, or the
    same file again, renumbered them and every decision quietly moved to a
    different paper. Ids are derived from the DOI or title now (see
    record_key), which stops it happening again, but it also means ids saved
    by an earlier version of this app no longer match.

    So every decision that points at nothing is re-attached here by DOI, then
    by title, using the doi and title columns the decisions file has always
    carried. Nothing is deleted: a decision that still finds no record is
    left where it is, because the paper may come back on the next import.

    Returns (moved, orphans)."""
    ids = {r["id"] for r in STATE["records"]}
    by_doi, by_title = {}, {}
    for r in STATE["records"]:
        if r["ndoi"]:
            by_doi.setdefault(r["ndoi"], r["id"])
        if r["ntitle"]:
            by_title.setdefault(r["ntitle"], r["id"])

    moved, orphans = 0, 0
    for rid in [k for k in STATE["decisions"] if k not in ids]:
        d = STATE["decisions"][rid]
        if not has_work(d):
            STATE["decisions"].pop(rid, None)   # an empty shell, not your work
            continue
        target = (by_doi.get(norm_doi(d.get("doi", ""))) or
                  by_title.get(norm_title(d.get("title", ""))))
        # Never overwrite a decision already made against the live record.
        if target and target not in ids:
            target = None
        if target and has_work(STATE["decisions"].get(target, {})):
            target = None
        if target:
            STATE["decisions"][target] = d
            STATE["decisions"].pop(rid, None)
            moved += 1
        else:
            orphans += 1
    return moved, orphans


# ==========================================================================
#  PRISMA
# ==========================================================================

# ==========================================================================
#  When may you stop screening?
#
#  Active learning reorders records; it never labels them. So screening ends
#  when YOU decide it does, and that decision has to be defensible. This
#  follows the SAFE procedure (Boetje & van de Schoot, Systematic Reviews
#  2024), which combines four conditions rather than trusting any single one:
#
#    1. every key paper you already know about has been found
#    2. at least twice the estimated number of relevant records is screened
#    3. at least 10% of the corpus is screened
#    4. no relevant record has appeared in the last N screened
#
#  The thresholds in the paper are examples, not law, so both are adjustable.
#  Nothing here stops you screening: it reports where you stand, and records
#  the point if you choose to stop, so PRISMA can account for what was left.
# ==========================================================================

SAFE_CONSECUTIVE = 50       # condition 4: records since the last relevant one
SAFE_MIN_PERCENT = 10       # condition 3: share of the corpus screened
# Fallback only. The key papers normally come from the corpus itself: any
# record whose key_paper column says "core", which is what a hand-made list's
# status column becomes when it is merged in. This pattern is used only when
# no record carries that marking.
GOLDSET_GLOB = "*goldset*.csv"


def screening_order():
    """The records you screened yourself, oldest decision first.

    Only decisions about records IN THIS CORPUS count. The decisions file
    outlives any one harvest — carry-over deliberately preserves judgements
    across a re-run of the search — so it holds entries for records the
    current search did not return. Counting those would inflate every figure
    here: the share screened, the number found relevant, and the run of
    consecutive irrelevant records that the stopping rule turns on.

    Duplicates are out too, for the same reason they are out of the total:
    screened and total have to count the same population or the percentage
    is meaningless.

    Returns (sequence, skipped) where skipped separates the two reasons a
    decision is not counted, because they mean different things to you: one
    is work on records this search no longer returns, the other is work on a
    record that turned out to be a duplicate of one you kept."""
    dups, _ = dup_state()
    here = {r["id"] for r in STATE["records"]}
    seq, elsewhere, on_dupes = [], 0, 0
    for rid, d in STATE["decisions"].items():
        if not d.get("ta_decision"):
            continue
        if rid not in here:
            elsewhere += 1
            continue
        if dups.get(rid) == "duplicate":
            on_dupes += 1
            continue
        seq.append((d.get("updated_at", ""), rid, d["ta_decision"]))
    seq.sort(key=lambda x: (x[0], x[1]))
    return seq, {"elsewhere": elsewhere, "on_duplicates": on_dupes}


def key_paper_status():
    """How many of the papers you already know about have been found.

    The gold set is read from the corpus itself: any record whose key_paper
    column says "core" counts, which is what the status column of a
    hand-added list becomes when you import it. There used to be two separate
    gold sets — a CSV read only here, and the records you imported — which
    could disagree with each other. Now importing the list is what defines
    it.

    A loose *goldset*.csv beside the app is read as a fallback, so nothing
    breaks if your list has not been merged in yet.

    A key paper that is not in the corpus at all is a search problem, not a
    screening one, so the two are counted separately."""
    # Duplicates are excluded. The core marking is copied onto whichever
    # record survives deduplication (see merge_duplicates), so counting both
    # the hand-added row and the harvested record it duplicates would count
    # the same key paper twice — 41 of 33, which is what this did before the
    # filter was added.
    dups, _ = dup_state()
    gold = [{"doi": r.get("doi"), "title": r.get("title"),
             "authors": r.get("authors"), "year": r.get("year")}
            for r in STATE["records"]
            if (r.get("key_paper") or "") == "core"
            and dups.get(r["id"]) != "duplicate"]
    where = "the corpus (records marked core)"
    if not gold:
        hits = sorted(SCRIPT_DIR.glob(GOLDSET_GLOB),
                      key=lambda p: p.stat().st_mtime, reverse=True)
        if not hits:
            return None
        path = hits[0]
        try:
            with open(path, newline="", encoding="utf-8-sig") as fh:
                gold = [g for g in csv.DictReader(
                            fh, delimiter=sniff_delimiter(path))
                        if (g.get("status") or "").strip() in ("core", "")]
        except Exception:
            return None
        where = path.name
    if not gold:
        return None
    # A hand-added key paper is in the corpus because you put it there, so
    # matching it against itself would say the search found it. Records from
    # a database are therefore indexed separately: `in_corpus` answers "will
    # this get screened", `found_by_search` answers "is my query finding the
    # papers I already know about" — which is the question that tells you
    # whether to widen the search rather than keep screening.
    by_doi = {r["ndoi"]: r for r in STATE["records"] if r["ndoi"]}
    by_title = {r["ntitle"]: r for r in STATE["records"] if r["ntitle"]}
    db = [r for r in STATE["records"]
          if (r.get("source") or "openalex") not in OTHER_METHODS]
    db_doi = {r["ndoi"] for r in db if r["ndoi"]}
    db_titles = [r["ntitle"] for r in db if r["ntitle"]]
    db_tset = set(db_titles)
    titles = list(by_title)
    in_corpus, found, by_search = 0, 0, 0
    missing_corpus, not_yet, missed_by_search = [], [], []
    for g in gold:
        d, t = norm_doi(g.get("doi")), norm_title(g.get("title"))
        rec = (by_doi.get(d) if d else None) or (by_title.get(t) if t else None)
        if rec is None and t:
            close = difflib.get_close_matches(t, titles, n=1, cutoff=0.82)
            if close:
                rec = by_title[close[0]]
        label = f"{(g.get('authors') or '?')[:26]} ({g.get('year') or '?'})"

        hit = (d in db_doi) if d else False
        if not hit and t:
            hit = t in db_tset or bool(
                difflib.get_close_matches(t, db_titles, n=1, cutoff=0.82))
        if hit:
            by_search += 1
        else:
            missed_by_search.append(label)

        if rec is None:
            missing_corpus.append(label)
            continue
        in_corpus += 1
        if is_included(rec):
            found += 1
        else:
            not_yet.append(label)
    return {"total": len(gold), "in_corpus": in_corpus, "found": found,
            "by_search": by_search,
            "missing_from_corpus": missing_corpus[:12],
            "not_yet_included": not_yet[:12],
            "missed_by_search": missed_by_search[:12],
            "n_missing": len(missing_corpus), "n_not_yet": len(not_yet),
            "n_missed_by_search": len(missed_by_search),
            "where": where}


def stopping_status():
    """Where the four SAFE conditions stand right now."""
    dups, _ = dup_state()
    eligible = [r for r in STATE["records"] if dups.get(r["id"]) != "duplicate"]
    total = len(eligible)
    seq, skipped = screening_order()
    screened = len(seq)
    relevant = sum(1 for _, _, v in seq if v == "include")

    consecutive = 0
    for _, _, v in reversed(seq):
        if v == "include":
            break
        consecutive += 1

    cfg = STATE["stop_cfg"]
    manual = cfg.get("estimate")
    if manual:
        estimate, est_from = int(manual), "your estimate"
    elif screened:
        # Screening in prioritised order finds relevant records early, so the
        # observed rate overstates the true one. That makes this estimate too
        # HIGH, which makes condition 2 harder to meet — the safe direction.
        estimate = max(1, round(relevant / screened * total))
        est_from = "the rate so far (likely an overestimate)"
    else:
        estimate, est_from = 0, "nothing screened yet"

    keys = key_paper_status()
    pct = (screened / total * 100) if total else 0
    thresh = int(cfg.get("consecutive") or SAFE_CONSECUTIVE)
    minpct = float(cfg.get("min_percent") or SAFE_MIN_PERCENT)

    conds = [
        {"id": "keys", "label": "Every key paper found",
         "met": bool(keys) and keys["found"] >= keys["in_corpus"] and keys["in_corpus"] > 0,
         "detail": (f"{keys['found']} of {keys['in_corpus']} key papers that are "
                    f"in the corpus are included"
                    + (f"; {keys['n_missing']} are not in the corpus at all"
                       if keys and keys["n_missing"] else "")
                    + (f"; the database search returns only "
                       f"{keys['by_search']} of {keys['total']}"
                       if keys and keys["by_search"] < keys["total"] else "")
                    ) if keys else
                   "no key papers marked — merge your gold set into the "
                   "corpus with 10-merge_sources.py, or import it on the "
                   "Sources tab",
         "skipped": keys is None},
        {"id": "twice", "label": "Twice the estimated relevant screened",
         "met": bool(estimate) and screened >= 2 * estimate,
         "detail": f"{screened:,} screened, need {2 * estimate:,} "
                   f"(estimate {estimate:,} relevant, from {est_from})",
         "skipped": False},
        {"id": "pct", "label": f"At least {minpct:g}% of the corpus screened",
         "met": pct >= minpct,
         "detail": f"{pct:.1f}% screened ({screened:,} of {total:,})",
         "skipped": False},
        {"id": "run", "label": f"No relevant record in the last {thresh}",
         "met": consecutive >= thresh,
         "detail": f"{consecutive:,} screened since the last relevant one",
         "skipped": False},
    ]
    live = [c for c in conds if not c["skipped"]]
    return {
        "total": total, "screened": screened, "relevant": relevant,
        # Records with no decision at all, by rule or by hand — the work that
        # is genuinely left. `total - screened` was counting the records a
        # rule has already excluded as still to look at, which made this
        # disagree with the same figure on the PRISMA tab (842 against 830).
        "unscreened": sum(1 for r in eligible if not effective(r)[0]),
        "auto_decided": sum(1 for r in eligible
                            if effective(r)[2] == "auto"),
        "percent": round(pct, 1),
        "consecutive": consecutive, "threshold": thresh,
        "min_percent": minpct, "estimate": estimate, "estimate_from": est_from,
        "manual_estimate": cfg.get("estimate") or "",
        "skipped": skipped,
        "conditions": conds, "keys": keys,
        "all_met": bool(live) and all(c["met"] for c in live),
        "met_count": sum(1 for c in live if c["met"]), "live_count": len(live),
        "recall_curve": recall_curve(seq),
        "stop": STATE["stop"],
    }


def recall_curve(seq, points=60):
    """Relevant found against records screened, for the plateau check."""
    if not seq:
        return []
    out, found = [], 0
    step = max(1, len(seq) // points)
    for i, (_, _, v) in enumerate(seq, 1):
        if v == "include":
            found += 1
        if i % step == 0 or i == len(seq):
            out.append([i, found])
    return out


def is_included(rec):
    """Is this study in the review as things currently stand?

    One definition, used by the PRISMA final box, the per-theme breakdown and
    the export, so the three can never disagree. A study included at title and
    abstract counts until full-text screening rejects it; once you have done
    the full-text stage this converges on the strict answer."""
    d = STATE["decisions"].get(rec["id"], {})
    if d.get("ft_decision"):
        return d["ft_decision"] == "include"
    return effective(rec)[0] == "include"


def prisma_counts():
    recs = STATE["records"]
    dups, _ = dup_state()
    dup_removed = sum(1 for r in recs if dups.get(r["id"]) == "duplicate")
    eligible = [r for r in recs if dups.get(r["id"]) != "duplicate"]

    ta_inc = ta_exc = ta_may = 0
    auto_exc = man_exc = asr_exc = title_exc = 0
    reasons, auto_reasons = {}, {}
    pool = []
    for r in eligible:
        dec, why, how = effective(r)
        if dec == "include":
            ta_inc += 1
            pool.append(r)
        elif dec == "maybe":
            ta_may += 1
            pool.append(r)
        elif dec == "exclude":
            ta_exc += 1
            key = why or "Reason not recorded"
            reasons[key] = reasons.get(key, 0) + 1
            if how == "auto":
                auto_exc += 1
                auto_reasons[key] = auto_reasons.get(key, 0) + 1
            elif how == "asreview":
                asr_exc += 1
            elif how == "title-only":
                # Decided from the title without opening the abstract. A
                # legitimate shortcut where a title settles it, and one a
                # reviewer is entitled to see the size of.
                title_exc += 1
            else:
                man_exc += 1

    ft_inc = ft_exc = ft_done = 0
    ft_reasons = {}
    for r in pool:
        d = STATE["decisions"].get(r["id"], {})
        if d.get("ft_decision") == "exclude":
            ft_exc += 1
            ft_done += 1
            key = d.get("ft_reason") or "Reason not recorded"
            ft_reasons[key] = ft_reasons.get(key, 0) + 1
        elif d.get("ft_decision") == "include":
            ft_done += 1
    ft_inc = sum(1 for r in pool if is_included(r))

    lab = {l["id"]: 0 for l in STATE["labels"]}
    for r in pool:
        if not is_included(r):
            continue
        d = STATE["decisions"].get(r["id"], {})
        chosen = d.get("labels")
        ids = ([x.strip() for x in chosen.split("|") if x.strip()]
               if chosen else r["theme_guess"])
        for lid in ids:
            if lid in lab:
                lab[lid] += 1

    # If screening was stopped under a rule, the records never reached are
    # NOT excluded — nobody looked at them. Reporting them as exclusions
    # would misstate the review; they get their own line instead.
    stopped = STATE["stop"].get("recorded")
    not_screened = (len(eligible) - ta_inc - ta_exc - ta_may) if stopped else 0

    # PRISMA 2020 keeps studies found through databases apart from studies
    # found by other methods — citation searching, hand searching. Mixing
    # them would misreport how the review found what it found.
    srcs = source_counts()
    db_ident = sum(s["identified"] for s in srcs if not s["other_method"])
    om_ident = sum(s["identified"] for s in srcs if s["other_method"])
    identified_total = db_ident + om_ident
    # Rows that were identified but are not records: the same paper twice in
    # one export, and rows with neither a title nor a DOI. Kept on their own
    # line so the top of the flow reconciles with the bottom.
    not_stored = sum(s["not_stored"] for s in srcs)
    # Every duplicate is now a duplicate IN the corpus, which means it is one
    # you can see and override on the Duplicates tab. This figure and that
    # tab can no longer disagree.
    duplicates_total = dup_removed

    return {
        "identified": identified_total, "duplicates": duplicates_total,
        "duplicates_in_corpus": dup_removed,
        "not_stored": not_stored,
        "records_held": len(recs),
        "sources": srcs, "db_identified": db_ident,
        "other_identified": om_ident,
        "stopped": bool(stopped), "stop": STATE["stop"],
        "not_screened": max(0, not_screened),
        "screened": len(eligible), "ta_included": ta_inc,
        "ta_excluded": ta_exc, "ta_maybe": ta_may,
        "ta_auto_excluded": auto_exc, "ta_manual_excluded": man_exc,
        "ta_asreview_excluded": asr_exc,
        "ta_title_excluded": title_exc,
        "ta_todo": len(eligible) - ta_inc - ta_exc - ta_may,
        "ft_sought": len(pool), "ft_included": ft_inc, "ft_excluded": ft_exc,
        "ft_todo": len(pool) - ft_done,
        # What is left over once the full-text stage has taken its share:
        # records still marked "maybe", which are neither excluded nor
        # included. This is the number that makes the diagram look as though
        # it does not add up, so it is reported rather than hidden.
        "ft_undecided": max(0, len(pool) - ft_exc - ft_inc),
        # Counted as included for now on the strength of the abstract, but the
        # full text has not been read. A different thing from a "maybe", and
        # the two must not be added together: they are different records.
        "ft_pending": sum(
            1 for r in pool
            if not STATE["decisions"].get(r["id"], {}).get("ft_decision")
            and is_included(r)),
        "reasons": reasons, "auto_reasons": auto_reasons,
        "ft_reasons": ft_reasons, "labels": lab,
    }


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def prisma_svg(c):
    W, BW, BH, GAP = 1020, 460, 92, 34
    srcs = c.get("sources") or []
    dbs = [s for s in srcs if not s["other_method"]]
    other = [s for s in srcs if s["other_method"]]

    # Every source named inside the identification box, always — one line per
    # database and one per hand-added set, split into PRISMA 2020's two
    # identification arms. The old version printed this as a footnote and only
    # when there was more than one source, so a diagram built from four
    # databases could still read "Records identified" with a single total and
    # no way to see where any of it came from.
    ident = "Records identified"
    ident_lines = []
    if dbs:
        ident_lines.append(("From databases", ""))
        for s in dbs:
            ident_lines.append((f"      {s['label']}", f"n = {s['identified']:,}"))
    if other:
        ident_lines.append(("From other methods", ""))
        for s in other:
            ident_lines.append((f"      {s['label']}", f"n = {s['identified']:,}"))
    if other and dbs and not ident_lines:
        ident = (f"Records identified from databases (n = {c['db_identified']:,}) "
                 f"and other methods (n = {c['other_identified']:,})")
    # Screening in progress is not the same as screening finished, and the
    # diagram has to say which. Without these two lines the boxes look like
    # they do not add up: 589 sought minus 562 excluded is 27, not the 23
    # included, and the missing 4 are simply full texts you have not read
    # yet. Each line disappears once that stage is complete.
    screened_label = "Records screened on title and abstract"
    if c.get("ta_todo"):
        screened_label += f"  —  {c['ta_todo']:,} not yet screened"
    sought_label = "Full texts sought for retrieval"
    bits = []
    if c.get("ft_pending"):
        bits.append(f"{c['ft_pending']:,} not yet read")
    if c.get("ft_undecided"):
        bits.append(f"{c['ft_undecided']:,} still marked maybe")
    if bits:
        sought_label += "  —  " + ", ".join(bits)
    main = [(ident, f"n = {c['identified']:,}", ident_lines),
            (screened_label, f"n = {c['screened']:,}", []),
            (sought_label, f"n = {c['ft_sought']:,}", []),
            ("Studies included in the review", f"n = {c['ft_included']:,}", [])]
    how = [f"{c['ta_auto_excluded']:,} by rule",
           f"{c['ta_manual_excluded']:,} by hand"]
    if c.get("ta_asreview_excluded"):
        how.append(f"{c['ta_asreview_excluded']:,} in ASReview")
    if c.get("ta_title_excluded"):
        how.append(f"{c['ta_title_excluded']:,} on the title alone")
    # PRISMA's "removed before screening" box. Everything identified that
    # never reached screening belongs here, and the two reasons are kept
    # apart: duplicates found across the databases, and rows that could not
    # become records at all — the same paper listed twice inside one export,
    # or a row with neither a title nor a DOI.
    removed_before = c["duplicates"] + c.get("not_stored", 0)
    dup_label = "Duplicate records removed"
    if c.get("not_stored"):
        dup_label = (f"Records removed before screening "
                     f"({c['duplicates']:,} duplicates, "
                     f"{c['not_stored']:,} unusable or repeated rows)")
    side = [(dup_label, f"n = {removed_before:,}"),
            (f"Excluded at title/abstract ({', '.join(how)})",
             f"n = {c['ta_excluded']:,}"),
            ("Full texts excluded", f"n = {c['ft_excluded']:,}")]
    if c.get("not_screened"):
        s = c.get("stop") or {}
        side[1] = (f"Excluded at title/abstract ({', '.join(how)}); "
                   f"{c['not_screened']:,} not individually screened after "
                   f"the stopping rule was met",
                   f"n = {c['ta_excluded']:,} + {c['not_screened']:,}")
    themes = [(l["name"], c["labels"].get(l["id"], 0)) for l in STATE["labels"]]
    themes = [t for t in themes if t[1]] or themes
    TH = 56 + 24 * len(themes) if themes else 0
    SH = 42 + 20 * len(srcs) if srcs else 0
    # Boxes are no longer all the same height: the identification box grows by
    # one line per source, so positions are accumulated rather than indexed.
    heights = [BH + (22 * len(lines) + 10 if lines else 0) for _, _, lines in main]
    tops, acc = [], 80
    for h in heights:
        tops.append(acc)
        acc += h + GAP
    H = acc - GAP + 60 + SH + TH
    p = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
         f'viewBox="0 0 {W} {H}" font-family="Segoe UI,Helvetica,Arial,sans-serif">',
         f'<rect width="{W}" height="{H}" fill="#ffffff"/>',
         f'<text x="24" y="34" font-size="17" font-weight="600" fill="#111">'
         f'PRISMA 2020 flow</text>',
         f'<text x="24" y="56" font-size="12" fill="#666">'
         f'{esc(STATE["source"])} &#183; {datetime.date.today():%d %B %Y}</text>']
    x = 40
    for i, (label, n, lines) in enumerate(main):
        y, h = tops[i], heights[i]
        p.append(f'<rect x="{x}" y="{y}" width="{BW}" height="{h}" rx="8" '
                 f'fill="#f7f9fc" stroke="#9fb3c8" stroke-width="1.5"/>')
        p.append(f'<text x="{x+18}" y="{y+32}" font-size="14" fill="#102a43">'
                 f'{esc(label)}</text>')
        p.append(f'<text x="{x+18}" y="{y+60}" font-size="16" font-weight="600" '
                 f'fill="#0072B2">{esc(n)}</text>')
        for j, (txt, num) in enumerate(lines):
            yy = y + 84 + j * 22
            heading = not num
            p.append(f'<text x="{x+18}" y="{yy}" font-size="12.5" '
                     f'{"font-weight=\"600\" " if heading else ""}'
                     f'fill="{"#5b6b7c" if heading else "#102a43"}">'
                     f'{esc(txt)}</text>')
            if num:
                p.append(f'<text x="{x+BW-18}" y="{yy}" font-size="12.5" '
                         f'font-weight="600" fill="#102a43" text-anchor="end">'
                         f'{esc(num)}</text>')
        if i < len(main) - 1:
            my = y + h
            p.append(f'<line x1="{x+BW/2}" y1="{my}" x2="{x+BW/2}" y2="{my+GAP}" '
                     f'stroke="#9fb3c8" stroke-width="1.5" '
                     f'marker-end="url(#a)"/>')
            sx, sy = x + BW + 70, my - 24
            label2, n2 = side[i]
            p.append(f'<rect x="{sx}" y="{sy}" width="{W-sx-20}" height="{BH-18}" '
                     f'rx="8" fill="#fff7f0" stroke="#e0b089" stroke-width="1.5"/>')
            p.append(f'<text x="{sx+16}" y="{sy+30}" font-size="13" '
                     f'fill="#6b3f12">{esc(label2)}</text>')
            p.append(f'<text x="{sx+16}" y="{sy+54}" font-size="15" '
                     f'font-weight="600" fill="#D55E00">{esc(n2)}</text>')
            p.append(f'<line x1="{x+BW/2}" y1="{my+GAP/2}" x2="{sx}" '
                     f'y2="{my+GAP/2}" stroke="#e0b089" stroke-width="1.5" '
                     f'marker-end="url(#b)"/>')
    # What each source contributed once duplicates were resolved. The box
    # above reports what was identified, which is the figure PRISMA asks for;
    # this says how much of it survived, which is the figure you are asked
    # about in review.
    sy = acc - GAP + 16
    if srcs:
        p.append(f'<text x="{x}" y="{sy}" font-size="13" font-weight="600" '
                 f'fill="#102a43">After duplicates were removed, by source</text>')
        for i, s in enumerate(srcs):
            yy = sy + 22 + i * 20
            extra = (f", {s['not_stored']:,} unusable" if s["not_stored"] else "")
            p.append(f'<text x="{x+6}" y="{yy}" font-size="12.5" fill="#102a43">'
                     f'{esc(s["label"])}</text>')
            p.append(f'<text x="{x+500}" y="{yy}" font-size="12.5" '
                     f'fill="#102a43" text-anchor="end">'
                     f'{s["after_dedup"]:,} of {s["identified"]:,}{extra}</text>')
        sy += 22 + 20 * len(srcs)

    # The included studies, broken down by theme. A paper carrying two themes
    # is counted under both, so these do not sum to the included total.
    if themes:
        ty = sy + 20
        p.append(f'<text x="{x}" y="{ty}" font-size="13" font-weight="600" '
                 f'fill="#102a43">Included studies by theme</text>')
        widest = max((n for _, n in themes), default=1) or 1
        for i, (name, n) in enumerate(themes):
            yy = ty + 22 + i * 24
            colour = PALETTE[i % len(PALETTE)]
            p.append(f'<rect x="{x}" y="{yy-11}" width="{max(3, n / widest * 190):.0f}" '
                     f'height="15" rx="3" fill="{colour}"/>')
            p.append(f'<text x="{x+200}" y="{yy+1}" font-size="12.5" '
                     f'fill="#102a43">{esc(name)}</text>')
            p.append(f'<text x="{x+600}" y="{yy+1}" font-size="12.5" '
                     f'font-weight="600" fill="#102a43" '
                     f'text-anchor="end">n = {n:,}</text>')
        p.append(f'<text x="{x}" y="{ty + 26 + len(themes) * 24}" font-size="11" '
                 f'fill="#6b7a8c">A study carrying more than one theme is '
                 f'counted under each, so these need not sum to '
                 f'{c["ft_included"]:,}.</text>')
    p.append('<defs>'
             '<marker id="a" markerWidth="9" markerHeight="9" refX="8" refY="3" '
             'orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#9fb3c8"/></marker>'
             '<marker id="b" markerWidth="9" markerHeight="9" refX="8" refY="3" '
             'orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#e0b089"/></marker>'
             '</defs></svg>')
    return "".join(p)


# ==========================================================================
#  active learning: TF-IDF + naive Bayes, pure Python
# ==========================================================================

STOP = set("""the a an and or but if of to in on for with without by from as at
this that these those is are was were be been being have has had do does did
we our us they their it its i you your he she his her not no nor can could
will would shall should may might must about into over under between among
than then there here when where which who whom whose what how why all any
both each few more most other some such only own same so too very s t just
don now study paper research results findings using used use based data
analysis approach model method methods also however therefore thus
""".split())
TOKEN = re.compile(r"[a-z][a-z-]{2,}")


def tokenize(text):
    return [w for w in TOKEN.findall((text or "").lower())
            if w not in STOP and len(w) < 26]


# ASReview's own active-learning machinery, used when it is installed.
# It is a separate package under the Apache 2.0 licence
# (https://github.com/asreview/asreview) and is imported, not copied: the
# model that ranks your records is then literally theirs, which is what makes
# it citable in your methods section.
#
#     pip install asreview
#
# Nothing here depends on it. Without it the built-in ranker below runs
# instead and the app behaves exactly as before.
ASREVIEW_MODELS = {
    "nb": ("Naive Bayes", "NaiveBayes"),
    "svm": ("SVM", "SVM"),
    "logistic": ("Logistic regression", "Logistic"),
    "rf": ("Random forest", "RandomForest"),
}


def asreview_available():
    try:
        import asreview                                    # noqa: F401
        return True
    except Exception:
        return False


def asreview_version():
    try:
        import asreview
        return getattr(asreview, "__version__", "")
    except Exception:
        return ""


def rank_with_asreview(recs, pos_ids, neg_ids, model="nb"):
    """Rank every record using ASReview's classifier and feature extractor.

    One pass of their ActiveLearningCycle: fit on the records you have
    judged, then rank the whole corpus. No project file and no simulation —
    the labels stay in your decisions file and the model only reorders what
    you see. Returns scores in 0..1, highest first, or None if ASReview is
    not installed or cannot run.

    A failure here is never fatal: the caller falls back to the built-in
    ranker, because a ranking is a convenience and losing it must not stop
    you screening."""
    try:
        import numpy as np
        import pandas as pd
        from asreview import ActiveLearningCycle
        from asreview.models import classifiers as _clf
        from asreview.models.feature_extractors import Tfidf
        from asreview.models.queriers import Max
        from asreview.models.balancers import Balanced
    except Exception:
        return None
    try:
        cls_name = ASREVIEW_MODELS.get(model, ASREVIEW_MODELS["nb"])[1]
        classifier = getattr(_clf, cls_name)()
        # Their Tfidf reads a frame with these two columns by name.
        frame = pd.DataFrame({
            "title": [str(r.get("title") or "") for r in recs],
            "abstract": [str(r.get("abstract") or "")
                         + " " + str(r.get("keywords") or "") for r in recs],
        })
        pos, neg = set(pos_ids), set(neg_ids)
        y = np.array([1 if r["id"] in pos else 0 if r["id"] in neg else -1
                      for r in recs])
        cycle = ActiveLearningCycle(querier=Max(), classifier=classifier,
                                    balancer=Balanced(),
                                    feature_extractor=Tfidf())
        X = cycle.transform(frame)
        seen = y != -1
        cycle.fit(X[seen], y[seen])
        order = cycle.rank(X)
        # rank() returns indices best-first. Turn that into a 0..1 score so
        # it drops into the same place as the built-in ranker's output.
        n = len(order) or 1
        scores = {}
        for place, idx in enumerate(order):
            scores[recs[int(idx)]["id"]] = round(1.0 - place / n, 4)
        return scores
    except Exception as e:
        print(f"  ASReview ranking failed ({type(e).__name__}: {e}); "
              f"using the built-in ranker", file=sys.stderr)
        return None


def rank_records(model=None):
    """Order unscreened records by how much they look like your includes.

    ASReview's model when that package is installed, otherwise the built-in
    multinomial naive Bayes over TF-IDF. Either way it needs examples of both
    answers: with only includes it would rank everything highly and teach you
    nothing. Returns (scores, n_pos, n_neg, engine)."""
    recs, dec = STATE["records"], STATE["decisions"]
    pos, neg = [], []
    for r in recs:
        d = dec.get(r["id"], {})
        v = d.get("ft_decision") or d.get("ta_decision")
        if v == "include":
            pos.append(r)
        elif v == "exclude" and d.get("ta_decision"):
            neg.append(r)        # only your own excludes, never a rule's
    if len(pos) < 3 or len(neg) < 3:
        return {}, len(pos), len(neg), ""

    want = model or STATE.get("ranker") or "nb"
    if want != "builtin":
        s = rank_with_asreview(recs, [r["id"] for r in pos],
                               [r["id"] for r in neg], want)
        if s is not None:
            label = ASREVIEW_MODELS.get(want, ASREVIEW_MODELS["nb"])[0]
            return s, len(pos), len(neg), f"ASReview {asreview_version()} — {label}"

    scores, p, n = _rank_builtin(recs, pos, neg)
    return scores, p, n, "built-in (TF-IDF + naive Bayes)"


def compare_rankers(holdout=0.3, seeds=(1, 2, 3)):
    """Measure the models against each other on YOUR decisions.

    Hide a share of the records you have judged, train on the rest, then see
    where the hidden relevant ones land in the queue the model produces. A
    model that puts them near the top is a model that will save you work.

    Two numbers per model. `recall` is the share of hidden relevant records
    that turn up in the first 10% of the queue — screening at random would
    give 10%, so anything above that is the model earning its keep. `median`
    is how far down the queue the typical hidden relevant record sits, as a
    percentage; lower is better.

    This is a measurement, not a decision: it never changes a label and never
    touches your decisions file."""
    recs, dec = STATE["records"], STATE["decisions"]
    pos_i, neg_i = [], []
    for i, r in enumerate(recs):
        d = dec.get(r["id"], {})
        v = d.get("ft_decision") or d.get("ta_decision")
        if v == "include":
            pos_i.append(i)
        elif v == "exclude" and d.get("ta_decision"):
            neg_i.append(i)
    if len(pos_i) < 10 or len(neg_i) < 10:
        return {"error": f"Needs at least 10 includes and 10 of your own "
                         f"excludes to measure anything (you have "
                         f"{len(pos_i)} and {len(neg_i)})."}

    import random as _random
    rows = []
    engines = ([(k, v[0], v[1]) for k, v in ASREVIEW_MODELS.items()]
               if asreview_available() else [])
    engines.append(("builtin", "Built-in (no ASReview)", None))

    for key, label, cls_name in engines:
        rec_s, place_s = [], []
        failed = ""
        for s in seeds:
            rng = _random.Random(s)
            hp = set(rng.sample(pos_i, max(1, int(len(pos_i) * holdout))))
            hn = set(rng.sample(neg_i, max(1, int(len(neg_i) * holdout))))
            held = hp | hn
            train_pos = [recs[i]["id"] for i in pos_i if i not in held]
            train_neg = [recs[i]["id"] for i in neg_i if i not in held]
            if cls_name:
                scores = rank_with_asreview(recs, train_pos, train_neg, key)
            else:
                scores, _, _ = _rank_builtin(
                    recs, [recs[i] for i in pos_i if i not in held],
                    [recs[i] for i in neg_i if i not in held])
            if not scores:
                failed = "could not run"
                break
            # Rank only what the model was NOT trained on, which is the queue
            # you would actually be screening.
            pool = [i for i in range(len(recs))
                    if i in held or (i not in pos_i and i not in neg_i)]
            pool.sort(key=lambda i: -scores.get(recs[i]["id"], 0.0))
            at = {i: k for k, i in enumerate(pool)}
            cut = max(1, int(len(pool) * 0.10))
            rec_s.append(sum(1 for i in hp if at.get(i, 1 << 30) < cut) / len(hp))
            place_s += [at[i] / len(pool) for i in hp if i in at]
        if failed or not rec_s:
            rows.append({"key": key, "label": label, "error": failed or "no result"})
            continue
        place_s.sort()
        rows.append({
            "key": key, "label": label,
            "recall": round(sum(rec_s) / len(rec_s) * 100, 1),
            "median": round(place_s[len(place_s) // 2] * 100, 1),
        })
    ok = [r for r in rows if "recall" in r]
    ok.sort(key=lambda r: (-r["recall"], r["median"]))
    return {"rows": ok + [r for r in rows if "recall" not in r],
            "pos": len(pos_i), "neg": len(neg_i),
            "holdout": int(holdout * 100),
            "best": ok[0]["key"] if ok else ""}


def _rank_builtin(recs, pos, neg):
    """The original pure-Python ranker. No dependencies, always available."""
    docs = {r["id"]: tokenize(f"{r['title']} {r['abstract']} {r['keywords']}")
            for r in recs}
    df = Counter()
    for toks in docs.values():
        df.update(set(toks))
    n_docs = max(1, len(docs))
    import math
    idf = {w: math.log(n_docs / (1 + c)) + 1.0 for w, c in df.items()}

    def vec(ids):
        v = Counter()
        for i in ids:
            for w in docs.get(i, []):
                v[w] += idf.get(w, 1.0)
        return v

    vp, vn = vec([r["id"] for r in pos]), vec([r["id"] for r in neg])
    vocab = set(vp) | set(vn)
    tp, tn = sum(vp.values()), sum(vn.values())
    V = max(1, len(vocab))
    lp = {w: math.log((vp.get(w, 0) + 1) / (tp + V)) for w in vocab}
    ln = {w: math.log((vn.get(w, 0) + 1) / (tn + V)) for w in vocab}
    prior = math.log(len(pos) / max(1, len(neg)))

    scores = {}
    for rid, toks in docs.items():
        s = prior
        for w in toks:
            if w in lp:
                s += lp[w] - ln[w]
        scores[rid] = s / (1 + math.sqrt(len(toks) or 1))
    if scores:
        lo, hi = min(scores.values()), max(scores.values())
        rng = (hi - lo) or 1.0
        scores = {k: round((v - lo) / rng, 4) for k, v in scores.items()}
    return scores, len(pos), len(neg)


# ==========================================================================
#  Ask
# ==========================================================================

ASK_SYSTEM = """You are helping a researcher analyse the corpus of a \
systematic literature review on sustainability signals in lending-based \
crowdfunding.

Rules you must follow:
- Answer only from the records provided below. They are the only evidence \
you have.
- Cite every factual claim with the record id in square brackets, like \
[W2604622284]. A claim without a citation is not acceptable.
- If the records do not answer the question, say so plainly and say what \
would be needed. Do not fill the gap from general knowledge.
- Do not recommend whether a paper should be included in or excluded from \
the review. That decision is the researcher's alone.
- Be concise and concrete. Quantify when the records let you."""


def retrieve(question, scope_ids, k=None):
    k = k or STATE.get("ask_k") or MAX_CONTEXT_RECORDS
    recs = [r for r in STATE["records"] if r["id"] in scope_ids] \
        if scope_ids else STATE["records"]
    if not recs:
        return []
    import math
    docs = {r["id"]: tokenize(f"{r['title']} {r['abstract']} {r['keywords']}")
            for r in recs}
    df = Counter()
    for t in docs.values():
        df.update(set(t))
    n = max(1, len(docs))
    idf = {w: math.log(n / (1 + c)) + 1.0 for w, c in df.items()}
    q = tokenize(question)
    if not q:
        return recs[:k]
    qv = Counter()
    for w in q:
        qv[w] += idf.get(w, 1.0)
    qn = math.sqrt(sum(v * v for v in qv.values())) or 1.0
    scored = []
    for r in recs:
        tv = Counter()
        for w in docs[r["id"]]:
            tv[w] += idf.get(w, 1.0)
        if not tv:
            continue
        tn_ = math.sqrt(sum(v * v for v in tv.values())) or 1.0
        dot = sum(qv[w] * tv.get(w, 0) for w in qv)
        if dot:
            scored.append((dot / (qn * tn_), r))
    scored.sort(key=lambda x: -x[0])
    return [r for _, r in scored[:k]]


def ask_claude(question, hits, api_key, model):
    """Never raises. Returns (answer, error)."""
    if not api_key:
        return "", ("No API key set. Open Settings and paste an Anthropic key, "
                    "or set ANTHROPIC_API_KEY before starting the app.")
    if not hits:
        return "", "No records matched that question closely enough to quote."
    cap = STATE.get("ask_chars") or MAX_ABSTRACT_CHARS
    blocks = []
    for r in hits:
        ab = (r["abstract"] or "")[:cap]
        blocks.append(
            f"[{r['id']}] {r['title']}\n"
            f"Authors: {r['authors'][:200]}\nYear: {r['year']}  "
            f"Journal: {r['journal']}  Citations: {r['cited_by']}\n"
            f"Abstract: {ab or '(none)'}")
    payload = {
        "model": model, "max_tokens": 1600,
        "system": ASK_SYSTEM,
        "messages": [{"role": "user", "content":
                      "RECORDS:\n\n" + "\n\n---\n\n".join(blocks) +
                      f"\n\nQUESTION: {question}"}],
    }
    req = urllib.request.Request(
        ANTHROPIC_URL, data=json.dumps(payload).encode("utf-8"),
        headers={"content-type": "application/json", "x-api-key": api_key,
                 "anthropic-version": "2023-06-01"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            body = json.loads(r.read().decode("utf-8"))
        parts = [b.get("text", "") for b in body.get("content", [])
                 if b.get("type") == "text"]
        return "\n".join(parts).strip(), ""
    except urllib.error.HTTPError as e:
        try:
            detail = json.loads(e.read().decode("utf-8")) \
                .get("error", {}).get("message", "")
        except Exception:
            detail = ""
        if e.code == 401:
            return "", "The API key was rejected (401). Check it in Settings."
        if e.code == 429:
            return "", "Rate limited (429). Wait a moment and ask again."
        return "", f"API error {e.code}. {detail}"
    except Exception as e:
        return "", f"Could not reach the API: {e}"


# ==========================================================================
#  export
# ==========================================================================

EXPORT_COLUMNS = [
    # `source` and `source_file` make the export self-documenting: a reviewer
    # asking which database an included study came from can read it off the
    # row instead of taking the PRISMA diagram on trust. `key_paper` carries
    # your core/background marking through.
    "id", "source", "source_file", "key_paper",
    "doi", "title", "authors", "year", "journal", "type", "language",
    "cited_by", "topic", "subfield", "field", "institutions",
    "country", "country_source", "themes", "theme_source",
    "ta_decision", "ta_reason", "ta_source", "ft_decision", "ft_reason",
    "dup_status", "duplicate_of", "merged_fields",
    "instrument_type", "notes", "screener",
    "updated_at", "url", "apa_citation", "abstract", "abstract_source",
]


def export_rows(which="all"):
    dups, chosen_keep = dup_state()
    label_name = {l["id"]: l["name"] for l in STATE["labels"]}
    # Which record each duplicate was folded into, so a dropped row says what
    # replaced it rather than just "duplicate".
    kept_by_id = {}
    for g in STATE["dupes"]:
        keep = chosen_keep.get(g["key"], "")
        for i in g["ids"]:
            if i != keep:
                kept_by_id[i] = keep
    out = []
    for r in STATE["records"]:
        d = STATE["decisions"].get(r["id"], {})
        dec, why, how = effective(r)
        if which == "included" and not is_included(r):
            continue
        if which == "screened" and dups.get(r["id"]) == "duplicate":
            continue
        chosen = d.get("labels")
        tids = ([x.strip() for x in chosen.split("|") if x.strip()]
                if chosen else r["theme_guess"])
        out.append({
            "id": r["id"],
            "source": SOURCE_LABELS.get(r.get("source"), r.get("source", "")),
            "source_file": r.get("source_file", ""),
            "key_paper": r.get("key_paper", ""),
            "doi": r["doi"], "title": r["title"],
            "authors": r["authors"], "year": r["year"], "journal": r["journal"],
            "type": r["type"], "language": r["language"],
            "cited_by": r["cited_by"], "topic": r["topic"],
            "subfield": r["subfield"], "field": r["field"],
            "institutions": r["institutions"],
            "country": d.get("country") or r["country_guess"],
            "country_source": "edited" if d.get("country")
                              else (r["country_basis"] or ""),
            "themes": " | ".join(label_name.get(t, t) for t in tids),
            "theme_source": "edited" if chosen else "guess",
            "ta_decision": dec, "ta_reason": why, "ta_source": how,
            "ft_decision": d.get("ft_decision", ""),
            "ft_reason": d.get("ft_reason", ""),
            "dup_status": dups.get(r["id"], ""),
            "duplicate_of": kept_by_id.get(r["id"], ""),
            "merged_fields": " | ".join(
                f"{k} from {v}" for k, v in (r.get("merged_from") or {}).items()),
            "instrument_type": d.get("instrument_type", ""),
            "notes": d.get("notes", ""), "screener": d.get("screener", ""),
            "updated_at": d.get("updated_at", ""),
            "url": r["url"], "apa_citation": r["apa"], "abstract": r["abstract"],
            "abstract_source": ("added by hand" if r.get("abstract_added")
                                else r["abstract_source"]),
        })
    return out


PAGE = r"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>BarcalaReviewApp</title>
<style>
:root{
  --bg:#f4f6f9; --panel:#fff; --line:#dbe2ea; --ink:#17212b; --dim:#5b6b7c;
  --blue:#0072B2; --orange:#E69F00; --green:#009E73; --pink:#CC79A7;
  --red:#D55E00; --soft:#eef3f8;
}
*{box-sizing:border-box}
body{margin:0;font:14px/1.5 "Segoe UI",system-ui,Helvetica,Arial,sans-serif;
     background:var(--bg);color:var(--ink)}
header{background:var(--panel);border-bottom:1px solid var(--line);
       padding:10px 18px;position:sticky;top:0;z-index:30}
.hrow{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
h1{font-size:16px;margin:0;font-weight:650}
.src{color:var(--dim);font-size:12px}
nav{display:flex;gap:4px;margin-top:9px;flex-wrap:wrap}
nav button{background:none;border:1px solid transparent;border-radius:7px;
  padding:6px 13px;cursor:pointer;font:inherit;color:var(--dim)}
nav button:hover{background:var(--soft)}
nav button.on{background:var(--blue);color:#fff;font-weight:600}
main{padding:16px 18px 60px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:10px;
      padding:14px 16px;margin-bottom:14px}
.card h2{font-size:14px;margin:0 0 10px;font-weight:650}
.muted{color:var(--dim)}
.small{font-size:12px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
input,select,textarea{font:inherit;padding:6px 9px;border:1px solid var(--line);
  border-radius:7px;background:#fff;color:var(--ink)}
input[type=text],input[type=search]{min-width:200px}
textarea{width:100%;min-height:64px;resize:vertical}
button.b{background:var(--blue);color:#fff;border:none;border-radius:7px;
  padding:7px 13px;cursor:pointer;font:inherit}
button.g{background:var(--green)} button.o{background:var(--orange);color:#1a1a1a}
button.r{background:var(--red)} button.n{background:#fff;color:var(--ink);
  border:1px solid var(--line)}
button.b:hover{filter:brightness(1.08)}
button:disabled{opacity:.5;cursor:not-allowed}
.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:11px;
  background:var(--soft);color:var(--dim);white-space:nowrap}
.pill.auto{background:#fdf0e3;color:#8a4b00}
.pill.man{background:#e4f3ec;color:#0a5c42}
.pill.guess{background:#eef1f6;color:#4a5668;border:1px dashed #b9c4d2}
.pill.inc{background:#def2e9;color:#0a5c42}
.pill.exc{background:#fde8e0;color:#8a2b00}
.pill.may{background:#fdf3d6;color:#7a5c00}
.split{display:grid;grid-template-columns:minmax(360px,1.1fr) minmax(380px,1fr);
       gap:14px;align-items:start}
@media(max-width:1000px){.split{grid-template-columns:1fr}}
.list{max-height:70vh;overflow:auto;border:1px solid var(--line);border-radius:9px}
.item{padding:9px 12px;border-bottom:1px solid var(--line);cursor:pointer}
.item:last-child{border-bottom:none}
.item:hover{background:var(--soft)}
.item.sel{background:#e8f1fa;box-shadow:inset 3px 0 0 var(--blue)}
.item .t{font-weight:600;font-size:13px;margin-bottom:2px}
.item .m{font-size:11.5px;color:var(--dim)}
.item.ticked{background:#eef6ef;box-shadow:inset 3px 0 0 var(--green)}
.item.ticked.sel{background:#e4f0ef}
input.tick{margin-right:7px;vertical-align:middle;cursor:pointer}
table{border-collapse:collapse;width:100%;font-size:13px}
th,td{text-align:left;padding:6px 9px;border-bottom:1px solid var(--line);
      vertical-align:top}
th{font-weight:600;color:var(--dim);font-size:12px}
.num{text-align:right;font-variant-numeric:tabular-nums}
.ab{max-height:260px;overflow:auto;background:var(--soft);padding:10px;
    border-radius:8px;font-size:13px;line-height:1.55}
label.ck{display:flex;gap:7px;align-items:center;cursor:pointer;font-size:13px}
.rule{border:1px solid var(--line);border-radius:9px;padding:11px 13px;
      margin-bottom:9px}
.rule.on{border-color:var(--blue);background:#f6fafd}
.rule .top{display:flex;justify-content:space-between;gap:12px;align-items:center}
.bar{height:15px;border-radius:3px}
.key{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;margin:6px 0 10px}
.key i{width:11px;height:11px;border-radius:3px;display:inline-block;
       margin-right:5px;vertical-align:-1px}
.flash{position:fixed;right:16px;bottom:16px;background:#17212b;color:#fff;
  padding:9px 14px;border-radius:8px;font-size:13px;opacity:0;
  transition:opacity .2s;z-index:50}
.flash.on{opacity:.95}
.kbd{font:11px ui-monospace,Menlo,Consolas,monospace;background:var(--soft);
  border:1px solid var(--line);border-radius:4px;padding:0 4px}
a{color:var(--blue)}
.grid3{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:12px}
.scrollbox{max-height:420px;overflow:auto}
.scrollbox table{font-size:12px}
.scrollbox th{position:sticky;top:0;background:var(--panel)}
.stat{background:var(--soft);border-radius:9px;padding:10px 12px}
.stat .n{font-size:21px;font-weight:650}
.stat .l{font-size:11.5px;color:var(--dim)}
.panels{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,330px),1fr));
  gap:12px}
.panel{border:1px solid var(--line);border-radius:9px;padding:11px 13px}
.panel .pt{font-weight:650;font-size:13px}
.panel table{width:100%}
.panel td{border:none}
</style></head><body>

<header>
  <div id="saveerr" style="display:none;background:#fde8e0;color:#8a2b00;
      border:1px solid #e0a080;border-radius:8px;padding:9px 13px;
      margin-bottom:9px;font-size:13px"></div>
  <div class="hrow">
    <h1>BarcalaReviewApp</h1>
    <span class="src" id="src"></span>
    <span style="flex:1"></span>
    <label class="small muted">Screener
      <input type="text" id="screener" style="min-width:120px" placeholder="your name">
    </label>
  </div>
  <nav id="tabs"></nav>
</header>

<main>
  <div id="view"></div>
</main>
<div class="flash" id="flash"></div>

<script>
const TABS=["Screen","Sources","Duplicates","Rules","Stopping","Dashboard",
            "Bibliometrics","PRISMA","Export","Ask","Settings"];
let D=null, tab="Screen", sel=null, shown=300, scores=null;
let F={q:"",dec:"",theme:"",country:"",lang:"",src:"",yr:"",flag:"",dup:"live",sort:""};
let DF={group:"year",split:"",dec:"",theme:"",country:"",lang:"",yr:""};

const $=s=>document.querySelector(s);
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",
  ">":"&gt;",'"':"&quot;"}[c]));
function flash(m){const f=$("#flash");f.textContent=m;f.classList.add("on");
  clearTimeout(f._t);f._t=setTimeout(()=>f.classList.remove("on"),1800);}

async function api(path,opt){const r=await fetch(path,opt);
  let j={}; try{j=await r.json();}catch(e){}
  if(j && j.error) flash(j.error);
  return j;}

async function load(){
  D=await api("/api/data");
  $("#src").textContent=D.source+" · "+D.records.length.toLocaleString()+" records";
  $("#screener").value=D.screener||"";
  render();
}

// ---------- helpers on a record ----------
const labName=id=>{const l=(D.labels||[]).find(x=>x.id===id);return l?l.name:id;};
// A short badge for the list. The first word of a theme name is usually the
// code you gave it ("T2 Signals" -> "T2"), but only if those first words are
// distinct — themes called "Theme 1" and "Theme 2" would all show "Theme".
const labShort=id=>{
  const labs=D.labels||[];
  const firsts=labs.map(l=>(l.name||"").trim().split(/\s+/)[0]);
  const unique=new Set(firsts).size===firsts.length;
  const l=labs.find(x=>x.id===id);
  if(!l) return id.toUpperCase();
  return unique?((l.name||"").trim().split(/\s+/)[0]||id.toUpperCase())
               :id.toUpperCase();
};
function themesOf(r){ return r.labels ? r.labels.split("|").map(s=>s.trim()).filter(Boolean)
                                      : (r.tguess||[]); }
function themeEdited(r){ return !!r.labels; }
function countryOf(r){ return r.country || r.cguess || ""; }
function countryEdited(r){ return !!r.country; }
function isOldUncited(r){
  const y=parseInt(r.year||0,10); if(!y) return false;
  return (D.year - y) >= 5 && (parseInt(r.cited_by||0,10)===0);
}
function effPill(r){
  if(r.eff==="include") return '<span class="pill inc">include</span>';
  if(r.eff==="maybe")   return '<span class="pill may">maybe</span>';
  if(r.eff==="exclude") return '<span class="pill exc">exclude</span>'+
      (r.eff_src==="auto"?' <span class="pill auto">rule</span>':
                          ' <span class="pill man">by hand</span>');
  return '<span class="pill">not screened</span>';
}

function matches(r,f){
  if(f.dup==="live" && r.dup==="duplicate") return false;
  if(f.dup==="dup" && r.dup!=="duplicate") return false;
  if(f.q){const q=f.q.toLowerCase();
    if(!((r.title||"")+" "+(r.abstract||"")+" "+(r.authors||"")+" "+
         (r.journal||"")+" "+(r.doi||"")).toLowerCase().includes(q)) return false;}
  if(f.dec==="todo"   && r.eff) return false;
  if(f.dec==="include"&& r.eff!=="include") return false;
  if(f.dec==="maybe"  && r.eff!=="maybe") return false;
  if(f.dec==="exclude"&& r.eff!=="exclude") return false;
  if(f.dec==="auto"   && !(r.eff==="exclude"&&r.eff_src==="auto")) return false;
  if(f.dec==="manual" && r.eff_src!=="manual") return false;
  if(f.theme && !themesOf(r).includes(f.theme)) return false;
  if(f.country && !countryOf(r).includes(f.country)) return false;
  if(f.lang && (r.language||"")!==f.lang) return false;
  if(f.yr){const y=parseInt(r.year||0,10);
    if(f.yr==="old"&&!(y&&y<=D.year-5)) return false;
    if(f.yr==="new"&&!(y&&y>D.year-5)) return false;}
  if(f.flag==="noabs"&&r.abstract) return false;
  if(f.flag==="nodoi"&&r.doi) return false;
  if(f.flag==="noneu"&&(r.europe||"").toLowerCase()==="yes") return false;
  if(f.flag==="notes"&&(r.language||"")==="en") return false;
  if(f.flag==="olduncited"&&!isOldUncited(r)) return false;
  if(f.flag==="retracted"&&!r.retracted) return false;
  if(f.flag==="needsinfo"&&r.doi&&r.abstract) return false;
  if(f.src==="new"  && r.src==="openalex") return false;
  if(f.src==="newtodo"&&(r.src==="openalex"||r.eff)) return false;
  if(f.src&&!["","new","newtodo"].includes(f.src)&&r.src!==f.src) return false;
  return true;
}
const filtered=f=>sortList(D.records.filter(r=>matches(r,f)),f.sort);

// Sorting is applied at render time rather than by shuffling the records, so
// switching between orders never loses the model scores or your place.
function sortList(list,how){
  if(!how) return list;
  const sc=r=>(scores&&scores[r.id]!=null)?scores[r.id]:-1;
  const n=r=>parseInt(r.cited_by||0,10)||0;
  const y=r=>parseInt(r.year||0,10)||0;
  const out=list.slice();
  if(how==="score")       out.sort((a,b)=>sc(b)-sc(a));
  else if(how==="cites")  out.sort((a,b)=>n(b)-n(a));
  else if(how==="score_cites") out.sort((a,b)=>(sc(b)-sc(a))||(n(b)-n(a)));
  else if(how==="cites_score") out.sort((a,b)=>(n(b)-n(a))||(sc(b)-sc(a)));
  else if(how==="new")    out.sort((a,b)=>y(b)-y(a));
  else if(how==="old")    out.sort((a,b)=>y(a)-y(b));
  return out;
}

// ---------- chrome ----------
function render(){
  $("#tabs").innerHTML=TABS.map(t=>
    `<button class="${t===tab?"on":""}" onclick="go('${t}')">${t}</button>`).join("");
  const v=$("#view");
  if(tab==="Screen") v.innerHTML=screenTab();
  else if(tab==="Duplicates") v.innerHTML=dupTab();
  else if(tab==="Rules") v.innerHTML=rulesTab();
  else if(tab==="Sources") v.innerHTML=sourcesTab();
  else if(tab==="Stopping"){ v.innerHTML=stoppingShell(); loadStopping(); }
  else if(tab==="Dashboard") v.innerHTML=dashTab();
  else if(tab==="Bibliometrics"){ v.innerHTML=biblioShell(); loadBiblio(); }
  else if(tab==="PRISMA") {v.innerHTML=prismaTab(); loadPrisma();}
  else if(tab==="Export") v.innerHTML=exportTab();
  else if(tab==="Ask") v.innerHTML=askTab();
  else { v.innerHTML=settingsTab(); askEstimate(); loadBulk(); }
  if(tab==="Dashboard"){ drawPanels(); drawChart(); }
  if(tab==="Screen") bulkBar();
}
window.go=t=>{tab=t;shown=300;render();};

// ---------- Bibliometrics ----------
// A description of the literature itself, separate from the Dashboard, which
// is about where your screening stands. Counts only: nothing here is a
// judgement and nothing is written.
window.biblioWhich="screened";
function biblioShell(){
  const pick=(v,l)=>`<button class="b ${biblioWhich===v?"":"n"}"
    onclick="biblioWhich='${v}';loadBiblio()">${l}</button>`;
  return `<div class="card">
    <h2>Bibliometrics</h2>
    <p class="small muted">A profile of the literature you have gathered, for
    the descriptive part of your results. The Dashboard tracks your screening;
    this describes the field. Counts only &mdash; nothing here changes a
    record.</p>
    <div class="row" style="gap:6px;flex-wrap:wrap">
      ${pick("screened","after duplicates removed")}
      ${pick("included","included studies only")}
      ${pick("all","everything, duplicates and all")}
      <span style="flex:1"></span>
      <a class="b n" style="text-decoration:none;padding:7px 13px"
         href="#" onclick="downloadBiblio();return false">download as CSV</a>
    </div>
  </div><div id="bib" class="muted">loading&hellip;</div>`;
}
window.bibData=null;
async function loadBiblio(){
  const el=$("#bib"); if(el) el.innerHTML='<div class="card muted">counting&hellip;</div>';
  const j=await api("/api/bibliometrics?which="+encodeURIComponent(biblioWhich));
  bibData=j;
  const el2=$("#bib"); if(!el2) return;
  const bar=(rows,unit)=>{
    if(!rows||!rows.length) return '<p class="small muted">nothing to show</p>';
    const max=Math.max(...rows.map(r=>r.n))||1;
    return `<table>${rows.map(r=>`<tr>
      <td style="width:46%">${esc(r.name).slice(0,70)}</td>
      <td><div style="background:var(--blue);height:13px;border-radius:3px;
        width:${Math.max(2,r.n/max*100)}%"></div></td>
      <td class="num" style="width:68px">${r.n.toLocaleString()}${unit||""}</td>
      </tr>`).join("")}</table>`;};
  const panel=(title,note,rows,unit)=>`<div class="card">
    <h2>${title}</h2>${note?`<p class="small muted">${note}</p>`:""}
    ${bar(rows,unit)}</div>`;
  el2.innerHTML=`
  <div class="card"><div class="grid3">
    <div class="stat"><div class="n">${j.records.toLocaleString()}</div><div class="l">records</div></div>
    <div class="stat"><div class="n">${j.citations_total.toLocaleString()}</div><div class="l">citations in total</div></div>
    <div class="stat"><div class="n">${j.citations_median}</div><div class="l">median citations</div></div>
    <div class="stat"><div class="n">${j.h_index}</div><div class="l">h-index of this set</div></div>
    <div class="stat"><div class="n">${j.uncited.toLocaleString()}</div><div class="l">never cited</div></div>
    <div class="stat"><div class="n">${j.with_abstract.toLocaleString()}</div><div class="l">with an abstract</div></div>
  </div>
  <p class="small muted" style="margin-top:9px">The h-index here describes the
  SET: ${j.h_index} of these papers have at least ${j.h_index} citations each.
  Citation counts come from the database that supplied each record and are a
  snapshot, so report the date you took them.</p></div>
  <div class="grid2">
    ${panel("Publications per year","",j.years)}
    ${panel("Average citations per paper, by year",
      "Recent years look weak because citations take time to accumulate, not because the work is weaker.",j.cites_per_year)}
    ${panel("Most prolific authors","By number of papers in this set.",j.authors)}
    <div class="card"><h2>Most cited authors</h2>
      <p class="small muted">Total citations across their papers in this set.
      Citations belong to a paper, so every co-author of a heavily cited one
      carries its whole count &mdash; the paper count makes that visible.</p>
      ${(j.author_impact||[]).length?`<table>${(()=>{const mx=Math.max(...j.author_impact.map(r=>r.n))||1;
        return j.author_impact.map(r=>`<tr>
          <td style="width:42%">${esc(r.name).slice(0,60)}</td>
          <td><div style="background:var(--blue);height:13px;border-radius:3px;
            width:${Math.max(2,r.n/mx*100)}%"></div></td>
          <td class="num" style="width:70px">${r.n.toLocaleString()}</td>
          <td class="num small muted" style="width:82px">${r.papers} paper${r.papers===1?"":"s"}</td>
          </tr>`).join("");})()}</table>`:'<p class="small muted">nothing to show</p>'}
    </div>
    ${panel("Journals","",j.journals)}
    ${panel("Institutions","",j.institutions)}
    ${panel("Keywords","",j.keywords)}
    ${panel("Countries","As recorded or guessed for each record.",j.countries)}
    ${panel("Fields","",j.fields)}
    ${panel("Publication types","",j.types)}
    ${panel("Co-authorship pairs","Authors appearing together most often.",j.collaborations)}
    ${panel("Where the records came from","",j.sources)}
  </div>`;
}
window.downloadBiblio=()=>{
  if(!bibData){flash("nothing loaded yet");return;}
  const rows=[["panel","name","value"]];
  [["publications per year","years"],["citations per year","cites_per_year"],
   ["prolific authors","authors"],["cited authors","author_impact"],
   ["journals","journals"],["institutions","institutions"],
   ["keywords","keywords"],["countries","countries"],["fields","fields"],
   ["types","types"],["languages","languages"],["open access","open_access"],
   ["co-authorship","collaborations"],["sources","sources"]].forEach(([lab,key])=>
    (bibData[key]||[]).forEach(r=>rows.push([lab,r.name,r.n])));
  [["summary","records",bibData.records],["summary","citations total",bibData.citations_total],
   ["summary","citations median",bibData.citations_median],["summary","h-index",bibData.h_index],
   ["summary","never cited",bibData.uncited]].forEach(r=>rows.push(r));
  const csv="\\ufeff"+rows.map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(";")).join("\\r\\n");
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
  a.download=`bibliometrics_${bibData.which}_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
};

// ---------- Screen ----------
function screenTab(){
  const list=filtered(F);
  const langs=[...new Set(D.records.map(r=>r.language).filter(Boolean))].sort();
  const countries=[...new Set(D.records.flatMap(r=>
      countryOf(r).split("|").map(s=>s.trim()).filter(Boolean)))].sort();
  const done=D.records.filter(r=>r.dup!=="duplicate"&&r.eff).length;
  const live=D.records.filter(r=>r.dup!=="duplicate").length;
  return `
  <div class="card">
    <div class="row">
      <input type="search" id="q" placeholder="search title, abstract, author, DOI"
        value="${esc(F.q)}" oninput="F.q=this.value;shown=300;render();$('#q').focus()">
      <select onchange="F.dec=this.value;shown=300;render()">
        ${opt([["","any decision"],["todo","not screened"],["include","included"],
               ["maybe","maybe"],["exclude","excluded"],
               ["auto","excluded by a rule"],["manual","decided by hand"]],F.dec)}
      </select>
      <select onchange="F.theme=this.value;shown=300;render()">
        <option value="">any theme</option>
        ${(D.labels||[]).map(l=>`<option value="${l.id}" ${F.theme===l.id?"selected":""}>${esc(l.name)}</option>`).join("")}
      </select>
      <select onchange="F.country=this.value;shown=300;render()">
        <option value="">any country</option>
        ${countries.map(c=>`<option ${F.country===c?"selected":""}>${esc(c)}</option>`).join("")}
      </select>
      <select onchange="F.lang=this.value;shown=300;render()">
        <option value="">any language</option>
        ${langs.map(c=>`<option ${F.lang===c?"selected":""}>${esc(c)}</option>`).join("")}
      </select>
      <select onchange="F.src=this.value;shown=300;render()">
        ${opt([["","any source"],
               ["newtodo","newly added, not yet screened"],
               ["new","everything except the OpenAlex harvest"],
               ...(D.sources||[]).map(s=>[s.key,s.label])],F.src)}
      </select>
      <select onchange="F.flag=this.value;shown=300;render()">
        ${opt([["","no flag filter"],["noabs","no abstract"],["nodoi","no DOI"],
               ["needsinfo","missing a DOI or an abstract"],
               ["noneu","no European signal"],["notes","not English"],
               ["olduncited","5+ years old, never cited"],
               ["retracted","retracted"]],F.flag)}
      </select>
      <select onchange="F.dup=this.value;shown=300;render()">
        ${opt([["live","hide duplicates"],["all","show everything"],
               ["dup","only duplicates"]],F.dup)}
      </select>
      <button class="b n" onclick="F={q:'',dec:'',theme:'',country:'',lang:'',src:'',yr:'',flag:'',dup:'live',sort:F.sort};shown=300;render()">clear</button>
    </div>
    <div class="row small muted" style="margin-top:8px">
      <span><b>${list.length.toLocaleString()}</b> shown</span>
      <span>${done.toLocaleString()} of ${live.toLocaleString()} screened</span>
      <span><span class="kbd">i</span> include <span class="kbd">m</span> maybe
            <span class="kbd">e</span> exclude <span class="kbd">j</span>/<span class="kbd">k</span> move</span>
      <span style="flex:1"></span>
      <label class="small muted">order by
        <select onchange="F.sort=this.value;shown=300;render()">
          ${opt([["","the order they were harvested"],
                 ["score","how likely I am to include it"],
                 ["cites","citations"],
                 ["score_cites","likely to include, then citations"],
                 ["cites_score","citations, then likely to include"],
                 ["new","newest first"],["old","oldest first"]],F.sort)}
        </select></label>
      <button class="b n" onclick="rank()">${scores?"re-learn":"learn from my decisions"}</button>
      ${rankInfo&&rankInfo.asreview?`<select onchange="rank(this.value)"
        title="which model does the ranking">
        ${opt(Object.entries(rankInfo.models||{})
              .concat([["builtin","built-in (no ASReview)"]]),
              rankInfo.model)}
      </select>`:""}
      ${scores?`<button class="b n" onclick="compareModels()"
        title="measure the models against your own decisions">compare models</button>`:""}
    </div>
    <div id="cmp">${cmpOut||""}</div>
    ${rankEngine?`<p class="small muted" style="margin-top:7px">Ranked by
      <b>${esc(rankEngine)}</b>. The model only changes the ORDER you see
      records in &mdash; it never labels anything, and your decisions are the
      only thing it learns from.</p>`:""}
    ${(F.sort||"").includes("score")&&!scores?`<p class="small muted"
      style="margin-top:7px">No model yet &mdash; press
      <b>learn from my decisions</b>. It needs at least 3 includes and 3 of
      your own excludes, and ignores anything a rule excluded.</p>`:""}
    ${rankInfo&&!rankInfo.asreview?`<p class="small muted" style="margin-top:7px">
      Using the built-in ranker. For ASReview's own models, install it once:
      <code>pip install asreview</code>, then restart this app.</p>`:""}
  </div>
  <div id="bulkbar"></div>
  <div class="row" style="margin-top:8px">
    <button class="b n" onclick="pickAllShown()">${picked.size?"select / clear all shown":"select records by title"}</button>
    <span class="small muted">tick the box on a title to add it to a batch</span>
  </div>
  <div class="split">
    <div>
      <div class="list" id="list">
        ${list.slice(0,shown).map(itemHtml).join("")||
          '<div style="padding:18px" class="muted">Nothing matches these filters.</div>'}
      </div>
      ${list.length>shown?`<div style="margin-top:9px"><button class="b n"
        onclick="shown+=500;render()">show 500 more (${(list.length-shown).toLocaleString()} left)</button></div>`:""}
    </div>
    <div id="detail">${sel?detailHtml(byId(sel)):
      '<div class="card muted">Pick a record on the left.</div>'}</div>
  </div>`;
}
function opt(pairs,cur){return pairs.map(([v,l])=>
  `<option value="${v}" ${cur===v?"selected":""}>${esc(l)}</option>`).join("");}
const byId=id=>D.records.find(r=>r.id===id);

function itemHtml(r){
  const sc=scores&&scores[r.id]!=null?` <span class="pill">${(scores[r.id]*100).toFixed(0)}</span>`:"";
  const th=themesOf(r).map(labShort).join(" ");
  return `<div class="item ${sel===r.id?"sel":""} ${picked.has(r.id)?"ticked":""}"
      onclick="pick('${r.id}')">
    <div class="t"><input type="checkbox" class="tick" data-id="${r.id}" ${picked.has(r.id)?"checked":""}
      onclick="event.stopPropagation();togglePick('${r.id}')"
      title="select for a decision on several at once">
      ${esc(r.title||"(no title)")}</div>
    <div class="m">${esc(r.year||"?")} &middot; ${esc((r.journal||"no journal").slice(0,52))}
      &middot; ${r.cited_by} cites ${sc}</div>
    <div class="m" style="margin-top:3px">${effPill(r)}
      ${th?`<span class="pill">${esc(th)}</span>`:""}
      ${countryOf(r)?`<span class="pill">${esc(countryOf(r).split("|")[0].trim())}</span>`:""}
      ${!r.abstract?'<span class="pill">no abstract</span>':""}
      ${r.dup==="duplicate"?'<span class="pill exc">duplicate</span>':""}
    </div></div>`;
}
window.pick=id=>{sel=id;$("#detail").innerHTML=detailHtml(byId(id));
  document.querySelectorAll(".item").forEach(e=>e.classList.remove("sel"));
  event&&event.currentTarget&&event.currentTarget.classList.add("sel");};

// ---------- deciding on several records at once ----------
// Screening on titles alone is the right move when a title settles it, and
// wrong when it does not. So this is opt-in per record: tick the ones a
// title has already decided, judge the rest on the abstract as usual. Every
// record in a batch gets the same reason, recorded individually, so the
// PRISMA reasons table stays correct and each decision is still auditable.
window.picked=new Set();
// Ticking a box must not rebuild the list. With several hundred records on
// screen a full re-render loses your scroll position on every click, which
// makes selecting twenty titles unusable. So the row is updated in place and
// only the action bar is redrawn.
window.togglePick=id=>{
  picked.has(id)?picked.delete(id):picked.add(id);
  const box=document.querySelector(`input.tick[data-id="${id}"]`);
  if(box){ box.checked=picked.has(id);
           box.closest(".item").classList.toggle("ticked",picked.has(id)); }
  bulkBar();
};
window.pickAllShown=()=>{
  const list=filtered(F).slice(0,shown);
  const allOn=list.every(r=>picked.has(r.id));
  list.forEach(r=>allOn?picked.delete(r.id):picked.add(r.id));
  document.querySelectorAll("input.tick").forEach(b=>{
    const on=picked.has(b.dataset.id);
    b.checked=on; b.closest(".item").classList.toggle("ticked",on);});
  bulkBar();
};
window.clearPicked=()=>{
  picked.clear();
  document.querySelectorAll("input.tick").forEach(b=>{
    b.checked=false; b.closest(".item").classList.remove("ticked");});
  bulkBar();
};
function bulkBar(){
  const el=$("#bulkbar"); if(!el) return;
  if(!picked.size){ el.innerHTML=""; return; }
  el.innerHTML=`<div class="card" style="margin-top:10px;border-color:var(--blue)">
    <div class="row" style="flex-wrap:wrap;gap:8px;align-items:center">
      <b>${picked.size.toLocaleString()} selected</b>
      <span class="small muted">decide on the title alone &mdash; use this only
        where the title settles it</span>
      <span style="flex:1"></span>
      <button class="b g" onclick="decidePicked('include')">include all</button>
      <button class="b o" onclick="decidePicked('maybe')">maybe</button>
      <select id="bulkreason" style="min-width:230px">
        <option value="">reason for excluding&hellip;</option>
        ${(D.criteria.exclude||[]).map(r=>`<option>${esc(r)}</option>`).join("")}
      </select>
      <button class="b r" onclick="decidePicked('exclude')">exclude all</button>
      <button class="b n" onclick="clearPicked()">clear</button>
    </div></div>`;
}
window.decidePicked=async(decision)=>{
  const ids=[...picked];
  if(!ids.length){flash("nothing selected");return;}
  let reason="";
  if(decision==="exclude"){
    reason=$("#bulkreason")?$("#bulkreason").value:"";
    if(!reason){flash("pick a reason first");return;}
  }
  if(!confirm(`${decision} ${ids.length} record(s)`
      +(reason?` as "${reason}"`:"")+"?\n\nEach one is recorded separately and "
      +"you can change any of them afterwards.")) return;
  const patch={ta_decision:decision,ta_reason:reason,ta_source:"title-only"};
  await api("/api/bulk",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({ids,patch})});
  picked.clear();
  await load();
  flash(`${ids.length} decided on the title`);
};

function detailHtml(r){
  if(!r) return '<div class="card muted">Pick a record.</div>';
  const reasons=(D.criteria.exclude||[]);
  const th=themesOf(r);
  return `<div class="card">
    <div style="font-weight:650;font-size:15px;margin-bottom:4px">${esc(r.title)}</div>
    <div class="small muted">${esc(r.authors||"(no authors)")}</div>
    <div class="small muted" style="margin:5px 0 9px">
      ${esc(r.year||"?")} &middot; ${esc(r.journal||"no journal")}
      &middot; ${r.cited_by} citations &middot; ${esc(r.type||"?")}
      &middot; ${esc(r.language||"lang?")}
      ${r.doi?` &middot; <a href="https://doi.org/${esc(r.doi)}" target="_blank" rel="noopener">${esc(r.doi)}</a>`:" &middot; no DOI"}
      &middot; <span class="pill">${esc(r.src_label)}</span>
      ${r.edited.length?`<span class="pill man">edited: ${r.edited.join(", ")}</span>`:""}
    </div>
    <details ${(!r.doi||!r.abstract||r.edited.length)?"open":""}
      style="margin-bottom:9px">
      <summary class="small muted">fill in or correct the details${
        (!r.doi||!r.abstract)?' <span class="pill exc">something is missing</span>':""}</summary>
      <p class="small muted" style="margin:6px 0">What you type here replaces
      what the source gave, and is stored separately &mdash; the harvest and
      every import stay as they arrived. Clear a box to go back to the
      original.</p>
      <div class="grid3">
        ${[["doi_manual","DOI",r.doi],["year_manual","Year",r.year],
           ["journal_manual","Journal",r.journal],
           ["authors_manual","Authors",r.authors],
           ["title_manual","Title",r.title]].map(([f,lab,val])=>`
          <div><div class="small muted">${lab}</div>
          <input type="text" style="width:100%" value="${esc(val||"")}"
            onchange="save('${r.id}',{${f}:this.value})"></div>`).join("")}
        <div><div class="small muted">Look it up by DOI</div>
          <button class="b n" style="width:100%"
            onclick="fillFromDoi('${r.id}')">fetch from Crossref</button></div>
      </div>
    </details>
    <div style="margin-bottom:9px">${effPill(r)}
      ${r.eff_reason?`<span class="pill">${esc(r.eff_reason)}</span>`:""}
      ${r.auto&&r.eff_src==="manual"?`<span class="pill auto">a rule would have said: ${esc(r.auto)}</span>`:""}
    </div>
    ${r.abstract?`<div class="ab">${esc(r.abstract)}
       ${r.abstract_added?'<div style="margin-top:7px"><span class="pill man">abstract added by hand</span> <a href="#" onclick="editAbs(\'' + r.id + '\');return false">edit</a></div>':""}</div>`
      :`<div class="ab" style="background:#fff7f0">
         <b>No abstract.</b> OpenAlex, Semantic Scholar and Crossref all had none —
         usually a publisher that does not release them. Paste it in here and it
         is used everywhere: the rules, the ranking, the dashboard and every export.
         <div style="margin-top:8px"><textarea id="absbox" placeholder="paste the abstract"></textarea>
         <button class="b" style="margin-top:6px"
           onclick="save('${r.id}',{abstract_manual:$('#absbox').value})">save abstract</button></div>
       </div>`}
    <div id="absedit"></div>

    <div class="row" style="margin-top:12px">
      <button class="b g" onclick="decide('${r.id}','include','')">Include</button>
      <button class="b o" onclick="decide('${r.id}','maybe','')">Maybe</button>
      <select id="rsn">${reasons.map(x=>`<option>${esc(x)}</option>`).join("")}</select>
      <button class="b r" onclick="decide('${r.id}','exclude',$('#rsn').value)">Exclude</button>
      ${r.ta?`<button class="b n" onclick="decide('${r.id}','','')">undo my decision</button>`:""}
    </div>

    <div style="margin-top:14px" class="grid3">
      <div>
        <div class="small muted">Country studied
          ${countryEdited(r)?'<span class="pill man">edited</span>':
            (r.cguess?`<span class="pill guess">guess from ${esc(r.cbasis)}</span>`:"")}</div>
        <input type="text" style="width:100%" value="${esc(countryOf(r))}"
          onchange="save('${r.id}',{country:this.value})" placeholder="e.g. Germany">
      </div>
      <div>
        <div class="small muted">Instrument</div>
        <input type="text" style="width:100%" value="${esc(r.instrument_type||"")}"
          onchange="save('${r.id}',{instrument_type:this.value})"
          placeholder="crowdlending, P2P...">
      </div>
      <div>
        <div class="small muted">Themes
          ${themeEdited(r)?'<span class="pill man">edited</span>':
            (r.tguess.length?'<span class="pill guess">guess</span>':"")}</div>
        ${(D.labels||[]).map(l=>`<label class="ck"><input type="checkbox"
           ${th.includes(l.id)?"checked":""}
           onchange="toggleTheme('${r.id}','${l.id}',this.checked)">
           <span style="color:${l.colour}">&#9632;</span> ${esc(l.name)}</label>`).join("")}
      </div>
    </div>

    <div style="margin-top:12px">
      <div class="small muted">Notes</div>
      <textarea onchange="save('${r.id}',{notes:this.value})">${esc(r.notes||"")}</textarea>
    </div>

    <details style="margin-top:10px"><summary class="small muted">full-text stage</summary>
      <div class="row" style="margin-top:8px">
        <button class="b g" onclick="save('${r.id}',{ft_decision:'include',ft_reason:''})">FT include</button>
        <select id="frsn">${reasons.map(x=>`<option>${esc(x)}</option>`).join("")}</select>
        <button class="b r" onclick="save('${r.id}',{ft_decision:'exclude',ft_reason:$('#frsn').value})">FT exclude</button>
        ${r.ft?`<span class="pill">${esc(r.ft)}${r.ft_reason?": "+esc(r.ft_reason):""}</span>`:""}
      </div>
    </details>
    <div class="small muted" style="margin-top:10px">
      id ${esc(r.id)} &middot; ${esc(r.field||"no field")}
      ${r.topic?" &middot; "+esc(r.topic):""}
      ${r.institutions?"<br>"+esc(r.institutions.slice(0,180)):""}
    </div>
  </div>`;
}

window.fillFromDoi=async(id)=>{
  const r=byId(id);
  if(!r.doi){flash("this record has no DOI — type one first");return;}
  flash("asking Crossref…");
  const j=await api("/api/lookup-doi",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({doi:r.doi})});
  if(j.error){flash(j.error);return;}
  const c=j.record, patch={};
  // only fill what is actually blank — never overwrite what you have
  if(!r.title&&c.title)       patch.title_manual=c.title;
  if(!r.authors&&c.authors)   patch.authors_manual=c.authors;
  if(!r.year&&c.year)         patch.year_manual=c.year;
  if(!r.journal&&c.journal)   patch.journal_manual=c.journal;
  if(!r.abstract&&c.abstract) patch.abstract_manual=c.abstract;
  if(!Object.keys(patch).length){flash("Crossref added nothing new");return;}
  await save(id,patch);
  flash("filled: "+Object.keys(patch).map(k=>k.replace("_manual","")).join(", "));
};
window.editAbs=id=>{
  const r=byId(id);
  $("#absedit").innerHTML=`<div style="margin-top:9px">
    <textarea id="absbox2">${esc(r.abstract||"")}</textarea>
    <button class="b" style="margin-top:6px"
      onclick="save('${id}',{abstract_manual:$('#absbox2').value})">save</button>
    <button class="b n" style="margin-top:6px"
      onclick="save('${id}',{abstract_manual:''})">remove mine</button></div>`;
};
window.decide=(id,d,reason)=>save(id,{ta_decision:d,ta_reason:reason});
window.toggleTheme=(id,lid,on)=>{
  const r=byId(id); let t=themesOf(r).slice();
  if(on){ if(!t.includes(lid)) t.push(lid); } else t=t.filter(x=>x!==lid);
  save(id,{labels:t.join("|")});
};
function saveError(msg){
  const b=$("#saveerr");
  if(!msg){ b.style.display="none"; return; }
  b.innerHTML="<b>Not saved.</b> "+esc(msg)+
    ' <a href="#" onclick="saveError(\'\');return false">dismiss</a>';
  b.style.display="block";
}
async function save(id,patch){
  const j=await api("/api/save",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({id,patch})});
  // A failed write must not look like a success. The banner stays until
  // dismissed or until the next save goes through.
  if(j.error){ saveError(j.error); return; }
  saveError("");
  if(j.record){
    const i=D.records.findIndex(r=>r.id===id);
    if(i>=0) D.records[i]=j.record;
    if(j.preview) D.preview=j.preview;
    if(tab==="Screen"){ render(); flash("saved"); }
    else render();
  }
}
window.rankEngine="";
window.rankInfo=null;
window.cmpOut="";
window.compareModels=async()=>{
  const el=$("#cmp"); if(el) el.innerHTML=
    '<p class="small muted" style="margin-top:8px">measuring&hellip; this takes '
    +'a few seconds, it trains each model several times</p>';
  const j=await api("/api/compare-models");
  if(j.error){ if(el) el.innerHTML=`<p class="small muted" style="margin-top:8px">${esc(j.error)}</p>`; return; }
  cmpOut=`<div class="card" style="margin-top:10px">
    <h2>Which model suits your review</h2>
    <p class="small muted">${j.holdout}% of your ${j.pos} includes and ${j.neg}
    excludes were hidden, each model trained on the rest, and we checked where
    the hidden relevant papers landed in the queue it produced. Averaged over
    three splits. <b>Screening in a random order would find 10%</b> &mdash;
    everything above that is the model saving you work.</p>
    <table><tr><th>model</th>
      <th class="num">relevant found in the first 10%</th>
      <th class="num">typical position of a relevant paper</th></tr>
    ${j.rows.map(r=>r.error?`<tr><td>${esc(r.label)}</td>
        <td class="num muted" colspan="2">${esc(r.error)}</td></tr>`:
      `<tr${r.key===j.best?' style="background:var(--soft)"':''}>
        <td>${esc(r.label)}${r.key===j.best?" <b>&larr; best here</b>":""}</td>
        <td class="num">${r.recall}%</td>
        <td class="num">top ${r.median}% of the queue</td></tr>`).join("")}
    </table>
    <p class="small muted" style="margin-top:9px">Measured on your decisions as
    they stand, so it shifts as you screen more &mdash; worth re-running
    occasionally. Nothing here changed a label.</p></div>`;
  render();
};
async function rank(model){
  flash("learning...");
  const j=await api("/api/rank"+(model?("?model="+encodeURIComponent(model)):""));
  rankInfo=j;
  if(j.pos<3||j.neg<3){
    flash(`needs 3 includes and 3 of your own excludes (have ${j.pos} and ${j.neg})`);
    return;
  }
  scores=j.scores;
  rankEngine=j.engine||"";
  if(!F.sort) F.sort="score";     // sorting happens at render time
  render();
  flash(`learned from ${j.pos} includes and ${j.neg} excludes — ${j.engine}`);
}
document.addEventListener("keydown",e=>{
  if(tab!=="Screen"||!sel) return;
  if(/input|textarea|select/i.test(e.target.tagName)) return;
  const list=filtered(F), i=list.findIndex(r=>r.id===sel);
  if(e.key==="i"){decide(sel,"include","");}
  else if(e.key==="m"){decide(sel,"maybe","");}
  else if(e.key==="e"){decide(sel,"exclude",(D.criteria.exclude||[""])[0]);}
  else if(e.key==="j"&&i<list.length-1){sel=list[i+1].id;render();}
  else if(e.key==="k"&&i>0){sel=list[i-1].id;render();}
});

// ---------- Sources ----------
function sourcesTab(){
  const S=D.sources||[];
  const extras=(D.extra||[]);
  return `<div class="card">
    <h2>Where your records came from</h2>
    <p class="small muted">PRISMA 2020 asks for this per database, and keeps
    studies found through databases apart from studies found by other methods
    such as papers you added by hand. <b>Every row of every import is kept</b>,
    including papers another database already gave you: that is what makes a
    duplicate visible on the Duplicates tab, where the most complete version
    wins &mdash; so a Web of Science record can replace a harvested one that
    has no abstract. <code>identified</code> is what each source contributed;
    <code>after deduplication</code> is what survived.</p>
    <table><tr><th>source</th><th class="num">identified</th>
      <th class="num">stored</th>
      <th class="num">after deduplication</th><th></th></tr>
    ${S.map(s=>`<tr><td>${esc(s.label)}</td>
      <td class="num">${s.identified.toLocaleString()}</td>
      <td class="num">${s.stored.toLocaleString()}</td>
      <td class="num">${s.after_dedup.toLocaleString()}</td>
      <td class="small muted">${s.other_method?"other methods":"database"}</td>
      </tr>`).join("")||'<tr><td class="muted">loading…</td><td></td><td></td><td></td><td></td></tr>'}
    ${S.length?`<tr><td><b>total</b></td>
      <td class="num"><b>${S.reduce((a,s)=>a+s.identified,0).toLocaleString()}</b></td>
      <td class="num"><b>${S.reduce((a,s)=>a+s.stored,0).toLocaleString()}</b></td>
      <td class="num"><b>${S.reduce((a,s)=>a+s.after_dedup,0).toLocaleString()}</b></td>
      <td></td></tr>`:""}
    </table>
  </div>

  <div class="card"><h2>Import a database export</h2>
    <p class="small muted">Web of Science (.xls or .xlsx), Lens (.csv),
    Scopus (.csv), or any table with a title and DOI column. The format is
    detected from its columns. Paste the path, or drop the file in your
    Papers folder and type its name.</p>
    <div class="row">
      <input type="text" id="imppath" style="flex:1;min-width:320px"
        placeholder="07-web_of_science.xls">
      <select id="impsrc">
        <option value="">detect the format</option>
        <option value="wos">Web of Science</option>
        <option value="lens">Lens</option>
        <option value="scopus">Scopus</option>
        <option value="other-db">other database</option>
      </select>
      <button class="b n" onclick="impGo(false)">Preview</button>
    </div>
    <p class="small muted" style="margin-top:7px">Preview changes nothing. It
    reports how many records are new and how many you already have.</p>
    <div id="impout"></div>
  </div>

  <div class="card"><h2>Add a paper by hand</h2>
    <p class="small muted">For papers a search cannot reach &mdash; found by
    citation searching, or known to you already. Give the DOI and the details
    are fetched from Crossref; otherwise type what you have. These are
    reported in PRISMA under <i>other methods</i>, not as a database hit.</p>
    <div class="row">
      <input type="text" id="mdoi" style="flex:1;min-width:260px"
        placeholder="10.1016/j.jbusvent.2013.06.005">
      <button class="b n" onclick="doiLookup()">Look up</button>
      <select id="msrc">
        <option value="citation">citation searching</option>
        <option value="manual">added by hand</option>
      </select>
    </div>
    <div class="grid3" style="margin-top:10px">
      <div><div class="small muted">Title</div>
        <input type="text" id="mtitle" style="width:100%"></div>
      <div><div class="small muted">Authors</div>
        <input type="text" id="mauthors" style="width:100%"></div>
      <div><div class="small muted">Year</div>
        <input type="text" id="myear" style="width:100%"></div>
      <div><div class="small muted">Journal</div>
        <input type="text" id="mjournal" style="width:100%"></div>
      <div><div class="small muted">Citations</div>
        <input type="text" id="mcited" style="width:100%"></div>
      <div><div class="small muted">Note (why you added it)</div>
        <input type="text" id="mnote" style="width:100%"
          placeholder="cited in Basha et al. 2021"></div>
    </div>
    <div style="margin-top:8px"><div class="small muted">Abstract</div>
      <textarea id="mabstract"></textarea></div>
    <div class="row" style="margin-top:8px">
      <button class="b" onclick="addPaper()">Add to the corpus</button>
      <button class="b n" onclick="clearPaper()">clear</button>
    </div>
    <div id="addout" class="small" style="margin-top:8px"></div>
  </div>

  ${extras.length?`<div class="card"><h2>Imported and hand-added records
    (${extras.length.toLocaleString()})</h2>
    <p class="small muted">Held in <code>${esc((D.paths||{}).extra||"")}</code>,
    separate from your harvest, so re-running the search never loses them.</p>
    <table><tr><th>id</th><th>source</th><th>record</th><th></th></tr>
    ${extras.slice(0,200).map(e=>`<tr>
      <td class="small">${esc(e.id)}</td>
      <td class="small">${esc(e.source_label||e.source)}</td>
      <td><div style="font-size:13px">${esc((e.title||"").slice(0,86))}</div>
        <div class="small muted">${esc(e.year||"")} ${esc((e.journal||"").slice(0,46))}
        ${e.note?" &middot; "+esc(e.note):""}</div></td>
      <td><button class="b n" style="padding:3px 9px"
        onclick="removeExtra('${esc(e.id)}')">remove</button></td></tr>`).join("")}
    </table>
    ${extras.length>200?`<p class="small muted">showing the first 200</p>`:""}
  </div>`:""}`;
}
window.impGo=async(apply)=>{
  const path=$("#imppath").value.trim();
  if(!path){flash("paste the path of the export file");return;}
  const src=$("#impsrc").value;
  $("#impout").innerHTML='<p class="muted">reading…</p>';
  const j=await api(apply?"/api/import-apply":"/api/import-preview",
    {method:"POST",headers:{"content-type":"application/json"},
     body:JSON.stringify({path,source:src||undefined})});
  if(j.error){$("#impout").innerHTML=
    `<p class="muted" style="white-space:pre-wrap">${esc(j.error)}</p>`;return;}
  let h=`<div class="card" style="margin-top:10px">
    <div class="small muted">${esc(j.file)} &middot; read as
      <b>${esc(j.label)}</b> &middot; ${j.rows.toLocaleString()} rows</div>
    <div class="grid3" style="margin-top:10px">
      <div class="stat"><div class="n">${j.new.toLocaleString()}</div>
        <div class="l">${j.applied?"stored":"will be stored"}</div></div>
      <div class="stat"><div class="n">${j.overlap.toLocaleString()}</div>
        <div class="l">also in another database</div></div>
      <div class="stat"><div class="n">${(j.already_imported+j.dupes_within_file+j.unusable).toLocaleString()}</div>
        <div class="l">not stored</div></div>
    </div>
    <p class="small muted" style="margin-top:8px">
      ${j.overlap?`${j.overlap.toLocaleString()} of these are papers another
       database also returned. They are stored anyway, with this source on
       them, and resolved on the <b>Duplicates</b> tab &mdash; that is where
       the most complete version is chosen, so a version with an abstract
       your harvest lacks can win. PRISMA counts them as identified here and
       removed as duplicates there.<br>`:""}
      ${j.already_imported?`${j.already_imported.toLocaleString()} were
       imported from this source before and are untouched &mdash; importing
       the same file twice changes nothing.<br>`:""}
      ${j.dupes_within_file?`${j.dupes_within_file.toLocaleString()} appear
       twice inside this one file.<br>`:""}
      ${j.unusable?`${j.unusable.toLocaleString()} have neither a title nor a
       DOI, so they cannot be screened.<br>`:""}</p>
    <p class="small muted" style="margin-top:6px">Columns used:
      ${Object.entries(j.columns).map(([k,v])=>`${k} &larr; <code>${esc(v)}</code>`).join(" &middot; ")}
      ${j.unmapped.length?`<br>Not found in this file: ${j.unmapped.join(", ")}`:""}</p>`;
  if(j.sample&&j.sample.length&&!j.applied){
    h+=`<table style="margin-top:8px"><tr><th>a few of the new ones</th>
      <th class="num">year</th></tr>${j.sample.map(s=>`<tr>
      <td>${esc(s.title)}<div class="small muted">${esc(s.journal)}</div></td>
      <td class="num">${esc(s.year)}</td></tr>`).join("")}</table>`;
  }
  h += j.applied
    ? `<p style="margin-top:10px"><b>Done.</b> Corpus is now
       ${j.corpus_now.toLocaleString()} records. Duplicates and rules have been
       recomputed.</p>`
    : (j.new?`<div class="row" style="margin-top:10px">
        <button class="b" onclick="impGo(true)">Store these ${j.new} record(s)</button>
       </div>`:`<p class="small muted" style="margin-top:10px">Nothing to
       store &mdash; every usable row in this file is already in the corpus
       from this same source.</p>`);
  $("#impout").innerHTML=h+"</div>";
  if(j.applied) await load();
};
window.doiLookup=async()=>{
  const doi=$("#mdoi").value.trim();
  if(!doi){flash("paste a DOI");return;}
  $("#addout").textContent="looking up…";
  const j=await api("/api/lookup-doi",{method:"POST",
    headers:{"content-type":"application/json"},body:JSON.stringify({doi})});
  if(j.error){$("#addout").innerHTML=`<span class="muted">${esc(j.error)}</span>`;return;}
  const r=j.record;
  $("#mtitle").value=r.title||""; $("#mauthors").value=r.authors||"";
  $("#myear").value=r.year||"";   $("#mjournal").value=r.journal||"";
  $("#mabstract").value=r.abstract||""; $("#mcited").value=r.cited_by||"";
  $("#mdoi").value=r.doi||doi;
  $("#addout").innerHTML=`<span class="pill inc">found</span> check the fields, `
    + (r.abstract?"abstract included":"<b>no abstract</b> — paste one if you have it")
    + ", then add it.";
};
window.addPaper=async()=>{
  const body={doi:$("#mdoi").value,title:$("#mtitle").value,
    authors:$("#mauthors").value,year:$("#myear").value,
    journal:$("#mjournal").value,abstract:$("#mabstract").value,
    cited_by:$("#mcited").value,note:$("#mnote").value,
    source:$("#msrc").value};
  const j=await api("/api/add-paper",{method:"POST",
    headers:{"content-type":"application/json"},body:JSON.stringify(body)});
  if(j.error){$("#addout").innerHTML=`<span class="muted">${esc(j.error)}</span>`;return;}
  clearPaper();
  $("#addout").innerHTML=`<span class="pill inc">added as ${esc(j.id)}</span> ${esc(j.title)}`;
  await load(); render();
};
window.clearPaper=()=>{
  ["mdoi","mtitle","mauthors","myear","mjournal","mabstract","mcited","mnote"]
    .forEach(i=>{const e=$("#"+i); if(e) e.value="";});
};
window.removeExtra=async(id)=>{
  await api("/api/remove-extra",{method:"POST",
    headers:{"content-type":"application/json"},body:JSON.stringify({id})});
  await load(); flash("removed"); render();
};

// ---------- Stopping ----------
function stoppingShell(){
  return `<div id="stopbox" class="muted">loading...</div>`;
}
async function loadStopping(){
  const s=await api("/api/stopping");
  const box=$("#stopbox"); if(!box) return;
  const tick=c=>c.skipped?'<span class="pill">cannot check</span>'
    :c.met?'<span class="pill inc">met</span>':'<span class="pill exc">not yet</span>';
  const st=s.stop&&s.stop.recorded?s.stop:null;

  let html=`<div class="card">
    <h2>May you stop screening?</h2>
    <p class="small muted">Active learning reorders records; it never labels
    them. So screening ends when you decide it does. These are the four
    conditions of the SAFE procedure (Boetje &amp; van de Schoot, 2024), which
    deliberately combines several checks rather than trusting any one. The
    thresholds are examples in the paper, not law &mdash; adjust them below
    and say in your methods what you used.</p>
    <div class="grid3">
      <div class="stat"><div class="n">${s.screened.toLocaleString()}</div><div class="l">individually screened</div></div>
      <div class="stat"><div class="n">${s.relevant.toLocaleString()}</div><div class="l">relevant found</div></div>
      <div class="stat"><div class="n">${s.percent}%</div><div class="l">of the corpus</div></div>
      <div class="stat"><div class="n">${s.consecutive.toLocaleString()}</div><div class="l">since the last relevant</div></div>
      <div class="stat"><div class="n">${s.unscreened.toLocaleString()}</div><div class="l">still to decide</div></div>
    </div>
    ${s.auto_decided?`<p class="small muted" style="margin-top:8px">A further
      <b>${s.auto_decided.toLocaleString()}</b> were excluded by a rule rather
      than screened by hand. They are not in "still to decide" and not in
      "individually screened" &mdash; the SAFE counters measure the screening
      you did yourself.</p>`:""}
    ${(s.skipped.elsewhere||s.skipped.on_duplicates)?`
      <p class="small muted" style="margin-top:10px">Not counted above, though
      kept in your decisions file:
      ${s.skipped.elsewhere?`<b>${s.skipped.elsewhere.toLocaleString()}</b>
        decision(s) about records this search did not return &mdash; work
        carried over from an earlier harvest.`:""}
      ${s.skipped.on_duplicates?`<b>${s.skipped.on_duplicates.toLocaleString()}</b>
        decision(s) about records since marked duplicate; the copy that was
        kept carries the decision instead.`:""}
      A percentage only means something if the top and bottom of it count the
      same records.</p>`:""}
    <table style="margin-top:14px">
      <tr><th style="width:92px">status</th><th>condition</th></tr>
      ${s.conditions.map(c=>`<tr><td>${tick(c)}</td>
        <td><b>${esc(c.label)}</b><div class="small muted">${esc(c.detail)}</div></td>
        </tr>`).join("")}
    </table>
    <p style="margin-top:12px">${s.all_met
      ? '<span class="pill inc">all checkable conditions met</span>'
      : `<span class="pill may">${s.met_count} of ${s.live_count} met</span>`}</p>
  </div>`;

  if(s.recall_curve&&s.recall_curve.length>1){
    const C=s.recall_curve, W=820,H=200,P=38;
    const mx=C[C.length-1][0], my=Math.max(1,C[C.length-1][1]);
    const pts=C.map(([x,y])=>`${P+x/mx*(W-P-12)},${H-P-y/my*(H-P-24)}`).join(" ");
    html+=`<div class="card"><h2>Relevant found as screening went on</h2>
      <p class="small muted">A curve that has flattened is the visual check
      SAFE asks for: new relevant records have stopped appearing. A curve
      still climbing says keep going, whatever the counters say.</p>
      <svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" font-size="11">
        <line x1="${P}" y1="${H-P}" x2="${W-12}" y2="${H-P}" stroke="#9fb3c8"/>
        <line x1="${P}" y1="12" x2="${P}" y2="${H-P}" stroke="#9fb3c8"/>
        <polyline points="${pts}" fill="none" stroke="#0072B2" stroke-width="2.5"/>
        <text x="${P}" y="${H-12}" fill="#5b6b7c">0</text>
        <text x="${W-12}" y="${H-12}" text-anchor="end" fill="#5b6b7c">${mx.toLocaleString()} screened</text>
        <text x="6" y="18" fill="#5b6b7c">${my.toLocaleString()} relevant</text>
      </svg></div>`;
  }

  if(s.keys){
    const k=s.keys;
    html+=`<div class="card"><h2>Key papers</h2>
      <p class="small muted">The ${k.total} papers marked <code>core</code> in
      ${esc(k.where)}. Three different questions, kept apart because they call
      for different action: whether your <b>search</b> finds a known paper
      (widen the query), whether it is <b>in the corpus</b> at all, and whether
      you have <b>included</b> it (keep screening).</p>
      <div class="grid3">
        <div class="stat"><div class="n">${k.by_search}/${k.total}</div>
          <div class="l">found by the database search</div></div>
        <div class="stat"><div class="n">${k.in_corpus}/${k.total}</div>
          <div class="l">in the corpus, however added</div></div>
        <div class="stat"><div class="n">${k.found}/${k.in_corpus}</div>
          <div class="l">screened and included</div></div>
      </div>
      ${k.missed_by_search.length?`<p class="small" style="margin-top:10px">
        <b>Your search does not return these</b> (${k.n_missed_by_search} in
        all) &mdash; a query problem, not a screening one:
        ${k.missed_by_search.map(esc).join(" &middot; ")}</p>`:""}
      ${k.not_yet_included.length?`<p class="small" style="margin-top:6px">
        <b>In the corpus but not yet included:</b> ${k.not_yet_included.map(esc).join(" &middot; ")}</p>`:""}
      ${k.missing_from_corpus.length?`<p class="small muted" style="margin-top:6px">
        <b>Not in the corpus at all:</b> ${k.missing_from_corpus.map(esc).join(" &middot; ")}
        &mdash; import your list on the Sources tab to add them by hand.</p>`:""}
    </div>`;
  }

  html+=`<div class="card"><h2>Thresholds</h2>
    <div class="row">
      <label class="small muted">records since the last relevant
        <input type="number" id="scon" min="1" max="1000" value="${s.threshold}"
          style="width:80px"></label>
      <label class="small muted">minimum % of the corpus
        <input type="number" id="spct" min="0" max="100" step="0.5"
          value="${s.min_percent}" style="width:80px"></label>
      <label class="small muted">estimated relevant in total
        <input type="number" id="sest" min="0" placeholder="${s.estimate}"
          value="${esc(s.manual_estimate)}" style="width:90px"></label>
      <button class="b n" onclick="saveStopCfg()">save</button>
    </div>
    <p class="small muted" style="margin-top:7px">Leave the estimate blank and
    it is taken from the rate so far &mdash; currently ${s.estimate.toLocaleString()},
    from ${esc(s.estimate_from)}. Because prioritised screening finds relevant
    records early, that rate runs high, so the estimate is too large and the
    condition is harder to meet. That errs on the safe side. SAFE itself gets
    this number from a small random sample screened before prioritisation.</p>
  </div>`;

  if(st){
    html+=`<div class="card" style="border-color:#0072B2">
      <h2>Stopping point recorded</h2>
      <p class="small muted">${esc(st.at)}${st.screener?" by "+esc(st.screener):""}
      ${st.all_met?'<span class="pill inc">all conditions were met</span>'
                  :'<span class="pill may">recorded with conditions unmet</span>'}</p>
      <table>
        <tr><td>screened</td><td class="num">${st.screened.toLocaleString()} of ${st.total.toLocaleString()} (${st.percent}%)</td></tr>
        <tr><td>relevant found</td><td class="num">${st.relevant.toLocaleString()}</td></tr>
        <tr><td>not individually screened</td><td class="num">${st.unscreened.toLocaleString()}</td></tr>
        <tr><td>run of irrelevant at the end</td><td class="num">${st.consecutive.toLocaleString()} (threshold ${st.threshold})</td></tr>
      </table>
      ${st.note?`<p class="small" style="margin-top:8px"><b>Note:</b> ${esc(st.note)}</p>`:""}
      <h2 style="margin-top:14px">For your methods section</h2>
      <div class="ab" id="methods">${esc(methodsText(st))}</div>
      <div class="row" style="margin-top:10px">
        <button class="b n" onclick="navigator.clipboard.writeText($('#methods').textContent);flash('copied')">copy</button>
        <button class="b n" onclick="stopUndo()">undo &mdash; carry on screening</button>
      </div>
    </div>`;
  } else {
    html+=`<div class="card"><h2>Record the stopping point</h2>
      <p class="small muted">This does not exclude anything. It records where
      you stopped and why, so PRISMA can report the records nobody looked at
      as exactly that, rather than silently counting them as excluded. You can
      undo it and carry on at any time.</p>
      <div class="row">
        <input type="text" id="snote" style="flex:1;min-width:280px"
          placeholder="optional note, e.g. second screener agreed">
        <button class="b" onclick="stopRecord()">Record stopping point</button>
      </div>
      ${s.all_met?"":`<p class="small muted" style="margin-top:8px">
        ${s.live_count-s.met_count} condition(s) are not met. You can still
        record it &mdash; the unmet ones are saved alongside, so the record is
        honest about where you stopped.</p>`}
    </div>`;
  }
  box.innerHTML=html;
}
function methodsText(st){
  const cond=st.conditions.filter(c=>!c.skipped);
  return `Screening was prioritised using active learning and stopped `
    + `following the SAFE procedure (Boetje & van de Schoot, 2024). `
    + `Of ${st.total.toLocaleString()} unique records, ${st.screened.toLocaleString()} `
    + `(${st.percent}%) were screened individually on title and abstract, `
    + `yielding ${st.relevant.toLocaleString()} relevant records. Screening `
    + `stopped after ${st.consecutive.toLocaleString()} consecutive records `
    + `with no new relevant record (threshold: ${st.threshold}), with at least `
    + `${st.min_percent}% of the corpus screened and at least twice the `
    + `estimated number of relevant records (${st.estimate.toLocaleString()}, `
    + `derived from ${st.estimate_from}) examined`
    + (st.keys&&st.keys.in_corpus?`, and ${st.keys.found} of `
        + `${st.keys.in_corpus} known key papers retrieved and included`:``)
    + `. The remaining ${st.unscreened.toLocaleString()} records were not `
    + `individually screened and are reported as such in the PRISMA flow `
    + `diagram.`
    + (cond.some(c=>!c.met)?` Note: at the point of stopping, the following `
        + `condition(s) were not met: `
        + cond.filter(c=>!c.met).map(c=>c.label.toLowerCase()).join("; ")+`.`:``)
    + (st.note?` ${st.note}`:``);
}
window.saveStopCfg=async()=>{
  await api("/api/stop-config",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({consecutive:$("#scon").value,
      min_percent:$("#spct").value, estimate:$("#sest").value})});
  loadStopping(); flash("thresholds saved");
};
window.stopRecord=async()=>{
  await api("/api/stop-record",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({note:$("#snote").value})});
  loadStopping(); flash("stopping point recorded");
};
window.stopUndo=async()=>{
  await api("/api/stop-undo",{method:"POST",
    headers:{"content-type":"application/json"},body:"{}"});
  loadStopping(); flash("you can carry on screening");
};

// ---------- Duplicates ----------
window.dupFilter="all";
window.dupShow=40;
window.setDupFilter=(v)=>{window.dupFilter=v;window.dupShow=40;render();};
window.moreDupes=()=>{window.dupShow+=60;render();};

function dupTab(){
  const groups=(D.dupes||[]).map(g=>{
    const rs=g.ids.map(byId).filter(Boolean);
    const srcs=[...new Set(rs.map(r=>r.src_label))];
    // Did keeping the most complete version actually gain anything? If the
    // kept record has an abstract the others lack, that is a paper you can
    // now screen properly instead of on its title alone.
    const kept=rs.find(r=>r.dup==="keep");
    const gained=!!(kept&&Object.keys(kept.merged_from||{}).length);
    return {...g,rs,srcs,gained,
            manual:rs.some(r=>r.dup_manual)};
  }).filter(g=>g.rs.length>1);

  if(!groups.length) return '<div class="card">No duplicate groups found.</div>';
  const nDup=D.records.filter(r=>r.dup==="duplicate").length;
  const nCross=groups.filter(g=>g.srcs.length>1).length;
  const nGain=groups.filter(g=>g.gained).length;

  const f=window.dupFilter;
  const shown=groups.filter(g=>
    f==="cross"?g.srcs.length>1:
    f==="gained"?g.gained:
    f==="manual"?g.manual:true);
  const page=shown.slice(0,window.dupShow);

  const btn=(v,label,n)=>`<button class="b ${f===v?"":"n"}"
    onclick="setDupFilter('${v}')">${label} (${n})</button>`;

  return `<div class="card">
    <h2>${groups.length} duplicate groups &middot; ${nDup} records marked as duplicates</h2>
    <p class="small muted">The most complete record in each group is kept
    automatically: a DOI counts for more than anything else, then an abstract,
    then authors, journal and year, with citations as a tiebreak. Change the
    choice and yours is kept instead, for good. Removals flow into the PRISMA
    diagram, and the figure there and the figure here are the same number.</p>
    <p class="small muted"><b>Nothing is thrown away.</b> The kept record keeps
    everything it has and takes only what it is missing from the others &mdash;
    so a record can end up with an abstract from Web of Science and a topic
    from OpenAlex at once. Borrowed values are marked in green below, and
    named in the export's <code>merged_fields</code> column.</p>
    <p class="small muted"><b>${nCross}</b> of these groups span more than one
    database &mdash; the same paper returned by two of your searches. In
    <b>${nGain}</b> of them the kept record gained a field it did not have.</p>
    <div class="row" style="margin-top:10px;gap:6px;flex-wrap:wrap">
      ${btn("all","all groups",groups.length)}
      ${btn("cross","across databases",nCross)}
      ${btn("gained","kept version is better",nGain)}
      ${btn("manual","you chose",groups.filter(g=>g.manual).length)}
    </div>
  </div>` + page.map(g=>{
    return `<div class="card"><div class="small muted" style="margin-bottom:7px">
      ${esc(g.basis)} &middot; ${g.srcs.map(esc).join(" + ")}
      ${g.manual?' &middot; <b>your choice</b>':""}</div>
      <table><tr><th style="width:62px">keep</th><th>record</th>
      <th style="width:110px">source</th>
      <th style="width:170px">has</th><th class="num" style="width:62px">cites</th></tr>
      ${g.rs.map(r=>{
        const have=[];
        if(r.doi)have.push("DOI"); if(r.abstract)have.push("abstract");
        if(r.authors)have.push("authors"); if(r.journal)have.push("journal");
        if(r.year)have.push("year"); if(r.topic)have.push("topic");
        const borrowed=Object.entries(r.merged_from||{});
        return `<tr${r.dup==="keep"?' style="background:var(--soft)"':""}>
        <td><input type="radio" name="g${esc(g.key)}" ${r.dup==="keep"?"checked":""}
             onchange="keepOne('${esc(g.key)}','${r.id}')"></td>
        <td><div style="font-weight:600;font-size:13px">${esc(r.title||"(no title)")}</div>
          <div class="small muted">${esc(r.year||"?")} &middot;
          ${esc(r.journal||"no journal")} &middot; ${esc(r.doi||"no DOI")}
          &middot; <span class="small">${esc(r.id)}</span></div>
          ${borrowed.length?`<div class="small" style="color:var(--green);margin-top:3px">
            took ${borrowed.map(([k,v])=>`<b>${esc(k)}</b> from ${esc(v)}`).join(", ")}
            </div>`:""}</td>
        <td class="small">${esc(r.src_label)}</td>
        <td class="small">${have.join(", ")||"<i>almost nothing</i>"}</td>
        <td class="num">${r.cited_by}</td></tr>`;}).join("")}
      </table></div>`;}).join("")
   + (shown.length>page.length?`<div class="card" style="text-align:center">
      <button class="b" onclick="moreDupes()">Show more &mdash;
      ${page.length} of ${shown.length} shown</button></div>`:"");
}
window.keepOne=async(key,keep)=>{
  await api("/api/dupkeep",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({key,keep})});
  await load(); flash("kept");
};

// ---------- Rules ----------
function rulesTab(){
  return `<div class="card">
    <h2>Rules that feed the PRISMA diagram</h2>
    <p class="small muted">Each rule excludes records automatically and its
    reason is carried into PRISMA. Your own decision on a record always wins
    over a rule, and keeps winning. The counts below are live: they show what
    each rule removes right now, or would remove if you switched it on, after
    the rules above it have taken their share.</p>
  </div>` + (D.rules||[]).map(r=>`
    <div class="rule ${r.enabled?"on":""}">
      <div class="top">
        <label class="ck"><input type="checkbox" ${r.enabled?"checked":""}
          onchange="setRule('${r.id}',{enabled:this.checked})">
          <b>${esc(r.name)}</b></label>
        <span><span class="pill ${r.enabled?"auto":""}">${
          r.enabled?"removes":"would remove"} ${(D.preview[r.id]||0).toLocaleString()}</span></span>
      </div>
      <div class="small muted" style="margin-top:6px">
        PRISMA reason: &ldquo;${esc(r.reason)}&rdquo;</div>
      <div class="small muted" style="margin-top:4px">${esc(r.note||"")}</div>
      ${r.id==="old_uncited"?`<div class="row small" style="margin-top:7px">
        at least <input type="number" min="1" max="30" style="width:66px"
          value="${(r.params.years||5)}"
          onchange="setRule('old_uncited',{params:{years:+this.value}})"> years old
        and never cited</div>`:""}
      ${r.id==="lang"?`<div class="row small" style="margin-top:7px">
        keep these languages <input type="text" style="width:150px"
          value="${esc((r.params.allow||[]).join(", "))}"
          onchange="setRule('lang',{params:{allow:this.value.split(',').map(s=>s.trim()).filter(Boolean)}})">
        <label class="ck"><input type="checkbox" ${r.params.keep_unknown?"checked":""}
          onchange="setRule('lang',{params:{keep_unknown:this.checked}})">
          keep records whose language is unknown</label></div>`:""}
      ${r.id==="wrongtype"?`<div class="row small" style="margin-top:7px">
        block these types <input type="text" style="width:260px"
          value="${esc((r.params.block||[]).join(", "))}"
          onchange="setRule('wrongtype',{params:{block:this.value.split(',').map(s=>s.trim()).filter(Boolean)}})">
        </div>`:""}
      <div class="small" style="margin-top:7px">
        <a href="#" onclick="F={q:'',dec:'',theme:'',country:'',lang:'',src:'',yr:'',flag:'',dup:'live',sort:F.sort};
           F.dec='auto';go('Screen');return false">see what it caught</a></div>
    </div>`).join("");
}
window.setRule=async(id,patch)=>{
  const rules=(D.rules||[]).map(r=>r.id===id?Object.assign({},r,patch,
    patch.params?{params:Object.assign({},r.params,patch.params)}:{}):r);
  const j=await api("/api/rules",{method:"POST",
    headers:{"content-type":"application/json"},body:JSON.stringify({rules})});
  if(j.ok){ await load(); flash("rules updated"); }
};

// ---------- Dashboard ----------
const VARS=[["year","Year"],["country","Country"],["theme","Theme"],
  ["language","Language"],["type","Publication type"],["journal","Journal"],
  ["field","Field"],["subfield","Subfield"],["oa_status","Open access"],
  ["europe","European signal"],["decision","Decision"],
  ["source","Decision made by"],["db","Database it came from"],
  ["cites","Citations"],["abs","Abstract"]];

function varValues(r,key){
  switch(key){
    case "year": return [r.year||"unknown"];
    case "country": {const c=countryOf(r);
      return c?c.split("|").map(s=>s.trim()).filter(Boolean):["unknown"];}
    case "theme": {const t=themesOf(r);
      return t.length?t.map(labName):["none"];}
    case "language": return [r.language||"unknown"];
    case "type": return [r.type||"unknown"];
    case "journal": return [r.journal||"unknown"];
    case "field": return [r.field||"unknown"];
    case "subfield": return [r.subfield||"unknown"];
    case "oa_status": return [r.oa_status||"unknown"];
    case "europe": return [(r.europe||"").toLowerCase()==="yes"?"European":"no signal"];
    case "decision": return [r.eff||"not screened"];
    case "source": return [r.eff_src==="auto"?"a rule":r.eff_src==="manual"?"by hand":"not screened"];
    case "db": return [r.src_label||"unknown"];
    case "cites": {const c=parseInt(r.cited_by||0,10);
      return [c===0?"0":c<5?"1-4":c<10?"5-9":c<50?"10-49":"50+"];}
    case "abs": return [r.abstract?"has abstract":"no abstract"];
  }
  return ["?"];
}

function dashTab(){
  const langs=[...new Set(D.records.map(r=>r.language).filter(Boolean))].sort();
  const countries=[...new Set(D.records.flatMap(r=>
      countryOf(r).split("|").map(s=>s.trim()).filter(Boolean)))].sort();
  const list=filtered(Object.assign({},DF,{q:"",flag:"",src:"",dup:"live"}));
  const screened=list.filter(r=>r.eff).length;
  const inc=list.filter(r=>r.eff==="include").length;
  const auto=list.filter(r=>r.eff==="exclude"&&r.eff_src==="auto").length;
  const med=(()=>{const a=list.map(r=>+r.cited_by||0).sort((x,y)=>x-y);
    return a.length?a[Math.floor(a.length/2)]:0;})();
  return `<div class="card">
    <h2>Dashboard</h2>
    <div class="row">
      <label class="small muted">group by
        <select onchange="DF.group=this.value;render()">
          ${VARS.map(([v,l])=>`<option value="${v}" ${DF.group===v?"selected":""}>${l}</option>`).join("")}
        </select></label>
      <label class="small muted">split by
        <select onchange="DF.split=this.value;render()">
          <option value="">nothing</option>
          ${VARS.map(([v,l])=>`<option value="${v}" ${DF.split===v?"selected":""}>${l}</option>`).join("")}
        </select></label>
      <span style="flex:1"></span>
      <select onchange="DF.dec=this.value;render()">
        ${opt([["","any decision"],["todo","not screened"],["include","included"],
               ["maybe","maybe"],["exclude","excluded"],["auto","excluded by a rule"],
               ["manual","decided by hand"]],DF.dec)}
      </select>
      <select onchange="DF.theme=this.value;render()">
        <option value="">any theme</option>
        ${(D.labels||[]).map(l=>`<option value="${l.id}" ${DF.theme===l.id?"selected":""}>${esc(l.name)}</option>`).join("")}
      </select>
      <select onchange="DF.country=this.value;render()">
        <option value="">any country</option>
        ${countries.map(c=>`<option ${DF.country===c?"selected":""}>${esc(c)}</option>`).join("")}
      </select>
      <select onchange="DF.lang=this.value;render()">
        <option value="">any language</option>
        ${langs.map(c=>`<option ${DF.lang===c?"selected":""}>${esc(c)}</option>`).join("")}
      </select>
      <select onchange="DF.yr=this.value;render()">
        ${opt([["","any year"],["old","5+ years old"],["new","last 5 years"]],DF.yr)}
      </select>
      <button class="b n" onclick="DF={group:DF.group,split:DF.split,dec:'',theme:'',country:'',lang:'',yr:''};render()">clear</button>
    </div>
    <div class="grid3" style="margin-top:12px">
      <div class="stat"><div class="n">${list.length.toLocaleString()}</div><div class="l">records in view</div></div>
      <div class="stat"><div class="n">${screened.toLocaleString()}</div><div class="l">screened</div></div>
      <div class="stat"><div class="n">${inc.toLocaleString()}</div><div class="l">included</div></div>
      <div class="stat"><div class="n">${auto.toLocaleString()}</div><div class="l">excluded by a rule</div></div>
      <div class="stat"><div class="n">${med}</div><div class="l">median citations</div></div>
    </div>
  </div>
  <div class="card"><h2>Overview</h2>
    <p class="small muted">The standard panels, all obeying the filters above.</p>
    <div class="panels" id="panels"></div></div>
  <div class="grid2">
    <div class="card"><div id="chart" class="scrollbox"></div></div>
    <div class="card"><h2>The numbers</h2>
      <div id="ctable" class="scrollbox"></div></div>
  </div>`;
}

function dashList(){
  return filtered(Object.assign({},DF,{q:"",flag:"",src:"",dup:"live"}));
}
function countBy(list,key,top){
  // Web of Science writes NETHERLANDS where OpenAlex writes Netherlands, so
  // the same country arrived twice and was counted twice. Fold on case and
  // show the mixed-case spelling, matching the Bibliometrics tab.
  const c={}, seen={};
  list.forEach(r=>varValues(r,key).forEach(v=>{
    const k=String(v).trim().toLowerCase();
    if(!k) return;
    c[k]=(c[k]||0)+1;
    (seen[k]=seen[k]||{})[String(v).trim()]=(seen[k][String(v).trim()]||0)+1;
  }));
  const show=k=>{
    const w=seen[k]||{};
    const mixed=Object.entries(w).filter(([t])=>t!==t.toUpperCase());
    const pool=mixed.length?mixed:Object.entries(w);
    return pool.sort((a,b)=>b[1]-a[1])[0]?.[0]||k;
  };
  let e=Object.entries(c).map(([k,v])=>[show(k),v]);
  const numeric=e.every(([k])=>/^\d{4}$/.test(k));
  e.sort(numeric?((a,b)=>a[0].localeCompare(b[0])):((a,b)=>b[1]-a[1]));
  if(top&&!numeric&&e.length>top) e=e.slice(0,top);
  return e;
}
function miniBars(data,colour){
  if(!data.length) return '<p class="small muted">nothing here</p>';
  const max=Math.max(...data.map(d=>d[1]))||1;
  return `<table style="font-size:12px">${data.map(([k,v],i)=>`<tr>
    <td style="width:46%;padding:3px 6px 3px 0">${esc(k.length>26?k.slice(0,25)+"…":k)}</td>
    <td style="padding:3px 0"><div class="bar" style="width:${Math.max(2,v/max*100)}%;
      background:${colour||D.palette[i%D.palette.length]}"></div></td>
    <td class="num" style="width:52px;padding:3px 0 3px 6px">${v.toLocaleString()}</td>
    </tr>`).join("")}</table>`;
}
function miniCols(data,colour){
  if(!data.length) return '<p class="small muted">nothing here</p>';
  const max=Math.max(...data.map(d=>d[1]))||1, W=330, H=120;
  const bw=W/data.length;
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" font-size="9">
    ${data.map(([k,v],i)=>{
      const h=Math.max(1,v/max*(H-28));
      return `<rect x="${i*bw+1.5}" y="${H-20-h}" width="${Math.max(2,bw-3)}"
        height="${h}" rx="2" fill="${colour}"><title>${esc(k)}: ${v}</title></rect>
        ${data.length<=18||i%2===0?`<text x="${i*bw+bw/2}" y="${H-7}"
          text-anchor="middle" fill="#5b6b7c">${esc(k).slice(-2)}</text>`:""}`;
    }).join("")}</svg>`;
}
function drawPanels(){
  const el=$("#panels"); if(!el) return;
  const L=dashList(), P=D.palette;
  const panel=(t,note,inner)=>`<div class="panel"><div class="pt">${esc(t)}</div>
    <div class="small muted" style="margin:-3px 0 7px">${esc(note)}</div>${inner}</div>`;
  el.innerHTML=[
    panel("Decisions","where screening stands",miniBars(countBy(L,"decision"))),
    panel("Records by year","publication year",
          miniCols(countBy(L,"year"),P[0])),
    panel("Themes","a paper can carry more than one",
          miniBars(countBy(L,"theme"))),
    panel("Countries","top 10, from text then affiliation",
          miniBars(countBy(L,"country",10),P[2])),
    panel("Journals","top 10",miniBars(countBy(L,"journal",10),P[4])),
    panel("Citations","as harvested",miniBars(countBy(L,"cites"),P[1])),
    panel("Language","as recorded by OpenAlex",miniBars(countBy(L,"language",8),P[3])),
    panel("Fields","top 8",miniBars(countBy(L,"field",8),P[5])),
  ].join("");
}

function tally(){
  const list=filtered(Object.assign({},DF,{q:"",flag:"",src:"",dup:"live"}));
  const counts={}, splits=new Set();
  list.forEach(r=>{
    const gs=varValues(r,DF.group);
    const ss=DF.split?varValues(r,DF.split):["all"];
    gs.forEach(g=>{
      counts[g]=counts[g]||{};
      ss.forEach(s=>{counts[g][s]=(counts[g][s]||0)+1;splits.add(s);});
    });
  });
  let keys=Object.keys(counts);
  const numeric=keys.every(k=>/^\d{4}$/.test(k));
  if(numeric) keys.sort();
  else keys.sort((a,b)=>sum(counts[b])-sum(counts[a]));
  if(keys.length>22&&!numeric) keys=keys.slice(0,22);
  return {keys,counts,splits:[...splits].sort(),total:list.length};
}
const sum=o=>Object.values(o).reduce((a,b)=>a+b,0);

function drawChart(){
  const el=$("#chart"); if(!el) return;
  const {keys,counts,splits}=tally();
  if(!keys.length){el.innerHTML='<p class="muted">Nothing matches these filters.</p>';
    $("#ctable").innerHTML="";return;}
  const P=D.palette, max=Math.max(...keys.map(k=>sum(counts[k])))||1;
  // Rows are the height the Overview panels use. The SVG used to be
  // width="100%" over an 860-wide viewBox, so it stretched to the card and
  // every 26px row was drawn at about 39px — one chart taller than the whole
  // of the rest of the dashboard.
  const rowH=21, W=560, labelW=150, barW=W-labelW-54;
  const H=keys.length*rowH+10;
  let svg=`<svg width="100%" viewBox="0 0 ${W} ${H}" font-family="inherit"
    font-size="11.5" style="max-width:${W}px;display:block">`;
  keys.forEach((k,i)=>{
    const y=i*rowH+4; let x=labelW;
    svg+=`<text x="${labelW-7}" y="${y+11}" text-anchor="end" fill="#17212b">${
      esc(k.length>22?k.slice(0,21)+"…":k)}</text>`;
    splits.forEach((s,si)=>{
      const v=counts[k][s]||0; if(!v) return;
      const w=v/max*barW;
      svg+=`<rect class="bar" x="${x}" y="${y+1}" width="${w}" height="13" rx="3"
             fill="${P[si%P.length]}"><title>${esc(k)} · ${esc(s)}: ${v}</title></rect>`;
      x+=w;
    });
    svg+=`<text x="${x+6}" y="${y+12}" fill="#5b6b7c">${sum(counts[k]).toLocaleString()}</text>`;
  });
  svg+="</svg>";
  const key=DF.split?`<div class="key">${splits.map((s,i)=>
    `<span><i style="background:${P[i%P.length]}"></i>${esc(s)}</span>`).join("")}</div>`:"";
  const gl=(VARS.find(v=>v[0]===DF.group)||["",""])[1];
  const sl=DF.split?(VARS.find(v=>v[0]===DF.split)||["",""])[1]:"";
  el.innerHTML=`<h2>${esc(gl)}${sl?" split by "+esc(sl):""}</h2>${key}${svg}`;

  $("#ctable").innerHTML=`<table><tr><th>${esc(gl)}</th>${
    splits.map(s=>`<th class="num">${esc(s)}</th>`).join("")}<th class="num">total</th></tr>
    ${keys.map(k=>`<tr><td>${esc(k)}</td>${splits.map(s=>
      `<td class="num">${(counts[k][s]||0).toLocaleString()}</td>`).join("")}
      <td class="num"><b>${sum(counts[k]).toLocaleString()}</b></td></tr>`).join("")}
    </table>`;
}

// ---------- PRISMA ----------
function prismaTab(){
  return `<div class="card"><h2>PRISMA 2020</h2>
    <p class="small muted">Automatic exclusions and your own are both here;
    the title/abstract box says how many of each. A study you included at
    title and abstract counts as included until full-text screening says
    otherwise, so <b>included</b> and <b>full texts still to read</b> overlap
    until you have finished that stage.</p>
    <div id="pwarn"></div>
    <div class="row"><a class="b n" style="text-decoration:none;padding:7px 13px"
      href="/api/prisma.svg" download>download the diagram (SVG)</a></div>
    <div id="pm" style="margin-top:12px" class="muted">loading...</div></div>
    <div class="card"><h2>Included studies by theme</h2><div id="pt"></div></div>
    <div class="card"><h2>Why records were excluded</h2><div id="pr"></div></div>`;
}
async function loadPrisma(){
  const c=await api("/api/prisma");
  const img=document.createElement("img");
  img.src="/api/prisma.svg?"+Date.now();
  img.style.maxWidth="100%";
  const pm=$("#pm"); if(!pm) return;
  pm.innerHTML=""; pm.appendChild(img);
  // A batch written all in one second is reported here, not only in
  // Settings, because this is the diagram that carries its consequence.
  api("/api/bulk-batches").then(b=>{
    const w=$("#pwarn"); if(!w||!b.batches||!b.batches.length) return;
    const ft=b.batches.reduce((a,x)=>a+x.claims_fulltext,0);
    const n=b.batches.reduce((a,x)=>a+x.n,0);
    w.innerHTML=`<div style="background:#fff7f0;border:1px solid #e0b089;
      border-radius:8px;padding:11px;margin:10px 0">
      <b>${n.toLocaleString()} of these decisions were written in bulk</b>, in
      ${b.batches.length===1?"a single second":b.batches.length+" batches"},
      with no reason recorded.
      ${ft?`${ft.toLocaleString()} of them carry a full-text decision, so the
      diagram above is reporting ${ft.toLocaleString()} full texts as retrieved
      and assessed.`:""}
      Re-file them under <b>Settings &rarr; Decisions written in bulk</b>.</div>`;
  });
  const rows=(obj,auto)=>Object.entries(obj).sort((a,b)=>b[1]-a[1]).map(([k,v])=>
    `<tr><td>${esc(k)}</td><td class="num">${v.toLocaleString()}</td>
     <td class="small muted">${auto&&auto[k]?auto[k].toLocaleString()+" by rule":""}</td></tr>`).join("");
  $("#pr").innerHTML=`
   <table><tr><th>Title/abstract stage</th><th class="num">n</th><th></th></tr>
   ${rows(c.reasons,c.auto_reasons)||'<tr><td class="muted">none yet</td><td></td><td></td></tr>'}</table>
   <table style="margin-top:14px"><tr><th>Full-text stage</th><th class="num">n</th><th></th></tr>
   ${rows(c.ft_reasons)||'<tr><td class="muted">none yet</td><td></td><td></td></tr>'}</table>
   <div class="grid3" style="margin-top:14px">
     <div class="stat"><div class="n">${c.ta_todo.toLocaleString()}</div><div class="l">still to screen</div></div>
     <div class="stat"><div class="n">${c.ta_maybe.toLocaleString()}</div><div class="l">maybe</div></div>
     <div class="stat"><div class="n">${c.ft_todo.toLocaleString()}</div><div class="l">full texts still to read</div></div>
   </div>`;
  const tot=c.ft_included;
  const rowsT=(D.labels||[]).map((l,i)=>{
    const n=c.labels[l.id]||0;
    return `<tr><td><span style="color:${l.colour}">&#9632;</span> ${esc(l.name)}</td>
      <td class="num"><b>${n.toLocaleString()}</b></td>
      <td class="num muted">${tot?((n/tot*100).toFixed(0)+"%"):"—"}</td></tr>`;
  }).join("");
  $("#pt").innerHTML=`
    <p class="small muted">The last step of the diagram, broken out. A study
    carrying two themes is counted under each, so these need not sum to the
    ${tot.toLocaleString()} included.</p>
    <table><tr><th>Theme</th><th class="num">included</th>
      <th class="num">of ${tot.toLocaleString()}</th></tr>${rowsT}</table>`;
}

// ---------- Export ----------
function exportTab(){
  const inc=D.records.filter(r=>r.ft==="include"||
    (r.eff==="include"&&!r.ft)).length;
  return `<div class="card">
  <h2>The final dataset</h2>
  <p class="small muted">Everything you have decided, for the studies you
  included: the bibliographic record, your country and theme, the decision and
  its reason, your notes, and the abstract &mdash; including any you typed in
  yourself. Downloads straight to your browser.</p>
  <div class="row">
    <a class="b" style="text-decoration:none;padding:9px 16px;font-weight:600"
       href="/api/download?which=included&fmt=xlsx">Download the final set
       (Excel, ${inc.toLocaleString()} studies)</a>
    <a class="b n" style="text-decoration:none;padding:9px 16px"
       href="/api/download?which=included&fmt=csv">as CSV</a>
  </div>
  <div class="row" style="margin-top:9px">
    <a class="b n" style="text-decoration:none;padding:7px 13px"
       href="/api/download?which=screened&fmt=xlsx">everything screened (Excel)</a>
    <a class="b n" style="text-decoration:none;padding:7px 13px"
       href="/api/download?which=all&fmt=xlsx">everything incl. duplicates (Excel)</a>
    <a class="b n" style="text-decoration:none;padding:7px 13px"
       href="/api/prisma.svg" download>PRISMA diagram (SVG)</a>
  </div></div>

  <div class="card"><h2>ASReview round trip</h2>
  <p class="small muted">ASReview screens titles and abstracts and returns
  relevant / irrelevant. The export leaves out duplicates and anything a rule
  has already excluded, since those carry a PRISMA reason already. Your
  existing decisions go along as prior knowledge so it starts warm.</p>
  <div class="row">
    <a class="b" style="text-decoration:none;padding:8px 15px"
       href="/api/asreview.csv">Download for ASReview</a>
    <a class="b n" style="text-decoration:none;padding:8px 15px"
       href="/api/asreview.csv?priors=0">without my decisions</a>
  </div>
  <div style="margin-top:14px">
    <div class="small muted">Then bring its labels back. Paste the full path
    of the file ASReview exported (.csv, .tsv, .xlsx or .ris).</div>
    <div class="row" style="margin-top:6px">
      <input type="text" id="arpath" style="flex:1;min-width:340px"
        placeholder="C:\Users\...\asreview_export.csv">
      <button class="b n" onclick="arGo(false)">Preview</button>
    </div>
    <p class="small muted" style="margin-top:7px">Preview changes nothing.
    Nothing you have already decided is ever overwritten, and themes, country,
    notes, typed-in abstracts and full-text decisions are not touched at all.
    The decisions file is backed up before anything is written.</p>
    <div id="arout"></div>
  </div></div>

  <div class="card"><h2>Or write a copy next to the corpus</h2>
  <p class="small muted">Same content, saved into your Papers folder with
  today's date in the name, rather than downloaded.</p>
  <table><tr><th>what</th><th>CSV</th><th>Excel</th></tr>
  ${[["all","everything, duplicates included"],
     ["screened","everything except duplicates"],
     ["included","only the studies you included"]].map(([w,l])=>
    `<tr><td>${l}</td>
     <td><button class="b n" onclick="doExport('${w}','csv')">.csv</button></td>
     <td><button class="b n" onclick="doExport('${w}','xlsx')">.xlsx</button></td></tr>`).join("")}
  </table>
  <div id="exout" class="small muted" style="margin-top:10px"></div></div>`;
}
window.arGo=async(apply,prefer)=>{
  const path=$("#arpath").value.trim();
  if(!path){flash("paste the path of the file ASReview exported");return;}
  $("#arout").innerHTML='<p class="muted">reading...</p>';
  const j=await api(apply?"/api/asreview-apply":"/api/asreview-preview",
    {method:"POST",headers:{"content-type":"application/json"},
     body:JSON.stringify({path,prefer:!!prefer})});
  if(j.error){$("#arout").innerHTML=`<p class="muted">${esc(j.error)}</p>`;return;}
  const conf=j.conflict||[];
  let html=`<div class="card" style="margin-top:10px"><div class="small muted">${esc(j.note)}</div>
    <div class="grid3" style="margin-top:10px">
      <div class="stat"><div class="n">${j.apply.toLocaleString()}</div>
        <div class="l">${j.applied_now?"applied":"would be applied"}</div></div>
      <div class="stat"><div class="n">${j.agree.toLocaleString()}</div>
        <div class="l">already decided, same answer</div></div>
      <div class="stat"><div class="n">${j.conflicts.toLocaleString()}</div>
        <div class="l">already decided, different answer</div></div>
      <div class="stat"><div class="n">${j.unlabelled.toLocaleString()}</div>
        <div class="l">not screened in ASReview</div></div>
      <div class="stat"><div class="n">${j.unmatched.toLocaleString()}</div>
        <div class="l">could not be matched</div></div>
      <div class="stat"><div class="n">${(j.duplicate||0).toLocaleString()}</div>
        <div class="l">duplicates, skipped</div></div>
    </div>`;
  if(conf.length){
    html+=`<h2 style="margin-top:14px">Where you and ASReview disagree</h2>
      <p class="small muted">Your decision stands. These are worth a look —
      a disagreement between two passes is a consistency check you would not
      otherwise get.</p>
      <table><tr><th>record</th><th>you said</th><th>ASReview said</th></tr>
      ${conf.map(c=>`<tr><td>${esc(c.title)}</td>
        <td><span class="pill ${c.mine==="include"?"inc":"exc"}">${esc(c.mine)}</span></td>
        <td><span class="pill ${c.theirs==="include"?"inc":"exc"}">${esc(c.theirs)}</span></td>
        </tr>`).join("")}</table>`;
  }
  if(j.applied_now){
    html+=`<p style="margin-top:12px"><b>Done.</b> ${j.applied} label(s) written`
      + (j.overridden?`, ${j.overridden} of your own decisions replaced`:"")
      + (j.backup?`. Your previous decisions were copied to
         <code>${esc(j.backup)}</code> first.`:".")+`</p>`;
  }else{
    html+=`<div class="row" style="margin-top:12px">
      <button class="b" onclick="arGo(true,false)">Apply ${j.apply} label(s)</button>`;
    if(conf.length) html+=`<button class="b n" onclick="arGo(true,true)"
        title="Replaces your own decision where the two disagree">
        Apply, and let ASReview win the ${j.conflicts} disagreement(s)</button>`;
    html+=`</div>`;
  }
  $("#arout").innerHTML=html+"</div>";
  if(j.applied_now) await load();
};
window.doExport=async(which,fmt)=>{
  $("#exout").textContent="writing...";
  const j=await api(`/api/export?which=${which}&fmt=${fmt}`);
  $("#exout").textContent=j.ok?`wrote ${j.file} (${j.rows.toLocaleString()} rows)`
                              :(j.error||"failed");
};

// ---------- Ask ----------
function askTab(){
  return `<div class="card"><h2>Ask about the corpus</h2>
  <p class="small muted">Answers come only from the records it quotes, and
  every claim is cited with a record id. It will not tell you what to include
  or exclude &mdash; that is your decision. Only the records it quotes are
  sent, and only when you press Ask.</p>
  <div class="row">
    <input type="text" id="qq" style="flex:1;min-width:320px"
      placeholder="e.g. what signals are reported as predicting funding success?">
    <select id="scope">
      <option value="all">whole corpus</option>
      <option value="filtered">what the Screen filters show</option>
      <option value="included">only my includes</option>
    </select>
    <button class="b" onclick="doAsk()">Ask</button>
  </div>
  <div id="ans" style="margin-top:12px"></div></div>`;
}
window.doAsk=async()=>{
  const q=$("#qq").value.trim(); if(!q){flash("type a question");return;}
  const s=$("#scope").value;
  let scope=[];
  if(s==="filtered") scope=filtered(F).map(r=>r.id);
  if(s==="included") scope=D.records.filter(r=>r.eff==="include").map(r=>r.id);
  $("#ans").innerHTML='<span class="muted">thinking...</span>';
  const j=await api("/api/ask",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({question:q,scope})});
  if(j.error&&!j.answer){$("#ans").innerHTML=`<p class="muted">${esc(j.error)}</p>`;return;}
  // Every id in the answer becomes a link to the abstract it came from, so
  // checking a claim does not mean switching tabs and hunting for it.
  let body=esc(j.answer).replace(/\n/g,"<br>");
  j.used.forEach(u=>{
    body=body.split("["+u.id+"]").join(
      `<a href="#src-${u.id}" class="pill" style="text-decoration:none">${u.id}</a>`);
  });
  $("#ans").innerHTML=`<div class="ab" style="max-height:none">${body}</div>
    <p class="small muted" style="margin-top:9px">
      Answered from ${j.used.length} record(s)${
        j.clipped?`, ${j.clipped} of which had the abstract cut at ${j.cap.toLocaleString()} characters`:""}.
      A claim with no id is unsupported &mdash; check each one against the text below.</p>
    <details style="margin-top:6px"><summary class="small muted">
      the ${j.used.length} abstracts it read</summary>
      ${j.used.map(u=>`<div id="src-${u.id}" style="margin-top:10px">
        <div class="small"><span class="pill">${esc(u.id)}</span>
          <b>${esc(u.title)}</b> <span class="muted">${esc(u.year||"")}</span></div>
        <div class="ab" style="max-height:170px;margin-top:4px">${
          esc(u.abstract||"(no abstract)")}</div></div>`).join("")}
    </details>`;
};

// ---------- Settings ----------
function settingsTab(){
  return `<div class="card"><h2>Themes</h2>
    <p class="small muted">Used for the guesses, the dashboard and PRISMA.</p>
    <div id="labs">${(D.labels||[]).map((l,i)=>`<div class="row" style="margin-bottom:6px">
      <input type="color" value="${l.colour}" onchange="editLabel(${i},'colour',this.value)">
      <input type="text" style="flex:1" value="${esc(l.name)}"
        onchange="editLabel(${i},'name',this.value)"></div>`).join("")}</div>
  </div>
  <div class="card"><h2>Exclusion reasons</h2>
    <p class="small muted">One per line. These fill the Exclude dropdown and
    appear in PRISMA.</p>
    <textarea id="exc" style="min-height:150px">${esc((D.criteria.exclude||[]).join("\n"))}</textarea>
    <div style="margin-top:8px"><button class="b" onclick="saveCriteria()">save</button></div>
  </div>
  <div class="card"><h2>Inclusion criteria</h2>
    <textarea id="inc" style="min-height:120px">${esc((D.criteria.include||[]).join("\n"))}</textarea>
    <div style="margin-top:8px"><button class="b" onclick="saveCriteria()">save</button></div>
  </div>
  <div class="card"><h2>Ask &mdash; API key</h2>
    <p class="small muted">Held in memory for this session only and never
    written to disk, because this folder syncs to OneDrive. Set
    ANTHROPIC_API_KEY before starting the app to skip this.</p>
    <div class="row"><input type="password" id="key" placeholder="sk-ant-..."
      style="flex:1;min-width:260px">
      <button class="b" onclick="saveKey()">use for this session</button></div>
    <div id="keyst" class="small muted" style="margin-top:7px"></div>
  </div>
  <div class="card"><h2>Ask &mdash; how much it reads</h2>
    <p class="small muted">Your whole corpus cannot be sent: ${D.records.length.toLocaleString()}
    records come to roughly ${Math.round(D.records.length*360/1000).toLocaleString()}k tokens
    against a 200k context window. So each question is answered from the
    closest records instead. That is also why a negative answer is never
    trustworthy &mdash; it means &ldquo;not in these records&rdquo;, not
    &ldquo;not in your corpus&rdquo;. Use the dashboard for counts.</p>
    <div class="row">
      <label class="small muted">records per question
        <input type="number" id="askk" min="1" max="200" value="${D.ask_k}"
          style="width:80px" oninput="askEstimate()"></label>
      <label class="small muted">characters of each abstract
        <input type="number" id="askc" min="200" max="20000" step="100"
          value="${D.ask_chars}" style="width:100px" oninput="askEstimate()"></label>
      <button class="b" onclick="saveAsk()">save</button>
    </div>
    <div id="askest" class="small muted" style="margin-top:8px"></div>
  </div>
  <div class="card"><h2>Who is screening</h2>
    <p class="small muted">Recorded on every decision you make, so a second
    screener's judgements can be told apart from yours. Remembered between
    sessions in <code>review_config.json</code>.</p>
    <div class="row">
      <label class="small muted" style="flex:1">your name
        <input type="text" id="cfgscreener" style="width:100%"
          value="${esc(D.screener||"")}" placeholder="e.g. V. Barcala"></label>
      <label class="small muted" style="flex:1">email for Crossref and OpenAlex
        <input type="text" id="cfgmail" style="width:100%"
          value="${esc(D.mailto||"")}" placeholder="you@university.edu"></label>
      <button class="b" onclick="saveWho()">save</button>
    </div>
    <p class="small muted" style="margin-top:7px">The email is sent to
    Crossref and OpenAlex so their polite pools know who is calling &mdash;
    it gets you faster, more reliable lookups. It is not used for anything
    else and never leaves those requests.</p>
  </div>

  <div class="card"><h2>Your review's configuration</h2>
    <p class="small muted">Themes, criteria and rules above are stored for
    THIS corpus. The defaults they start from, and the keyword patterns that
    guess a country or a theme, live in <code>review_config.json</code> in
    the same folder as this app. Edit that file to set up a different review,
    or to give this one its own vocabulary.</p>
    <table>
      <tr><td><code>themes</code></td><td class="small muted">your labels and their colours</td></tr>
      <tr><td><code>criteria</code></td><td class="small muted">inclusion and exclusion lists</td></tr>
      <tr><td><code>rules</code></td><td class="small muted">the automatic exclusions, including any <code>kw:</code> keyword rule you add</td></tr>
      <tr><td><code>country_patterns</code></td><td class="small muted">name &rarr; regular expression, for the country guess</td></tr>
      <tr><td><code>theme_patterns</code></td><td class="small muted">theme id &rarr; regular expression, for the theme guess</td></tr>
      <tr><td><code>screener</code>, <code>mailto</code></td><td class="small muted">set above</td></tr>
    </table>
    <p class="small muted" style="margin-top:8px">Leave a pattern block empty
    and the built-in ones are used. Restart the app after editing the file.</p>
  </div>

  <div class="card"><h2>Decisions written in bulk</h2>
    <p class="small muted">Every decision records the second it was made in.
    When many share one second they cannot have been made by reading the
    records one at a time &mdash; they came from a bulk action, an import or a
    script. That is ordinary screening when a filter is applied to a whole
    selection. It matters only because of what the batch wrote: a full-text
    decision claims the full texts were retrieved and read, and PRISMA
    reports it that way. A batch disappears from here once every record in
    it carries a reason.</p>
    <div id="bulkout" class="small muted">checking&hellip;</div>
  </div>

  <div class="card"><h2>Files</h2>
    <p class="small muted">Everything is written beside your corpus, named
    after it. The corpus file itself is only ever read, never written to.</p>
    <table><tr><th>file</th><th>what is in it</th></tr>
    ${[["corpus","the harvest you opened &mdash; read only, never modified"],
       ["decisions","every judgement: include/exclude, reasons, themes, country, notes, typed-in abstracts, duplicate choices"],
       ["extra","every record imported from Web of Science, Lens or a hand-made list, with the source on it"],
       ["imports","a log of each file you imported and what it contained"],
       ["rules","which rules are switched on and their settings"],
       ["labels","your theme names and colours"],
       ["criteria","your inclusion and exclusion criteria"],
       ["stop","the stopping-rule thresholds and, once recorded, the decision to stop"]]
      .map(([k,what])=>D.paths&&D.paths[k]?`<tr>
        <td><code>${esc((D.paths[k]||"").split(/[\\\\/]/).pop())}</code></td>
        <td class="small muted">${what}</td></tr>`:"").join("")}
    </table>
    <p class="small muted" style="margin-top:8px">Exports you download, and
    backups taken before a bulk change, land in the same folder with the date
    in the name.</p>
    <p class="small muted">Folder: <code>${esc(D.folder||"")}</code></p>
  </div>`;
}

window.loadBulk=async()=>{
  const el=$("#bulkout"); if(!el) return;
  const j=await api("/api/bulk-batches");
  if(!j.batches||!j.batches.length){
    el.innerHTML="No unexplained bulk batches. Every decision either stands "
      +"on its own or carries a reason.";return;}
  el.innerHTML=j.batches.map(b=>{
    const when=b.stamp.replace("T"," ");
    const ta=Object.entries(b.ta).map(([k,v])=>`${v} &times; ${esc(k)}`).join(", ");
    const ft=Object.entries(b.ft).map(([k,v])=>`${v} &times; ${esc(k)}`).join(", ");
    // Offer the reason that matches what these records actually have in
    // common first, rather than whichever criterion happens to be listed
    // first. Each rule carries its own reason text, so if one of your
    // exclusion criteria says the same thing, that is the one to default to.
    const want=((b.checks||[])[0]||{}).reason||"";
    const all=(D.criteria.exclude||[]).slice();
    const hit=all.findIndex(r=>r.toLowerCase()===want.toLowerCase());
    if(hit>0) all.unshift(all.splice(hit,1)[0]);
    else if(want&&hit<0) all.unshift(want);
    const reasons=all.map(r=>`<option>${esc(r)}</option>`).join("");
    const rules=(b.checks||[]).map(c=>
      `<option value="${esc(c.id)}">${esc(c.name)} &mdash; fits ${c.hits} of ${c.of}</option>`).join("");
    return `<div style="border:1px solid var(--line);border-radius:8px;padding:12px;margin-bottom:12px">
      <div style="color:var(--ink)"><b>${b.n.toLocaleString()} decisions written in
        one second</b> &mdash; ${esc(when)}</div>
      <div style="margin-top:6px">title/abstract: ${ta} &nbsp;&middot;&nbsp;
        full text: ${ft}</div>
      <div style="margin-top:4px">${b.unexplained.toLocaleString()} of them record no reason at all.</div>
      ${b.claims_fulltext?`<div style="margin-top:6px;color:var(--red)">
        <b>${b.claims_fulltext.toLocaleString()} carry a full-text decision</b>, so PRISMA
        is reporting that many full texts as retrieved and assessed.</div>`:""}
      ${(b.checks||[]).length?`<div style="margin-top:6px">What these records have in
        common: ${b.checks.map(c=>`${esc(c.name)} (${c.hits}/${c.of})`).join(" &middot; ")}</div>`:""}
      ${(b.sample||[]).length?`<div class="small muted" style="margin-top:6px">
        e.g. ${b.sample.map(s=>esc(s.title)).join(" &middot; ")}</div>`:""}
      <div class="row" style="margin-top:10px;flex-wrap:wrap;gap:6px">
        <select id="r-${esc(b.stamp)}" style="min-width:260px">${reasons}</select>
        <select id="k-${esc(b.stamp)}" style="min-width:240px">
          <option value="">apply to every record in the batch</option>${rules}</select>
      </div>
      <div class="row" style="margin-top:8px;flex-wrap:wrap;gap:6px">
        <button class="b" onclick="bulkFix('${esc(b.stamp)}','exclude')">
          Re-file as excluded at title/abstract</button>
        <button class="b n" onclick="bulkFix('${esc(b.stamp)}','explain')">
          Keep as it is, just add this reason</button>
        <button class="b n" onclick="bulkFix('${esc(b.stamp)}','clear')">
          Return all ${b.n.toLocaleString()} to unscreened</button>
      </div>
      <p class="small muted" style="margin-top:7px">Choosing a rule in the second
      box guards the reason: a record the rule does not fit would be given a
      reason that is untrue of it, so it is handed back unscreened instead.
      The decisions file is backed up before anything is written, and the
      timestamps are kept, so your screening order is unchanged.</p>
    </div>`;}).join("");
};
window.bulkFix=async(stamp,action)=>{
  const reason=($("#r-"+stamp)||{}).value||"";
  const rule=($("#k-"+stamp)||{}).value||"";
  if(action==="clear"&&!confirm("Return every record in this batch to "
      +"unscreened? The file is backed up first.")) return;
  const j=await api("/api/bulk-fix",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({stamp,action,reason,rule})});
  if(j.error){flash(j.error);return;}
  flash(`${j.changed} re-filed${j.handed_back?`, ${j.handed_back} handed back`:""}`
        +`${j.backup?` — backup ${j.backup}`:""}`);
  await load(); loadBulk();
};
window.saveWho=async()=>{
  const j=await api("/api/screener",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({name:$("#cfgscreener").value,
                         mailto:$("#cfgmail").value})});
  if(j.error){flash(j.error);return;}
  await load(); flash("saved to review_config.json");
};
window.editLabel=async(i,f,v)=>{
  const labels=D.labels.map((l,j)=>j===i?Object.assign({},l,{[f]:v}):l);
  await api("/api/labels",{method:"POST",headers:{"content-type":"application/json"},
    body:JSON.stringify({labels})});
  await load(); flash("saved");
};
window.saveCriteria=async()=>{
  const criteria={include:$("#inc").value.split("\n").map(s=>s.trim()).filter(Boolean),
                  exclude:$("#exc").value.split("\n").map(s=>s.trim()).filter(Boolean)};
  await api("/api/criteria",{method:"POST",headers:{"content-type":"application/json"},
    body:JSON.stringify({criteria})});
  await load(); flash("saved");
};
window.askEstimate=()=>{
  const k=+($("#askk").value||0), c=+($("#askc").value||0);
  if(!k||!c){$("#askest").textContent="";return;}
  const perRec=Math.round(c/4)+60, total=perRec*k;
  const withAbs=D.records.filter(r=>r.abstract);
  const whole=withAbs.length?withAbs.filter(r=>r.abstract.length<=c).length:0;
  const pct=withAbs.length?Math.round(whole/withAbs.length*100):0;
  $("#askest").innerHTML=
    `About <b>${total.toLocaleString()} tokens</b> per question `
    + (total>190000?'<span class="pill exc">too big for a 200k context</span>':
       total>60000?'<span class="pill may">large</span>':
       '<span class="pill inc">comfortable</span>')
    + `. At ${c.toLocaleString()} characters, ${pct}% of your abstracts `
    + `(${whole.toLocaleString()} of ${withAbs.length.toLocaleString()}) are sent whole.`;
};
window.saveAsk=async()=>{
  const j=await api("/api/ai-key",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({k:+$("#askk").value,chars:+$("#askc").value})});
  if(j.ok){ await load(); flash(`now reading ${j.k} records, ${j.chars} chars each`); }
};
window.saveKey=async()=>{
  const j=await api("/api/ai-key",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({key:$("#key").value})});
  $("#key").value="";
  $("#keyst").textContent=j.ready?"key set for this session":"no key set";
};
$("#screener").addEventListener("change",async e=>{
  await api("/api/screener",{method:"POST",
    headers:{"content-type":"application/json"},
    body:JSON.stringify({name:e.target.value})});
  flash("screener set");
});

load();
</script></body></html>
"""


# ==========================================================================
#  ASReview round trip
#
#  ASReview screens titles and abstracts and nothing else: its labels are
#  relevant / irrelevant, with no themes, no country, no reasons. So the
#  export carries only what it reads, plus openalex_id so every record can be
#  found again, and the import touches only the title/abstract decision.
#  Everything you have built here — themes, country, notes, typed-in
#  abstracts, full-text decisions, duplicate choices — is never written to by
#  the import and cannot be lost by it.
#
#  Column names follow ASReview's documented list: title, abstract, authors,
#  doi, keywords, and `included` for labels, where 1 is relevant, 0 is
#  irrelevant and a blank means not seen.
# ==========================================================================

ASREVIEW_COLUMNS = ["openalex_id", "doi", "title", "abstract", "authors",
                    "year", "keywords", "included"]
# Every spelling ASReview accepts for the label column, so its export reads
# back whichever one it happened to write.
ASREVIEW_LABEL_COLUMNS = ["included", "label", "label_included",
                          "included_label", "final_included",
                          "included_final", "included_flag", "include",
                          "asreview_label"]


def asreview_rows(with_priors=True):
    """The records worth screening, in the shape ASReview reads.

    Duplicates and anything a rule has already excluded are left out: they
    carry a PRISMA reason already, and there is nothing to gain from
    screening them a second time."""
    dups, _ = dup_state()
    out = []
    for r in STATE["records"]:
        if dups.get(r["id"]) == "duplicate":
            continue
        dec, _why, how = effective(r)
        if dec == "exclude" and how == "auto":
            continue
        prior = ""
        if with_priors:
            d = STATE["decisions"].get(r["id"], {})
            if d.get("ta_decision") == "include":
                prior = "1"
            elif d.get("ta_decision") == "exclude":
                prior = "0"
        out.append({
            "openalex_id": r["id"], "doi": r["doi"], "title": r["title"],
            "abstract": r["abstract"], "authors": r["authors"],
            "year": r["year"], "keywords": r["keywords"], "included": prior,
        })
    return out


def _ris_records(text):
    """Minimal RIS reader. ASReview stores its labels in the N1 tag."""
    recs, cur = [], {}
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        m = re.match(r"^([A-Z][A-Z0-9])  - ?(.*)$", line)
        if not m:
            continue
        tag, val = m.group(1), m.group(2).strip()
        if tag == "TY":
            if cur:
                recs.append(cur)
            cur = {}
        cur.setdefault(tag, []).append(val)
        if tag == "ER":
            recs.append(cur)
            cur = {}
    if cur:
        recs.append(cur)
    return recs


def read_asreview_file(path):
    """Read whatever ASReview exported. Returns (rows, note).

    rows carry id / doi / title / label, where label is 1, 0 or None."""
    # Windows "Copy as path" wraps the path in double quotes, and pasting
    # that verbatim is the obvious thing to do, so strip them rather than
    # fail on it. Same for stray whitespace.
    raw = str(path or "").strip().strip('"').strip("'").strip()
    if not raw:
        raise RuntimeError("Paste the path of the file ASReview exported.")
    p = Path(raw).expanduser()
    if not p.is_absolute():
        beside = SCRIPT_DIR / p
        p = beside if beside.exists() else p.resolve()
    if not p.exists():
        raise RuntimeError(
            f"No file at {p}\n"
            "Either drop the file into your Papers folder and type just its "
            "name, or paste the full path (File Explorer: right-click the "
            "file, Copy as path).")
    suffix = p.suffix.lower()

    out = []
    if suffix in (".ris", ".txt"):
        text = p.read_text(encoding="utf-8", errors="replace")
        for rec in _ris_records(text):
            notes = " ".join(rec.get("N1", []) + rec.get("N2", []))
            low = notes.lower()
            label = (1 if "asreview_relevant" in low else
                     0 if "asreview_irrelevant" in low else None)
            out.append({
                "id": (rec.get("ID") or [""])[0].strip(),
                "doi": norm_doi((rec.get("DO") or rec.get("DI") or [""])[0]),
                "title": " ".join(rec.get("TI") or rec.get("T1") or []),
                "label": label})
        return out, f"{p.name}: RIS, {len(out)} records"

    if suffix in (".csv", ".tsv", ".tab", ".xlsx", ".xlsm"):
        rows = read_table(p)
        if not rows:
            raise RuntimeError(f"{p.name} has no rows.")
        cols = {c.strip().lower(): c for c in rows[0]}
        label_col = next((cols[c] for c in ASREVIEW_LABEL_COLUMNS
                          if c in cols), None)
        id_col = cols.get("openalex_id") or cols.get("id")
        doi_col = cols.get("doi")
        title_col = cols.get("title") or cols.get("primary_title")
        if not label_col:
            raise RuntimeError(
                f"{p.name} has no label column. Looked for: "
                f"{', '.join(ASREVIEW_LABEL_COLUMNS[:5])} and a few others.\n"
                f"The file has: {', '.join(list(rows[0])[:12])}")
        if not (id_col or doi_col or title_col):
            raise RuntimeError(
                f"{p.name} has nothing to match on — it needs at least one of "
                "openalex_id, doi or title.")
        for r in rows:
            raw = str(r.get(label_col, "")).strip().lower()
            label = (1 if raw in ("1", "1.0", "true", "yes", "relevant",
                                  "included")
                     else 0 if raw in ("0", "0.0", "false", "no", "irrelevant",
                                       "excluded")
                     else None)
            out.append({
                "id": (r.get(id_col) or "").strip() if id_col else "",
                "doi": norm_doi(r.get(doi_col)) if doi_col else "",
                "title": (r.get(title_col) or "") if title_col else "",
                "label": label})
        return out, f"{p.name}: {len(out)} rows, labels from '{label_col}'"

    raise RuntimeError(
        f"Cannot read {p.suffix or 'that'} files. ASReview exports .csv, "
        ".tsv, .xlsx or .ris.")


def match_asreview(rows):
    """Work out what importing would do, changing nothing.

    Matching runs id, then DOI, then normalised title — the same ladder the
    carry-over uses, so it survives ASReview renaming or reordering."""
    by_id = {r["id"]: r for r in STATE["records"]}
    by_doi = {r["ndoi"]: r for r in STATE["records"] if r["ndoi"]}
    by_title = {r["ntitle"]: r for r in STATE["records"] if r["ntitle"]}
    plan = {"apply": [], "conflict": [], "agree": [], "unlabelled": [],
            "unmatched": [], "duplicate": []}
    dups, _ = dup_state()
    seen = set()
    for row in rows:
        rec = (by_id.get(row["id"]) if row["id"] else None) \
            or (by_doi.get(row["doi"]) if row["doi"] else None) \
            or (by_title.get(norm_title(row["title"])) if row["title"] else None)
        if rec is None:
            plan["unmatched"].append(row)
            continue
        if rec["id"] in seen:
            continue
        seen.add(rec["id"])
        if row["label"] is None:
            plan["unlabelled"].append(rec["id"])
            continue
        incoming = "include" if row["label"] == 1 else "exclude"
        if dups.get(rec["id"]) == "duplicate":
            plan["duplicate"].append((rec, incoming))
            continue
        mine = STATE["decisions"].get(rec["id"], {}).get("ta_decision", "")
        if not mine:
            plan["apply"].append((rec, incoming))
        elif mine == incoming:
            plan["agree"].append((rec, incoming))
        else:
            plan["conflict"].append((rec, mine, incoming))
    return plan


def apply_asreview(plan, prefer_asreview=False):
    """Write the labels in. Backs the decisions file up first.

    Only ta_decision, ta_reason and ta_source are ever touched. Themes,
    country, notes, typed-in abstracts, full-text decisions and duplicate
    choices are not written to at all, so none of them can be lost here."""
    path = STATE["paths"]["decisions"]
    backup = ""
    if os.path.exists(path):
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = f"{os.path.splitext(path)[0]}_backup_{stamp}.csv"
        import shutil
        shutil.copy2(path, backup)

    now = datetime.datetime.now().isoformat(timespec="seconds")

    def write(rec, decision, overwrote=False):
        d = STATE["decisions"].setdefault(rec["id"],
                                          {k: "" for k in DECISION_FIELDS})
        d["ta_decision"] = decision
        d["ta_source"] = "asreview"
        d["ta_reason"] = ("Excluded during ASReview screening"
                          if decision == "exclude" else "")
        d["doi"], d["title"] = rec["doi"], rec["title"]
        d["screener"] = STATE["screener"] or d.get("screener", "")
        d["updated_at"] = now
        return overwrote

    n = 0
    for rec, decision in plan["apply"]:
        write(rec, decision)
        n += 1
    overridden = 0
    if prefer_asreview:
        for rec, _mine, incoming in plan["conflict"]:
            write(rec, incoming, True)
            overridden += 1
    save_decisions(path, STATE["decisions"])
    return {"applied": n, "overridden": overridden,
            "backup": os.path.basename(backup) if backup else ""}


def export_to_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=EXPORT_COLUMNS, delimiter=";",
                           quoting=csv.QUOTE_ALL, lineterminator="\r\n",
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def build_workbook(rows):
    """One workbook builder, used both for saving to disk and for a download."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "review"
    ws.append(EXPORT_COLUMNS)
    for r in rows:
        ws.append([str(r.get(c, "") or "")[:32767] for c in EXPORT_COLUMNS])
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="44546A")
    for i in range(1, len(EXPORT_COLUMNS) + 1):
        c = ws.cell(row=1, column=i)
        c.font, c.fill = head, fill
    widths = {"title": 60, "abstract": 90, "authors": 38, "apa_citation": 60,
              "notes": 34, "journal": 30, "institutions": 34, "themes": 28,
              "ta_reason": 28, "ft_reason": 28, "country": 20}
    for i, name in enumerate(EXPORT_COLUMNS, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(name, 15)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EXPORT_COLUMNS))}{len(rows)+1}"
    return wb


def export_to_xlsx(path, rows):
    wb = build_workbook(rows)
    if wb is None:
        return False
    wb.save(path)
    return True


def xlsx_bytes(rows):
    """The same workbook as bytes, for streaming straight to the browser."""
    wb = build_workbook(rows)
    if wb is None:
        return None
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==========================================================================
#  HTTP
# ==========================================================================

def bibliometrics(which="screened"):
    """A bibliometric profile of the corpus, the kind Lens shows.

    `which` is all / screened (duplicates removed) / included.

    Counts only. Nothing here is a judgement and nothing is written: it is a
    description of the literature you have, for the results section of a
    review that reports on its own field."""
    dups, _ = dup_state()
    recs = STATE["records"]
    if which == "screened":
        recs = [r for r in recs if dups.get(r["id"]) != "duplicate"]
    elif which == "included":
        recs = [r for r in recs if dups.get(r["id"]) != "duplicate"
                and is_included(r)]

    def split_list(v):
        """Authors, institutions and keywords arrive joined by | or ;."""
        s = str(v or "")
        parts = re.split(r"\s*[|;]\s*", s) if ("|" in s or ";" in s) else \
                ([s] if s.strip() else [])
        return [p.strip() for p in parts if p.strip()]

    # Web of Science shouts its metadata — NETHERLANDS where OpenAlex writes
    # Netherlands — so the same country, journal or institution arrives in two
    # spellings and gets counted twice. Names are folded case-insensitively
    # and reported in the spelling that occurs most often.
    spellings = {}

    def fold(name):
        k = re.sub(r"\s+", " ", str(name or "")).strip().lower()
        if not k:
            return ""
        spellings.setdefault(k, Counter())[str(name).strip()] += 1
        return k

    def best_spelling(key):
        """The spelling to show. Web of Science writes NETHERLANDS, so the
        shoutiest version is often the commonest; a mixed-case alternative is
        preferred whenever one exists."""
        seen = spellings.get(key)
        if not seen:
            return key
        mixed = [(n, w) for w, n in seen.items() if not w.isupper()]
        if mixed:
            return max(mixed)[1]
        return seen.most_common(1)[0][0]

    def unfold(ctr, k=15):
        return [{"name": best_spelling(key), "n": n}
                for key, n in ctr.most_common(k)]

    years, authors, journals, insts, fields, kws = (
        Counter(), Counter(), Counter(), Counter(), Counter(), Counter())
    countries, types, langs, oa = Counter(), Counter(), Counter(), Counter()
    cites_by_year, n_by_year = Counter(), Counter()
    author_cites, author_papers = Counter(), Counter()
    total_c, uncited = 0, 0
    pairs = Counter()                 # co-authorship, for the collaboration view

    for r in recs:
        y = as_int(r.get("year"))
        c = as_int(r.get("cited_by"))
        total_c += c
        if c == 0:
            uncited += 1
        if y:
            years[y] += 1
            cites_by_year[y] += c
            n_by_year[y] += 1
        au = [fold(x) for x in split_list(r.get("authors"))[:25]]
        au = [x for x in au if x]
        for a in au:
            authors[a] += 1
            author_cites[a] += c
            author_papers[a] += 1
        for i, a in enumerate(au[:8]):
            for b in au[i + 1:8]:
                pairs[tuple(sorted((a, b)))] += 1
        if fold(r.get("journal")):
            journals[fold(r.get("journal"))] += 1
        for i in split_list(r.get("institutions"))[:12]:
            if fold(i):
                insts[fold(i)] += 1
        if fold(r.get("field")):
            fields[fold(r.get("field"))] += 1
        for k in split_list(r.get("keywords"))[:20]:
            if fold(k):
                kws[fold(k)] += 1
        for cy in split_list(r.get("country_guess") or r.get("countries")):
            if fold(cy):
                countries[fold(cy)] += 1
        types[(r.get("type") or "unknown").lower()] += 1
        langs[(r.get("language") or "unknown").lower()] += 1
        oa[(r.get("oa_status") or "unknown").lower()] += 1

    cites = sorted((as_int(r.get("cited_by")) for r in recs), reverse=True)
    n = len(recs) or 1
    # h-index of the SET, the usual bibliometric summary of a collection.
    h = 0
    for i, c in enumerate(cites, start=1):
        if c >= i:
            h = i
        else:
            break
    top = unfold
    return {
        "which": which, "records": len(recs),
        "with_abstract": sum(1 for r in recs if r.get("abstract")),
        "with_doi": sum(1 for r in recs if r.get("doi")),
        "citations_total": total_c,
        "citations_mean": round(total_c / n, 1),
        "citations_median": cites[len(cites) // 2] if cites else 0,
        "uncited": uncited, "h_index": h,
        "years": [{"name": str(y), "n": years[y]} for y in sorted(years)],
        "cites_per_year": [{"name": str(y),
                            "n": round(cites_by_year[y] / max(1, n_by_year[y]), 1)}
                           for y in sorted(cites_by_year)],
        "authors": top(authors),
        # Citations are a property of a PAPER, so every co-author of a heavily
        # cited one carries its whole count. The paper count is reported
        # alongside so a row showing one paper's citations is recognisable as
        # exactly that.
        "author_impact": [
            {"name": best_spelling(a),
             "n": author_cites[a], "papers": author_papers[a]}
            for a, _ in author_cites.most_common(15)],
        "journals": top(journals), "institutions": top(insts),
        "fields": top(fields, 10), "keywords": top(kws, 25),
        "countries": top(countries, 15), "types": top(types, 10),
        "languages": top(langs, 8), "open_access": top(oa, 8),
        "collaborations": [{"name": f"{best_spelling(a)} + {best_spelling(b)}", "n": c}
                           for (a, b), c in pairs.most_common(12)],
        "sources": [{"name": SOURCE_LABELS.get(s, s), "n": c} for s, c in
                    Counter((r.get("source") or "") for r in recs).most_common()],
    }


def public_records():
    """Records plus everything the browser needs, computed once per request."""
    dups, chosen = dup_state()
    out = []
    for r in STATE["records"]:
        d = STATE["decisions"].get(r["id"], {})
        dec, why, how = effective(r)
        auto_why, auto_rule = auto_reason(r)
        out.append({
            "id": r["id"], "title": r["title"], "authors": r["authors"],
            "year": r["year"], "journal": r["journal"], "doi": r["doi"],
            "abstract": r["abstract"], "language": r["language"],
            "type": r["type"], "cited_by": r["cited_by"],
            "topic": r["topic"], "subfield": r["subfield"], "field": r["field"],
            "institutions": r["institutions"], "countries": r["countries"],
            "europe": r["europe"], "oa_status": r["oa_status"],
            "url": r["url"], "apa": r["apa"],
            "abstract_source": ("added by hand" if r.get("abstract_added")
                                else r["abstract_source"]),
            "abstract_added": bool(r.get("abstract_added")),
            "edited": r.get("edited_fields") or [],
            "merged_from": r.get("merged_from") or {},
            "src": r.get("source") or "openalex",
            "src_label": SOURCE_LABELS.get(r.get("source") or "openalex",
                                           r.get("source") or "openalex"),
            "retracted": r["retracted"],
            "cguess": r["country_guess"], "cbasis": r["country_basis"],
            "tguess": r["theme_guess"],
            "country": d.get("country", ""),
            "labels": d.get("labels", ""),
            "ta": d.get("ta_decision", ""), "ta_reason": d.get("ta_reason", ""),
            "ft": d.get("ft_decision", ""), "ft_reason": d.get("ft_reason", ""),
            "notes": d.get("notes", ""),
            "instrument_type": d.get("instrument_type", ""),
            "dup": dups.get(r["id"], ""),
            "dup_manual": d.get("dup_status", ""),
            "auto": auto_why, "auto_rule": auto_rule,
            "eff": dec, "eff_reason": why, "eff_src": how,
        })
    return out, chosen


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8",
              extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj))

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    # ---------------- GET ----------------

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        try:
            if u.path == "/":
                return self._send(200, PAGE, "text/html; charset=utf-8")
            if u.path == "/api/data":
                with LOCK:
                    recs, chosen = public_records()
                    return self._json({
                        "records": recs, "labels": STATE["labels"],
                        "criteria": STATE["criteria"], "rules": STATE["rules"],
                        "dupes": STATE["dupes"], "dup_chosen": chosen,
                        "preview": rule_preview(),
                        "screener": STATE["screener"],
                        "mailto": STATE.get("mailto",""),
                        "sources": source_counts(),
                        "extra": [dict(e, source_label=SOURCE_LABELS.get(
                            e.get("source"), e.get("source", "")))
                            for e in (STATE.get("extra") or [])],
                        "ask_k": STATE["ask_k"], "ask_chars": STATE["ask_chars"],
                        "source": STATE["source"], "year": THIS_YEAR,
                        "palette": PALETTE,
                        "paths": {k: os.path.basename(v)
                                  for k, v in STATE["paths"].items()},
                        # The folder is sent once, separately, so the Files
                        # panel can tell you where to look without repeating
                        # the whole path on every row.
                        "folder": os.path.dirname(
                            STATE["paths"]["decisions"]) or ".",
                    })
            if u.path == "/api/bulk-batches":
                with LOCK:
                    return self._json({"batches": bulk_batches(),
                                       "min": BULK_MIN})

            if u.path == "/api/stopping":
                with LOCK:
                    return self._json(stopping_status())
            if u.path == "/api/prisma":
                with LOCK:
                    return self._json(prisma_counts())
            if u.path == "/api/prisma.svg":
                with LOCK:
                    svg = prisma_svg(prisma_counts())
                return self._send(200, svg, "image/svg+xml; charset=utf-8",
                                  {"Content-Disposition":
                                   'attachment; filename="prisma.svg"'})
            if u.path == "/api/rank":
                want = (q.get("model") or [""])[0]
                with LOCK:
                    if want:
                        STATE["ranker"] = want
                    scores, npos, nneg, engine = rank_records()
                return self._json({"scores": scores, "pos": npos, "neg": nneg,
                                   "engine": engine,
                                   "model": STATE.get("ranker") or "nb",
                                   "asreview": asreview_available(),
                                   "asreview_version": asreview_version(),
                                   "models": {k: v[0] for k, v
                                              in ASREVIEW_MODELS.items()}})
            if u.path == "/api/bibliometrics":
                which = (q.get("which") or ["screened"])[0]
                with LOCK:
                    return self._json(bibliometrics(which))

            if u.path == "/api/compare-models":
                with LOCK:
                    return self._json(compare_rankers())

            if u.path == "/api/ai-status":
                return self._json({"ready": bool(STATE["api_key"]),
                                   "model": STATE["model"]})
            if u.path == "/api/asreview.csv":
                priors = (q.get("priors") or ["1"])[0] != "0"
                with LOCK:
                    rows = asreview_rows(priors)
                buf = io.StringIO()
                w = csv.DictWriter(buf, fieldnames=ASREVIEW_COLUMNS,
                                   quoting=csv.QUOTE_ALL, lineterminator="\r\n",
                                   extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)
                stem = Path(STATE["source"]).with_suffix("").name
                return self._send(
                    200, "﻿" + buf.getvalue(), "text/csv; charset=utf-8",
                    {"Content-Disposition":
                     f'attachment; filename="{stem}_for_asreview.csv"'})

            if u.path == "/api/download":
                which = (q.get("which") or ["included"])[0]
                fmt = (q.get("fmt") or ["xlsx"])[0]
                with LOCK:
                    rows = export_rows(which)
                stamp = datetime.date.today().isoformat()
                stem = Path(STATE["source"]).with_suffix("").name
                name = f"{stem}_{which}_{stamp}"
                if fmt == "xlsx":
                    blob = xlsx_bytes(rows)
                    if blob is None:
                        return self._json(
                            {"error": "Writing .xlsx needs openpyxl. "
                                      "Run: pip install openpyxl"}, 400)
                    return self._send(
                        200, blob,
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet",
                        {"Content-Disposition":
                         f'attachment; filename="{name}.xlsx"'})
                buf = io.StringIO()
                w = csv.DictWriter(buf, fieldnames=EXPORT_COLUMNS,
                                   delimiter=";", quoting=csv.QUOTE_ALL,
                                   lineterminator="\r\n",
                                   extrasaction="ignore")
                w.writeheader()
                w.writerows(rows)
                return self._send(
                    200, "﻿" + buf.getvalue(), "text/csv; charset=utf-8",
                    {"Content-Disposition": f'attachment; filename="{name}.csv"'})

            if u.path == "/api/export":
                which = (q.get("which") or ["all"])[0]
                fmt = (q.get("fmt") or ["csv"])[0]
                with LOCK:
                    rows = export_rows(which)
                    stem = Path(STATE["paths"]["decisions"]).with_suffix("")
                    stamp = datetime.date.today().isoformat()
                    base = f"{stem}_{which}_{stamp}"
                if fmt == "xlsx":
                    ok = export_to_xlsx(base + ".xlsx", rows)
                    if not ok:
                        return self._json(
                            {"error": "Writing .xlsx needs openpyxl. "
                                      "Run: pip install openpyxl"}, 400)
                    return self._json({"ok": True, "rows": len(rows),
                                       "file": os.path.basename(base + ".xlsx")})
                export_to_csv(base + ".csv", rows)
                return self._json({"ok": True, "rows": len(rows),
                                   "file": os.path.basename(base + ".csv")})
            return self._send(404, "not found", "text/plain; charset=utf-8")
        except Exception as e:
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)

    # ---------------- POST ----------------

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        try:
            data = self._body()

            if u.path == "/api/ask":
                question = (data.get("question") or "").strip()
                scope = set(data.get("scope") or [])
                if not question:
                    return self._json({"error": "Ask something first."}, 400)
                with LOCK:
                    hits = retrieve(question, scope)
                    key, model = STATE["api_key"], STATE["model"]
                answer, err = ask_claude(question, hits, key, model)
                cap = STATE.get("ask_chars") or MAX_ABSTRACT_CHARS
                clipped = sum(1 for h in hits if len(h["abstract"] or "") > cap)
                return self._json({
                    "answer": answer, "error": err,
                    "clipped": clipped, "cap": cap,
                    "used": [{"id": h["id"], "title": h["title"],
                              "year": h["year"],
                              "abstract": (h["abstract"] or "")[:cap]}
                             for h in hits]})

            with LOCK:
                if u.path == "/api/save":
                    rid = data.get("id")
                    if not rid:
                        return self._json({"error": "no id"}, 400)
                    self._patch(rid, data.get("patch") or {})
                    save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
                    recs, chosen = public_records()
                    one = next((r for r in recs if r["id"] == rid), None)
                    return self._json({"ok": True, "record": one,
                                       "preview": rule_preview()})

                if u.path == "/api/bulk":
                    ids = data.get("ids") or []
                    patch = data.get("patch") or {}
                    for rid in ids:
                        self._patch(rid, patch)
                    save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
                    return self._json({"ok": True, "n": len(ids)})

                if u.path == "/api/rules":
                    incoming = data.get("rules") or []
                    by_id = {r["id"]: r for r in STATE["rules"]}
                    for r in incoming:
                        tgt = by_id.get(r.get("id"))
                        if not tgt:
                            continue
                        tgt["enabled"] = bool(r.get("enabled"))
                        if isinstance(r.get("params"), dict):
                            tgt["params"].update(r["params"])
                    save_json(STATE["paths"]["rules"], STATE["rules"])
                    return self._json({"ok": True, "rules": STATE["rules"],
                                       "preview": rule_preview()})

                if u.path == "/api/dupkeep":
                    key, keep = data.get("key"), data.get("keep")
                    grp = next((g for g in STATE["dupes"] if g["key"] == key), None)
                    if not grp:
                        return self._json({"error": "unknown group"}, 400)
                    for rid in grp["ids"]:
                        d = STATE["decisions"].setdefault(
                            rid, {k: "" for k in DECISION_FIELDS})
                        d["dup_status"] = "keep" if rid == keep else "duplicate"
                        d["updated_at"] = datetime.datetime.now().isoformat(
                            timespec="seconds")
                        rec = next((x for x in STATE["records"]
                                    if x["id"] == rid), None)
                        if rec:
                            d["doi"], d["title"] = rec["doi"], rec["title"]
                    # A different record now stands for the group, so it has
                    # to take over the gaps the old one was filling.
                    merge_duplicates()
                    save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
                    return self._json({"ok": True})

                if u.path == "/api/bulk-fix":
                    try:
                        res = apply_bulk_fix(
                            str(data.get("stamp") or ""),
                            str(data.get("action") or ""),
                            clean(data.get("reason")),
                            str(data.get("rule") or ""))
                    except RuntimeError as e:
                        return self._json({"error": str(e)}, 400)
                    res["ok"] = True
                    return self._json(res)

                if u.path == "/api/labels":
                    STATE["labels"] = data.get("labels") or STATE["labels"]
                    save_json(STATE["paths"]["labels"], STATE["labels"])
                    return self._json({"ok": True, "labels": STATE["labels"]})

                if u.path == "/api/criteria":
                    STATE["criteria"] = data.get("criteria") or STATE["criteria"]
                    save_json(STATE["paths"]["criteria"], STATE["criteria"])
                    return self._json({"ok": True})

                if u.path in ("/api/asreview-preview", "/api/asreview-apply"):
                    try:
                        rows, note = read_asreview_file(data.get("path") or "")
                    except RuntimeError as e:
                        return self._json({"error": str(e)}, 400)
                    plan = match_asreview(rows)
                    out = {
                        "note": note,
                        "apply": len(plan["apply"]),
                        "agree": len(plan["agree"]),
                        "conflict": [
                            {"id": r["id"], "title": r["title"][:90],
                             "mine": m, "theirs": t}
                            for r, m, t in plan["conflict"][:50]],
                        "conflicts": len(plan["conflict"]),
                        "unlabelled": len(plan["unlabelled"]),
                        "unmatched": len(plan["unmatched"]),
                        "duplicate": len(plan["duplicate"]),
                    }
                    if u.path.endswith("preview"):
                        out["preview"] = True
                        return self._json(out)
                    res = apply_asreview(plan, bool(data.get("prefer")))
                    out.update(res)
                    out["applied_now"] = True
                    return self._json(out)

                if u.path in ("/api/import-preview", "/api/import-apply"):
                    raw = str(data.get("path") or "").strip().strip('"').strip("'")
                    if not raw:
                        return self._json({"error": "Paste the path of the "
                                                    "export file."}, 400)
                    p = Path(raw).expanduser()
                    if not p.is_absolute():
                        beside = SCRIPT_DIR / p
                        p = beside if beside.exists() else p.resolve()
                    if not p.exists():
                        return self._json({"error":
                            f"No file at {p}\nDrop it into your Papers folder "
                            "and type just its name, or paste the full path "
                            "(File Explorer: right-click, Copy as path)."}, 400)
                    try:
                        rows = read_any_table(p)
                    except RuntimeError as e:
                        return self._json({"error": str(e)}, 400)
                    except Exception as e:
                        return self._json({"error": f"Could not read {p.name}: "
                                                    f"{e}"}, 400)
                    if not rows:
                        return self._json({"error": f"{p.name} has no rows."}, 400)
                    src = data.get("source") or detect_source(rows)
                    mapped, picked = map_source_rows(rows, src)
                    if not mapped:
                        return self._json({"error":
                            f"{p.name} was read ({len(rows)} rows) but no title "
                            "or DOI column was recognised.\nColumns found: "
                            + ", ".join(list(rows[0])[:14])}, 400)

                    # Every usable row is STORED, including rows that another
                    # database already gave you. Dropping them at import was
                    # what broke the diagram: a record Web of Science also
                    # returned is a duplicate the review removed, and PRISMA
                    # has to report it — but if it was never stored, the
                    # Duplicates tab cannot show it, you cannot check it, and
                    # the better of the two versions can never win. On this
                    # corpus that hid 228 of 259 duplicates and threw away 23
                    # abstracts the harvest did not have.
                    #
                    # Overlap is resolved afterwards, in the corpus, by the
                    # same duplicate machinery you can already see and
                    # override.
                    held = {e.get("id") for e in (STATE.get("extra") or [])}
                    own = {r["id"] for r in STATE["records"]
                           if r.get("source") == src}
                    seen_here = set()
                    new, again, within, unusable = [], 0, 0, 0
                    for m in mapped:
                        k = record_key(src, m["doi"], m["title"])
                        if not k:
                            unusable += 1
                            continue
                        if k in seen_here:
                            within += 1          # the same paper twice in one file
                            continue
                        seen_here.add(k)
                        if k in held or k in own:
                            again += 1           # this file was imported before
                            continue
                        m["_key"] = k
                        new.append(m)

                    # How much of this file the corpus already holds from
                    # ANOTHER source. Reported so the preview is honest about
                    # what the import will add, but these rows are still
                    # stored and still counted as identified.
                    other_doi = {r["ndoi"] for r in STATE["records"]
                                 if r["ndoi"] and r.get("source") != src}
                    other_t = {r["ntitle"] for r in STATE["records"]
                               if r["ntitle"] and r.get("source") != src}
                    overlap = sum(1 for m in new
                                  if (norm_doi(m["doi"]) in other_doi
                                      if norm_doi(m["doi"]) else
                                      norm_title(m["title"]) in other_t))

                    out = {"file": p.name, "source": src,
                           "label": SOURCE_LABELS.get(src, src),
                           "rows": len(rows), "usable": len(mapped),
                           "new": len(new), "already_imported": again,
                           "overlap": overlap, "unusable": unusable,
                           "dupes_within_file": within,
                           "columns": {k: v for k, v in picked.items() if v},
                           "unmapped": [k for k, v in picked.items() if not v],
                           "sample": [{"title": m["title"][:90], "year": m["year"],
                                       "journal": m["journal"][:46]}
                                      for m in new[:8]]}
                    if u.path.endswith("preview"):
                        out["preview"] = True
                        return self._json(out)

                    now = datetime.datetime.now().isoformat(timespec="seconds")
                    extra = STATE.get("extra") or []
                    for m in new:
                        extra.append({
                            "id": m["_key"], "source": src,
                            "source_file": p.name, "added_at": now,
                            "doi": m["doi"], "title": m["title"],
                            "authors": m["authors"], "year": m["year"],
                            "journal": m["journal"], "abstract": m["abstract"],
                            "keywords": m["keywords"], "language": m["language"],
                            "type": m["type"], "issn": m["issn"],
                            "cited_by": m["cited_by"], "note": m.get("note", ""),
                            "key_paper": m.get("key_paper", ""),
                            "country": m.get("country", "")})
                    STATE["extra"] = extra
                    save_extra(STATE["paths"]["extra"], extra)
                    # Kept as an audit trail of what each file contained. No
                    # count in the diagram is read from it any more: PRISMA is
                    # computed from the stored records, so there is one source
                    # of truth and re-importing a file cannot inflate it.
                    STATE.setdefault("imports", []).append({
                        "source": src, "file": p.name, "at": now,
                        "rows": len(rows), "usable": len(mapped),
                        "new": len(new), "already_imported": again,
                        "overlap": overlap, "unusable": unusable,
                        "dupes_within_file": within})
                    save_json(STATE["paths"]["imports"], STATE["imports"])
                    nb, nx = rebuild_records()
                    out["applied"] = True
                    out["corpus_now"] = nb + nx
                    return self._json(out)

                if u.path == "/api/lookup-doi":
                    try:
                        return self._json({"ok": True,
                                           "record": crossref_lookup(data.get("doi"))})
                    except RuntimeError as e:
                        return self._json({"error": str(e)}, 400)

                if u.path == "/api/add-paper":
                    title = clean(data.get("title"))
                    doi = strip_doi(data.get("doi"))
                    if not title and not doi:
                        return self._json({"error": "A title or a DOI is "
                                                    "needed."}, 400)
                    nd_, nt_ = norm_doi(doi), norm_title(title)
                    for r in STATE["records"]:
                        if (nd_ and r["ndoi"] == nd_) or (nt_ and r["ntitle"] == nt_):
                            return self._json({"error":
                                f"That paper is already in the corpus as "
                                f"{r['id']}: {r['title'][:70]}"}, 400)
                    extra = STATE.get("extra") or []
                    src = data.get("source") or "citation"
                    rec = {"id": record_key(src, doi, title)
                                 or f"MAN{len(extra) + 1}", "source": src,
                           "source_file": "",
                           "key_paper": clean(data.get("key_paper")).lower(),
                           "country": clean(data.get("country")),
                           "added_at": datetime.datetime.now().isoformat(
                               timespec="seconds"),
                           "doi": doi, "title": title,
                           "authors": clean(data.get("authors")),
                           "year": clean(data.get("year")),
                           "journal": clean(data.get("journal")),
                           "abstract": clean(data.get("abstract")),
                           "keywords": clean(data.get("keywords")),
                           "language": clean(data.get("language")),
                           "type": clean(data.get("type")),
                           "issn": clean(data.get("issn")),
                           "cited_by": as_int(data.get("cited_by")),
                           "note": clean(data.get("note"))}
                    extra.append(rec)
                    STATE["extra"] = extra
                    save_extra(STATE["paths"]["extra"], extra)
                    rebuild_records()
                    return self._json({"ok": True, "id": rec["id"],
                                       "title": rec["title"]})

                if u.path == "/api/remove-extra":
                    rid = data.get("id")
                    extra = [e for e in (STATE.get("extra") or [])
                             if e.get("id") != rid]
                    STATE["extra"] = extra
                    save_extra(STATE["paths"]["extra"], extra)
                    rebuild_records()
                    return self._json({"ok": True})

                if u.path == "/api/stop-config":
                    for k in ("consecutive", "min_percent", "estimate"):
                        if k in data:
                            v = str(data[k]).strip()
                            STATE["stop_cfg"][k] = v if v else ""
                    save_json(STATE["paths"]["stop"],
                              {"cfg": STATE["stop_cfg"], "stop": STATE["stop"]})
                    return self._json(stopping_status())

                if u.path == "/api/stop-record":
                    st = stopping_status()
                    STATE["stop"] = {
                        "recorded": True,
                        "at": datetime.datetime.now().isoformat(timespec="seconds"),
                        "screener": STATE["screener"],
                        "screened": st["screened"], "relevant": st["relevant"],
                        "total": st["total"], "unscreened": st["unscreened"],
                        "percent": st["percent"],
                        "consecutive": st["consecutive"],
                        "threshold": st["threshold"],
                        "min_percent": st["min_percent"],
                        "estimate": st["estimate"],
                        "estimate_from": st["estimate_from"],
                        "keys": st["keys"],
                        "conditions": [{"label": c["label"], "met": c["met"],
                                        "detail": c["detail"],
                                        "skipped": c["skipped"]}
                                       for c in st["conditions"]],
                        "all_met": st["all_met"],
                        "note": (data.get("note") or "").strip(),
                    }
                    save_json(STATE["paths"]["stop"],
                              {"cfg": STATE["stop_cfg"], "stop": STATE["stop"]})
                    return self._json(stopping_status())

                if u.path == "/api/stop-undo":
                    STATE["stop"] = {}
                    save_json(STATE["paths"]["stop"],
                              {"cfg": STATE["stop_cfg"], "stop": STATE["stop"]})
                    return self._json(stopping_status())

                if u.path == "/api/screener":
                    # Remembered between sessions. Every decision records who
                    # made it, which is what lets a second screener's work be
                    # told apart from yours.
                    STATE["screener"] = (data.get("name") or "").strip()
                    cfg = STATE.get("config") or {}
                    cfg["screener"] = STATE["screener"]
                    if "mailto" in data:
                        STATE["mailto"] = clean(data.get("mailto"))
                        cfg["mailto"] = STATE["mailto"]
                    STATE["config"] = cfg
                    try:
                        save_config(cfg)
                    except OSError as e:
                        return self._json({"error": f"could not save: {e}"}, 400)
                    return self._json({"ok": True, "screener": STATE["screener"],
                                       "mailto": STATE.get("mailto", "")})

                if u.path == "/api/ai-key":
                    # Held in memory for this session only. Never written to
                    # disk: this folder syncs to OneDrive.
                    if "key" in data:
                        STATE["api_key"] = (data.get("key") or "").strip()
                    if data.get("model"):
                        STATE["model"] = data["model"].strip()
                    if data.get("k"):
                        STATE["ask_k"] = max(1, min(200, int(data["k"])))
                    if data.get("chars"):
                        STATE["ask_chars"] = max(200, min(20000,
                                                          int(data["chars"])))
                    return self._json({"ok": True,
                                       "ready": bool(STATE["api_key"]),
                                       "model": STATE["model"],
                                       "k": STATE["ask_k"],
                                       "chars": STATE["ask_chars"]})

            return self._send(404, "not found", "text/plain; charset=utf-8")
        except Exception as e:
            return self._json({"error": f"{type(e).__name__}: {e}"}, 500)

    def _patch(self, rid, patch):
        d = STATE["decisions"].setdefault(rid, {k: "" for k in DECISION_FIELDS})
        for k, v in patch.items():
            if k in DECISION_FIELDS:
                d[k] = "" if v is None else str(v)
        rec = next((x for x in STATE["records"] if x["id"] == rid), None)
        if rec:
            d["doi"], d["title"] = rec["doi"], rec["title"]
            touched = [c for c in MANUAL_OVERRIDES.values() if c in patch]
            if touched:
                # Fold the correction in now, so the rules, the ranking and
                # the dashboard all see it on this very request.
                apply_manual_fields()
                # A corrected DOI or title changes what deduplication matches
                # on, so the groups have to be found again.
                if "doi_manual" in patch or "title_manual" in patch:
                    STATE["dupes"] = find_duplicates(STATE["records"])
                # Always after apply_manual_fields: a field you just cleared
                # may now be fillable from a duplicate, and one you just typed
                # must not be overwritten by one.
                merge_duplicates()
        d["screener"] = STATE["screener"]
        # Milliseconds, not seconds: screening with the keyboard puts several
        # decisions in the same second, and the stopping rule reads this
        # column as the order you screened in. At second resolution the sort
        # falls back to record id — alphabetical, not chronological — and the
        # run of consecutive irrelevant records comes out wrong.
        d["updated_at"] = datetime.datetime.now().isoformat(
            timespec="milliseconds")


# ==========================================================================
#  main
# ==========================================================================

def main():
    ap = argparse.ArgumentParser(
        description="Step 3 — the review workbench.")
    ap.add_argument("input", nargs="?", default="",
                    help="A corpus file. Omit it and today's newest is used.")
    ap.add_argument("--port", type=int, default=PORT)
    ap.add_argument("--no-browser", action="store_true")
    ap.add_argument("--merge-decisions", action="store_true",
                    help="Bring in judgements made against another corpus "
                         "file in this folder, for records not screened here.")
    args = ap.parse_args()

    # Absolute, so the Files panel can show you the folder these live in.
    src = find_corpus(args.input or None).resolve()
    stem = str(src.with_suffix(""))
    STATE["paths"] = {"corpus": str(src),
                      "decisions": stem + "_decisions.csv",
                      "labels": stem + "_labels.json",
                      "criteria": stem + "_criteria.json",
                      "rules": stem + "_rules.json",
                      "stop": stem + "_stop.json",
                      "extra": stem + "_extra_records.csv",
                      "imports": stem + "_import_log.json"}
    STATE["source"] = src.name
    STATE["records"] = load_corpus(src)
    STATE["extra"] = load_extra(STATE["paths"]["extra"])
    STATE["imports"] = load_json(STATE["paths"]["imports"], []) or []
    rebuild_records()
    # Your review's own settings. The config file beside the app supplies the
    # defaults; anything already saved for THIS corpus wins over it, so a
    # second review in the same folder starts from the config rather than
    # inheriting the first one's themes.
    cfg = load_config()
    STATE["config"] = cfg
    STATE["screener"] = cfg.get("screener") or ""
    STATE["mailto"] = cfg.get("mailto") or ""
    cp, tp = config_patterns(cfg)
    STATE["country_patterns"] = cp
    STATE["theme_patterns"] = tp
    STATE["labels"] = load_json(STATE["paths"]["labels"],
                                cfg.get("themes") or DEFAULT_LABELS)
    STATE["criteria"] = load_json(STATE["paths"]["criteria"],
                                  cfg.get("criteria") or DEFAULT_CRITERIA)

    saved = load_json(STATE["paths"]["rules"], None)
    rules = [dict(r) for r in (cfg.get("rules") or DEFAULT_RULES)]
    if saved:
        by_id = {r["id"]: r for r in rules}
        for s in saved:
            t = by_id.get(s.get("id"))
            if t:
                t["enabled"] = bool(s.get("enabled"))
                if isinstance(s.get("params"), dict):
                    t["params"].update(s["params"])
    STATE["rules"] = rules

    saved_stop = load_json(STATE["paths"]["stop"], {}) or {}
    STATE["stop_cfg"] = saved_stop.get("cfg") or {}
    STATE["stop"] = saved_stop.get("stop") or {}

    STATE["decisions"] = load_decisions(STATE["paths"]["decisions"])
    # Judgements from another corpus file are only ever merged in when this
    # corpus has none of its own, or when you ask for it with
    # --merge-decisions. Doing it silently would be worse than not doing it:
    # a decisions file can hold the result of a bulk operation rather than
    # screening you did record by record, and quietly spreading that across
    # every corpus in the folder would be unrecoverable. So when there is
    # something to merge, the app says so and leaves the choice to you.
    brought = carry_over(STATE["records"], STATE["paths"]["decisions"],
                         STATE["decisions"], quiet=bool(STATE["decisions"])
                         and not args.merge_decisions)
    if brought and (not STATE["decisions"] or args.merge_decisions):
        STATE["decisions"].update(brought)
        save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
        print(f"  merged in {len(brought):,} decisions from other corpus files")
        brought = {}
    moved, orphans = reattach_decisions()
    if moved:
        print(f"  re-attached {moved:,} decisions whose record id had changed")
        save_decisions(STATE["paths"]["decisions"], STATE["decisions"])
    typed = apply_manual_fields()
    STATE["dupes"] = find_duplicates(STATE["records"])
    borrowed = merge_duplicates()

    print(f"\nLoaded {len(STATE['records']):,} records from {src.name}")
    missing = sum(1 for r in STATE["records"] if not r["abstract"])
    print(f"  abstracts: {len(STATE['records']) - missing:,} present, "
          f"{missing:,} missing"
          + (f" ({typed:,} you typed in yourself)" if typed else ""))
    dups, _ = dup_state()
    n_dup = sum(1 for v in dups.values() if v == "duplicate")
    print(f"  duplicate groups: {len(STATE['dupes'])}  "
          f"({n_dup} records auto-marked as duplicates)")
    if borrowed:
        nrec = sum(1 for r in STATE["records"] if r.get("merged_from"))
        print(f"  filled {borrowed:,} empty fields on {nrec:,} kept records "
              f"from the duplicates they stand for")

    prev = rule_preview()
    print("\n  Rules (each one overridable on any record):")
    for r in STATE["rules"]:
        mark = "ON " if r["enabled"] else "off"
        print(f"    [{mark}] {r['name']:34s} would remove {prev.get(r['id'], 0):>5,}")

    guessed_c = sum(1 for r in STATE["records"] if r["country_guess"])
    guessed_t = sum(1 for r in STATE["records"] if r["theme_guess"])
    print(f"\n  Guessed a country for {guessed_c:,} records, "
          f"a theme for {guessed_t:,}. Both editable.")
    done = sum(1 for d in STATE["decisions"].values() if d.get("ta_decision"))
    print(f"  Decisions on file: {len(STATE['decisions']):,} ({done:,} screened)")
    if orphans:
        print(f"  {orphans:,} decision(s) point at a paper no longer in the "
              f"corpus. They are kept, not counted, and will re-attach if "
              f"that paper comes back on an import.")

    class Server(socketserver.ThreadingTCPServer):
        allow_reuse_address = True
        daemon_threads = True

    srv = port = None
    for p in range(args.port, args.port + 10):
        try:
            srv = Server(("127.0.0.1", p), Handler)
            port = p
            break
        except OSError:
            continue
    if srv is None:
        raise SystemExit(
            f"\nPorts {args.port}-{args.port + 9} are all busy.\n"
            f"  The workbench may already be open — try "
            f"http://127.0.0.1:{args.port}")
    if port != args.port:
        print(f"\n  (port {args.port} was busy — using {port})")
    print(f"\n  Open http://127.0.0.1:{port}      Ctrl+C here to stop\n")
    with srv:
        if not args.no_browser:
            threading.Timer(
                0.8, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
        try:
            srv.serve_forever()
        except KeyboardInterrupt:
            print("\nStopped. Everything is saved.")


if __name__ == "__main__":
    main()