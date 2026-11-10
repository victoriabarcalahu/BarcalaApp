#!/usr/bin/env python3
"""
STEP 2 of 3 — fill in every field OpenAlex left blank, not just the abstract.

  python 02-complete.py

Reads the file step 1 wrote and writes the same name with "_complete" added:

    01-openalex_raw.csv  ->  01-openalex_raw_complete.csv

What it completes
-----------------
Every column step 1 produces, in four passes, because different fields live
in different places:

  doi                                      pass 1 (Crossref title search)
  topic, subfield, field, institutions,
    author_countries                       pass 2 (OpenAlex by DOI)
  abstract, title, authors, year,
    journal, type, language,
    cited_by_count                         passes 2, 3 and 4
  europe_affiliation, europe_text, europe  recomputed at the end

Step 1 always writes a citation count, so in normal use cited_by_count is
OpenAlex's figure throughout and this step leaves it alone. It is only ever
filled here if you feed in a file where that cell is blank, and then the
source is recorded in enriched_fields — worth checking, because Semantic
Scholar counts over a different corpus and its number will not match
OpenAlex's for the same paper.

Why four passes
---------------
  PASS 1 — Crossref title search, for records with NO DOI.
     Runs first, because a recovered DOI is the key every later pass needs.
     A candidate is accepted only when the titles match 92% or better, so a
     near-miss is left alone rather than quietly attached to the wrong
     paper.

  PASS 2 — OpenAlex by DOI, 25 at a time.
     This is the only source for topic, subfield and field: they are
     OpenAlex's own taxonomy and no other service publishes them. It also
     returns institutions with proper ISO country codes. Most useful for
     records that only just acquired a DOI in pass 1 — those were invisible
     to OpenAlex before.

  PASS 3 — Semantic Scholar batch, 400 DOIs per request.
     Best coverage for the abstracts commercial publishers withhold.
     Elsevier does not deposit abstracts where OpenAlex can index them, so
     Journal of Cleaner Production, Journal of Business Research and
     Technological Forecasting all arrive blank. Measured on a 1,793-record
     harvest: Finance Research Letters 1 of 13 had an abstract, Electronic
     Commerce Research 0 of 8.

  PASS 4 — Crossref by DOI, for whatever is still blank.
     Also carries author affiliations, which fills institutions where
     OpenAlex has none.

One honest limit
----------------
topic, subfield and field are filled ONLY from OpenAlex. Semantic Scholar's
fieldsOfStudy and Crossref's subject headings are different taxonomies, and
mixing them into the same column would quietly corrupt any count you later
make per field. If OpenAlex has no topic for a work, that cell stays empty
and the summary says how many. An empty cell you know about is worth more
than a filled one you cannot trust.

Country codes recovered from Crossref affiliation strings are matched by
country name and marked `derived` in enriched_fields, because they are
inferred rather than asserted by the source.

Nothing already present is ever overwritten; only blank fields are filled.
`abstract_source` and `enriched_fields` record what came from where, which
belongs in your methods section: an abstract fetched from Crossref has a
different provenance from one the search returned.

Progress is checkpointed. If it is interrupted, run it again and it picks up
where it stopped. No service here needs an API key.

Standard library only.
"""

# ===========================================================================
#  CONFIGURATION
# ===========================================================================

# Leave blank to use whatever step 1 wrote (it reads OUTPUT_FILE out of
# 01-harvest.py), or name a file here to override that.
INPUT_FILE = ""

SUFFIX = "_complete"

# The delimiter of the file being READ is detected automatically, so this only
# controls what gets written. Leave it blank to reuse whatever the input used,
# which keeps the whole pipeline consistent. Otherwise ";", "," or "\t".
CSV_DELIMITER = ""
WRITE_XLSX = True

MAILTO = "victoria.barcala@hu.nl"
PACE = 0.3                 # seconds between requests
TITLE_MATCH = 0.92         # how close a title must be to accept a match

# ---------------------------------------------------------------------------
# ELSEVIER — optional, and the only way to fill the biggest remaining gap.
#
# Elsevier does not deposit abstracts to Crossref, so OpenAlex never has them.
# Measured on your own corpus: Electronic Commerce Research and Applications
# 0 of 8, Journal of Behavioral and Experimental Finance 0 of 6, Finance
# Research Letters 1 of 13. Springer Nature deposits abstracts only for its
# open-access content, which is why Financial Innovation (OA) has all of them
# and Journal of Business Ethics (subscription) has one in four.
#
# Elsevier's own Abstract Retrieval API can return them. Two steps:
#   1. Register for a free key at https://dev.elsevier.com and paste it below.
#   2. Run this ON THE HU NETWORK, or ask the HU library for an institutional
#      token (X-ELS-Insttoken) and paste that too. The key removes the
#      technical barrier, not the licensing one: off-campus and without a
#      token, Elsevier returns little or nothing.
#
# Leave both blank and this pass is skipped silently. Nothing else changes.
# https://dev.elsevier.com/apikey/manage---------------------------------------------------------------------------
ELSEVIER_API_KEY = "9e706209759782a94b3c8626e4dce0d8"
ELSEVIER_INSTTOKEN = ""

# Turn individual passes off if you want to save time on a rerun.
RUN_TITLE_SEARCH = True    # pass 1
RUN_OPENALEX = True        # pass 2  <- the only source of topic/subfield/field
RUN_SEMANTIC_SCHOLAR = True  # pass 3
RUN_CROSSREF = True        # pass 4
RUN_ELSEVIER = True        # pass 5, only does anything with a key set above

# ===========================================================================

import argparse
import csv
import difflib
import html
import json
import os
import random
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
OPENALEX = "https://api.openalex.org/works"
S2_BATCH = "https://api.semanticscholar.org/graph/v1/paper/batch"
CROSSREF_DOI = "https://api.crossref.org/works/"
CROSSREF_Q = "https://api.crossref.org/works"
S2_BATCH_SIZE = 400
OA_BATCH_SIZE = 25
OA_SELECT = ("id,doi,title,display_name,publication_year,type,language,"
             "authorships,primary_location,primary_topic,cited_by_count,"
             "abstract_inverted_index")

# Fields any source can supply. cited_by_count comes from OpenAlex or
# Semantic Scholar; Crossref does not publish one. A count of 0 is a real
# value, so "0" is never treated as a blank to be filled.
BIBLIO = ["title", "authors", "year", "journal", "type", "language",
          "abstract", "cited_by_count"]
# Fields only OpenAlex publishes.
OPENALEX_ONLY = ["topic", "subfield", "field"]
# Affiliation fields; OpenAlex gives codes, Crossref gives names only.
AFFIL = ["institutions", "author_countries"]
COMPLETABLE = BIBLIO + OPENALEX_ONLY + AFFIL
DERIVED = ["europe_affiliation", "europe_text", "europe"]

CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
WS = re.compile(r"\s+")

EU_SET = {"NL", "DE", "FR", "IT", "ES", "BE", "AT", "SE", "NO", "DK", "FI",
          "PL", "PT", "IE", "GB", "CZ", "HU", "RO", "GR", "EE", "LV", "LT",
          "SI", "HR", "LU", "SK", "BG", "CH", "IS", "CY", "MT"}

# For reading a country out of a Crossref affiliation string, which has no
# country code of its own. Longest names first so "United Kingdom" is tried
# before any substring of it could match.
COUNTRY_NAMES = [
    ("united kingdom", "GB"), ("great britain", "GB"), ("netherlands", "NL"),
    ("switzerland", "CH"), ("luxembourg", "LU"), ("lithuania", "LT"),
    ("czech republic", "CZ"), ("czechia", "CZ"), ("slovakia", "SK"),
    ("slovenia", "SI"), ("portugal", "PT"), ("romania", "RO"),
    ("bulgaria", "BG"), ("denmark", "DK"), ("germany", "DE"),
    ("hungary", "HU"), ("belgium", "BE"), ("estonia", "EE"),
    ("iceland", "IS"), ("ireland", "IE"), ("finland", "FI"),
    ("austria", "AT"), ("croatia", "HR"), ("england", "GB"),
    ("scotland", "GB"), ("latvia", "LV"), ("norway", "NO"),
    ("poland", "PL"), ("sweden", "SE"), ("cyprus", "CY"), ("france", "FR"),
    ("greece", "GR"), ("italy", "IT"), ("spain", "ES"), ("malta", "MT"),
    ("wales", "GB"),
    # a few non-European ones, so a country is not wrongly read as European
    ("united states", "US"), ("usa", "US"), ("china", "CN"), ("india", "IN"),
    ("indonesia", "ID"), ("australia", "AU"), ("canada", "CA"),
    ("brazil", "BR"), ("japan", "JP"), ("singapore", "SG"),
    ("south korea", "KR"), ("taiwan", "TW"), ("malaysia", "MY"),
    ("turkey", "TR"), ("israel", "IL"), ("south africa", "ZA"),
]

EUROPE_TEXT = re.compile(
    r"netherlands|dutch|german|\bfrance\b|french|ital(y|ian)|spain|spanish|"
    r"portug(al|uese)|belgi(um|an)|austria|switzerland|swiss|swed(en|ish)|"
    r"norw(ay|egian)|denmark|danish|finland|finnish|poland|polish|czech|"
    r"hungar(y|ian)|greece|greek|ireland|irish|estonia|latvia|lithuania|"
    r"slovenia|croatia|luxembourg|slovakia|bulgaria|iceland|cyprus|malta|"
    r"romania|united kingdom|\buk\b|britain|british|england|scotland|wales|"
    r"europe|european|\beu\b|eurozone", re.I)


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return WS.sub(" ", CTRL.sub(" ", s)).strip()


def strip_jats(x):
    """Crossref returns abstracts as JATS XML."""
    if not x:
        return ""
    s = re.sub(r"<jats:title>.*?</jats:title>", " ", x, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    return clean(re.sub(r"^\s*abstract\s*[:.\-]?\s*", "", s, flags=re.I))


def norm_doi(s):
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", str(s or "").strip(),
                  flags=re.I).strip().lower()


def norm_title(s):
    return WS.sub(" ", re.sub(r"[^a-z0-9 ]+", " ", str(s or "").lower())).strip()


def reconstruct_abstract(inv):
    if not inv:
        return ""
    pos = [(i, w) for w, idxs in inv.items() for i in idxs]
    pos.sort(key=lambda p: p[0])
    return clean(" ".join(w for _, w in pos))


def countries_from_affiliations(text):
    """Read ISO codes out of affiliation strings. Inferred, not asserted."""
    low = (text or "").lower()
    out = []
    for name, code in COUNTRY_NAMES:
        if name in low and code not in out:
            out.append(code)
    return out


def sniff_delimiter(path):
    """Work out which delimiter a CSV uses, by counting candidates in the
    header line. The header has no quoted commas in it, so counting there is
    reliable in a way that counting over the whole file would not be."""
    try:
        with open(path, encoding="utf-8-sig") as fh:
            header = fh.readline()
    except Exception:
        return ","
    if not header:
        return ","
    counts = {d: header.count(d) for d in (";", "\t", ",")}
    best = max(counts, key=lambda d: counts[d])
    return best if counts[best] else ","


XLSX_WIDTH = {
    "openalex_id": 14, "doi": 26, "title": 60, "authors": 38, "year": 7,
    "journal": 32, "type": 12, "abstract": 80, "language": 9,
    "institutions": 38, "topic": 28, "subfield": 20, "field": 18,
    "cited_by_count": 11, "retrieved": 12,
    "author_countries": 16, "europe_affiliation": 12, "europe_text": 11,
    "europe": 8, "abstract_source": 16, "enriched_fields": 44,
}


def write_table(path, columns, rows, delimiter=",", want_xlsx=True):
    """Write the CSV fully quoted, and an .xlsx beside it when openpyxl is
    installed. Full quoting is what makes a comma or semicolon inside an
    abstract harmless."""
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=columns, delimiter=delimiter,
                           quoting=csv.QUOTE_ALL, lineterminator="\r\n",
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    if not want_xlsx:
        return None
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None
    xlsx = Path(path).with_suffix(".xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "corpus"
    ws.append(list(columns))
    for r in rows:
        ws.append([str(r.get(c, "") or "")[:32767] for c in columns])
    head, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="44546A")
    for i in range(1, len(columns) + 1):
        c = ws.cell(row=1, column=i)
        c.font, c.fill = head, fill
        c.alignment = Alignment(vertical="center")
    for i, name in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = XLSX_WIDTH.get(name, 18)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(rows) + 1}"
    ws.row_dimensions[1].height = 24
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=False)
    try:
        wb.save(xlsx)
    except PermissionError:
        print(f"\n  Could not write {xlsx.name} — it is probably open in Excel. "
              "Close it and run again.", file=sys.stderr)
        return None
    return xlsx


def find_input():
    if INPUT_FILE:
        p = Path(INPUT_FILE)
        p = (SCRIPT_DIR / p) if not p.is_absolute() else p
        if not p.exists():
            raise SystemExit(f"\nINPUT_FILE is set to {p} but that file is "
                             "not there.")
        return p
    step1 = SCRIPT_DIR / "01-harvest.py"
    if step1.exists():
        try:
            m = re.search(r'^OUTPUT_FILE\s*=\s*["\']([^"\']+)["\']',
                          step1.read_text(encoding="utf-8"), re.M)
            if m:
                p = SCRIPT_DIR / m.group(1)
                if p.exists():
                    print(f"Using the file 01-harvest.py writes: {p.name}\n")
                    return p
                print(f"01-harvest.py says its output is {m.group(1)}, but that "
                      "file is not here yet.\nRun step 1 first, or set "
                      "INPUT_FILE at the top of this script.\n", file=sys.stderr)
        except Exception:
            pass
    cands = []
    for p in SCRIPT_DIR.glob("*.csv"):
        if p.name.startswith("~$") or p.stem.endswith(SUFFIX):
            continue
        try:
            with open(p, newline="", encoding="utf-8-sig") as fh:
                h = {c.strip().lower() for c in
                     next(csv.reader(fh, delimiter=sniff_delimiter(p)), [])}
        except Exception:
            continue
        if {"item type", "date added", "abstract note"} & h:
            continue                      # a Zotero export, not a harvest
        if "title" in h and {"abstract", "doi", "openalex_id"} & h:
            cands.append(p)
    if not cands:
        raise SystemExit(
            f"\nNo harvest file found in:\n  {SCRIPT_DIR}\n\n"
            "  Run step 1 first:  python 01-harvest.py\n"
            "  Or set INPUT_FILE at the top of this script.")
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    print(f"Using the newest harvest found: {cands[0].name}\n")
    return cands[0]


def request(url, data=None, retries=5, timeout=90):
    hdrs = {"User-Agent": f"slr-complete/2.0 (mailto:{MAILTO})"}
    body = json.dumps(data).encode() if data is not None else None
    if body is not None:
        hdrs["Content-Type"] = "application/json"
    last = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=body, headers=hdrs,
                                         method="POST" if body else "GET")
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (400, 401, 403, 404, 414):
                return None          # not held there, or the request was refused
            wait = max(2 ** attempt, 5) + random.uniform(0, 1.5)
            last = e
            print(f"    rate-limited ({e.code}); waiting {wait:.0f}s "
                  f"(attempt {attempt + 1}/{retries})", file=sys.stderr)
            time.sleep(wait)
        except Exception as e:
            last = e
            wait = 2 ** attempt
            print(f"    connection problem; retry in {wait}s ({e})", file=sys.stderr)
            time.sleep(wait)
    print(f"    gave up: {last}", file=sys.stderr)
    return None


# --- pass 1: Crossref title search, for records with no DOI ----------------

def parse_crossref(msg):
    if not msg:
        return {}
    rec = {"_from": "crossref"}
    t = msg.get("title") or []
    if t:
        rec["title"] = clean(t[0])
    au, affs = [], []
    for a in (msg.get("author") or []):
        nm = " ".join(x for x in [a.get("given"), a.get("family")] if x) \
             or a.get("name") or ""
        if nm:
            au.append(clean(nm))
        for af in (a.get("affiliation") or []):
            if af.get("name"):
                affs.append(clean(af["name"]))
    if au:
        rec["authors"] = " | ".join(au)
    if affs:
        rec["institutions"] = " | ".join(dict.fromkeys(affs))
        codes = countries_from_affiliations(" ".join(affs))
        if codes:
            rec["author_countries"] = " | ".join(codes)
            rec["_derived_countries"] = True
    dp = ((msg.get("issued") or {}).get("date-parts") or [[None]])[0]
    if dp and dp[0]:
        rec["year"] = str(dp[0])
    ct = msg.get("container-title") or []
    if ct:
        rec["journal"] = clean(ct[0])
    if msg.get("type"):
        rec["type"] = clean(msg["type"])
    if msg.get("language"):
        rec["language"] = clean(msg["language"])[:2].lower()
    ab = strip_jats(msg.get("abstract"))
    if ab:
        rec["abstract"] = ab
    if msg.get("DOI"):
        rec["doi"] = clean(msg["DOI"])
    return rec


def pass_title_search(items):
    out, rejected = {}, 0
    for n, it in enumerate(items, 1):
        q = {"query.bibliographic": it["title"], "rows": "5", "mailto": MAILTO}
        if str(it.get("year", "")).isdigit():
            y = int(it["year"])
            q["filter"] = (f"from-pub-date:{y - 1}-01-01,"
                           f"until-pub-date:{y + 1}-12-31")
        res = request(f"{CROSSREF_Q}?{urllib.parse.urlencode(q)}")
        best, score = None, 0.0
        for cand in ((res or {}).get("message", {}).get("items") or []):
            ct = (cand.get("title") or [""])[0]
            s = difflib.SequenceMatcher(None, norm_title(it["title"]),
                                        norm_title(ct)).ratio()
            if s > score:
                best, score = cand, s
        if best and score >= TITLE_MATCH:
            rec = parse_crossref(best)
            rec["_match"] = round(score, 3)
            rec["_from"] = "crossref-title"
            out[it["key"]] = rec
        elif best:
            rejected += 1
        if n % 20 == 0 or n == len(items):
            print(f"  Crossref by title: {n}/{len(items)}, {len(out)} matched, "
                  f"{rejected} below the {TITLE_MATCH:.0%} threshold")
        time.sleep(PACE)
    return out


# --- pass 2: OpenAlex by DOI — the only source of topic/subfield/field -----

def parse_openalex(w):
    rec = {"_from": "openalex"}
    names, insts, codes = [], [], []
    for a in (w.get("authorships") or []):
        au = a.get("author") or {}
        if au.get("display_name"):
            names.append(clean(au["display_name"]))
        codes += [c.upper() for c in (a.get("countries") or []) if c]
        for i in (a.get("institutions") or []):
            if i.get("display_name"):
                insts.append(clean(i["display_name"]))
            if i.get("country_code"):
                codes.append(i["country_code"].upper())
    if names:
        rec["authors"] = " | ".join(names)
    if insts:
        rec["institutions"] = " | ".join(dict.fromkeys(insts))
    if codes:
        rec["author_countries"] = " | ".join(dict.fromkeys(codes))
    src = (w.get("primary_location") or {}).get("source") or {}
    if src.get("display_name"):
        rec["journal"] = clean(src["display_name"])
    pt = w.get("primary_topic") or {}
    if pt.get("display_name"):
        rec["topic"] = clean(pt["display_name"])
    if (pt.get("subfield") or {}).get("display_name"):
        rec["subfield"] = clean(pt["subfield"]["display_name"])
    if (pt.get("field") or {}).get("display_name"):
        rec["field"] = clean(pt["field"]["display_name"])
    t = clean(w.get("title") or w.get("display_name"))
    if t:
        rec["title"] = t
    if w.get("publication_year"):
        rec["year"] = str(w["publication_year"])
    if w.get("type"):
        rec["type"] = clean(w["type"])
    if w.get("language"):
        rec["language"] = clean(w["language"])
    if w.get("cited_by_count") is not None:
        rec["cited_by_count"] = str(int(w["cited_by_count"] or 0))
    ab = reconstruct_abstract(w.get("abstract_inverted_index"))
    if ab:
        rec["abstract"] = ab
    return rec


def pass_openalex(dois):
    """Batched DOI lookup. Skips DOIs holding a comma or pipe, which are the
    two characters OpenAlex uses to separate filter values."""
    out = {}
    usable = [d for d in dois if "," not in d and "|" not in d]
    skipped = len(dois) - len(usable)
    if skipped:
        print(f"  {skipped} DOI(s) contain a comma or pipe and cannot go in an "
              "OpenAlex filter; they are left to the other passes")
    for i in range(0, len(usable), OA_BATCH_SIZE):
        chunk = usable[i:i + OA_BATCH_SIZE]
        f = "doi:" + "|".join(chunk)
        url = (f"{OPENALEX}?filter={urllib.parse.quote(f)}"
               f"&select={OA_SELECT}&per-page={OA_BATCH_SIZE}"
               f"&mailto={urllib.parse.quote(MAILTO)}")
        res = request(url)
        for w in ((res or {}).get("results") or []):
            d = norm_doi(w.get("doi"))
            if d:
                out[d] = parse_openalex(w)
        print(f"  OpenAlex by DOI: {min(i + OA_BATCH_SIZE, len(usable))}"
              f"/{len(usable)}, {len(out)} returned data")
        if i + OA_BATCH_SIZE < len(usable):
            time.sleep(PACE)
    return out


# --- pass 3: Semantic Scholar ----------------------------------------------

def pass_semantic_scholar(dois):
    fields = ("abstract,title,authors.name,authors.affiliations,year,venue,"
              "journal,publicationTypes,citationCount")
    out = {}
    for i in range(0, len(dois), S2_BATCH_SIZE):
        chunk = dois[i:i + S2_BATCH_SIZE]
        print(f"  Semantic Scholar: batch {i // S2_BATCH_SIZE + 1}, "
              f"{len(chunk)} DOIs")
        res = request(f"{S2_BATCH}?fields={fields}",
                      data={"ids": [f"DOI:{d}" for d in chunk]})
        if not res:
            continue
        for d, it in zip(chunk, res):
            if not it:
                continue
            rec = {"_from": "semanticscholar"}
            if it.get("abstract"):
                rec["abstract"] = clean(it["abstract"])
            if it.get("title"):
                rec["title"] = clean(it["title"])
            if it.get("year"):
                rec["year"] = str(it["year"])
            if it.get("citationCount") is not None:
                rec["cited_by_count"] = str(int(it["citationCount"] or 0))
            names, affs = [], []
            for a in (it.get("authors") or []):
                if a.get("name"):
                    names.append(clean(a["name"]))
                for af in (a.get("affiliations") or []):
                    if af:
                        affs.append(clean(af))
            if names:
                rec["authors"] = " | ".join(names)
            if affs:
                rec["institutions"] = " | ".join(dict.fromkeys(affs))
                codes = countries_from_affiliations(" ".join(affs))
                if codes:
                    rec["author_countries"] = " | ".join(codes)
                    rec["_derived_countries"] = True
            jr = it.get("journal") or {}
            if jr.get("name"):
                rec["journal"] = clean(jr["name"])
            elif it.get("venue"):
                rec["journal"] = clean(it["venue"])
            if len(rec) > 1:
                out[d] = rec
        print(f"    {sum(1 for d in chunk if d in out)} returned data")
        if i + S2_BATCH_SIZE < len(dois):
            time.sleep(1.0)
    return out


# --- pass 4: Crossref by DOI ------------------------------------------------

def pass_crossref_doi(dois):
    out = {}
    for n, d in enumerate(dois, 1):
        res = request(f"{CROSSREF_DOI}{urllib.parse.quote(d)}"
                      f"?mailto={urllib.parse.quote(MAILTO)}")
        if res:
            rec = parse_crossref(res.get("message"))
            if len(rec) > 1:
                out[d] = rec
        if n % 25 == 0 or n == len(dois):
            print(f"  Crossref by DOI: {n}/{len(dois)}, {len(out)} returned data")
        time.sleep(PACE)
    return out


# ---------------------------------------------------------------------------

ELSEVIER_ABSTRACT = "https://api.elsevier.com/content/abstract/doi/"


def dig_elsevier_abstract(payload):
    """Pull the abstract out of an Elsevier response.

    Written to survive the response shape moving: the text sits in coredata
    as dc:description for most records, but older ones carry it deeper in the
    bibrecord, and dc:description is sometimes a string and sometimes a dict.
    Each shape is tried in turn and anything unrecognised returns nothing
    rather than raising."""
    try:
        resp = payload.get("abstracts-retrieval-response") or {}
    except AttributeError:
        return ""

    def text(v):
        if isinstance(v, str):
            return clean(v)
        if isinstance(v, dict):
            for k in ("$", "para", "ce:para", "abstract"):
                if k in v:
                    return text(v[k])
            return ""
        if isinstance(v, list):
            return clean(" ".join(x for x in (text(i) for i in v) if x))
        return ""

    core = resp.get("coredata") or {}
    for key in ("dc:description", "dc:abstract", "abstract"):
        got = text(core.get(key))
        if got:
            return re.sub(r"^\s*abstract\s*[:.\-]?\s*", "", got, flags=re.I)
    try:
        head = resp["item"]["bibrecord"]["head"]
        got = text(head.get("abstracts"))
        if got:
            return re.sub(r"^\s*abstract\s*[:.\-]?\s*", "", got, flags=re.I)
    except (KeyError, TypeError):
        pass
    return ""


def pass_elsevier(dois):
    """One request per DOI against Elsevier's Abstract Retrieval API."""
    out, refused, empty = {}, 0, 0
    hdrs = {"X-ELS-APIKey": ELSEVIER_API_KEY,
            "Accept": "application/json",
            "User-Agent": f"slr-complete/2.0 (mailto:{MAILTO})"}
    if ELSEVIER_INSTTOKEN:
        hdrs["X-ELS-Insttoken"] = ELSEVIER_INSTTOKEN
    for n, d in enumerate(dois, 1):
        url = f"{ELSEVIER_ABSTRACT}{urllib.parse.quote(d)}?httpAccept=application/json"
        req = urllib.request.Request(url, headers=hdrs)
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                payload = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                print(f"    Elsevier refused the request ({e.code}). The key is\n"
                      "    rejected, or you are off the HU network without an\n"
                      "    institutional token. Skipping this pass.",
                      file=sys.stderr)
                return out
            if e.code == 429:
                print("    Elsevier rate limit reached; stopping this pass "
                      "here.", file=sys.stderr)
                return out
            refused += 1
            time.sleep(PACE)
            continue
        except Exception:
            refused += 1
            time.sleep(PACE)
            continue
        ab = dig_elsevier_abstract(payload)
        if ab:
            out[d] = {"abstract": ab, "_from": "elsevier"}
        else:
            empty += 1
        if n % 20 == 0 or n == len(dois):
            print(f"  Elsevier: {n}/{len(dois)}, {len(out)} abstracts "
                  f"({empty} held none, {refused} failed)")
        time.sleep(max(PACE, 0.25))
    return out


def main():
    ap = argparse.ArgumentParser(
        description="Step 2 — complete every field. Settings are at the top "
                    "of the file.")
    ap.add_argument("--limit", type=int, default=0,
                    help="Only look up this many records, for a quick test.")
    args = ap.parse_args()

    src = find_input()
    in_delim = sniff_delimiter(src)
    with open(src, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh, delimiter=in_delim))
    if not rows:
        raise SystemExit(f"{src.name} has no rows.")
    shown = {",": "comma", ";": "semicolon", "\t": "tab"}
    print(f"  delimiter detected: {shown.get(in_delim, repr(in_delim))}")
    if len(rows[0]) < 3:
        raise SystemExit(
            f"\n{src.name} parsed into only {len(rows[0])} column(s), which "
            "means the\n  delimiter was read wrongly. The header looks like:\n"
            f"    {list(rows[0])}\n"
            "  Check that the file was written by step 1 and not re-saved by "
            "Excel\n  with a different separator.")

    cols = list(rows[0].keys())
    for extra in ("abstract_source", "enriched_fields"):
        if extra not in cols:
            cols.append(extra)
    present = [f for f in COMPLETABLE if f in rows[0]]
    for r in rows:
        r.setdefault("abstract_source", "")
        r.setdefault("enriched_fields", "")
        if clean(r.get("abstract")) and not r["abstract_source"]:
            r["abstract_source"] = "openalex"

    out_path = src.with_name(src.stem + SUFFIX + ".csv")
    ckpt = str(out_path) + ".progress.json"
    found = {"doi": {}, "title": {}}
    if os.path.exists(ckpt):
        try:
            found = json.load(open(ckpt, encoding="utf-8"))
            print(f"Resuming: {len(found['doi'])} by DOI, "
                  f"{len(found['title'])} by title already fetched\n")
        except Exception:
            found = {"doi": {}, "title": {}}

    def save():
        json.dump(found, open(ckpt, "w", encoding="utf-8"))

    def gaps(r):
        return [f for f in present if not clean(r.get(f))]

    print(f"Read {len(rows):,} records from {src.name}")
    print(f"  complete already: {sum(1 for r in rows if not gaps(r)):,}")
    gc = Counter(f for r in rows for f in gaps(r))
    if gc:
        print("\n  what is blank:")
        for f, n in gc.most_common():
            mark = "  <- OpenAlex only" if f in OPENALEX_ONLY else ""
            print(f"    {f:18s} {n:5d}{mark}")
    print()

    def merge_into(store, rec_map):
        """Keep the first value seen for each field, since passes run in
        order of how much we trust them for that field."""
        for k, rec in rec_map.items():
            tgt = store.setdefault(k, {})
            for f, v in rec.items():
                tgt.setdefault(f, v)

    try:
        # ---- pass 1: recover DOIs, so later passes have a key -------------
        no_doi = [r for r in rows
                  if not norm_doi(r.get("doi")) and clean(r.get("title"))]
        if no_doi and RUN_TITLE_SEARCH:
            items = [{"key": norm_title(r.get("title"))[:120],
                      "title": clean(r.get("title")),
                      "year": clean(r.get("year"))} for r in no_doi]
            items = list({i["key"]: i for i in items
                          if i["key"] and i["key"] not in found["title"]}.values())
            if args.limit:
                items = items[:args.limit]
            if items:
                print(f"PASS 1 — {len(items)} records with no DOI, searched by title")
                merge_into(found["title"], pass_title_search(items))
                save()
                print()

        # Apply pass 1's DOIs now, so passes 2-4 can use them.
        recovered = 0
        for r in rows:
            if norm_doi(r.get("doi")):
                continue
            rec = found["title"].get(norm_title(r.get("title"))[:120])
            if rec and rec.get("doi") and "doi" in r:
                r["doi"] = rec["doi"]
                recovered += 1
        if recovered:
            print(f"  {recovered} DOIs recovered and now usable by the next passes\n")

        def needing(fields):
            """DOIs of records still missing any of these fields."""
            out = []
            for r in rows:
                d = norm_doi(r.get("doi"))
                if not d:
                    continue
                if any(f in r and not clean(r.get(f)) for f in fields):
                    out.append(d)
            return list(dict.fromkeys(out))

        # ---- pass 2: OpenAlex, for topic/subfield/field and affiliations ---
        if RUN_OPENALEX:
            todo = [d for d in needing(present)
                    if d not in found["doi"]
                    or not any(f in found["doi"][d] for f in OPENALEX_ONLY)]
            if args.limit:
                todo = todo[:args.limit]
            if todo:
                print(f"PASS 2 — {len(todo)} DOIs to OpenAlex "
                      "(topic, subfield, field, institutions)")
                merge_into(found["doi"], pass_openalex(todo))
                save()
                print()

        # ---- pass 3: Semantic Scholar, mainly abstracts --------------------
        if RUN_SEMANTIC_SCHOLAR:
            todo = [d for d in needing(BIBLIO + AFFIL)
                    if "abstract" not in found["doi"].get(d, {})]
            if args.limit:
                todo = todo[:args.limit]
            if todo:
                print(f"PASS 3 — {len(todo)} DOIs to Semantic Scholar")
                merge_into(found["doi"], pass_semantic_scholar(todo))
                save()
                print()

        # ---- pass 4: Crossref, for whatever is left ------------------------
        if RUN_CROSSREF:
            todo = []
            for r in rows:
                d = norm_doi(r.get("doi"))
                if not d:
                    continue
                have = found["doi"].get(d, {})
                if any(f in r and not clean(r.get(f)) and f not in have
                       for f in BIBLIO + AFFIL):
                    todo.append(d)
            todo = list(dict.fromkeys(todo))
            if args.limit:
                todo = todo[:args.limit]
            if todo:
                print(f"PASS 4 — {len(todo)} DOIs to Crossref")
                merge_into(found["doi"], pass_crossref_doi(todo))
                save()
                print()

        # ---- pass 5: Elsevier, for the abstracts nothing else can reach ----
        if RUN_ELSEVIER and ELSEVIER_API_KEY:
            todo = []
            for r in rows:
                d = norm_doi(r.get("doi"))
                if not d or clean(r.get("abstract")):
                    continue
                if "abstract" in found["doi"].get(d, {}):
                    continue
                todo.append(d)
            todo = list(dict.fromkeys(todo))
            if args.limit:
                todo = todo[:args.limit]
            if todo:
                print(f"PASS 5 — {len(todo)} DOIs to Elsevier")
                merge_into(found["doi"], pass_elsevier(todo))
                save()
                print()
        elif RUN_ELSEVIER:
            still = sum(1 for r in rows
                        if not clean(r.get("abstract")) and norm_doi(r.get("doi")))
            if still:
                print(f"PASS 5 — skipped. {still} records still have no abstract "
                      "and a DOI.\n  Elsevier does not give its abstracts to "
                      "Crossref, so OpenAlex never\n  has them. Set "
                      "ELSEVIER_API_KEY at the top of this script and run it\n"
                      "  on the HU network to try for those.\n")
    except KeyboardInterrupt:
        save()
        print("\nStopped. Progress saved — run it again to continue.",
              file=sys.stderr)
        raise SystemExit(1)

    # ---- apply everything, never overwriting what is already there --------
    filled, touched, derived_n = Counter(), 0, 0
    for r in rows:
        d = norm_doi(r.get("doi"))
        recs = []
        if d and d in found["doi"]:
            recs.append(found["doi"][d])
        tk = norm_title(r.get("title"))[:120]
        if tk in found["title"]:
            recs.append(found["title"][tk])
        if not recs:
            continue
        got, sources = [], []
        for f in present + (["doi"] if "doi" in r else []):
            if clean(r.get(f)):
                continue
            for rec in recs:
                if f in rec:
                    r[f] = rec[f]
                    got.append(f)
                    filled[f] += 1
                    sources.append(rec.get("_from", "lookup"))
                    if f == "author_countries" and rec.get("_derived_countries"):
                        derived_n += 1
                        got[-1] = "author_countries(derived)"
                    break
        if got:
            touched += 1
            note = "|".join(got) + " (" + "+".join(dict.fromkeys(sources))
            m = next((x.get("_match") for x in recs if x.get("_match")), None)
            note += f" {m:.0%})" if m else ")"
            r["enriched_fields"] = note
            if "abstract" in got:
                r["abstract_source"] = next(
                    (rec.get("_from", "lookup") for rec in recs
                     if "abstract" in rec), "lookup")
    for r in rows:
        if not r.get("abstract_source"):
            r["abstract_source"] = "none"

    # ---- recompute the European markers from whatever we now have --------
    if all(c in rows[0] for c in DERIVED):
        for r in rows:
            codes = {c.strip().upper()
                     for c in (r.get("author_countries") or "").split("|")
                     if c.strip()}
            eu_a = bool(EU_SET & codes)
            eu_t = bool(EUROPE_TEXT.search(
                f"{r.get('title', '')} {r.get('abstract', '')}"))
            r["europe_affiliation"] = "yes" if eu_a else ""
            r["europe_text"] = "yes" if eu_t else ""
            r["europe"] = "yes" if (eu_a or eu_t) else ""

    out_delim = CSV_DELIMITER or in_delim
    xlsx = write_table(out_path, cols, rows, out_delim, WRITE_XLSX)
    if os.path.exists(ckpt):
        os.remove(ckpt)

    def have(f):
        return sum(1 for r in rows if clean(r.get(f)))

    print("Wrote")
    if xlsx:
        print(f"  {xlsx.name}   <- open this one, it has no delimiter at all")
    elif WRITE_XLSX:
        print("  (.xlsx skipped — run: pip install openpyxl)")
    print(f"  {out_path.name}   "
          f"({shown.get(out_delim, repr(out_delim))}-separated, every field quoted)")
    print(f"  records improved: {touched}")
    if filled:
        print("\n  fields filled:")
        for f, n in filled.most_common():
            print(f"    {f:18s} {n:5d}")
    if derived_n:
        was = "list was" if derived_n == 1 else "lists were"
        print(f"\n  {derived_n} country {was} read out of affiliation text "
              "rather than\n  supplied as codes, so they are inferred. They are "
              "marked\n  'author_countries(derived)' in enriched_fields, and the "
              "europe\n  columns were recomputed from them.")
    print("\n  completeness now:")
    for f in present:
        n = have(f)
        bar = "" if n == len(rows) else f"   {len(rows) - n} still blank"
        print(f"    {f:18s} {n:5d}/{len(rows)}  "
              f"({n / len(rows) * 100:3.0f}%){bar}")
    print("\n  abstract provenance:")
    for k, v in Counter(r["abstract_source"] for r in rows).most_common():
        print(f"    {k:18s} {v:5d}")
    if all(c in rows[0] for c in DERIVED):
        print(f"\n  European (any signal): "
              f"{sum(1 for r in rows if r['europe'] == 'yes'):,}/{len(rows):,}")
    stuck = [f for f in OPENALEX_ONLY if f in rows[0] and have(f) < len(rows)]
    if stuck:
        print(f"\n  {', '.join(stuck)} could not be completed for every record. "
              "Those\n  come from OpenAlex's own taxonomy, so where OpenAlex has "
              "no topic for\n  a work nothing else can supply one. Left blank "
              "rather than filled from\n  a different classification that would "
              "not be comparable.")

    missing = len(rows) - have("abstract")
    if missing:
        print(f"\n  {missing} records still have no abstract. Where they sit:")
        gaps = Counter((r.get("journal") or "(no journal)").strip()
                       for r in rows if not clean(r.get("abstract")))
        per_journal = Counter((r.get("journal") or "(no journal)").strip()
                              for r in rows)
        for j, n in gaps.most_common(10):
            print(f"    {n:>4} of {per_journal[j]:<4} {j[:52]}")
        print("\n  A journal missing nearly all of them is the publisher's "
              "choice, not a\n  fault here. Elsevier deposits no abstracts to "
              "Crossref and Springer\n  Nature deposits them only for "
              "open-access content, so OpenAlex never\n  receives them. Two ways "
              "forward: set ELSEVIER_API_KEY at the top of\n  this script and "
              "rerun on the HU network, or type the ones that matter\n  in by "
              "hand in step 3, where they then feed the rules, the ranking and\n"
              "  every export. Screening the rest on title alone is worth a "
              "line in\n  your methods.")
    print("\n  Next:  python 03-review_app.py")


if __name__ == "__main__":
    main()