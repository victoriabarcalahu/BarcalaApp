#!/usr/bin/env python3
"""
STEP 1 of 3 — download the records from OpenAlex.

  python 01-harvest.py --count-only     <- do this first, it is quick
  python 01-harvest.py

Everything you are likely to change is in the CONFIGURATION block directly
below: the output file name, the years, the publication types, the European
filter, and the two term lists that make up the search.

  Step 1  01-harvest.py          OpenAlex  ->  the file named below
  Step 2  02-complete.py         fills in missing abstracts and metadata
                                 ->  same name with "_complete" added
  Step 3  03-preclassify.py      applies inclusion/exclusion criteria and
                                 labels, for you to check in the review app

Standard library only. Everything is written next to this script.
"""

# ===========================================================================
#  CONFIGURATION — edit this block
# ===========================================================================

OUTPUT_FILE = "01-openalex_raw.csv"

# --- how the file is written -----------------------------------------------
# The comma problem: abstracts contain commas (100% of them) and semicolons
# (13%), so no delimiter is safe on its own. Two things fix it.
#
#   WRITE_XLSX = True   also writes an .xlsx next to the CSV. A spreadsheet
#                       file has no delimiter at all, so the question cannot
#                       arise. Open this one. Needs: pip install openpyxl
#
#   CSV_DELIMITER       for the .csv, which other tools and step 2 read.
#                         ";"   Dutch and most European Excel installations
#                         ","   the international default
#                         "\t"  tab-separated; tabs appear in 0% of abstracts
#                               and Excel opens it cleanly in any locale
#
# Every field is written fully quoted, which is what actually makes a comma
# or semicolon inside an abstract harmless.
WRITE_XLSX = True
CSV_DELIMITER = ";"

YEAR_FROM = 2010
YEAR_TO = 2026

# OpenAlex type names. "article" and "book-chapter" as you asked.
# Others you could add: "review", "preprint", "book", "dissertation".
TYPES = ["article", "book-chapter"]

# The search is  (any MECHANISM term)  AND  (any TOPIC term),
# matched against title, abstract and keywords.
#
# Rules OpenAlex imposes, so that editing these does not silently break it:
#   - No commas inside a term. The comma separates filters, so one inside
#     the search cuts the query short with no error.
#   - No wildcards. No * and no ?. OpenAlex stems automatically, so
#     "signalling" already reaches "signal" and "signals".
#   - A longer phrase can only narrow. "equity crowdfunding" cannot match
#     anything "crowdfunding" misses, so adding it changes nothing.

MECHANISM_TERMS = [
    "lending-based crowdfunding",
    "crowdlending",
    "crowd lending",
    "peer-to-peer lending",
    "peer to peer lending",
    "p2p lending",
]

TOPIC_TERMS = [
    # sustainability
    "sustainability", "sustainable", "esg", "SDG", "impact investing",
    "green",
    "responsible investment", "ethical investment",
    "impact investment", "social impact", "environmental impact",
    "renewable energy", "low-carbon", "energy transition",
    "social entrepreneurship", "greenwashing",
    # T1 investor motives
    "investor motivation", "investor behaviour", "investor behavior",
    "lender behaviour", "lender behavior", "warm glow", "prosocial",
    "pro-social", "altruism", "human values", "investment decision",
    "willingness to invest", "investor profile", "investor confidence",
    # T2 signals
    "signalling", "signaling", "quality signal", "information asymmetry",
    "certification", "third-party", "campaign success", "funding success",
    "rhetorical signals", "credit risk",
    # T3 impact, reporting, regulation
    "impact measurement", "impact reporting", "sustainability reporting",
    "non-financial reporting", "impact metrics", "crowdfunding regulation",
    "harmonization", "harmonisation",
]

# ---------------------------------------------------------------------------
# European filter.  "none" | "affiliation"
#
#   "none"         Harvest everything, and write europe_affiliation,
#                  europe_text and europe columns so you can filter while
#                  screening. RECOMMENDED.
#
#   "affiliation"  Add institutions.country_code to the query, so OpenAlex
#                  returns only papers with at least one author at a European
#                  institution.
#
# What "affiliation" costs, measured on your existing 1,793-record corpus:
#   - 443 records (25%) have NO author-country data in OpenAlex at all.
#     The filter discards every one of them and does not tell you.
#   - It also discards papers that study European platforms but whose
#     authors sit elsewhere, which is a real part of this literature.
# Geography is cheap to apply later and expensive to apply too early, which
# is why the default leaves it out. Set it here if you want it anyway.
# ---------------------------------------------------------------------------
EUROPE_FILTER = "none"

EUROPE_COUNTRIES = [
    "NL", "DE", "FR", "IT", "ES", "BE", "AT", "SE", "NO", "DK", "FI", "PL",
    "PT", "IE", "GB", "CZ", "HU", "RO", "GR", "EE", "LV", "LT", "SI", "HR",
    "LU", "SK", "BG", "CH", "IS", "CY", "MT",
]

MAILTO = "victoria.barcala@hu.nl"
OPENALEX_API_KEY = ""          # only needed past 10,000 results per chunk
PACE = 1.0                     # seconds between requests

# ===========================================================================
#  Below here you should not need to change anything.
# ===========================================================================

import argparse
import csv
import datetime
import difflib
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

RETRIEVED = datetime.date.today().isoformat()
API = "https://api.openalex.org/works"
SCRIPT_DIR = Path(__file__).resolve().parent
BASIC_PAGING_LIMIT = 10000
URL_LIMIT = 2000               # stay under the common 2,048 ceiling
GOLDSET_FILE = "goldset_from_review.csv"      # optional; skipped if absent

SELECT = ",".join([
    "id", "doi", "title", "display_name", "publication_year", "type",
    "language", "authorships", "primary_location", "primary_topic",
    "cited_by_count", "abstract_inverted_index",
])

# Exactly the fields you asked for, then four that geography screening needs.
COLUMNS = [
    "openalex_id", "doi", "title", "authors", "year", "journal", "type",
    "abstract", "language", "institutions", "topic", "subfield", "field",
    "cited_by_count",
    # The date the count was taken. A citation count is only meaningful with
    # one, it changes weekly, and PRISMA asks for the search date anyway.
    "retrieved",
    # added so you can filter for Europe while screening instead of in the query
    "author_countries", "europe_affiliation", "europe_text", "europe",
]

EU_SET = {c.upper() for c in EUROPE_COUNTRIES}
EUROPE_TEXT = re.compile(
    r"netherlands|dutch|german|\bfrance\b|french|ital(y|ian)|spain|spanish|"
    r"portug(al|uese)|belgi(um|an)|austria|switzerland|swiss|swed(en|ish)|"
    r"norw(ay|egian)|denmark|danish|finland|finnish|poland|polish|czech|"
    r"hungar(y|ian)|greece|greek|ireland|irish|estonia|latvia|lithuania|"
    r"slovenia|croatia|luxembourg|slovakia|bulgaria|iceland|cyprus|malta|"
    r"romania|united kingdom|\buk\b|britain|british|england|scotland|wales|"
    r"europe|european|\beu\b|eurozone", re.I)

CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
WS = re.compile(r"\s+")


class HardFail(Exception):
    pass


# Column widths for the .xlsx, so it is readable without dragging borders.
XLSX_WIDTH = {
    "openalex_id": 14, "doi": 26, "title": 60, "authors": 38, "year": 7,
    "journal": 32, "type": 12, "abstract": 80, "language": 9,
    "institutions": 38, "topic": 28, "subfield": 20, "field": 18,
    "cited_by_count": 11, "retrieved": 12,
    "author_countries": 16, "europe_affiliation": 12, "europe_text": 11,
    "europe": 8,
}


def write_table(path, columns, rows, delimiter=",", want_xlsx=True):
    """Write the CSV fully quoted, and an .xlsx beside it when openpyxl is
    installed. Returns the .xlsx path, or None if it was not written."""
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


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return WS.sub(" ", CTRL.sub(" ", s)).strip()


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


def check_config():
    problems = []
    for name, terms in (("MECHANISM_TERMS", MECHANISM_TERMS),
                        ("TOPIC_TERMS", TOPIC_TERMS)):
        if not terms:
            problems.append(f"{name} is empty.")
        for t in terms:
            if "," in t:
                problems.append(f'{name}: "{t}" contains a comma, which would '
                                "truncate the query. Remove it.")
            if "*" in t or "?" in t:
                problems.append(f'{name}: "{t}" contains a wildcard. OpenAlex '
                                "does not support them; it stems automatically.")
            if not t.strip():
                problems.append(f"{name} contains a blank entry.")
    if EUROPE_FILTER not in ("none", "affiliation"):
        problems.append(f'EUROPE_FILTER is "{EUROPE_FILTER}"; use "none" or '
                        '"affiliation".')
    if YEAR_FROM > YEAR_TO:
        problems.append(f"YEAR_FROM ({YEAR_FROM}) is after YEAR_TO ({YEAR_TO}).")
    if not TYPES:
        problems.append("TYPES is empty.")
    if problems:
        raise SystemExit("\nProblems in the CONFIGURATION block:\n  - "
                         + "\n  - ".join(problems))


def or_block(terms):
    return " OR ".join(f'"{t}"' for t in terms)


def build_filter(topic_terms):
    q = f"({or_block(MECHANISM_TERMS)}) AND ({or_block(topic_terms)})"
    parts = [f"title_and_abstract.search:{q}",
             f"publication_year:{YEAR_FROM}-{YEAR_TO}",
             "type:" + "|".join(TYPES)]
    if EUROPE_FILTER == "affiliation":
        parts.append("institutions.country_code:"
                     + "|".join(c.lower() for c in EUROPE_COUNTRIES))
    return ",".join(parts)


def url_for(filter_str, extra=""):
    return (f"{API}?filter={urllib.parse.quote(filter_str)}&select={SELECT}"
            f"{extra}&mailto={urllib.parse.quote(MAILTO)}")


def plan_chunks():
    """Split TOPIC_TERMS into as few groups as keep every URL under the limit.

    Searching (M AND (a OR b)) then (M AND (c OR d)) and merging gives exactly
    the same set as (M AND (a OR b OR c OR d)); it just takes more requests.
    """
    n = 1
    while n <= len(TOPIC_TERMS):
        size = (len(TOPIC_TERMS) + n - 1) // n
        groups = [TOPIC_TERMS[i:i + size] for i in range(0, len(TOPIC_TERMS), size)]
        longest = max(len(url_for(build_filter(g), "&per-page=200&page=1"))
                      for g in groups)
        if longest <= URL_LIMIT:
            return groups, longest
        n += 1
    raise SystemExit("Even one term per request is too long. Shorten MAILTO or "
                     "the MECHANISM_TERMS list.")


def fetch(url, retries=7):
    full = url + (f"&api_key={urllib.parse.quote(OPENALEX_API_KEY)}"
                  if OPENALEX_API_KEY else "")
    last = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(
                full, headers={"User-Agent": f"slr-harvest/3 (mailto:{MAILTO})"})
            with urllib.request.urlopen(req, timeout=90) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            try:
                body = e.read().decode("utf-8", "replace")[:300]
            except Exception:
                body = ""
            if e.code in (401, 403):
                raise HardFail(
                    f"HTTP {e.code} — OpenAlex wants an API key.\n  {body}\n"
                    "  Free key: https://openalex.org/settings/api, then put it\n"
                    "  in OPENALEX_API_KEY at the top of this script.")
            if e.code == 414:
                raise HardFail(
                    "HTTP 414 — the URL is too long for your network.\n"
                    "  Lower URL_LIMIT near the top of this script (try 1500)\n"
                    "  and run it again; it will split into more chunks.")
            if e.code in (400, 404):
                raise HardFail(f"HTTP {e.code} — OpenAlex rejected the request.\n"
                               f"  {body}\n  Check the terms you edited.")
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
    raise HardFail(f"Gave up after {retries} attempts: {last}")


def row_from_work(w):
    names, insts, countries = [], [], []
    for a in (w.get("authorships") or []):
        au = a.get("author") or {}
        if au.get("display_name"):
            names.append(clean(au["display_name"]))
        countries += [c for c in (a.get("countries") or [])]
        for i in (a.get("institutions") or []):
            if i.get("display_name"):
                insts.append(clean(i["display_name"]))
            if i.get("country_code"):
                countries.append(i["country_code"])
    src = (w.get("primary_location") or {}).get("source") or {}
    pt = w.get("primary_topic") or {}
    title = clean(w.get("title") or w.get("display_name"))
    abstract = reconstruct_abstract(w.get("abstract_inverted_index"))
    ctry = list(dict.fromkeys(c.upper() for c in countries if c))
    eu_a = bool(EU_SET & set(ctry))
    eu_t = bool(EUROPE_TEXT.search(f"{title} {abstract}"))
    return {
        "openalex_id": (w.get("id") or "").replace("https://openalex.org/", ""),
        "doi": re.sub(r"^https?://(dx\.)?doi\.org/", "", w.get("doi") or "",
                      flags=re.I),
        "title": title,
        "authors": " | ".join(names),
        "year": w.get("publication_year") or "",
        "journal": clean(src.get("display_name")),
        "type": w.get("type") or "",
        "abstract": abstract,
        "language": w.get("language") or "",
        "institutions": " | ".join(dict.fromkeys(insts)),
        "topic": clean(pt.get("display_name")),
        "subfield": clean((pt.get("subfield") or {}).get("display_name")),
        "field": clean((pt.get("field") or {}).get("display_name")),
        # 0 is a real value, not a blank, so it is written as 0.
        "cited_by_count": int(w.get("cited_by_count") or 0),
        "retrieved": RETRIEVED,
        "author_countries": " | ".join(ctry),
        "europe_affiliation": "yes" if eu_a else "",
        "europe_text": "yes" if eu_t else "",
        "europe": "yes" if (eu_a or eu_t) else "",
    }


def harvest_chunk(label, filter_str, total):
    rows, per = [], 200
    if total > BASIC_PAGING_LIMIT and not OPENALEX_API_KEY:
        raise HardFail(
            f"{label} has {total:,} results, past the {BASIC_PAGING_LIMIT:,} "
            "that work without a key.\n  Free key: "
            "https://openalex.org/settings/api, then set OPENALEX_API_KEY.")
    last = (min(total, BASIC_PAGING_LIMIT) + per - 1) // per
    for page in range(1, last + 1):
        pay = fetch(url_for(filter_str, f"&per-page={per}&page={page}"))
        batch = pay.get("results") or []
        rows.extend(row_from_work(w) for w in batch)
        print(f"    page {page}/{last}: {len(rows):,}/{total:,}")
        if not batch:
            break
        if page < last:
            time.sleep(PACE)
    return rows


def check_goldset(rows, path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        gold = [g for g in csv.DictReader(fh)
                if (g.get("status") or "").strip() == "core"]
    if not gold:
        return
    by_doi = {norm_doi(r["doi"]) for r in rows if r["doi"]}
    by_title = {norm_title(r["title"]) for r in rows if r["title"]}
    tl = list(by_title)
    found, missing = [], []
    for g in gold:
        d, t = norm_doi(g.get("doi")), norm_title(g.get("title"))
        hit = bool(d and d in by_doi) or bool(t and t in by_title)
        if not hit and t and difflib.get_close_matches(t, tl, n=1, cutoff=0.82):
            hit = True
        (found if hit else missing).append(g)
    print(f"\n  RECALL CHECK against {os.path.basename(path)}")
    print(f"    found {len(found)}/{len(gold)} of your review's core references "
          f"({len(found) / len(gold) * 100:.0f}%)")
    if missing:
        print(f"    missing ({len(missing)}):")
        for g in missing[:12]:
            print(f"      {(g.get('authors') or '')[:24]:26s} "
                  f"{str(g.get('year') or ''):6s} {(g.get('title') or '')[:42]}")
        if len(missing) > 12:
            print(f"      ... and {len(missing) - 12} more")


def main():
    ap = argparse.ArgumentParser(
        description="Step 1 — harvest from OpenAlex. Settings are at the top "
                    "of the file.")
    ap.add_argument("--count-only", action="store_true",
                    help="Report how many results there are and stop.")
    args = ap.parse_args()

    check_config()
    out = SCRIPT_DIR / Path(OUTPUT_FILE).name
    try:
        probe = SCRIPT_DIR / f".write_test_{os.getpid()}"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except Exception as e:
        raise SystemExit(f"\nCannot write into {SCRIPT_DIR}\n  {e}")

    groups, longest = plan_chunks()
    print("STEP 1 — harvest from OpenAlex")
    print(f"  output file     {out}")
    print(f"  years           {YEAR_FROM}-{YEAR_TO}")
    print(f"  types           {', '.join(TYPES)}")
    print(f"  search          ({len(MECHANISM_TERMS)} mechanism terms) AND "
          f"({len(TOPIC_TERMS)} topic terms)")
    print(f"  matched against title + abstract")
    if EUROPE_FILTER == "affiliation":
        print(f"  EUROPEAN FILTER on: author affiliation in "
              f"{len(EUROPE_COUNTRIES)} countries")
        print("                      note this drops records with no "
              "affiliation data")
    else:
        print("  European filter off; europe columns are written instead")
    if len(groups) > 1:
        print(f"\n  The topic list makes a URL too long for one request, so it "
              f"runs as\n  {len(groups)} chunks of up to "
              f"{max(len(g) for g in groups)} terms "
              f"(longest URL {longest} characters).\n  Merging them gives "
              "exactly the same set as one query would.")
    print()

    try:
        counts, filters = [], []
        print("  Counting:")
        for i, g in enumerate(groups, 1):
            f = build_filter(g)
            c = fetch(url_for(f, "&per-page=1")).get("meta", {}).get("count", 0)
            counts.append(c)
            filters.append(f)
            print(f"    chunk {i}/{len(groups)}  {c:>7,}")
            if i < len(groups):
                time.sleep(PACE)
        if len(groups) > 1:
            print(f"    {'sum (before merging)':13s} {sum(counts):>7,}")
        if args.count_only:
            if any(c > BASIC_PAGING_LIMIT for c in counts) and not OPENALEX_API_KEY:
                print(f"\n  At least one chunk is past the {BASIC_PAGING_LIMIT:,} "
                      "keyless limit.\n  Get a free key at "
                      "https://openalex.org/settings/api and set\n  "
                      "OPENALEX_API_KEY, or narrow YEAR_FROM / TYPES.")
            print("\n  Nothing was written. Run without --count-only to harvest.")
            return

        rows = []
        for i, (f, c) in enumerate(zip(filters, counts), 1):
            if not c:
                continue
            print(f"\n  Chunk {i}/{len(groups)} ({c:,})")
            rows.extend(harvest_chunk(f"chunk {i}", f, c))
            if i < len(groups):
                time.sleep(PACE)
    except HardFail as e:
        raise SystemExit(f"\nStopped: {e}")

    merged, order = {}, []
    for r in rows:
        k = norm_doi(r["doi"]) or norm_title(r["title"]) or r["openalex_id"]
        if k not in merged:
            merged[k] = r
            order.append(k)
    uniq = [merged[k] for k in order]
    if len(groups) > 1:
        print(f"\n  {len(rows):,} across chunks -> {len(uniq):,} unique records")
    else:
        print(f"\n  {len(uniq):,} unique records")

    gpath = SCRIPT_DIR / GOLDSET_FILE
    if gpath.exists():
        check_goldset(uniq, str(gpath))

    xlsx = write_table(out, COLUMNS, uniq, CSV_DELIMITER, WRITE_XLSX)

    def n(p):
        return sum(1 for r in uniq if p(r))

    shown = {",": "comma", ";": "semicolon", "\t": "tab"}.get(
        CSV_DELIMITER, repr(CSV_DELIMITER))
    print(f"\n  Wrote {len(uniq):,} records")
    if xlsx:
        print(f"    {xlsx.name}   <- open this one, it has no delimiter at all")
    elif WRITE_XLSX:
        print("    (.xlsx skipped — run: pip install openpyxl)")
    print(f"    {out.name}   ({shown}-separated, every field quoted)")
    print(f"    with abstract        {n(lambda r: r['abstract']):>6,}  "
          f"({n(lambda r: r['abstract']) / max(1, len(uniq)) * 100:.0f}%)")
    print(f"    with DOI             {n(lambda r: r['doi']):>6,}")
    print(f"    with authors         {n(lambda r: r['authors']):>6,}")
    print(f"    with journal         {n(lambda r: r['journal']):>6,}")
    print(f"    English              {n(lambda r: r['language'] == 'en'):>6,}")
    print(f"    European (any sign)  {n(lambda r: r['europe'] == 'yes'):>6,}")
    print(f"      by affiliation     {n(lambda r: r['europe_affiliation'] == 'yes'):>6,}")
    print(f"      by text mention    {n(lambda r: r['europe_text'] == 'yes'):>6,}")
    print(f"    no country data      {n(lambda r: not r['author_countries']):>6,}"
          "   <- an affiliation filter would discard these")
    cites = sorted(int(r["cited_by_count"] or 0) for r in uniq)
    if cites:
        mid = cites[len(cites) // 2]
        print(f"\n    citations (as of {RETRIEVED}):")
        print(f"      never cited        {sum(1 for c in cites if c == 0):>6,}")
        print(f"      median             {mid:>6,}")
        print(f"      most cited         {cites[-1]:>6,}")
        print(f"      10 or more         {sum(1 for c in cites if c >= 10):>6,}")
    top = Counter(r["field"] for r in uniq if r["field"]).most_common(6)
    if top:
        print("\n    main fields:")
        for f, c in top:
            print(f"      {f[:38]:40s} {c:,}")
    print(f"\n  Next:  python 02-complete.py")


if __name__ == "__main__":
    main()