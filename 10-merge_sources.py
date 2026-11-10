#!/usr/bin/env python3
"""
STEP 10 — put every database into one corpus file.

    python 10-merge_sources.py
    python 10-merge_sources.py --dry-run     # report, write nothing

Why this exists
---------------
A review used to live in two files that had to be re-joined every time the
app started: the harvest, plus an _extra_records.csv holding everything
imported afterwards. Which corpus happened to be open then decided where
decisions and imports were written. That is how a 36-row gold set got opened
instead of a 1,082-record corpus, how 446 judgements went missing, and how a
PRISMA diagram built from four databases still said every record came from
OpenAlex.

One file fixes all of it. Every row from every database, each carrying the
source it came from, in a corpus you can open in Excel and read.

What it does NOT do
-------------------
It does not deduplicate. The same paper returned by OpenAlex and Web of
Science stays as two rows, because resolving that is a judgement, not a
merge: the app groups them on the Duplicates tab, keeps the most complete
version, fills its gaps from the others, and lets you change the choice.
Deduplicating here would make that invisible again and would throw away the
abstracts that a second database is often the only one to have.

Re-running it is safe. The output is rebuilt from the inputs every time, so
adding a new export means dropping it in SOURCES and running this again.
Your decisions are not stored here and are never touched.
"""

import argparse
import csv
import datetime
import os
import re
import sys
from pathlib import Path

# ==========================================================================
#  CONFIG — the files that go in, and the file that comes out
# ==========================================================================

# An .xlsx is written. A CSV is written too only with --csv: having both a
# .csv and an .xlsx of the same corpus in the folder means two files the app
# has to choose between, for no benefit — they hold identical data and share
# one decisions file.
OUTPUT_FILE = "10-corpus_all_sources"

# (pattern, source key). The key is what PRISMA reports and what the Screen
# tab filters on. Keys in OTHER_METHODS land in PRISMA's second
# identification arm; everything else counts as a database.
#
# The pattern is matched against the folder, so the same line keeps working
# whether a file is .csv or .xlsx this week and whatever number it carries.
# If a pattern matches more than one file the newest wins, and the choice is
# printed — guessing silently is how the wrong file gets used.
SOURCES = [
    ("*openalex*complete*.*",      "openalex"),
    ("*web_of_science*.*",         "wos"),
    ("*lens*.*",                   "lens"),
    ("*goldset*.*",                "reflist"),
]

# Never read as a source, however the pattern matches: files this pipeline
# writes itself, and the app's own working files.
IGNORE = ("_decisions", "_extra_records", "_import_log", "_rules", "_labels",
          "_criteria", "_stop", "_backup", "_beforefix", "corpus_all_sources",
          "~$")

OTHER_METHODS = {"reflist", "manual", "citation"}

SOURCE_LABELS = {
    "openalex": "OpenAlex", "wos": "Web of Science", "lens": "Lens",
    "scopus": "Scopus", "other-db": "Other database",
    "reflist": "Added by hand (known papers)",
    "manual": "Added by hand", "citation": "Citation searching",
}

CSV_DELIMITER = ";"        # Dutch/European Excel opens this without an import

# ==========================================================================

# What each export calls things. The first name present in the file wins.
FIELD_MAP = {
    "openalex": {
        "id": ["openalex_id", "id"], "doi": ["doi"], "title": ["title"],
        "authors": ["authors"], "year": ["year", "publication_year"],
        "journal": ["journal", "source"], "type": ["type"],
        "abstract": ["abstract"], "language": ["language"],
        "institutions": ["institutions"], "topic": ["topic"],
        "subfield": ["subfield"], "field": ["field"],
        "cited_by_count": ["cited_by_count", "cited_by", "citations"],
        "retrieved": ["retrieved"], "author_countries": ["author_countries"],
        "europe_affiliation": ["europe_affiliation"],
        "europe_text": ["europe_text"], "europe": ["europe", "europe_any"],
        "abstract_source": ["abstract_source"], "keywords": ["keywords"],
        "issn": ["issn"], "url": ["url", "landing_page_url"],
        "oa_status": ["oa_status"], "apa_citation": ["apa_citation"],
    },
    "wos": {
        "id": ["UT (Unique ID)"], "doi": ["DOI"], "title": ["Article Title"],
        "authors": ["Authors"], "year": ["Publication Year"],
        "journal": ["Source Title"], "type": ["Document Type"],
        "abstract": ["Abstract"], "language": ["Language"],
        "cited_by_count": ["Times Cited, WoS Core", "Times Cited, All Databases"],
        "keywords": ["Author Keywords", "Keywords Plus"],
        "issn": ["ISSN", "eISSN"], "institutions": ["Institution"],
    },
    "lens": {
        "id": ["Lens ID"], "doi": ["DOI"], "title": ["Title"],
        "authors": ["Author/s"], "year": ["Publication Year"],
        "journal": ["Source Title"], "type": ["Publication Type"],
        "abstract": ["Abstract"], "cited_by_count": ["Citing Works Count"],
        "keywords": ["Keywords", "Fields of Study"], "issn": ["ISSNs"],
        "author_countries": ["Source Country"],
    },
    "scopus": {
        "id": ["EID"], "doi": ["DOI"], "title": ["Title"],
        "authors": ["Authors", "Author full names"], "year": ["Year"],
        "journal": ["Source title"], "type": ["Document Type"],
        "abstract": ["Abstract"], "cited_by_count": ["Cited by"],
        "language": ["Language of Original Document"],
        "keywords": ["Author Keywords", "Index Keywords"], "issn": ["ISSN"],
    },
    "reflist": {
        "doi": ["doi", "DOI"], "title": ["title", "Title"],
        "authors": ["authors", "Authors"], "year": ["year", "Year"],
        "journal": ["journal", "Journal"], "abstract": ["abstract", "Abstract"],
        "type": ["publication_type"], "language": ["language"],
        "keywords": ["keywords"],
        "cited_by_count": ["cited_by", "cited_by_count", "citations"],
        # core / background / candidate — what marks a key paper
        "key_paper": ["status", "key_paper"],
        "country": ["country", "Country"],
        "note": ["match_note", "note", "notes"],
    },
}
FIELD_MAP["other-db"] = {
    "doi": ["doi", "DOI"],
    "title": ["title", "Title", "Article Title", "primary_title"],
    "authors": ["authors", "Authors", "Author/s"],
    "year": ["year", "Year", "Publication Year", "publication_year"],
    "journal": ["journal", "Journal", "Source Title", "Source title"],
    "type": ["type", "Document Type", "Publication Type"],
    "abstract": ["abstract", "Abstract"], "language": ["language", "Language"],
    "cited_by_count": ["cited_by_count", "cited_by", "Cited by", "citations"],
    "keywords": ["keywords", "Keywords"], "issn": ["issn", "ISSN", "ISSNs"],
}

# The column order of the output. The app reads this natively.
COLUMNS = [
    "openalex_id", "source", "source_file", "doi", "title", "authors", "year",
    "journal", "type", "abstract", "language", "institutions", "topic",
    "subfield", "field", "cited_by_count", "keywords", "issn", "url",
    "oa_status", "apa_citation", "retrieved", "author_countries",
    "europe_affiliation", "europe_text", "europe", "abstract_source",
    "key_paper", "country", "note",
]

_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", _CTRL.sub(" ", s)).strip()


def strip_doi(s):
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", str(s or "").strip(),
                  flags=re.I).strip()


def norm_doi(s):
    return strip_doi(s).lower()


def norm_title(s):
    return re.sub(r"\s+", " ",
                  re.sub(r"[^a-z0-9 ]+", " ", str(s or "").lower())).strip()


def as_int(v, default=0):
    try:
        return int(float(re.sub(r"\.0$", "", str(v).strip())))
    except (ValueError, TypeError):
        return default


def sniff(path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        head = fh.readline()
    return ";" if head.count(";") >= head.count(",") else ","


def read_table(path):
    """Read .csv/.tsv/.xlsx/.xlsm, and .xls, which Web of Science exports."""
    p = Path(path)
    suf = p.suffix.lower()
    if suf == ".xls":
        try:
            import xlrd
        except ImportError:
            raise RuntimeError(
                f"{p.name} is an old-format Excel file. Reading one needs:\n"
                f"    pip install xlrd\n"
                f"Or re-export it from Web of Science as .xlsx.")
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
                # xlrd makes every number a float: 2021 arrives as "2021.0".
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                row[hdr[c]] = str(v).strip()
            out.append(row)
        return out
    if suf in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            raise RuntimeError("Reading .xlsx needs:  pip install openpyxl")
        wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
        ws = wb.worksheets[0]
        grid = list(ws.iter_rows(values_only=True))
        wb.close()
        if not grid:
            return []
        hdr = [(str(c).strip() if c is not None else f"col{i}")
               for i, c in enumerate(grid[0])]
        return [dict(zip(hdr, g)) for g in grid[1:]]
    csv.field_size_limit(10 ** 7)
    with open(p, newline="", encoding="utf-8-sig") as fh:
        return [dict(r) for r in csv.DictReader(fh, delimiter=sniff(p))]


def map_rows(rows, source, filename):
    """Turn one export into rows shaped like the merged corpus."""
    fmap = FIELD_MAP.get(source, FIELD_MAP["other-db"])
    cols = list(rows[0]) if rows else []
    lower = {str(c).strip().lower(): c for c in cols}

    def find(names):
        for n in names:
            if n in cols:
                return n
            if n.strip().lower() in lower:
                return lower[n.strip().lower()]
        return None

    picked = {f: find(names) for f, names in fmap.items()}
    out, skipped = [], 0
    for r in rows:
        g = lambda f: clean(r.get(picked[f])) if picked.get(f) else ""
        title, doi = g("title"), strip_doi(g("doi"))
        if not title and not doi:
            skipped += 1          # nothing to identify it by
            continue
        yr = g("year")
        m = re.search(r"(19|20)\d{2}", yr)
        rec = {c: "" for c in COLUMNS}
        rec.update({
            "openalex_id": g("id") or "", "source": source,
            "source_file": filename, "doi": doi, "title": title,
            "authors": g("authors"), "year": m.group(0) if m else "",
            "journal": g("journal"), "type": g("type"),
            "abstract": g("abstract"), "language": g("language"),
            "institutions": g("institutions"), "topic": g("topic"),
            "subfield": g("subfield"), "field": g("field"),
            "cited_by_count": as_int(g("cited_by_count")),
            "keywords": g("keywords"), "issn": g("issn"), "url": g("url"),
            "oa_status": g("oa_status"), "apa_citation": g("apa_citation"),
            "retrieved": g("retrieved"),
            "author_countries": g("author_countries"),
            "europe_affiliation": g("europe_affiliation"),
            "europe_text": g("europe_text"), "europe": g("europe"),
            "abstract_source": g("abstract_source"),
            "key_paper": g("key_paper").lower(), "country": g("country"),
            "note": g("note"),
        })
        out.append(rec)
    return out, picked, skipped


def row_id(rec, source, seen):
    """A stable id, from the row's own DOI or title.

    Derived rather than counted, so re-running this gives every row the id it
    had last time. Decisions are stored against ids; numbering by position
    would move them to different papers whenever the inputs changed."""
    import hashlib
    native = clean(rec.get("openalex_id"))
    if source == "openalex" and native:
        return native
    basis = norm_doi(rec.get("doi")) or norm_title(rec.get("title"))
    if not basis:
        return ""
    h = hashlib.sha1(basis.encode("utf-8")).hexdigest()[:10]
    rid = f"{source.upper().replace('-', '')}-{h}"
    if rid in seen:                       # same paper twice in one export
        return ""
    return rid


def main():
    ap = argparse.ArgumentParser(
        description="Merge every database export into one corpus file.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Report what would be merged; write nothing.")
    ap.add_argument("--csv", action="store_true",
                    help="Also write a .csv beside the .xlsx.")
    ap.add_argument("--out", default=OUTPUT_FILE)
    args = ap.parse_args()

    folder = Path(__file__).resolve().parent
    merged, seen = [], set()
    report = []

    def resolve(pattern):
        """The file this pattern means, newest first, writable files ignored."""
        hits = [q for q in folder.glob(pattern)
                if q.is_file() and q.suffix.lower() in
                (".csv", ".tsv", ".xlsx", ".xlsm", ".xls")
                and not any(k in q.name for k in IGNORE)]
        hits.sort(key=lambda q: q.stat().st_mtime, reverse=True)
        return hits

    print("Reading:")
    for pattern, source in SOURCES:
        hits = resolve(pattern)
        if not hits:
            print(f"  !! nothing matching {pattern} — skipped")
            report.append((pattern, source, 0, 0, 0, "not found"))
            continue
        p = hits[0]
        if len(hits) > 1:
            print(f"  .. {pattern} matched {len(hits)} files; using the newest,"
                  f" {p.name}  (others: {', '.join(q.name for q in hits[1:4])})")
        fname = p.name
        try:
            rows = read_table(p)
        except RuntimeError as e:
            print(f"  !! {fname}: {e}")
            report.append((fname, source, 0, 0, 0, "unreadable"))
            continue
        mapped, picked, no_id = map_rows(rows, source, fname)
        kept, repeated = 0, 0
        for rec in mapped:
            rid = row_id(rec, source, seen)
            if not rid:
                repeated += 1
                continue
            seen.add(rid)
            rec["openalex_id"] = rid
            merged.append(rec)
            kept += 1
        report.append((fname, source, len(rows), kept, repeated + no_id, ""))
        missing = [f for f, v in picked.items() if not v]
        print(f"  {fname:44s} {SOURCE_LABELS.get(source, source):28s} "
              f"{len(rows):5,} rows -> {kept:5,} kept"
              + (f", {repeated + no_id} repeated or unusable" if repeated + no_id else ""))
        if missing:
            print(f"      not in this file: {', '.join(missing[:9])}")

    if not merged:
        sys.exit("\nNothing was merged. Check the filenames in SOURCES.")

    print(f"\n  merged corpus: {len(merged):,} records")
    by_src = {}
    for r in merged:
        by_src[r["source"]] = by_src.get(r["source"], 0) + 1
    for s, n in sorted(by_src.items(), key=lambda kv: -kv[1]):
        arm = "other methods" if s in OTHER_METHODS else "database"
        print(f"     {SOURCE_LABELS.get(s, s):30s} {n:6,}   {arm}")

    # How much overlap the app will have to resolve. Reported, not removed:
    # which version of a duplicate to keep is a judgement, and the app shows
    # it to you on the Duplicates tab with the reason it chose.
    bydoi, bytitle = {}, {}
    for r in merged:
        d, t = norm_doi(r["doi"]), norm_title(r["title"])
        if d:
            bydoi.setdefault(d, []).append(r["source"])
        elif t:
            bytitle.setdefault(t, []).append(r["source"])
    dup_rows = sum(len(v) - 1 for v in bydoi.values() if len(v) > 1)
    dup_rows += sum(len(v) - 1 for v in bytitle.values() if len(v) > 1)
    print(f"\n  {dup_rows:,} of these are the same paper arriving from more "
          f"than one source.\n  They are kept, with their sources, and "
          f"resolved on the app's Duplicates tab.")
    abs_missing = sum(1 for r in merged if not r["abstract"])
    print(f"  {len(merged) - abs_missing:,} have an abstract, "
          f"{abs_missing:,} do not.")
    keyp = sum(1 for r in merged if r["key_paper"] == "core")
    if keyp:
        print(f"  {keyp} marked core — the key papers the stopping rule "
              f"measures your search against.")

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return

    stamp = datetime.date.today().isoformat()
    for r in merged:
        r["retrieved"] = r["retrieved"] or stamp

    if args.csv:
        out_csv = folder / f"{args.out}.csv"
        tmp = str(out_csv) + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=COLUMNS, delimiter=CSV_DELIMITER,
                               quoting=csv.QUOTE_ALL, lineterminator="\r\n",
                               extrasaction="ignore")
            w.writeheader()
            w.writerows(merged)
        try:
            os.replace(tmp, out_csv)
        except PermissionError:
            os.remove(tmp)
            sys.exit(f"\nCould not write {out_csv.name} — it is open in "
                     f"another program. Close it and run this again.")
        print(f"\n  wrote {out_csv.name}")
    else:
        # A stale .csv from an earlier run would be a second corpus for the
        # app to choose between, holding older data under the same name.
        stale = folder / f"{args.out}.csv"
        if stale.exists():
            try:
                stale.unlink()
                print(f"\n  removed the older {stale.name} "
                      f"(use --csv to keep one)")
            except OSError:
                print(f"\n  !! {stale.name} is older than the .xlsx and will "
                      f"confuse the app — delete it")

    try:
        import openpyxl
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("  (.xlsx skipped — pip install openpyxl)")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "corpus"
    ws.append(COLUMNS)
    NUM = {"year", "cited_by_count"}
    for r in merged:
        row = []
        for c in COLUMNS:
            v = r.get(c, "")
            if c in NUM and str(v).strip():
                try:
                    v = int(v)
                except (TypeError, ValueError):
                    pass
            elif isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
                v = "'" + v          # stop Excel evaluating it as a formula
            row.append(v)
        ws.append(row)
    for c in ws[1]:
        c.font = Font(bold=True)
    widths = {"title": 64, "abstract": 110, "authors": 40, "journal": 34,
              "doi": 30, "source": 12, "source_file": 32, "keywords": 36,
              "institutions": 40, "note": 34, "openalex_id": 20}
    for i, c in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 14)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(merged)+1}"
    out_x = folder / f"{args.out}.xlsx"
    try:
        wb.save(out_x)
    except PermissionError:
        sys.exit(f"\nCould not write {out_x.name} — close it in Excel and "
                 f"run this again. The .csv was written.")
    print(f"  wrote {out_x.name}")
    print(f"\nOpen it with:\n    python \"09-review_app.py\" \"{out_x.name}\"")


if __name__ == "__main__":
    main()
