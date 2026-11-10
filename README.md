# BarcalaApp
App to select papers for a literatured review in a structured way. Methodolodigally based on PRISMA and SAFE procedure and with ASReview. Everythink works locally and it is traceable. A pass key can be installed to include questions.
# BarcalaReviewApp

A systematic literature review workbench that runs on your own computer, keeps
every judgement you make in a plain-text audit trail, and builds the PRISMA
flow diagram **from that audit trail** rather than from numbers you type in.

It combines searches from several bibliographic databases into one corpus
without losing track of which database each record came from, finds and merges
duplicates between them, lets you screen titles and abstracts with
machine-learning assistance ([ASReview](https://github.com/asreview/asreview)),
applies and documents your own exclusion rules, tells you when it is defensible
to stop screening, and exports the whole decision trail for a reviewer or a
replication.

Nothing leaves your machine except the searches you ask for yourself. There is
no account, no server and no database: one Python file, your corpus, and a CSV
of decisions beside it.

---

## Contents

- [Why this exists](#why-this-exists)
- [Requirements and installation](#requirements-and-installation)
- [The method, step by step](#the-method-step-by-step)
  - [1. Define the scope and the search words](#1-define-the-scope-and-the-search-words)
  - [2. Search and download from each database](#2-search-and-download-from-each-database)
  - [3. Which fields to ask each database for](#3-which-fields-to-ask-each-database-for)
  - [4. Merge into one corpus that remembers its sources](#4-merge-into-one-corpus-that-remembers-its-sources)
  - [5. Add papers you already know about](#5-add-papers-you-already-know-about)
  - [6. Screen](#6-screen)
  - [7. Stop, and say why](#7-stop-and-say-why)
  - [8. Report and export](#8-report-and-export)
- [Running the app](#running-the-app)
- [Tailoring it to your review](#tailoring-it-to-your-review)
  - [Themes (sub-topic tags)](#themes-sub-topic-tags)
  - [Inclusion and exclusion criteria](#inclusion-and-exclusion-criteria)
  - [Automatic rules](#automatic-rules)
  - [Regions and keyword rules](#regions-and-keyword-rules)
- [The tabs, one by one](#the-tabs-one-by-one)
- [Machine-learning assistance](#machine-learning-assistance)
- [The stopping rule](#the-stopping-rule)
- [The two dashboards](#the-two-dashboards)
- [The Ask feature](#the-ask-feature)
- [Export](#export)
- [Where everything is stored](#where-everything-is-stored)
- [Reporting this method in a paper](#reporting-this-method-in-a-paper)
- [What this tool will not do](#what-this-tool-will-not-do)
- [What is in this repository](#what-is-in-this-repository)
- [Tests](#tests)
- [Citing](#citing)
- [Licence](#licence)

---

## Why this exists

Screening tools and PRISMA tools are usually separate things. You screen in one
place, then open a diagram tool and type in the numbers by hand — which means
the published diagram is a *claim* about the screening rather than a *view* of
it. If a reviewer asks why 830 records were excluded at title and abstract, or
which database the 44 studies from one source came from, or whether a count
includes the duplicates, the honest answer is often that it would take a day's
work to reconstruct.

Here every number in the PRISMA diagram is computed, on demand, from the
decision records. The per-source counts, the exclusion reasons, the records
removed by an automatic rule, the duplicates, the papers added by hand because
you already knew about them — all of it comes out of the same audit trail you
screened into. Change a decision and the diagram changes. You cannot get a
diagram that disagrees with your data, because there is nowhere to type one in.

Two other things follow from that design:

- **Source provenance survives everything.** A record carries the database it
  came from, and the file it came from, through deduplication, merging, export
  and the diagram. When two databases return the same paper, the merged record
  says which fields were borrowed from which source.
- **Nothing is thrown away quietly.** Every row of every import is kept.
  Duplicates are flagged, not deleted. Rule-based exclusions are reversible and
  are reported separately in the diagram. Bulk operations back up the decisions
  file first.

---

## Requirements and installation

**Python 3.10 or newer.** The app itself uses nothing but the standard
library — you can run it on a locked-down work laptop with no installation at
all.

Three optional extras, each enabling one thing:

```bash
pip install openpyxl     # read and write .xlsx  (strongly recommended)
pip install xlrd         # read the old .xls that Web of Science still exports
pip install asreview     # the machine-learning ranker (see below)
```

Without `openpyxl` you can still work, in CSV only. Without `asreview` the
ranker falls back to a simpler built-in one and the app says so.

Clone or download, and you are done:

```bash
git clone https://github.com/<your-account>/BarcalaReviewApp.git
cd BarcalaReviewApp
pip install -r requirements.txt     # optional
python BarcalaReviewApp.py
```

Put the app in the same folder as your corpus, or point it at one:

```bash
python BarcalaReviewApp.py path/to/corpus_all_sources.xlsx
```

It opens a browser at `http://127.0.0.1:8766`. That page is the whole
interface. Nothing is sent anywhere; the server is listening on your own
machine only.

---

## The method, step by step

### 1. Define the scope and the search words

Do this before you touch a database, and write it down — the search string is
part of what you publish, and it is the single biggest determinant of what your
review can possibly find.

Build the query as a **population × mechanism × outcome** grid, one column per
concept, synonyms within a column joined by `OR`, columns joined by `AND`:

| | concept | terms (`OR` within the row) |
|---|---|---|
| 1 | the phenomenon | the platform, instrument or practice you study, **with its common variants and misspellings** |
| 2 | the mechanism | what the papers must be about — the behaviour, signal or effect |
| 3 | the outcome | optional; only add it if row 1 × row 2 returns far too much |

Two pieces of advice that are worth more than the rest of this README:

> **Be generous in row 1 and strict in your screening, not the other way
> round.** A term missing from the search cannot be recovered by any amount of
> careful screening, by a better database, or by machine learning. A term that
> is too broad only costs you screening time — which is exactly what the
> ranking and the stopping rule are for.

> **Test the query against papers you already know.** Collect 20–40 papers you
> are certain belong in the review (your gold set — step 5). Run the query. If
> it returns fewer than about two thirds of them, the query is the problem, not
> the database. Look at the titles of the ones it missed and you will usually
> see the missing term immediately.

That second check has caught, in practice, a query that searched only for the
narrow technical name of a mechanism while most of the relevant literature used
the broad popular name in the title. No database could have fixed it. Running
the gold set through the query before screening, rather than after, is the
cheapest quality control in the whole process.

Record, for the paper: each concept, its full term list, the field each term
was searched in (title/abstract/keywords vs full text), the publication window,
the language restriction, the document types, and the date you ran the search.
The app's PRISMA tab holds a free-text box for exactly this.

### 2. Search and download from each database

Three databases with complementary coverage, which is why using more than one
is worth the deduplication work:

| database | access | strengths | weaknesses |
|---|---|---|---|
| **OpenAlex** | free, open API, no account | the broadest coverage; abstracts, topics, institutional affiliations, open-access status, citation counts | metadata quality varies; some abstracts missing |
| **Web of Science** | institutional subscription | curated, consistent, good for journal-level filtering | narrower; export caps per batch; shouted metadata (`NETHERLANDS`) |
| **Lens.org** | free account | aggregates several sources including patents and books; good coverage of what the other two miss | abstracts often absent; export needs an account |

Run **the same conceptual query** in each, adapted to each one's syntax, and
save each download under a name that says where it came from. The merge script
recognises these by filename pattern:

| file name must contain | treated as |
|---|---|
| `openalex` … `complete` | OpenAlex |
| `web_of_science` | Web of Science |
| `lens` | Lens |
| `goldset` | papers you already knew about (added by hand) |

So `01-openalex_raw_complete.xlsx`, `web_of_science_2026-11-03.xls`,
`lens-export.csv` and `goldset.xlsx` all land in the right arm of the diagram
without any further configuration. Any other name is fine too — you can import
it in the app and choose its source from a menu.

> Export caps are real. Web of Science exports 1,000 records at a time; if your
> search returns more, download in batches and keep every file — the app reads
> a whole folder and importing the same file twice changes nothing.

**Harvesting OpenAlex through its API instead.** The web interface is fine, but
a script is reproducible and handles paging for you. If you use one, **keep it
in this repository** next to the merge script: the search is part of the
method, and a query that exists only in your shell history cannot be published.
A harvest script needs to do three things — page through the API with a
`mailto`, flatten the nested `authorships` into text columns so affiliations
survive, and reconstruct abstracts from `abstract_inverted_index` — and should
write a file whose name contains `openalex` and `complete` so the merge picks
it up.

### 3. Which fields to ask each database for

Ask for more than you think you need. A field you did not download cannot be
screened on, cannot be charted, and cannot be filtered — and re-running a
search months later does not reproduce the same result set.

**The fields that matter, and what each one buys you:**

| field | why you want it |
|---|---|
| **title** | required; the primary screening surface |
| **abstract** | required in practice; without it a record can only be screened on its title, and the ML ranker has almost nothing to work with |
| **authors** | duplicate detection, author-level bibliometrics |
| **year** | the publication-window rule, the year chart |
| **journal / source title** | venue bibliometrics, journal-level exclusions |
| **DOI** | *the* duplicate key, and what the app uses to look up anything missing |
| **document type** | the wrong-publication-type rule (editorials, errata, front matter) |
| **language** | the language rule |
| **institutional affiliations** | the region rule, and institution/country charts — **this is the field most often lost**, see the warning below |
| **topic / subject category / keywords** | automatic theme tagging, field-level charts |
| **times cited** | the old-and-never-cited rule, and completeness scoring when merging duplicates |
| **open-access status, URL** | finding the full text later |

Per database:

- **OpenAlex** — via the API, request
  `id,doi,title,display_name,publication_year,type,language,abstract_inverted_index,authorships,primary_location,open_access,cited_by_count,topics,keywords,referenced_works`.
  `authorships` is what carries institutions and countries; it is nested, so
  flatten it to a text column when you write the spreadsheet. Set a `mailto`
  (the app's Settings tab holds it) — it is the polite pool and it is faster.
- **Web of Science** — choose *Export → Excel → Full Record*, not *Author,
  Title, Source*. Full Record is what includes the abstract and the addresses.
- **Lens** — in the export dialog, tick the abstract and the author
  affiliation boxes explicitly; they are off by default.

> **The warning.** Affiliations are the field people drop, because they are
> nested and ugly and the cleaning script that flattens a spreadsheet tends to
> lose them. If you drop them you lose the ability to filter by region at all —
> and discovering that after screening 1,000 records means either screening
> them again or abandoning a planned inclusion criterion. Check that the
> affiliation column is populated **before** you start screening. The app's
> Sources tab shows you how many records have one.

### 4. Merge into one corpus that remembers its sources

Put every download in one folder with `10-merge_sources.py` and run:

```bash
python 10-merge_sources.py --dry-run     # see what it found, write nothing
python 10-merge_sources.py
```

It reads every file it recognises, normalises the columns, stamps each row with
its database and its filename, and writes **one** spreadsheet:
`corpus_all_sources.xlsx`.

What it does, and why:

- **Every row is kept.** Cross-database duplicates are *not* removed here. They
  are found and merged inside the app, where you can see them, override which
  one is kept, and have them counted correctly in the diagram. A merge step
  that silently dropped them would make the per-database PRISMA counts
  unreproducible.
- **Each record gets a stable id** derived from its DOI, or from its
  normalised title when it has no DOI — `OPENALEX-3f2a1b…`. Because the id
  comes from the content and not from a row number, re-running the merge,
  adding a database, or re-sorting the file does not detach your decisions from
  their records.
- **It refuses to write a partial corpus.** If any source file fails to read,
  it stops with a loud message and a non-zero exit rather than quietly giving
  you a corpus missing a database. (This is not hypothetical: a Web of Science
  file silently absent from a merge is very hard to notice later, because the
  corpus looks perfectly normal.)
- `--csv` also writes a CSV; `--skip-missing` allows a source to be absent on
  purpose.

### 5. Add papers you already know about

Papers you know belong in the review but that no database returned — from a
conference paper's reference list, a colleague's recommendation, citation
chasing — are **a separate identification arm in PRISMA 2020**, and the app
keeps them that way.

Two routes:

- Put them in a file whose name contains `goldset` before you merge. They
  arrive tagged as *Added by hand (known papers)* and appear in the right-hand
  arm of the diagram, never counted as a database hit.
- Or add them one at a time in the app: **Sources → Add a paper**, by DOI. It
  fetches the metadata for you from Crossref/OpenAlex, so you type one DOI
  rather than twelve fields.

Mark the ones that are your yardstick as **core** papers (`key_paper`). The
Sources tab then reports how many of your core set each database actually
returned — which is the number that tells you whether your query works, and a
genuinely useful sentence for a methods section.

### 6. Screen

Run the app and work through the Screen tab. Title and abstract first, full
text second. Every decision is written to the decisions CSV the moment you make
it, with the reason, your name as screener, and a timestamp.

### 7. Stop, and say why

Use the Stopping tab (below). Screening to the end of a long corpus is rarely
necessary and almost never done; what matters is that your stopping point is a
*stated rule* rather than the day you ran out of patience.

### 8. Report and export

PRISMA tab for the diagram and the counts; Export tab for the full decision
trail. Both come from the same records.

---

## Running the app

```bash
python BarcalaReviewApp.py                      # today's newest corpus
python BarcalaReviewApp.py corpus.xlsx          # a specific one
python BarcalaReviewApp.py --port 8800          # if 8766 is taken
python BarcalaReviewApp.py --no-browser         # don't open a browser
python BarcalaReviewApp.py --merge-decisions    # see below
```

**How it chooses a corpus when you don't name one.** This caused real grief
during development, so it is now deliberate and visible: it prefers the corpus
you last had open (remembered in `.review_corpus.json`), then the one with the
most *hand-made* decisions beside it, then the largest. Bulk-written decisions
are discounted when it counts work, so a file carrying 1,800 decisions written
in a single second does not outrank the file where you actually did the
screening. The startup output always prints which file it opened and how many
decisions it found. **Read that line.** If it is the wrong file, quit and name
the right one.

**`--merge-decisions`** brings decisions made against *another* corpus file in
the same folder across to this one, for records not yet screened here. This is
how you move to a newly merged corpus without re-screening. It is deliberately
not automatic: a decisions file may hold the result of a bulk operation rather
than record-by-record screening, and spreading that silently across every
corpus in a folder would be unrecoverable. When there is something to merge the
app tells you and leaves the choice to you.

---

## Tailoring it to your review

Nothing about any particular subject is baked into the program. On first run it
writes `review_config.json` beside itself, holding the defaults for every
review you start in that folder. Copy `review_config.example.json` to
`review_config.json` and edit, or change everything in the **Settings** tab.

Settings saved for a *particular corpus* live in files beside that corpus and
win over the config, so a second review in the same folder starts from the
config rather than inheriting the first review's themes.

### Themes (sub-topic tags)

Up to a dozen tags you apply to included papers, for the synthesis stage —
"this one is about disclosure", "this one is about investor behaviour". They
are yours to name:

```json
"themes": [
  {"id": "t1", "name": "Signalling and disclosure", "colour": "#0072B2"},
  {"id": "t2", "name": "Investor behaviour",         "colour": "#E69F00"}
]
```

The default colours are the Okabe–Ito palette, which stays distinguishable for
colour-blind readers and prints legibly in greyscale. If you replace them, keep
that property.

You can have the app **suggest** tags from the text, with `theme_patterns`:

```json
"theme_patterns": {
  "t1": ["signal", "disclos", "certification"],
  "t2": ["investor", "backer", "herding"]
}
```

A suggested tag is marked as suggested, not applied. You confirm it. The
distinction is kept in the export (`theme_source`), so the synthesis can
report how many tags were automatic.

### Inclusion and exclusion criteria

Your own lists, which become the dropdown of reasons when you exclude
something:

```json
"criteria": {
  "include": ["Matches the population of interest", "..."],
  "exclude": ["Off topic", "Wrong setting or population", "..."]
}
```

Keep the exclusion list short and mutually exclusive — these become the
"excluded, with reasons" breakdown in the PRISMA diagram, and a list of
twenty overlapping reasons makes an unreadable diagram. Eight or so is right.

### Automatic rules

Rules exclude records mechanically, with a stated reason, **reversibly**, and
are reported as a separate line in the PRISMA diagram so a reader can see that
a human did not look at them. Each can be switched on and off in the Rules tab,
which shows you how many records it would catch before you enable it.

| rule | default | what it does |
|---|---|---|
| `retracted` | **on** | drops retracted publications |
| `lang` | **on** | keeps only `allow` languages; `keep_unknown` keeps records whose language the source does not record, because unknown is not the same as not-English |
| `wrongtype` | **on** | drops the types in `block` — editorials, errata, letters, front matter |
| `noabstract` | off | drops records with no abstract. Off for a reason: some are real papers whose publisher withholds abstracts. Read the titles first |
| `old_uncited` | off | drops papers `years` old with no citations at all |
| `region` | off | drops records with no connection to your regions |

A rule's exclusion can always be overridden on an individual record, and the
override is recorded.

> **On rules as a substitute for screening.** A rule that would remove 85% of
> your corpus is not a rule, it is a different review. Before enabling a
> restrictive one, use the Rules tab's preview to check what it would do to the
> papers you have *already included* — if it would have excluded a tenth of
> them, it is too blunt, and often the reason is missing metadata rather than a
> genuine mismatch.

### Regions and keyword rules

`country_patterns` maps a region name to the strings that identify it in an
affiliation or in the text, so you are not limited to a built-in country list:

```json
"country_patterns": {
  "Europe": ["Netherlands", "Germany", "France", "Utrecht University", "..."],
  "Nordics": ["Denmark", "Sweden", "Norway", "Finland", "Iceland"]
}
```

Then `region`'s `params.regions` names which of those to accept, and `basis`
says where to look: `affiliation` (author institutions), `text` (title,
abstract, keywords), or `both`.

You can add your own keyword rules with a `kw:` id — `{"id": "kw:survey",
"params": {"terms": ["survey experiment"]}}` — which catch records whose text
matches, with a reason of your choosing.

---

## The tabs, one by one

| tab | what it is for |
|---|---|
| **Screen** | the work itself. Filter, sort, read, decide. Select several obviously-relevant titles and include them in one action; the rest you open and read the abstract |
| **Sources** | how many records each database contributed, how many have abstracts and affiliations, which of your core papers each database found, and the Add-a-paper-by-DOI box |
| **Duplicates** | every duplicate group, which record was kept, what it borrowed, and an override |
| **Rules** | switch rules on and off, with a live count of what each would catch |
| **Stopping** | the stopping rule and whether its conditions are met |
| **Dashboard** | the corpus and your progress through it |
| **Bibliometrics** | the literature itself — years, venues, countries, institutions, authors, fields |
| **PRISMA** | the flow diagram and every number in it, computed from your decisions |
| **Export** | the decision trail, as CSV or Excel |
| **Ask** | a question about your own corpus, answered from it |
| **Settings** | screener name, themes, criteria, contact address, API key |

---

## Machine-learning assistance

Screening in corpus order wastes most of your attention on irrelevant records.
Active learning reorders the queue: it learns from the decisions you have
already made and puts what it predicts you will include next.

The app uses **[ASReview](https://github.com/asreview/asreview)** when it is
installed (`pip install asreview`), with its TF-IDF feature extractor and your
choice of classifier; it falls back to a simple built-in ranker when ASReview
is absent, and the Dashboard says which engine produced the current ranking.

**Choosing a classifier.** Don't take a default on trust — the Dashboard's
*Compare models* button benchmarks every available classifier **on your own
decisions**: it hides a random 30% of what you have already judged, trains on
the rest, and reports what fraction of the hidden relevant papers each model
would have put in the top 10% of the queue, averaged over several random
splits. Random ordering scores 10% by definition, so that is your floor. On a
real corpus the spread between the best and worst classifier was nearly
two-fold, and the best was not the most elaborate one — so this is worth the
thirty seconds it takes.

Ranking needs something to learn from: roughly 10 includes and 10 excludes
before it is better than chance. Screen the first stretch in corpus order.

**The ranker never decides anything.** It changes the order in which you see
records, nothing else. No decision is ever written by a model, and there is no
setting that would let one be.

---

## The stopping rule

Reporting "we screened until we stopped finding things" is not a method. The
Stopping tab implements the **SAFE** procedure (Boetje & van de Schoot, 2024)
as four conditions, all of which must hold:

1. a **minimum share** of the corpus has been screened (default 10%);
2. a run of **consecutive irrelevant records** has been reached (default 50);
3. the records you *have* screened include a reasonable number of includes, so
   the run is not simply a corpus with nothing in it;
4. your **core papers** have been found — if the papers you already knew about
   are still unscreened, you have not screened enough yet.

The tab shows each condition, whether it is met, and the current numbers, and
lets you record the decision to stop with a timestamp and a note. That record
goes into the export and the PRISMA tab, so the published figure for "records
screened" has a documented reason behind it.

Both thresholds are configurable (`stopping.consecutive`,
`stopping.min_percent`). Set them before you start, not when you are tired.

---

## The two dashboards

They answer different questions and are deliberately separate.

**Dashboard — your corpus and your progress.** How many records, how many
screened, included, excluded and by which reason; the queue remaining; which
database each came from; how the includes distribute across your themes; the
year spread of what you are screening; whether the ML ranking is active and
which engine is producing it. This is the tab you look at while working.

**Bibliometrics — the literature itself.** Publication years, journals,
countries, institutions, authors, fields and subfields, document types,
languages, open-access status, and the most-cited records — for any slice you
choose: the whole corpus, only what you included, only one theme. This is the
tab you look at when writing, and it produces most of the descriptive
statistics a review's first results paragraph needs ("the 104 included studies
appeared in 61 journals, concentrated in…").

Counts there are case-folded and the mixed-case spelling is preferred, so
`NETHERLANDS` from Web of Science and `Netherlands` from OpenAlex are one
country and not two — a detail that matters because the uncorrected version
silently halves your largest categories.

---

## The Ask feature

A question box over your own corpus: *"how many included papers use
experimental designs?"*, *"which included studies are about disclosure but not
tagged t1?"*, *"summarise what the included papers say about platform risk"*.
It answers from the records and the decisions in front of it — titles,
abstracts, your tags, your reasons — not from the model's own knowledge of the
literature, and it says when the corpus does not contain the answer.

It needs an [Anthropic API key](https://console.anthropic.com/). Two ways to
provide one, and **only** these two:

```bash
export ANTHROPIC_API_KEY="sk-ant-..."     # before starting the app
```

or paste it into the Settings tab, where it is held **in the server's memory
for that session only**.

> **The key is never written to disk.** Not to the config, not to a dotfile,
> not to a log. This is on purpose: a review folder normally lives in OneDrive
> or Dropbox, and a key written beside the corpus is a key uploaded to a cloud
> provider and shared with everyone the folder is shared with. The same applies
> to any other API key you use in this pipeline — put it in an environment
> variable, never in a script in the synced folder.

What is sent, when you ask a question: the titles, abstracts and your own
annotations for the records relevant to that question. Nothing else, and
nothing at all until you ask. If your corpus contains anything you are not
free to send to a third party, do not use this tab; everything else in the app
works without it.

**The model is forbidden, in its instructions, from making inclusion or
exclusion decisions**, and will decline to. Screening judgements are yours.
Use it to interrogate your corpus and to draft descriptive text, never to
decide.

---

## Export

The Export tab writes one row per record — all of them, or only the included
ones — as CSV or Excel. It is the audit trail, and it is what you archive and
what a reviewer or a replicator reads.

The columns, and why each is there:

- **identity** `id`, `doi`, `title`, `authors`, `year`, `journal`, `type`,
  `language`, `url`, `apa_citation`, `abstract`
- **provenance** `source` (which database), `source_file` (which download),
  `key_paper` (core/background), `abstract_source` (whether you typed the
  abstract in by hand)
- **classification** `topic`, `subfield`, `field`, `institutions`, `country`,
  `country_source`, `themes`, `theme_source` (yours or suggested),
  `cited_by`, `instrument_type`
- **decisions** `ta_decision`, `ta_reason`, `ta_source` (screened by hand,
  title-only, by a rule, in bulk), `ft_decision`, `ft_reason`, `notes`,
  `screener`, `updated_at`
- **duplicates** `dup_status`, `duplicate_of` (which record it was folded
  into), `merged_fields` (which fields were borrowed from which source)

`ta_source` and `merged_fields` are the two columns reviewers ask about, and
having them means you can answer without reconstructing anything.

---

## Where everything is stored

Everything is a plain file in **the folder your corpus is in**. There is no
database and no hidden state; you can read, back up, version and inspect all of
it with a text editor.

Given a corpus called `corpus_all_sources.xlsx`:

| file | what it holds |
|---|---|
| `corpus_all_sources.xlsx` | **the corpus**. Never modified by the app |
| `corpus_all_sources_decisions.csv` | **your work** — every judgement, reason, tag, note and typed-in correction, with screener and timestamp. This is the file to back up |
| `corpus_all_sources_extra_records.csv` | papers you added by hand in the app |
| `corpus_all_sources_import_log.json` | what was imported, when, from which file |
| `corpus_all_sources_labels.json` | this corpus's themes |
| `corpus_all_sources_criteria.json` | this corpus's inclusion/exclusion criteria |
| `corpus_all_sources_rules.json` | which rules are on, and their parameters |
| `corpus_all_sources_stop.json` | the stopping thresholds and the recorded decision to stop |
| `corpus_all_sources_decisions_<tag>_<timestamp>.csv` | automatic backups, written before anything rewrites decisions in bulk |

And beside the app itself:

| file | what it holds |
|---|---|
| `review_config.json` | your defaults for new reviews: screener name, contact address, themes, criteria, rules, region and theme patterns, stopping thresholds |
| `.review_corpus.json` | which corpus you last had open |

**Corrections are kept apart from the data.** When you fix a mangled title or
paste in a missing abstract, it goes into the decisions file, not into the
corpus — so the harvest stays exactly as it arrived from the database, and
re-running the merge can never silently drop a fix you made. The export shows
the corrected value and `abstract_source` says it was yours.

**What to back up:** the corpus and `*_decisions.csv`. Everything else can be
regenerated. **What to keep out of version control:** both of those, plus your
filled-in `review_config.json` — the supplied `.gitignore` already excludes
them, so your corpus and your screening never end up in a public repository by
accident.

---

## Reporting this method in a paper

`docs/METHODS.md` holds a methods passage you can adapt, with the references
formatted and the three citations you need (PRISMA 2020, ASReview, SAFE). The
things worth stating explicitly, because reviewers ask:

- the search string per database, the fields searched, and the date run;
- that records from all sources were pooled **before** deduplication, with
  provenance retained, and that duplicates were identified within the pooled
  set;
- which automatic rules were enabled, with their reasons and counts, reported
  separately from records a human screened;
- that the screening order was prioritised by active learning, which model, and
  that **no decision was made by a model**;
- the stopping rule, its thresholds, and the figures when you stopped;
- that the PRISMA counts were computed from the decision records, with the
  export available as supplementary material.

That last point is the one that distinguishes this from a typed-in diagram, and
it is worth a sentence.

---

## What this tool will not do

Stated plainly, because knowing a tool's limits is part of using it honestly:

- **It does not support two independent screeners.** There is one `screener`
  field and no reconciliation, no blinding and no agreement statistic. If your
  protocol requires dual independent screening with Cohen's κ, this tool does
  not implement it — screen in two copies and compare the exported decisions,
  or use a tool built for it.
- **It does not make decisions.** No rule, ranker or model writes an inclusion.
  Rules exclude mechanically and reversibly, and say so in the diagram.
- **It does not fix your search.** See step 1. This is the real limit on what
  any review can find, and no feature here changes it.
- **It does not retrieve full texts.** It records your full-text decisions; you
  find the PDFs.
- **It is a local single-user tool.** No accounts, no sharing, no concurrent
  editing. Two people editing one decisions file at once will lose work.

---

## What is in this repository

| file | what it is for |
|---|---|
| `BarcalaReviewApp.py` | the application. One file, standard library only |
| `10-merge_sources.py` | pools the database downloads into one corpus with provenance |
| `review_config.example.json` | copy to `review_config.json` and edit; your defaults |
| `requirements.txt` | the three optional extras |
| `README.md` | this file: the method, and how to run and tailor the app |
| `docs/METHODS.md` | a methods passage for a paper, with the three references |
| `docs/TROUBLESHOOTING.md` | when a count looks wrong, or work seems missing |
| `docs/AI_USAGE.md` | how the software was written, and what it does with AI at runtime |
| `SECURITY.md` | the local server, and why API keys stay out of the review folder |
| `CONTRIBUTING.md` | how to contribute, and the design commitments |
| `CODE_OF_CONDUCT.md` | Contributor Covenant 2.1 |
| `CHANGELOG.md` | what changed, including defects found and fixed |
| `CITATION.cff` | how to cite this, machine-readable |
| `LICENSE` | Apache 2.0 |
| `tests/` | the test suite and its synthetic 12-record fixture |
| `.github/workflows/` | CI: tests on Python 3.10–3.12, plus a committed-secret check |
| `.gitignore` | keeps your corpus, decisions and config out of git |

What is deliberately **not** here: your corpus, your decisions file, your
filled-in config, and any API key. See *Where everything is stored*.

---

## Tests

```bash
pip install pytest openpyxl
python -m pytest -q
```

33 tests over a 12-record synthetic corpus with a known number of duplicates,
languages, years and citations, so every expected figure can be worked out by
hand from the fixture rather than copied from a previous run. They cover
loading, deduplication and field merging, record ids, rules, the PRISMA
arithmetic, decision recording, export provenance, configuration neutrality,
bibliometrics and core-paper tracking. CI runs them on Python 3.10–3.12, and
also fails the build if an API key or a personal address is ever committed.

---

## Citing

See `CITATION.cff`. If you use the stopping rule or the ranker, cite those too:

- Page, M. J., et al. (2021). The PRISMA 2020 statement: an updated guideline
  for reporting systematic reviews. *BMJ*, 372, n71.
- van de Schoot, R., de Bruin, J., Schram, R., et al. (2021). An open source
  machine learning framework for efficient and transparent systematic
  screening. *Nature Machine Intelligence*, 3, 125–133.
- Boetje, J., & van de Schoot, R. (2024). The SAFE procedure: a practical
  stopping heuristic for active learning-based screening in systematic reviews
  and meta-analyses. *Systematic Reviews*, 13, 81.

---

## Licence

Apache 2.0 — see `LICENSE`. Contributions welcome; see `CONTRIBUTING.md`.
