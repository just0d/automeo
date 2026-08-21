# AutoMEO

Automated monthly macroeconomic reporting for Mongolia. One command turns the
official statistics — Customs, the National Statistics Office and the Bank of
Mongolia — into a deck-ready Excel workbook: **18 sheets, ~1,030 live
formulas, 14 ready-made charts**, bilingual labels, every figure traceable to
its source.

> 📘 **The Operations Manual is distributed separately** (`AutoMEO_Operations_Manual.docx`).
> It covers how to run it, where every number comes from, and how to repair it
> when a source changes. Read that first — this page is only the quick start.

---

## Quick start

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp seed/bop_manual.xlsx raw_data/                   # starter file; refresh monthly

python run_all.py                  # newest published month, detected automatically
python run_all.py 6 2026           # force a month
python run_all.py --offline        # rebuild from cache, no network
```

Python 3.10+. Three dependencies: `requests`, `pandas`, `openpyxl`.

## The monthly run

```bash
python run_all.py          # pauses once for the Balance of Payments download
python deck_digest.py      # writes deck_digest.md for the presentation step
python check_deck.py <deck>.pptx    # after the deck is built
```

Then fill the two yellow cells on the `Budget` sheet — the annual revenue plan
and the annual coal export plan, both set once a year.

The July-approved presentation structure is the canonical deck contract: 21
slides, with the sector section on slides 17–21. Use
[`docs/DECK_PROMPT.md`](docs/DECK_PROMPT.md) with the previous approved deck as
the base file; it maps every slide to its workbook source and preserves the
July chart/layout pattern.

**Read `Data_Vintage`, the first sheet, before the numbers.** Sources publish
on different calendars, so a report for month M is always a mix. That sheet
states which month each source actually carried; every sheet repeats its own
in a banner on row 4, and no lagged figure is ever presented as current.

## When something breaks

Every identifier the pipeline discovers can be pinned by hand — no code
editing:

```bash
python overrides.py --init     # commented template
python overrides.py            # checks it parses, prints what it pins
```

```json
{ "mongolbank": { "report_ids": { "2026-06": { "bulletin": 5217 } } } }
```

Chapter 5 of the Operations Manual ranks what is most likely to break, gives
the log line to look for, and walks through capturing a replacement identifier
from browser DevTools.

## Layout

```
run_all.py              orchestrator
raw_ingestor.py         downloads and validates every source
data_processor.py       extracts, cross-checks, computes
excel_builder.py        builds the workbook and its charts
overrides.py            hand-set identifiers
deck_digest.py          workbook -> plain text for the deck chat
check_deck.py           checks a finished .pptx
overrides.example.json  template for overrides.json
seed/                   starter Balance of Payments export
raw_data/               downloaded sources (git-ignored)
```

## Licence

MIT — see [LICENSE](LICENSE). The data is published by the National Statistics
Office of Mongolia, the Bank of Mongolia and the Mongolian Customs General
Administration and remains subject to their terms; this project only fetches
and arranges it.
