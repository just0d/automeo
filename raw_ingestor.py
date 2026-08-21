#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Downloads all raw source files into ./raw_files (Customs, NSO, Mongolbank
cards, Mongolbank monthly reports incl. the statistical bulletin) with
run-time id resolution, walk-back, retries, validation and a run manifest.
See the Data Dictionary document for source-by-source details.

Usage: python raw_ingestor.py [month] [year]
"""

import io
import json
import re
import sys
import time
import logging
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import requests
import urllib3

import overrides as _ov
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# CONFIGURATION
BASE_DIR = Path(__file__).resolve().parent
RAW_DATA_DIR = BASE_DIR / "raw_files"
MANIFEST_PATH = RAW_DATA_DIR / "manifest.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9,mn;q=0.8",
}
REQUEST_TIMEOUT = 45          # seconds per request
RETRY_TOTAL = 4               # total retry attempts per request
RETRY_BACKOFF = 2             # 2s, 4s, 8s ... between retries

# SOURCE 1: Customs (gaali.mn)
# The list API returns, for a given year, EVERY monthly statistical bulletin
# with a direct file.url -> no fragile 2-step dance needed (detail endpoint is
# kept as a fallback only).
# Hand-set ids, if any. Loaded once here so every constant below can be
# replaced without editing code -- see overrides.py and
# docs/TROUBLESHOOTING.md. An absent file leaves everything as-is.
OVERRIDES = _ov.load()
_ovget = lambda *p, default=None: _ov.section(OVERRIDES, *p,
                                              default=default)

CUSTOMS_LIST_API = _ovget("customs", "list_api",
    default="https://www.gaali.mn/shared-api/api/statistic-news")

# SOURCE 2 (RETIRED 2026-07): Ministry of Economy PDF
# The MED monthly PDF was only ever a reference document. Inflation, reserves
# and budget execution now come from the Mongolbank STATISTICAL BULLETIN
# Excel (stat.mongolbank.mn/bulletin), which is more official and detailed.
# The bulletin is just another monthly survey report, so it rides the same
# run-time id resolution as the banking reports (see MONGOLBANK_SURVEYS).

# SOURCE 3: NSO (1212.mn)
# Generic table-view endpoint: POST with {} returns the FULL data matrix of a
# table in json-stat2. Add more tables to NSO_TABLES as they are identified
# (e.g. CPI monthly, Household Socio-Economic Survey) -- the fetcher is generic.
NSO_API_TEMPLATE = (
    "https://www.1212.mn/api/table-view"
    "?lng=mn&sector={sector}&subsector={subsector}&id={table_id}"
)
NSO_TABLES = {
    # save_name          : (table_id,                sector,                    subsector[, extra query])
    "nso_gdp_raw":        ("DT_NSO_0500_004V1.px",  "Economy%2C%20environment", "National%20Accounts"),
    # Monthly Indicator of Economic Growth (MIEG), cumulative YoY % +
    # sector contributions; table id captured from DevTools 2026-07
    "nso_mieg_raw":       ("DT_NSO_0500_001V5.px",  "Economy%2C%20environment", "National%20Accounts"),

    # REAL SECTOR (deck slides 16-21). Captured from DevTools 2026-08.
    # Trade sales are a MONTHLY level; hotel and food income are already
    # published cumulative ('өссөн дүнгээр'); construction is quarterly
    # cumulative; industry sales are a monthly level by sub-sector. The
    # processor knows which is which -- see REAL_SECTOR_TABLES there.
    "nso_trade_sales":    ("DT_NSO_1600_002V1_month.px", "Industry%2C%20service",
                           "Trade%2C%20hotel%20and%20restaurant"),
    "nso_hotel_income":   ("DT_NSO_1602_003V1_M.px", "Industry%2C%20service",
                           "Trade%2C%20hotel%20and%20restaurant",
                           "&subtables=INCOME%20OF%20THE%20HOTEL%20SECTOR,"
                           "%20by%20Aimag%20and%20Capital%20City"),
    "nso_food_income":    ("DT_NSO_1602_004V1_M.px", "Industry%2C%20service",
                           "Trade%2C%20hotel%20and%20restaurant",
                           "&subtables=INCOME%20OF%20THE%20FOOD%20SECTOR,"
                           "%20by%20Aimag%20and%20Capital%20City"),
    "nso_construction":   ("DT_NSO_0902_002V1.px",  "Industry%2C%20service",
                           "Construction"),
    "nso_industry_sales": ("DT_NSO_1100_032V1.px",  "Industry%2C%20service",
                           "Industry"),
    # Transport (deck slide 21, alongside construction). Captured 2026-08:
    # national-level key indicators, quarterly, dims [Үзүүлэлт, Улирал].
    "nso_transport":      ("DT_NSO_1200_012V4_y.px", "Industry%2C%20service",
                           "Transportation",
                           "&subtables=KEY%20INDICATORS%20OF%20TRANSPORTATION"
                           "%20SECTOR,%20by%20national%20level"),
    # WEEKLY food and petrol prices in Ulaanbaatar (deck slide 9). This is
    # the only weekly source in the pipeline and the only one that can be
    # fresher than the reported month, so the processor clamps it like the
    # rest. It replaces what used to be two hand-typed yellow cells.
    # НИЙСЛЭЛИЙН ХҮНСНИЙ ГОЛ НЭР БОЛОН БЕНЗИН ТҮЛШНИЙ 7 ХОНОГИЙН ҮНИЙН МЭДЭЭ
    # A SECOND weekly table, by aimag, covering the same survey. NSO keeps
    # both, and they do not always update together -- in August 2026 the
    # capital-city table stalled in July while this one was current. The
    # processor reads whichever is fresher, so one going quiet no longer
    # freezes the price slide.
    "nso_weekly_prices_aimag": ("DT_NSO_0300_010V5.px",
                                "Economy%2C%20environment",
                                "Consumer%20Price%20Index"),
    "nso_weekly_prices":  ("DT_NSO_0600_001V4.px",  "Economy%2C%20environment",
                           "Consumer%20Price%20Index"),
}
# Hand-set NSO table ids. 1212.mn occasionally republishes a table under
# a new id; this replaces one without touching the code.
for _n, _spec in (_ovget("nso", "tables", default={}) or {}).items():
    _cur = NSO_TABLES.get(_n)
    if not isinstance(_spec, dict):
        continue
    _base = list(_cur) if _cur else ["", "", "", ""]
    while len(_base) < 4:
        _base.append("")
    for _i, _k in enumerate(("id", "sector", "subsector", "extra")):
        if _spec.get(_k) is not None:
            _base[_i] = _spec[_k]
    NSO_TABLES[_n] = tuple(_base)

NSO_POST_PAYLOAD = {"query": [], "response": {"format": "json-stat2"}}
# Some tables refuse the blank query; per-table explicit payloads are tried
# in order until one returns a json-stat2 body with a 'value' array. The
# MIEG selections mirror the request captured from DevTools (dimension
# codes: 'Статистик үзүүлэлт', 'Эдийн засгийн салбар', 'Он').
NSO_TABLE_PAYLOADS = {
    "nso_mieg_raw": [
        NSO_POST_PAYLOAD,
        {"query": [
            {"code": "Статистик үзүүлэлт",
             "selection": {"filter": "item", "values": ["0", "1", "2"]}},
            {"code": "Эдийн засгийн салбар",
             "selection": {"filter": "item",
                           "values": ["0", "1", "2", "3", "4", "5"]}},
            {"code": "Он",
             "selection": {"filter": "top", "values": ["36"]}}],
         "response": {"format": "json-stat2"}},
        {"query": [
            {"code": "Статистик үзүүлэлт",
             "selection": {"filter": "item", "values": ["0", "1", "2"]}},
            {"code": "Эдийн засгийн салбар",
             "selection": {"filter": "item",
                           "values": ["0", "1", "2", "3", "4", "5"]}},
            {"code": "Он",
             "selection": {"filter": "all", "values": ["*"]}}],
         "response": {"format": "json-stat2"}},
    ],
}


def _sel(**codes):
    """
    Explicit per-dimension selection, mirroring the captured requests.

    These tables are HIERARCHICAL: 'Бүс' carries Улсын дүн, the five regions
    AND all 21 aimags, and the regions already contain the aimags. A blank
    full-matrix query therefore returns parents and children together, and
    anything that sums the members double-counts (observed: contributions
    came to exactly 2x the headline). Selecting items 0-5 returns the total
    plus the five regions only, which is what the analyst's own workbooks
    use. The blank query is kept as a LAST resort, never first.
    """
    # A dimension listed with _ALL is OMITTED from the query: PxWeb returns
    # every value for any dimension not mentioned. Spelling it as
    # {"filter":"item","values":["*"]} is invalid and the server answers 500,
    # which is what happened on the first live run -- every table fell
    # through to the blank query and came back with the full hierarchy.
    explicit = [{"code": c, "selection": {"filter": "item", "values": v}}
                for c, v in codes.items() if v != _ALL]
    return [{"query": explicit, "response": {"format": "json-stat2"}},
            NSO_POST_PAYLOAD]


_ALL = ["*"]                                    # 'everything' -> omit the dim
REGION_ITEMS = ["0", "1", "2", "3", "4", "5"]   # Улсын дүн + 5 бүс


# The captured requests select all six 'Бүс' items plus every period. The
# blank full-matrix query is tried first because it is what the GDP and MIEG
# tables accept; the explicit selections mirror DevTools exactly and are the
# fallback if a table refuses it.
_ALL = ["*"]
NSO_TABLE_PAYLOADS.update({
    # Бүс: items 0-5 only -- total + the five regions, never the aimags
    "nso_trade_sales":    _sel(**{"Бүс": REGION_ITEMS, "Сар": _ALL}),
    "nso_hotel_income":   _sel(**{"Бүс": REGION_ITEMS, "Сар": _ALL}),
    "nso_food_income":    _sel(**{"Бүс": REGION_ITEMS, "Сар": _ALL}),
    # construction: regions collapsed to the national total by the processor
    "nso_construction":   _sel(**{"Үзүүлэлт": _ALL, "Бүс": REGION_ITEMS,
                                  "Улирал": _ALL}),
    "nso_industry_sales": _sel(**{"Дэд салбар": _ALL, "Сар": _ALL}),
    # Transport, captured 2026-08: national-level key indicators, quarterly.
    "nso_transport":      _sel(**{"Үзүүлэлт": _ALL, "Улирал": _ALL}),
    # Weekly prices: both dimensions in full -- 31 products x every week.
    "nso_weekly_prices":  _sel(**{"Бүтээгдэхүүн": _ALL, "Хугацаа": _ALL}),
    # the aimag table carries a third dimension, Бүс
    "nso_weekly_prices_aimag": _sel(**{"Бүтээгдэхүүн": _ALL, "Бүс": _ALL,
                                       "Хугацаа": _ALL}),
})

# SOURCE 4: Mongolbank (stat.mongolbank.mn)
MONGOLBANK_API_URL = _ovget("mongolbank", "api", "main", default="https://stat.mongolbank.mn/api/report/main")
# Section ids to sweep. 20 = external sector (verified). The rest of the range
# is probed cheaply; empty/erroring ids are skipped, non-empty ones are saved.
# This is how policy-rate / money / banking cards get discovered automatically.
MONGOLBANK_PARENT_IDS = list(range(1, 61))
MONGOLBANK_KNOWN_SECTIONS = {20: "external"}   # nicer filenames for known ids

MONGOLBANK_REPORT_DATA_API = _ovget("mongolbank", "api", "data", default="https://stat.mongolbank.mn/api/survey/data")
# Monthly Excel reports use a TWO-STEP flow (captured from DevTools):
#   Step 1: POST /api/survey/sublist?lang=mn&surveyid=<SID>
#           -> lists the monthly files ('2026 5-р сар', ...) with their
#              DYNAMIC ids (e.g. April id=5038, May id=5116).
#   Step 2: POST /api/survey/data?lang=mn&id=<dynamic id>&type=xlsx
# So ids are resolved at RUN TIME from the target month -- nothing goes
# stale when a new month is published.
MONGOLBANK_SURVEY_LIST_API = _ovget("mongolbank", "api", "sublist", default="https://stat.mongolbank.mn/api/survey/sublist")
MONGOLBANK_SURVEYS = {
    # report -> surveyid of its download page. 11 captured from DevTools for
    # the bank loan report. For a missing surveyid, open that report's page
    # on stat.mongolbank.mn and read the id from the page URL / the sublist
    # request in DevTools; auto-discovery below also tries to find it.
    "bank_loan_report":      {"surveyid": 11,
                              "keywords": ["зээлийн тайлан"]},
    "banking_balance_sheet": {"surveyid": None,
                              "keywords": ["нэгдсэн тайлан тэнцэл"]},
    "banking_survey":        {"surveyid": None,
                              "keywords": ["хадгаламжийн байгууллагын тойм"]},
    "bulletin":              {"surveyid": 1,   # captured from DevTools:
                              # sublist?surveyid=1 fires on page load of
                              # stat.mongolbank.mn/bulletin
                              "keywords": ["сарын бюллетень"]},
    # NOTE: Balance of Payments is NOT here by design -- its portal page is
    # a dynamic dashboard, so BoP is handled by the human-in-the-loop step
    # ingest_bop_manual() below (raw_files/bop_manual.xlsx).
}
for _n, _sid in (_ovget("mongolbank", "survey_ids", default={}) or {}).items():
    if _n in MONGOLBANK_SURVEYS:
        MONGOLBANK_SURVEYS[_n]["surveyid"] = _sid
    else:
        log_names = ", ".join(MONGOLBANK_SURVEYS)
        print(f"overrides: unknown Mongolbank report {_n!r} in "
              f"survey_ids; known reports are {log_names}")
MONGOLBANK_SURVEYID_SCAN = range(1, 61)   # sweep for auto-discovery
MONGOLBANK_STATIC_FALLBACK = {            # last-resort frozen snapshot ids
    "banking_survey": 4515,
    "banking_balance_sheet": 5098,
    "bulletin": 5356,                     # captured from DevTools 2026-08-13
    #          previous known-good ids: 5217 (2026-06), 5115 (2026-05)
}
MONGOLBANK_STATIC_FALLBACK.update(
    _ovget("mongolbank", "static_fallback", default={}) or {})
MB_REPORT_WALKBACK_MONTHS = 2

# SOURCE 4c: the portal's interactive balance-sheet page
# (stat.mongolbank.mn/finance -> Банкны салбарын тайлан тэнцэл). Captured
# from DevTools 2026-07. The downloadable report carries only the few date
# columns Mongolbank chooses to publish; this endpoint returns one column per
# month, which is what makes the deposit series usable.
#   interval "3" = monthly; id 86 / parentId 10 identify the report;
#   the indicator ids are the 34 rows of the deposit tree, in tree order.
# The date range is built from the target period at run time, so nothing here
# goes stale. If the indicator ids are ever renumbered the response will be
# short or empty and the processor falls back to the report's own columns --
# and the deposit tree's parent/child sum check would catch a wrong set.
MONGOLBANK_INDICATOR_API = _ovget("mongolbank", "api", "indicator", default="https://stat.mongolbank.mn/api/indicator/data")
DEPOSIT_REPORT_ID = str(_ovget("deposits", "report_id", default="86"))
DEPOSIT_PARENT_ID = int(_ovget("deposits", "parent_id", default=10))
DEPOSIT_INDICATOR_IDS = [str(i) for i in
                         (_ovget("deposits", "indicator_ids")
                          or range(60564, 60598))]        # 34 by default
DEPOSIT_HISTORY_YEARS = 1        # start of (target year - 1) .. target month
# The loan report is a one-month flow statement, so a back-series is what
# makes loan GROWTH measurable. 13 months = the target month plus a full year
# behind it, which is the minimum for a year-on-year comparison.
LOAN_HISTORY_MONTHS = 13
# How far either side of a resolved id to hunt for the bulletin when its own
# listing lags. Observed gaps: 1 (2026-05) and 3-20 (2026-06).
BULLETIN_PROBE_RADIUS = 25
# Each probe downloads ~1.9 MB, so the search is capped. 40 attempts is about
# 30 s and covers every gap observed so far.
BULLETIN_PROBE_BUDGET = 40

# NSO CPI PDF: automation DROPPED by design (2026-07)
# Slide 7 (meat & fuel prices) stays a MANUAL slide: the source is a heavy
# presentation PDF and automating its parsing was judged not worth the
# complexity. The Coverage sheet flags it accordingly.

# Balance of Payments: human-in-the-loop
# The BoP page on stat.mongolbank.mn is a dynamic dashboard (checkboxes +
# date pickers), not a file list. The operator downloads the Excel once per
# run and saves it as raw_files/bop_manual.xlsx; ingest_bop_manual() prompts
# for it interactively when missing.
BOP_MANUAL_FILE = "bop_manual.xlsx"

# LOGGING & SESSION
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("raw_ingestor")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def make_session() -> requests.Session:
    """Session with automatic retries/backoff on connection errors and 5xx."""
    s = requests.Session()
    retry = Retry(
        total=RETRY_TOTAL,
        backoff_factor=RETRY_BACKOFF,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST", "HEAD"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update(HEADERS)
    return s


SESSION = make_session()
MANIFEST = {"run_at": None, "target": {}, "sources": {}}


def _record(source: str, status: str, detail: str = "", files=None) -> None:
    MANIFEST["sources"].setdefault(source, {"files": [], "events": []})
    MANIFEST["sources"][source]["status"] = status
    MANIFEST["sources"][source]["events"].append(
        {"at": datetime.now().isoformat(timespec="seconds"),
         "status": status, "detail": detail})
    if files:
        MANIFEST["sources"][source]["files"] = sorted(
            set(MANIFEST["sources"][source]["files"]) | set(files))


# HELPERS
def _ensure_raw_dir() -> None:
    RAW_DATA_DIR.mkdir(parents=True, exist_ok=True)
    log.info("Raw data directory ready: %s", RAW_DATA_DIR)


def _customs_workbook_ok(content: bytes) -> bool:
    """
    True only for the monthly statistics workbook.

    gaali.mn occasionally files a different publication under a month slot --
    2025-12 served a 4-sheet historical series ('Импорт 1995-2025', ...)
    instead of the ~23-sheet monthly bulletin. Those files parse fine as xlsx
    but have none of the sheets the processor reads, so accepting one costs a
    month of the trade and border-price series. The monthly workbook is
    identified by its numeric sheet names ('1' = totals, '3' = HS sections).
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        names = [str(s).strip() for s in wb.sheetnames]
        wb.close()
    except Exception as exc:
        log.warning("customs workbook unreadable: %s", exc)
        return False

    # Match sheet names the same way the processor does -- exact, or the
    # number followed by a caption. Customs renamed '1' to
    # '1 Нийт бараа эргэлт' in the 2025-12 re-publication; demanding an exact
    # name here rejected a workbook the processor could read perfectly, and
    # silently cost a month of history. The validator must never be stricter
    # than the parser it is protecting.
    def has(num):
        return any(n == num or n.startswith(num + " ") for n in names)

    missing = [n for n in ("1", "3") if not has(n)]
    if missing:
        log.warning("customs workbook rejected: no sheet named %s "
                    "(has %s)", " or ".join(missing), names[:6])
        return False
    return True


def _validate(content: bytes, kind: str) -> bool:
    """Reject error pages saved as data files."""
    if kind == "xlsx":
        return content[:2] == b"PK"
    if kind == "pdf":
        return content[:5] == b"%PDF-"
    if kind == "json":
        try:
            json.loads(content.decode("utf-8"))
            return True
        except Exception:
            return False
    return True


def _save(filename: str, content: bytes, kind: str, probing=False) -> bool:
    """probing=True when this is one of several candidates being tried: a
    rejection is then the mechanism working, not a failure, and logging it at
    ERROR makes a healthy run look broken to whoever reads the log."""
    if not _validate(content, kind):
        log.log(logging.INFO if probing else logging.ERROR,
                "%s for %s (not a real %s) -> not saved.",
                "candidate rejected" if probing else "VALIDATION FAILED",
                filename, kind)
        return False
    path = RAW_DATA_DIR / filename
    path.write_bytes(content)
    log.info("Saved %s (%d bytes)", filename, len(content))
    return True


def _get(url, **kw):
    kw.setdefault("timeout", REQUEST_TIMEOUT)
    return SESSION.get(url, **kw)


def _post(url, **kw):
    kw.setdefault("timeout", REQUEST_TIMEOUT)
    return SESSION.post(url, **kw)


# TARGET PERIOD (auto-discovery)
_CUSTOMS_LIST_CACHE = {}


def customs_list_for_year(year: int):
    """
    Return ALL monthly customs bulletins for a year.

    The plain ?year= query paginates at 10 records and IGNORES every page/
    size parameter (verified live), which silently hid Nov/Dec. The ?month=
    filter, however, works -- so after the base year query, every month not
    present in the first page is fetched individually. Results are cached
    per process (auto-detection and ingestion share one set of calls).
    """
    if year in _CUSTOMS_LIST_CACHE:
        return _CUSTOMS_LIST_CACHE[year]
    records, have_months = [], set()

    def add(lst):
        for rec in lst:
            try:
                m = int(rec.get("month"))
            except (TypeError, ValueError):
                continue
            if 1 <= m <= 12 and m not in have_months:
                have_months.add(m)
                records.append(rec)

    try:
        r = _get(CUSTOMS_LIST_API, params={"year": year})
        r.raise_for_status()
        add((r.json() or {}).get("data", {}).get("list", []) or [])
    except Exception as exc:
        log.warning("Customs year query failed for %s: %s", year, exc)
    # top-up: fetch months hidden by the 10-record page limit
    for m in range(1, 13):
        if m in have_months:
            continue
        try:
            r = _get(CUSTOMS_LIST_API, params={"year": year, "month": m})
            r.raise_for_status()
            add((r.json() or {}).get("data", {}).get("list", []) or [])
        except Exception:
            continue
    _CUSTOMS_LIST_CACHE[year] = records
    return records


def get_target_period() -> tuple:
    """
    (month, year) to run for.
      * CLI override:  raw_ingestor.py <month> [year]
      * Otherwise AUTO: newest month published on the Customs list API
        (current year; falls back to the previous year around January).
    """
    if len(sys.argv) > 1 and sys.argv[1].strip().isdigit():
        month = int(sys.argv[1].strip())
        year = (int(sys.argv[2].strip())
                if len(sys.argv) > 2 and sys.argv[2].strip().isdigit()
                else datetime.now().year)
        if 1 <= month <= 12:
            log.info("Target period (CLI override): %d-%02d", year, month)
            return month, year
        log.warning("Invalid CLI month %r -> switching to auto-detect.", month)

    now = datetime.now()
    for year in (now.year, now.year - 1):
        records = customs_list_for_year(year)
        months = [int(r.get("month") or 0) for r in records
                  if str(r.get("month") or "").isdigit() or isinstance(r.get("month"), int)]
        months = [m for m in months if 1 <= m <= 12]
        if months:
            month = max(months)
            log.info("AUTO-DETECTED newest published period: %d-%02d "
                     "(from Customs list API)", year, month)
            return month, year

    # Last resort: previous calendar month.
    month = now.month - 1 or 12
    year = now.year if now.month > 1 else now.year - 1
    log.warning("Could not auto-detect -> defaulting to previous month %d-%02d",
                year, month)
    return month, year


# SOURCE 1: CUSTOMS -- all monthly workbooks, current + previous year
def _customs_record_file_url(rec: dict) -> str:
    """file.url straight from the list record; falls back to the detail API."""
    url = ((rec.get("file") or {}).get("url") or "").strip()
    if url:
        return url
    rec_id = rec.get("id")
    if not rec_id:
        return ""
    try:
        r = _get(f"{CUSTOMS_LIST_API}/{rec_id}")
        r.raise_for_status()
        return ((r.json() or {}).get("data", {}).get("file") or {}).get("url", "")
    except Exception as exc:
        log.warning("Customs detail fallback failed for id=%s: %s", rec_id, exc)
        return ""


def ingest_customs(target_month: int, target_year: int) -> None:
    """
    Download EVERY monthly bulletin workbook for the target year and the
    previous year (skipping files already cached in raw_files). The per-month
    history is what powers the monthly trade series and border unit prices.
    """
    log.info("---- [1/4] CUSTOMS (gaali.mn) ----")
    saved, wanted = [], 0
    try:
        for year in (target_year - 1, target_year):
            records = customs_list_for_year(year)
            if not records:
                _record("customs", "partial", f"no records listed for {year}")
                continue
            # every record per month, freshest first (list is DESC by id), so
            # a month that published a non-standard file can fall through to
            # the next candidate instead of losing the month entirely
            by_month = {}
            for rec in records:
                try:
                    m = int(rec.get("month"))
                except (TypeError, ValueError):
                    continue
                if 1 <= m <= 12:
                    by_month.setdefault(m, []).append(rec)
            for m in sorted(by_month):
                if year == target_year and m > target_month:
                    continue
                wanted += 1
                fname = f"customs_{year}_{m:02d}.xlsx"
                path = RAW_DATA_DIR / fname
                if path.exists():
                    if _customs_workbook_ok(path.read_bytes()):
                        log.info("cached   %s (skip download)", fname)
                        saved.append(fname)
                        continue
                    # a previously accepted wrong file: drop it and retry
                    log.warning("cached %s is not a monthly workbook -> "
                                "discarding and re-downloading", fname)
                    path.unlink()
                got = False
                for cand in by_month[m]:
                    url = _customs_record_file_url(cand)
                    if not url:
                        continue
                    log.info("download %s <- %s", fname, url)
                    try:
                        r = _get(url)
                        r.raise_for_status()
                        time.sleep(0.5)    # be polite to the CDN
                        if not _customs_workbook_ok(r.content):
                            log.warning("%s-%02d: candidate rejected, trying "
                                        "next record", year, m)
                            continue
                        if _save(fname, r.content, "xlsx"):
                            saved.append(fname)
                            got = True
                            break
                    except Exception as exc:
                        log.error("download failed for %s: %s", fname, exc)
                if not got:
                    # Not fatal for a history month: the series simply skips
                    # it. Fatal only if it is the target month, which the
                    # status check below reports.
                    log.warning("%s-%02d: no monthly workbook among %d listed "
                                "record(s) -- month skipped in the series",
                                year, m, len(by_month[m]))

        target_file = f"customs_{target_year}_{target_month:02d}.xlsx"
        status = "ok" if target_file in saved else "failed"
        _record("customs", status,
                f"{len(saved)}/{wanted} monthly workbooks present", saved)
        log.info("CUSTOMS: %d/%d workbooks present.", len(saved), wanted)
    except Exception as exc:
        log.error("CUSTOMS failed: %s", exc)
        _record("customs", "failed", str(exc), saved)


# SOURCE 3: NSO tables (generic fetcher)
def ingest_nso() -> None:
    log.info("---- [3/4] NSO (1212.mn) ----")
    for save_name, spec in NSO_TABLES.items():
        # a 4th element is an extra query string (the hotel and food tables
        # need &subtables=... to select the right sub-table)
        table_id, sector, subsector = spec[0], spec[1], spec[2]
        extra = spec[3] if len(spec) > 3 else ""
        url = NSO_API_TEMPLATE.format(table_id=table_id, sector=sector,
                                      subsector=subsector) + extra
        candidates = NSO_TABLE_PAYLOADS.get(save_name, [NSO_POST_PAYLOAD])
        last_err, done = None, False
        for i, body in enumerate(candidates):
            try:
                log.info("POST %s (payload %d/%d)", url, i + 1,
                         len(candidates))
                r = _post(url, json=body, verify=False)
                r.raise_for_status()
                payload = r.content
                js = json.loads(payload.decode("utf-8"))
                if "value" not in js or not js["value"]:
                    raise ValueError("JSON but no json-stat2 'value' array")
                fname = f"{save_name}.json"
                if _save(fname, payload, "json"):
                    _record("nso", "ok", f"{table_id} -> {fname}", [fname])
                    done = True
                    break
            except Exception as exc:
                last_err = exc
                log.warning("NSO[%s] payload %d failed: %s", save_name,
                            i + 1, exc)
        if not done:
            log.error("NSO[%s] failed with all %d payload(s): %s",
                      save_name, len(candidates), last_err)
            _record("nso", "failed", f"{table_id}: {last_err}")


# SOURCE 4: MONGOLBANK -- section sweep + xlsx reports
def _mongolbank_headers() -> dict:
    h = dict(HEADERS)
    h.update({
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://stat.mongolbank.mn",
        "Referer": "https://stat.mongolbank.mn/external",
        "X-Requested-With": "XMLHttpRequest",
    })
    return h


def ingest_mongolbank_sections() -> None:
    """
    Probe every parentId in MONGOLBANK_PARENT_IDS and save each section that
    returns cards. Card names are matched downstream, so new/unknown sections
    are picked up with zero code changes here.
    """
    log.info("---- [4/4] MONGOLBANK section sweep (%d ids) ----",
             len(MONGOLBANK_PARENT_IDS))
    saved, cards_total = [], 0
    for pid in MONGOLBANK_PARENT_IDS:
        try:
            r = _post(MONGOLBANK_API_URL, params={"lang": "mn", "parentId": pid},
                      headers=_mongolbank_headers(), timeout=20)
            if r.status_code != 200:
                continue
            js = r.json()
            cards = js.get("result") or []
            # keep only sections that actually carry data cards
            cards = [c for c in cards if (c.get("data") or "").strip()]
            if not cards:
                continue
            name = MONGOLBANK_KNOWN_SECTIONS.get(pid, f"section_{pid}")
            fname = f"mongolbank_{name}.json"
            if _save(fname, r.content, "json"):
                saved.append(fname)
                cards_total += len(cards)
                log.info("parentId=%-3s -> %2d cards (%s)", pid, len(cards), fname)
        except Exception:
            continue   # sweep is best-effort by design
    status = "ok" if saved else "failed"
    _record("mongolbank_cards", status,
            f"{len(saved)} sections / {cards_total} cards", saved)
    log.info("MONGOLBANK sweep: %d sections, %d cards.", len(saved), cards_total)


def _mb_sublist(surveyid: int):
    """POST the sublist API for one surveyid -> parsed JSON or None."""
    try:
        r = _post(MONGOLBANK_SURVEY_LIST_API,
                  params={"lang": "mn", "surveyid": surveyid},
                  headers=_mongolbank_headers(), timeout=20)
        if r.status_code != 200:
            return None
        return r.json()
    except Exception:
        return None


_MB_DASHES = str.maketrans({"‐": "-", "‑": "-", "‒": "-",
                            "–": "-", "—": "-", "−": "-"})


def _mb_entries(js) -> list:
    """
    Tolerantly collect {id, text, (year, month)} entries from a sublist
    payload, wherever the API nests them (schema-agnostic tree walk). The
    id key may be spelled id / Id / ID; month names like '2026 5-р сар'
    are parsed into numeric (year, month) with unicode-dash normalisation.
    """
    found = []
    per_pat = re.compile(r"(20\d{2})\s*(?:оны)?\s*(\d{1,2})\s*-\s*[рp]\s*сар")
    # id key may be spelled id / Id / ID / SURVEY_ID (real schema, from the
    # captured sublist payload) -- accept any '(survey_)id' spelling.
    id_pat = re.compile(r"^(survey_?)?id$", re.IGNORECASE)

    def walk(node):
        if isinstance(node, dict):
            id_key = next((k for k in node if id_pat.match(str(k))), None)
            if id_key is not None:
                texts = " ".join(str(v) for v in node.values()
                                 if isinstance(v, (str, int, float)))
                texts = texts.translate(_MB_DASHES)
                m = per_pat.search(texts)
                found.append({"id": node[id_key], "text": texts,
                              "period": (int(m.group(1)), int(m.group(2)))
                              if m else None,
                              "is_xls": bool(node.get("IS_XLS", 1)),
                              "raw": node})
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    walk(js)
    return found


def _mb_resolve_report_id(name: str, cfg: dict, dump_on_fail=True,
                          want=None):
    """
    Find the node whose sublist yields MONTHLY FILE entries for a report.

    The portal is two-level: sublist(<section id>) lists the section's
    REPORTS (by name), and sublist(<report id>) lists that report's monthly
    files. A configured surveyid may point at either level, so:
      1. sublist(surveyid); if entries parse as months -> done.
      2. otherwise look for a child node matching the report's keywords and
         recurse into its id.
      3. if no surveyid configured, sweep MONGOLBANK_SURVEYID_SCAN looking
         for keyword matches at either level.
    Returns (report_id, monthly_entries) or (None, []). Raw payloads are
    dumped to raw_files/mongolbank_sublist_debug_<sid>.json on failure so a
    mismatch is diagnosable without another DevTools session.
    """
    def monthly(sid):
        js = _mb_sublist(sid)
        if not js:
            return None, []
        entries = _mb_entries(js)
        months = [e for e in entries if e["period"]]
        return js, months if months else []

    def wanted(months):
        """True if this listing actually covers the month being requested."""
        if not want or not months:
            return bool(months)
        y, m = want
        for _ in range(MB_REPORT_WALKBACK_MONTHS + 1):
            if any(e["period"] == (y, m) for e in months):
                return True
            m -= 1
            if m == 0:
                m, y = 12, y - 1
        return False

    def try_sid(sid, depth=0):
        js, months = monthly(sid)
        # A node can yield month-like entries that are NOT the report's file
        # list -- the bulletin's section lists revision notices whose titles
        # carry old month names, which is why the June issue was missed while
        # a 2025-10 list looked "good enough". Accept a listing only when it
        # covers the requested month; otherwise keep descending, and fall
        # back to whatever was found only if nothing better turns up.
        if months and wanted(months):
            return sid, months
        stale = (sid, months) if months else None
        if js is None or depth >= 1:
            return stale or (None, [])
        # second level: child node whose text matches the report keywords.
        # Skip revision notices ('...гарсан өөрчлөлт' archive entries) --
        # their titles quote report names and are pure false positives.
        for e in _mb_entries(js):
            text = e["text"].lower()
            if "өөрчлөлт" in text:
                continue
            if any(kw.lower() in text for kw in cfg["keywords"]):
                rid, months = try_sid(e["id"], depth + 1)
                if months and wanted(months):
                    return rid, months
                if months and stale is None:
                    stale = (rid, months)
        if stale:
            return stale
        if dump_on_fail and js is not None:
            (RAW_DATA_DIR / f"mongolbank_sublist_debug_{sid}.json").write_text(
                json.dumps(js, ensure_ascii=False, indent=2),
                encoding="utf-8")
        return None, []

    # cached resolution from a previous run
    cache = RAW_DATA_DIR / "mongolbank_survey_map.json"
    cached = {}
    if cache.exists():
        try:
            cached = json.loads(cache.read_text(encoding="utf-8"))
        except Exception:
            cached = {}
    # Keep the best listing seen so far, but do not stop until one actually
    # covers the requested month. A cached surveyid that has gone stale must
    # not prevent the sweep from finding the node that now carries it.
    best = (None, [])

    def better(cand):
        nonlocal best
        rid, months = cand
        if not months:
            return False
        if wanted(months):
            best = cand
            return True
        if not best[1] or (max(e["period"] for e in months if e["period"])
                           > max(e["period"] for e in best[1] if e["period"])):
            best = cand
        return False

    if name in cached and better(try_sid(cached[name])):
        return best

    if cfg.get("surveyid") is not None and better(try_sid(cfg["surveyid"])):
        cached[name] = best[0]
        cache.write_text(json.dumps(cached, ensure_ascii=False, indent=2),
                         encoding="utf-8")
        return best

    log.info("MONGOLBANK[%s]: sweeping surveyids for keyword match...", name)
    for sid in MONGOLBANK_SURVEYID_SCAN:
        js = _mb_sublist(sid)
        if not js:
            continue
        blob = json.dumps(js, ensure_ascii=False).lower()
        if not any(kw.lower() in blob for kw in cfg["keywords"]):
            continue
        if better(try_sid(sid)):
            cached[name] = best[0]
            cache.write_text(json.dumps(cached, ensure_ascii=False, indent=2),
                             encoding="utf-8")
            return best
    if best[1]:
        newest = max(e["period"] for e in best[1] if e["period"])
        log.warning("MONGOLBANK[%s]: no listing covers %s; the freshest found "
                    "reaches %s.", name,
                    f"{want[0]}-{want[1]:02d}" if want else "the target",
                    f"{newest[0]}-{newest[1]:02d}")
    return best


ID_HISTORY_FILE = "mongolbank_id_history.json"


def _load_id_history() -> dict:
    """{'YYYY-MM': {report: id}} of everything resolved on previous runs."""
    p = RAW_DATA_DIR / ID_HISTORY_FILE
    if not p.exists():
        return {}
    try:
        h = json.loads(p.read_text(encoding="utf-8"))
        return h if isinstance(h, dict) else {}
    except Exception:
        return {}


def _save_id_history(hist: dict, year: int, month: int, ids: dict) -> None:
    if not ids:
        return
    hist = dict(hist)
    hist[f"{year}-{month:02d}"] = {k: int(v) for k, v in ids.items()}
    # keep it small: the last two years is far more than the projection needs
    for k in sorted(hist)[:-24]:
        hist.pop(k, None)
    try:
        (RAW_DATA_DIR / ID_HISTORY_FILE).write_text(
            json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        log.warning("could not write %s: %s", ID_HISTORY_FILE, exc)


def _prev_period_ids(hist: dict, year: int, month: int) -> dict:
    """Ids from the most recent period BEFORE the one being reported."""
    tag = f"{year}-{month:02d}"
    earlier = sorted(k for k in hist if k < tag)
    return hist.get(earlier[-1], {}) if earlier else {}


def _bulletin_period_of(content: bytes):
    """
    (year, month) the bulletin actually covers, read from its own data.

    The 'Inflation Nat' sheet lists one row per period: an explicit
    'YYYY MM' label starts a year and bare month numbers continue it. The
    newest such row IS the issue month. Knowing this lets a candidate be
    rejected for being the WRONG MONTH, not merely for not being a bulletin --
    without it, probing ids can happily re-download the previous issue.
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True,
                                    data_only=True)
        sheet = next((s for s in wb.sheetnames
                      if "inflation nat" in str(s).lower()), None)
        if sheet is None:
            wb.close()
            return None
        year = month = None
        for row in wb[sheet].iter_rows(max_col=1, values_only=True):
            v = str(row[0]).strip() if row and row[0] is not None else ""
            m = re.match(r"^(\d{4})\s+(\d{1,2})$", v)
            if m:
                year, month = int(m.group(1)), int(m.group(2))
            elif re.match(r"^\d{1,2}$", v) and year is not None:
                mm = int(v)
                if month is not None and mm < month:
                    year += 1
                month = mm
        wb.close()
        return (year, month) if year and month else None
    except Exception:
        return None


def _loan_report_ok(content: bytes) -> bool:
    """The loan report is identified by its borrower-type sheets."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        names = {str(s).strip().lower() for s in wb.sheetnames}
        wb.close()
    except Exception:
        return False
    return {"total", "private", "individual"} <= names


def _report_content_ok(name: str, content: bytes) -> bool:
    """A wrong survey id downloads a DIFFERENT report; verify the workbook
    is really the one we asked for (bulletin must carry its Budget sheet,
    the loan report its borrower-type sheets)."""
    if name == "bank_loan_report":
        return _loan_report_ok(content)
    if name != "bulletin":
        return True
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        names = " ".join(str(s).lower() for s in wb.sheetnames)
        wb.close()
        return "budget" in names and "inflation" in names
    except Exception:
        return False


def ingest_mongolbank_reports(target_month: int, target_year: int) -> None:
    """
    Download each configured monthly Excel report by resolving its DYNAMIC
    file id from the sublist API for the target month (with walk-back), so
    nothing breaks when a new month is published. Files are saved as
    mongolbank_<name>.xlsx (always the freshest resolved month).
    """
    log.info("---- [4b] MONGOLBANK EXCEL REPORTS (dynamic ids) ----")
    saved, pending, resolved_ids = [], [], {}

    for name, cfg in MONGOLBANK_SURVEYS.items():
        fname = f"mongolbank_{name}.xlsx"
        # resolve candidate file ids. A hand-set id for this exact month
        # goes first and short-circuits discovery entirely -- that is the
        # whole point of it: when the portal has changed enough that the
        # listing and the probe both fail, the operator reads the id out of
        # DevTools and the run just works.
        candidates = []          # [(file_id, period_tag), ...] in try-order
        pinned = _ovget("mongolbank", "report_ids",
                        f"{target_year}-{target_month:02d}", name)
        if pinned is not None:
            candidates.append((pinned, "PINNED in overrides.json"))
            log.warning("MONGOLBANK[%s]: using the id pinned in "
                        "overrides.json (%s) for %d-%02d; discovery is "
                        "skipped for this report.", name, pinned,
                        target_year, target_month)
        try:
            report_id, months = _mb_resolve_report_id(
                name, cfg, want=(target_year, target_month))
        except Exception as exc:
            log.error("MONGOLBANK[%s]: id resolution crashed (%s); "
                      "continuing with fallbacks.", name, exc)
            months = []
        if months:
            m, y = target_month, target_year
            for attempt in range(MB_REPORT_WALKBACK_MONTHS + 1):
                cands = [e for e in months if e["period"] == (y, m)]
                hit = next((e for e in cands if e["is_xls"]),
                           cands[0] if cands else None)
                if hit:
                    if attempt:
                        log.warning("MONGOLBANK[%s]: %d-%02d not published; "
                                    "walked back to %d-%02d.", name,
                                    target_year, target_month, y, m)
                    candidates.append((hit["id"], f"{y}-{m:02d}"))
                    break
                m -= 1
                if m == 0:
                    m, y = 12, y - 1
            if not candidates:
                newest = max((e["period"] for e in months), default=None)
                # Not fatal for the bulletin: its listing is known to lag by
                # months and three further routes are tried below. It IS
                # serious for any other report, which has only this one route.
                log.log(logging.WARNING if name == "bulletin"
                        else logging.ERROR,
                        "MONGOLBANK[%s]: no monthly entry for %d-%02d or "
                        "%d months back (newest available: %s)%s", name,
                        target_year, target_month,
                        MB_REPORT_WALKBACK_MONTHS, newest,
                        " -- falling through to the id probe."
                        if name == "bulletin" else ".")
        # The bulletin's own listing lags, so when it cannot be resolved its
        # id is hunted near the ids that DID resolve this month. Publication
        # ids are allocated close together but the gap is not fixed -- in
        # 2026-05 the bulletin sat 1 below the loan report, in 2026-06 it was
        # 20 below the loan report and only 3 below the balance sheet. So all
        # resolved anchors are used and the search widens by proximity, with
        # every candidate checked for the RIGHT MONTH before acceptance.
        if name == "bulletin" and resolved_ids:
            seen = {c[0] for c in candidates}
            # Safety net from history. Publication ids climb by roughly the
            # number of files Mongolbank posts in a month -- about 100-120,
            # and NOT a constant, so a fixed offset is useless. What is
            # usable: however far the reports that DID resolve moved since
            # last month, the bulletin moved by a similar amount. Projecting
            # last month's bulletin id by that delta lands within the probe
            # radius (June: projected 5235, actual 5217), which matters when
            # every other anchor is far away.
            hist = _load_id_history()
            prev = _prev_period_ids(hist, target_year, target_month)
            if prev.get("bulletin"):
                deltas = [resolved_ids[n] - prev[n] for n in resolved_ids
                          if prev.get(n)]
                if deltas:
                    deltas.sort()
                    med = deltas[len(deltas) // 2]
                    proj = prev["bulletin"] + med
                    if proj > 0 and proj not in seen:
                        seen.add(proj)
                        candidates.append(
                            (proj, f"projected from last month's bulletin "
                                   f"{prev['bulletin']} +{med}"))
                    for off in range(1, BULLETIN_PROBE_RADIUS + 1):
                        for cand in (proj - off, proj + off):
                            if cand > 0 and cand not in seen:
                                seen.add(cand)
                                candidates.append(
                                    (cand, f"near projected id {proj}"))
            # Rank by the SMALLEST distance to ANY anchor, not anchor by
            # anchor: the June bulletin sat 3 from the balance sheet but 20
            # from the loan report, and per-anchor ordering buried it behind
            # forty loan-adjacent misses.
            best_off = {}
            for anchor_name, base in resolved_ids.items():
                for off in range(1, BULLETIN_PROBE_RADIUS + 1):
                    for cand in (base - off, base + off):
                        if cand <= 0 or cand in seen:
                            continue
                        if cand not in best_off or off < best_off[cand][0]:
                            best_off[cand] = (off, f"near {anchor_name} "
                                                   f"id {base}")
            for cand, (off, tag) in sorted(best_off.items(),
                                           key=lambda kv: kv[1][0]):
                candidates.append((cand, tag))
        static_id = MONGOLBANK_STATIC_FALLBACK.get(name)
        if static_id is not None and static_id not in [c[0] for c in
                                                       candidates]:
            candidates.append((static_id, "static fallback"))
        if not candidates:
            pending.append(name)
            log.warning("MONGOLBANK[%s]: no id resolvable (see "
                        "mongolbank_sublist_debug_*.json in raw_files).", name)
            continue

        # download: first candidate that yields a VALID xlsx wins. For the
        # bulletin the month is checked too, and a right-report/wrong-month
        # hit is kept aside as a fallback while the search continues -- that
        # is what turns "we found a bulletin" into "we found June's".
        ok = False
        fallback = None                      # (content, id, period) if stale
        probes = 0
        for file_id, tag in candidates:
            if tag.startswith("near "):
                # Skip further probes past the budget, but never abandon the
                # remaining non-probe candidates -- the static fallback is
                # last in the list, and breaking here would leave the run
                # with no bulletin at all.
                if probes >= BULLETIN_PROBE_BUDGET:
                    if probes == BULLETIN_PROBE_BUDGET:
                        log.warning("MONGOLBANK[%s]: probe budget of %d "
                                    "reached; falling back.", name,
                                    BULLETIN_PROBE_BUDGET)
                        probes += 1
                    continue
                probes += 1
            try:
                log.info("MONGOLBANK[%s]: downloading id=%s (%s)", name,
                         file_id, tag)
                r = _post(MONGOLBANK_REPORT_DATA_API,
                          params={"lang": "mn", "id": file_id,
                                  "type": "xlsx"},
                          headers=_mongolbank_headers())
                r.raise_for_status()
                if (r.content[:2] == b"PK"
                        and not _report_content_ok(name, r.content)):
                    log.info("MONGOLBANK[%s]: id=%s is a valid xlsx but not "
                             "this report; next candidate.", name, file_id)
                    continue
                if name == "bulletin":
                    per = _bulletin_period_of(r.content)
                    if per and per != (target_year, target_month):
                        if fallback is None:
                            fallback = (r.content, file_id, per)
                        log.info("MONGOLBANK[bulletin]: id=%s is the "
                                 "%d-%02d issue, not %d-%02d; keeping it as a "
                                 "fallback and continuing.", file_id, per[0],
                                 per[1], target_year, target_month)
                        continue
                    if per:
                        log.info("MONGOLBANK[bulletin]: id=%s confirmed as "
                                 "the %d-%02d issue.", file_id, *per)
                # More candidates left? Then a rejection here is the search
                # working. Only the last one failing is worth an ERROR.
                more = file_id != candidates[-1][0]
                if _save(fname, r.content, "xlsx", probing=more):
                    saved.append(fname)
                    resolved_ids[name] = file_id
                    ok = True
                    break
                log.log(logging.INFO if more else logging.ERROR,
                        "MONGOLBANK[%s]: id=%s returned an invalid payload; "
                        "trying next candidate.", name, file_id)
            except Exception as exc:
                log.log(logging.INFO
                        if file_id != candidates[-1][0] else logging.ERROR,
                        "MONGOLBANK[%s]: id=%s failed (%s); trying next "
                        "candidate.", name, file_id, exc)
        if not ok and fallback is not None:
            content, file_id, per = fallback
            if _save(fname, content, "xlsx"):
                saved.append(fname)
                resolved_ids[name] = file_id
                ok = True
                log.warning("MONGOLBANK[bulletin]: %d-%02d was not found "
                            "anywhere; using the %d-%02d issue (id=%s). The "
                            "workbook will label those figures %d-%02d.",
                            target_year, target_month, per[0], per[1],
                            file_id, per[0], per[1])
        if not ok:
            pending.append(name)
            log.error("MONGOLBANK[%s]: ALL candidates failed -- %s will be "
                      "missing this run and its workbook cells flagged.",
                      name, fname)

    # Remember this month's resolved ids so next month can project from them
    # if a report's own listing is unusable.
    if resolved_ids:
        _save_id_history(_load_id_history(), target_year, target_month,
                         resolved_ids)
        log.info("resolved ids recorded for %d-%02d: %s", target_year,
                 target_month,
                 ", ".join(f"{k}={v}" for k, v in sorted(resolved_ids.items())))

    detail = f"{len(saved)}/{len(MONGOLBANK_SURVEYS)} reports"
    if pending:
        detail += f"; failed/pending: {', '.join(pending)}"
    _record("mongolbank_reports", "ok" if saved else "failed", detail, saved)


def ingest_loan_history(target_month: int, target_year: int) -> None:
    """
    Download the bank loan report for the target month and the months before
    it, one file per month (mb_loans_YYYY_MM.xlsx).

    Each report is a single-month flow statement (opening balance, loans
    disbursed, repaid, FX revaluation, closing balance), so a history is what
    turns it into a growth series and makes year-on-year comparison possible.
    Ids come from the same sublist the single-report path uses, months already
    on disk are skipped, and every download is verified to be the loan report
    before it is accepted.
    """
    log.info("---- [4c] MONGOLBANK LOAN REPORT HISTORY (%d months) ----",
             LOAN_HISTORY_MONTHS)
    cfg = MONGOLBANK_SURVEYS.get("bank_loan_report")
    if not cfg:
        log.warning("loan history: no survey config; skipped.")
        return

    wanted = []
    m, y = target_month, target_year
    for _ in range(LOAN_HISTORY_MONTHS):
        wanted.append((y, m))
        m -= 1
        if m == 0:
            m, y = 12, y - 1

    missing = [(y, m) for (y, m) in wanted
               if not (RAW_DATA_DIR / f"mb_loans_{y}_{m:02d}.xlsx").exists()]
    if not missing:
        log.info("loan history: all %d months cached.", len(wanted))
        _record("mongolbank_loan_history", "ok",
                f"{len(wanted)}/{len(wanted)} months cached")
        return

    try:
        _, months = _mb_resolve_report_id("bank_loan_report", cfg,
                                          want=(target_year, target_month))
    except Exception as exc:
        log.error("loan history: id resolution failed (%s); skipped.", exc)
        _record("mongolbank_loan_history", "failed", str(exc))
        return
    if not months:
        log.warning("loan history: sublist returned no monthly entries; "
                    "only the current report will be available.")
        _record("mongolbank_loan_history", "partial", "no sublist entries")
        return

    by_period = {}
    for e in months:
        by_period.setdefault(e["period"], []).append(e)

    saved = []
    for (y, m) in wanted:
        fname = f"mb_loans_{y}_{m:02d}.xlsx"
        path = RAW_DATA_DIR / fname
        if path.exists():
            saved.append(fname)
            continue
        entries = by_period.get((y, m), [])
        if not entries:
            log.warning("loan history: %d-%02d not listed -- month skipped.",
                        y, m)
            continue
        entries.sort(key=lambda e: not e.get("is_xls"))
        got = False
        for e in entries:
            try:
                r = _post(MONGOLBANK_REPORT_DATA_API,
                          params={"lang": "mn", "id": e["id"], "type": "xlsx"},
                          headers=_mongolbank_headers())
                r.raise_for_status()
                if not _loan_report_ok(r.content):
                    log.warning("loan history: %d-%02d id=%s is not the loan "
                                "report; trying next entry.", y, m, e["id"])
                    continue
                if _save(fname, r.content, "xlsx"):
                    saved.append(fname)
                    got = True
                    break
            except Exception as exc:
                log.warning("loan history: %d-%02d id=%s failed (%s).",
                            y, m, e["id"], exc)
            time.sleep(0.3)
        if not got:
            log.warning("loan history: %d-%02d could not be downloaded.", y, m)

    log.info("loan history: %d/%d months present.", len(saved), len(wanted))
    _record("mongolbank_loan_history",
            "ok" if len(saved) >= 2 else "partial",
            f"{len(saved)}/{len(wanted)} months", saved)


def ingest_deposits_export(target_month: int, target_year: int) -> None:
    """
    Fetch the deposit tree as a monthly series from the portal's indicator
    API, so the Banking sheet is not limited to the few date columns the
    downloadable report happens to carry.

    The response is saved verbatim. If it is a workbook it is stored as
    deposits_export.xlsx; if it is JSON it is stored as deposits_export.json
    and the processor reads whichever it can. A hand-downloaded export saved
    as deposits_manual.xlsx is used when this fetch is unavailable, so the
    deposit block never depends on this call succeeding.
    """
    log.info("---- [4d] MONGOLBANK DEPOSIT SERIES (indicator API) ----")
    payload = {
        "id": DEPOSIT_REPORT_ID,
        "parentId": DEPOSIT_PARENT_ID,
        "rCheck": 1,
        "cycle_data": {
            "interval": "3",                       # monthly
            "year_start": str(target_year - DEPOSIT_HISTORY_YEARS),
            "mq_start": "1", "day_start": "1",
            "year_end": str(target_year),
            "mq_end": str(target_month), "day_end": "1",
        },
        "indicators": DEPOSIT_INDICATOR_IDS,
    }
    try:
        r = _post(MONGOLBANK_INDICATOR_API, params={"lang": "mn"},
                  json=payload, headers=_mongolbank_headers())
        r.raise_for_status()
        body = r.content
    except Exception as exc:
        log.warning("deposit series fetch failed (%s) -- the Banking sheet "
                    "falls back to the report's own columns, or to "
                    "raw_files/deposits_manual.xlsx if you saved one.", exc)
        _record("mongolbank_deposits", "failed", str(exc))
        return

    if body[:2] == b"PK":
        ok = _save("deposits_export.xlsx", body, "xlsx")
        _record("mongolbank_deposits", "ok" if ok else "failed",
                "workbook response", ["deposits_export.xlsx"])
        return
    if _save("deposits_export.json", body, "json"):
        try:
            js = json.loads(body.decode("utf-8"))
            shape = (list(js)[:8] if isinstance(js, dict)
                     else f"list[{len(js)}]")
            log.info("deposit series saved as JSON (top level: %s)", shape)
        except Exception:
            pass
        _record("mongolbank_deposits", "ok", "json response",
                ["deposits_export.json"])
    else:
        log.warning("deposit series response was neither a workbook nor "
                    "JSON -- ignored.")
        _record("mongolbank_deposits", "failed", "unrecognised response")


def ingest_bop_manual() -> None:
    """
    Balance of Payments -- HUMAN-IN-THE-LOOP by design. The BoP portal page
    is a dynamic dashboard (indicator checkboxes + date pickers), so the
    operator downloads the Excel once per run:

        stat.mongolbank.mn/external -> Төлбөрийн тэнцэл -> export Excel
        -> save as raw_files/bop_manual.xlsx

    When the file is missing and the run is interactive, execution pauses
    with an input() prompt; in non-interactive runs (cron/CI) it is skipped
    with a clear manifest flag instead of blocking forever.
    """
    log.info("---- [5/5] BALANCE OF PAYMENTS (manual file) ----")
    path = RAW_DATA_DIR / BOP_MANUAL_FILE

    def _ok():
        """
        Valid Excel is not enough -- it has to be the RIGHT export.

        The portal offers several dashboards side by side and every one of
        them downloads as a nondescript .xlsx. The banking balance sheet is
        the easy mistake, because its Mongolian name 'Банкны салбарын тайлан
        ТЭНЦЭЛ' shares a word with 'ТӨЛБӨРИЙН ТЭНЦЭЛ'. Catching it here means
        the operator is told while they still have the browser open, instead
        of finding out from a blank slide.
        """
        if path.read_bytes()[:2] != b"PK":
            log.error("%s exists but is not a valid Excel file.", path.name)
            return False
        try:
            df = pd.read_excel(path, sheet_name=0, header=None, nrows=60)
        except Exception as exc:
            log.error("%s cannot be opened as a workbook: %s", path.name, exc)
            return False
        col0 = " ".join(str(v).casefold() for v in df.iloc[:, 0].tolist()
                        if isinstance(v, str))
        wanted = ("урсгал данс", "хөрөнгийн данс", "санхүүгийн данс",
                  "нөөц хөрөнгө")
        hits = sum(w in col0 for w in wanted)
        if hits < 2:
            what = ("Банкны салбарын тайлан тэнцэл — the banking BALANCE "
                    "SHEET, which the pipeline already downloads on its own"
                    if "нийт актив" in col0 or "банкны нөөц" in col0
                    else "not the balance of payments")
            log.error("WRONG FILE: %s is %s.", path.name, what)
            log.error("  Need: stat.mongolbank.mn/external -> "
                      "'Төлбөрийн тэнцэл' -> export to Excel.")
            log.error("  The right file's first column reads I. УРСГАЛ ДАНС, "
                      "II. ХӨРӨНГИЙН ДАНС, III. САНХҮҮГИЙН ДАНС.")
            return False
        _record("bop_manual", "ok", path.name, [path.name])
        log.info("BoP manual file present and verified: %s (%d of 4 account "
                 "headings found).", path.name, hits)
        return True

    if path.exists() and _ok():
        return
    if not sys.stdin.isatty():
        _record("bop_manual", "missing",
                f"{BOP_MANUAL_FILE} missing or not the BoP export "
                "(non-interactive run; not prompting). Slide 12 detail "
                "stays flagged.")
        log.warning("BoP file missing or wrong and the shell is "
                    "non-interactive -- skipping the prompt.")
        return
    # Two attempts: the first mistake is usually the neighbouring dashboard,
    # and by now the operator knows exactly which file to look for.
    for attempt in (1, 2):
        input(
            "\nDownload the ТӨЛБӨРИЙН ТЭНЦЭЛ (balance of payments) Excel "
            "from stat.mongolbank.mn/external,\nsave it as "
            f"'{RAW_DATA_DIR / BOP_MANUAL_FILE}', and press Enter"
            f"{' to try again' if attempt == 2 else ''}..."
        )
        if path.exists() and _ok():
            return
        if attempt == 1:
            log.warning("Still not the right file — one more try, then the "
                        "run continues without it.")
    _record("bop_manual", "missing",
            f"{BOP_MANUAL_FILE} missing or not the BoP export after two "
            "prompts. Slide 12 detail stays flagged; re-run when the right "
            "file is in place.")
    log.warning("Continuing without the BoP file. Nothing else is affected "
                "and Data_Vintage marks the BoP rows MISSING.")


# MAIN
def main() -> None:
    _ov.apply_and_log(OVERRIDES)
    log.info("========== RAW INGESTION v2 START (%s) ==========",
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    _ensure_raw_dir()

    target_month, target_year = get_target_period()
    MANIFEST["run_at"] = datetime.now().isoformat(timespec="seconds")
    MANIFEST["target"] = {"year": target_year, "month": target_month}

    ingest_customs(target_month, target_year)
    ingest_nso()
    ingest_mongolbank_sections()
    ingest_mongolbank_reports(target_month, target_year)
    ingest_loan_history(target_month, target_year)
    ingest_deposits_export(target_month, target_year)
    ingest_bop_manual()

    MANIFEST_PATH.write_text(json.dumps(MANIFEST, ensure_ascii=False, indent=2),
                             encoding="utf-8")
    log.info("Manifest written: %s", MANIFEST_PATH)

    failed = [s for s, v in MANIFEST["sources"].items()
              if v.get("status") == "failed"]
    if failed:
        log.warning("Sources with failures: %s (pipeline continues; the "
                    "processor uses whatever is cached).", ", ".join(failed))
    log.info("========== RAW INGESTION COMPLETE ==========")


if __name__ == "__main__":
    main()
