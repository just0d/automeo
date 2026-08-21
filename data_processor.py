#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Computes every metric and series from ./raw_files into
processed_metrics.json and processed_series.json.
See the Data Dictionary document for field-by-field lineage.

Usage: python data_processor.py [month] [year]
"""

import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# CONFIG
BASE_DIR = Path(__file__).resolve().parent
RAW_DATA_DIR = BASE_DIR / "raw_files"
NSO_JSON = RAW_DATA_DIR / "nso_gdp_raw.json"
METRICS_JSON = BASE_DIR / "processed_metrics.json"
SERIES_JSON = BASE_DIR / "processed_series.json"

# Customs: commodities by HS code (sheet '8', matched on the CODE column).
# Sheet 8 columns (header=None):
#   2=code 3=name 4=unit | 5=QtyPrevExp 6=AmtPrevExp 7=QtyCurExp 8=AmtCurExp
#                        | 9=QtyPrevImp 10=AmtPrevImp 11=QtyCurImp 12=AmtCurImp
EXPORT_COMMODITIES = {          # key -> (HS code, label, volume unit out)
    "coal":   ("2701", "Чулуун нүүрс",            "million tonnes"),
    "copper": ("2603", "Зэсийн хүдэр, баяжмал",   "thousand tonnes"),
    "gold":   ("7108", "Мөнгөжөөгүй алт",         "kg"),
    "iron":   ("2601", "Төмрийн хүдэр, баяжмал",  "million tonnes"),
}
IMPORT_COMMODITIES = {          # sheet 8, import side
    "petroleum":   ("2710", "Нефтийн бүтээгдэхүүн"),
    "cars":        ("8703", "Суудлын автомашин"),
    "electricity": ("2716", "Цахилгаан эрчим хүч"),
    "trucks":      ("8704", "Ачааны автомашин"),
}
# sheet '8.2 бүлгээр' (by HS chapter): col0=chapter col1=name,
#   2=QtyPrevExp 3=AmtPrevExp 4=QtyCurExp 5=AmtCurExp 6..9 = import side
CHAPTER_COMMODITIES = {
    "wool_cashmere": ("51", "Ноос, ноолуур"),
}
# sheet '3' (by HS section): col0 startswith 'NN ', cols 2/3=exp prev/cur,
# 4/5 = imp prev/cur
SECTION_IMPORTS = {
    "machinery":   ("16", "Машин, тоног төхөөрөмж"),
    "base_metals": ("15", "Үндсэн төмөрлөг"),
    "food":        ("04", "Хүнсний бүтээгдэхүүн"),
}

# NSO dimension names / positions (verified against the live table).
NSO_DIM_INDICATOR = "Статистик үзүүлэлт"
NSO_DIM_SECTOR = "Салбар"
NSO_DIM_PERIOD = "Он"
NSO_REAL_GDP_2015 = 3       # 'ДНБ, 2015 оны зэрэгцүүлэх үнээр'
NSO_TOTAL_SECTOR = 0        # 'ДНБ'

INFLATION_TARGET_PCT = 8.0  # Mongolbank official target (deck slide 6)

# Mongolbank card-name fragments -> metric keys (matched case-insensitive,
# 'contains'). The section sweep saves everything; we map what we recognise.
MONGOLBANK_CARD_MAP = [
    ("гадаад валютын улсын нөөц", "reserves_mln_usd_mb"),
    ("нийт гадаад өр",            "external_debt_mln_usd"),
    ("төлбөрийн тэнцэл",          "balance_of_payments_mln_usd"),
    ("гадаад худалдаа",           "trade_balance_mln_usd_mb"),
    ("шууд хөрөнгө оруулалт",     "fdi_mln_usd"),
    ("бодлогын хүү",              "policy_rate"),
    ("банкны салбарын зээл",      "bank_loans_tln_mnt"),
    ("ипотекийн зээл",            "mortgage_loans_tln_mnt"),
    ("мөнгөний үзүүлэлт",         "money_supply_tln_mnt"),
    ("төгрөгийн ханш",            "usd_mnt_rate"),
    ("ам.доллар",                 "usd_mnt_rate"),
    ("мөнгөний нийлүүлэлт",       "money_supply"),
    ("м2",                        "money_supply"),
    ("нийт зээл",                 "total_loans"),
    ("чанаргүй зээл",             "npl"),
    ("хугацаа хэтэрсэн",          "past_due_loans"),
    ("харилцах, хадгаламж",       "deposits"),
    ("инфляц",                    "inflation_mb"),
]

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s | %(levelname)-7s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("data_processor")


# PERIOD RESOLUTION
def resolve_target_period():
    if len(sys.argv) > 1 and sys.argv[1].strip().isdigit():
        month = int(sys.argv[1])
        year = (int(sys.argv[2]) if len(sys.argv) > 2
                and sys.argv[2].strip().isdigit() else datetime.now().year)
        return year, month
    best = None
    for f in RAW_DATA_DIR.glob("customs_*_*.xlsx"):
        m = re.match(r"customs_(\d{4})_(\d{2})\.xlsx", f.name)
        if m:
            best = max(best or (0, 0), (int(m.group(1)), int(m.group(2))))
    if best:
        return best
    now = datetime.now()
    return now.year, now.month


TARGET_YEAR, TARGET_MONTH = resolve_target_period()


def _f(x):
    """Loose float; '' / NaN / '-' -> None."""
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        if isinstance(x, str):
            x = x.replace(",", "").strip()
            if x in ("", "-"):
                return None
        return float(x)
    except (TypeError, ValueError):
        return None


def pct_change(cur, prev):
    if cur is None or prev in (None, 0):
        return None
    return round((cur / prev - 1.0) * 100.0, 1)


# CUSTOMS
def customs_files():
    """{(year, month): path} for every cached workbook."""
    out = {}
    for f in RAW_DATA_DIR.glob("customs_*_*.xlsx"):
        m = re.match(r"customs_(\d{4})_(\d{2})\.xlsx", f.name)
        if m:
            out[(int(m.group(1)), int(m.group(2)))] = f
    return out


def _read_sheet(xls, name, **kw):
    """Resolve sheet by exact name or prefix ('8.2' -> '8.2 бүлгээр')."""
    if name in xls.sheet_names:
        return pd.read_excel(xls, sheet_name=name, header=None, **kw)
    for s in xls.sheet_names:
        ss = str(s).strip()
        if ss == name or ss.startswith(name + " "):
            return pd.read_excel(xls, sheet_name=s, header=None, **kw)
    raise KeyError(f"sheet {name!r} not found "
                   f"(available: {xls.sheet_names})")


def _sheet1_totals(xls):
    """(export_prev, export_cur, import_prev, import_cur) in thousand USD."""
    df = _read_sheet(xls, "1")
    exp = imp = None
    for i in range(len(df)):
        c0 = str(df.iat[i, 0]) if pd.notna(df.iat[i, 0]) else ""
        c1 = str(df.iat[i, 1]) if pd.notna(df.iat[i, 1]) else ""
        line = (c0 + " " + c1).upper()
        if exp is None and ("А.ЭКСПОРТ" in line or "A.EXPORT" in line
                            or line.strip().startswith("A.ЭКСПОРТ")):
            exp = i
        if imp is None and ("Б.ИМПОРТ" in line or "B.IMPORT" in line):
            imp = i
    if exp is None or imp is None:
        raise LookupError("Export/Import rows not found in sheet 1")
    return (_f(df.iat[exp, 2]), _f(df.iat[exp, 3]),
            _f(df.iat[imp, 2]), _f(df.iat[imp, 3]))


def _first_code_row(df, code_col, code):
    """First row whose code cell equals `code` (files list the aggregate
    4-digit code before its 6/8-digit children)."""
    for r in range(len(df)):
        v = df.iat[r, code_col]
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s.endswith(".0"):        # excel sometimes floats the codes
            s = s[:-2]
        if s == code or s == code.lstrip("0"):
            return r
    return None


def _resolve_chapter_sheet(xls):
    """
    Locate the HS-chapter totals sheet, whose NAME the Customs office changes
    between months ('8.2 бүлгээр' in May 2026, '8.1 бүлгээр' in January 2026).
    Fallback chain, in order:
      1. a sheet named/prefixed '8.2'                (the usual name)
      2. any sheet whose name contains 'бүлгээр'     ('by chapter' suffix --
         this is what actually identifies the table, whatever the number)
      3. sheet '8' as a last resort (commodity-level; chapter codes may
         still resolve if Customs merged the tables)
    Returns (dataframe, sheet_name) or (None, None) -- never raises.
    """
    def _load(s):
        return pd.read_excel(xls, sheet_name=s, header=None), s
    try:
        for s in xls.sheet_names:
            ss = str(s).strip()
            if ss == "8.2" or ss.startswith("8.2 "):
                return _load(s)
        for s in xls.sheet_names:
            if "бүлгээр" in str(s).lower():
                return _load(s)
        for s in xls.sheet_names:
            if str(s).strip() == "8":
                return _load(s)
    except Exception as exc:
        log.warning("chapter sheet unreadable: %s", exc)
    return None, None


def _sheet8_commodity(df, code, side="export"):
    """{'qty_prev','amt_prev','qty_cur','amt_cur'} for an HS code."""
    r = _first_code_row(df, 2, code)
    if r is None:
        return None
    base = 5 if side == "export" else 9
    return {"qty_prev": _f(df.iat[r, base]), "amt_prev": _f(df.iat[r, base + 1]),
            "qty_cur": _f(df.iat[r, base + 2]), "amt_cur": _f(df.iat[r, base + 3])}


def extract_customs(files):
    """
    Returns (kpis, series) where series carries:
      trade_monthly, commodity_exports, commodity_imports, border_prices
    Money amounts are kept in THOUSAND USD (source unit); the Excel builder
    converts to million USD for presentation.
    """
    target_key = (TARGET_YEAR, TARGET_MONTH)
    # Historical consistency: a rebuild for month M must not include
    # workbooks newer than M in its series (prior-year months all qualify).
    files = {k: v for k, v in files.items() if k <= target_key}
    log.info("===== CUSTOMS EXTRACTION (%d workbook(s) <= %s) =====",
             len(files), f"{TARGET_YEAR}-{TARGET_MONTH:02d}")
    kpis, series = {}, {"trade_monthly": [], "commodity_exports": {},
                        "commodity_imports": {}, "border_prices": {}}
    if target_key not in files:
        log.error("Target workbook customs_%d_%02d.xlsx missing.",
                  TARGET_YEAR, TARGET_MONTH)
        return kpis, series

    # headline trade KPIs from the target workbook. A workbook that is
    # present but unreadable (interrupted download, error page saved with an
    # .xlsx name) must not abort the run: the trade cells stay flagged and
    # every other source still reaches the workbook.
    try:
        xls = pd.ExcelFile(files[target_key])
    except Exception as exc:
        log.error("Target workbook customs_%d_%02d.xlsx is unreadable (%s) -- "
                  "delete it from raw_files and re-run the ingestor. Trade "
                  "cells stay flagged.", TARGET_YEAR, TARGET_MONTH, exc)
        return kpis, series
    try:
        e_prev, e_cur, i_prev, i_cur = _sheet1_totals(xls)
        kpis.update({
            "total_export": round(e_cur, 2),
            "total_import": round(i_cur, 2),
            "trade_balance": round(e_cur - i_cur, 2),
            "total_export_prev": round(e_prev, 2),
            "total_import_prev": round(i_prev, 2),
            "trade_balance_prev": round(e_prev - i_prev, 2),
            "export_yoy_pct": pct_change(e_cur, e_prev),
            "import_yoy_pct": pct_change(i_cur, i_prev),
            "trade_balance_ratio": (round((e_cur - i_cur) / (e_prev - i_prev), 2)
                                    if (e_prev - i_prev) else None),
        })
        log.info("Trade: exp %.1f (%+.0f%%) imp %.1f (%+.0f%%) bal %.1f",
                 e_cur, kpis["export_yoy_pct"], i_cur,
                 kpis["import_yoy_pct"], e_cur - i_cur)
    except Exception as exc:
        log.warning("sheet 1 totals failed: %s", exc)

    # commodity breakdowns (target workbook)
    try:
        s8 = _read_sheet(xls, "8")
        for key, (code, label, unit) in EXPORT_COMMODITIES.items():
            d = _sheet8_commodity(s8, code, "export")
            if d:
                d.update({"label": label, "hs_code": code, "volume_unit": unit,
                          "amt_yoy_pct": pct_change(d["amt_cur"], d["amt_prev"]),
                          "qty_yoy_pct": pct_change(d["qty_cur"], d["qty_prev"])})
                series["commodity_exports"][key] = d
            else:
                log.warning("export commodity %s (HS %s) not found", key, code)
        for key, (code, label) in IMPORT_COMMODITIES.items():
            d = _sheet8_commodity(s8, code, "import")
            if d:
                d.update({"label": label, "hs_code": code,
                          "amt_yoy_pct": pct_change(d["amt_cur"], d["amt_prev"])})
                series["commodity_imports"][key] = d
    except Exception as exc:
        log.warning("sheet 8 extraction failed: %s", exc)

    # HS-chapter commodities (wool/cashmere). The sheet name is unstable
    # ('8.2 бүлгээр' vs '8.1 бүлгээр'), so resolve via fallback chain; on
    # total failure the metrics are emitted as blanks (Excel-safe NaN) so
    # excel_builder still constructs the row/formulas, and we log a clean
    # WARNING -- never a traceback.
    s82, chapter_sheet = _resolve_chapter_sheet(xls)
    for key, (chapter, label) in CHAPTER_COMMODITIES.items():
        d = {"label": label, "hs_code": f"ch{chapter}",
             "qty_prev": None, "amt_prev": None,
             "qty_cur": None, "amt_cur": None, "amt_yoy_pct": None}
        try:
            row = (_first_code_row(s82, 0, chapter)
                   if s82 is not None else None)
            if row is not None:
                d.update({"qty_prev": _f(s82.iat[row, 2]),
                          "amt_prev": _f(s82.iat[row, 3]),
                          "qty_cur": _f(s82.iat[row, 4]),
                          "amt_cur": _f(s82.iat[row, 5])})
                d["amt_yoy_pct"] = pct_change(d["amt_cur"], d["amt_prev"])
                log.info("chapter %s (%s) resolved via sheet %r",
                         chapter, key, chapter_sheet)
            else:
                log.warning("chapter %s (%s) not found -- tried '8.2*', "
                            "'*бүлгээр*', '8'. Metrics set to blank/NaN; "
                            "Excel row still built.", chapter, key)
        except Exception as exc:
            log.warning("chapter %s (%s) extraction skipped: %s",
                        chapter, key, exc)
        series["commodity_exports"][key] = d

    try:
        s3 = _read_sheet(xls, "3")
        for key, (sec, label) in SECTION_IMPORTS.items():
            row = None
            for r in range(len(s3)):
                v = s3.iat[r, 0]
                if pd.notna(v) and str(v).strip().startswith(sec + " "):
                    row = r
                    break
            if row is not None:
                d = {"label": label, "hs_code": f"sec{sec}",
                     "amt_prev": _f(s3.iat[row, 4]), "amt_cur": _f(s3.iat[row, 5])}
                d["amt_yoy_pct"] = pct_change(d["amt_cur"], d["amt_prev"])
                series["commodity_imports"][key] = d
    except Exception as exc:
        log.warning("sheet 3 extraction failed: %s", exc)

    try:
        s3f = _read_sheet(xls, "3")
        secs = []
        for r3 in range(len(s3f)):
            v3 = s3f.iat[r3, 0]
            if pd.isna(v3):
                continue
            label3 = re.sub(r"\s+", " ", str(v3).strip())
            # HS sections are numbered ('01 ...'); Customs also files an
            # unnumbered 'Бусад' (Other) line that carries real import value
            # -- without it the section totals fall short of the headline.
            if not (re.match(r"^\d{2}\s", label3) or label3.startswith("Бусад")):
                continue
            en3 = s3f.iat[r3, 1]
            secs.append({
                "mn": label3[:70],
                "en": re.sub(r"\s+", " ", str(en3).strip())[:60]
                      if pd.notna(en3) else "",
                "exp_prev": _f(s3f.iat[r3, 2]),
                "exp_cur": _f(s3f.iat[r3, 3]),
                "imp_prev": _f(s3f.iat[r3, 4]),
                "imp_cur": _f(s3f.iat[r3, 5])})
        series["trade_sections"] = secs
        log.info("Trade sections extracted: %d rows", len(secs))
        # Self-check: the section rows must add up to the headline totals.
        for side, head in (("exp_cur", kpis.get("total_export")),
                           ("imp_cur", kpis.get("total_import"))):
            tot = sum(s[side] or 0 for s in secs)
            if head:
                gap = tot - head
                if abs(gap) > max(1.0, abs(head) * 1e-6):
                    log.warning("trade sections %s sum %.1f != headline %.1f "
                                "(gap %.1f) -- a source row may be unmatched.",
                                side, tot, head, gap)
                else:
                    log.info("trade sections %s reconciles to headline "
                             "(%.1f)", side, tot)
    except Exception as exc:
        log.warning("trade sections extraction failed: %s", exc)

    # convenience KPI aliases (kept for backward compatibility with v1)
    ce = series["commodity_exports"]
    for k, alias in [("coal", "coal_value"), ("copper", "copper_value"),
                     ("gold", "gold_value"), ("iron", "iron_value"),
                     ("wool_cashmere", "wool_cashmere_value")]:
        if k in ce and ce[k].get("amt_cur") is not None:
            kpis[alias] = round(ce[k]["amt_cur"], 1)
    if "coal" in ce and ce["coal"].get("qty_cur"):
        kpis["coal_volume_mt"] = round(ce["coal"]["qty_cur"] / 1e9, 2)
    if "copper" in ce and ce["copper"].get("qty_cur"):
        kpis["copper_volume_kt"] = round(ce["copper"]["qty_cur"] / 1e6, 2)
    if "gold" in ce and ce["gold"].get("qty_cur"):
        kpis["gold_volume_kg"] = round(ce["gold"]["qty_cur"], 1)

    # monthly cumulative series across ALL cached workbooks
    monthly = {}   # (year, month) -> {"export": kUSD, "import": kUSD}
    for (y, m), path in sorted(files.items()):
        try:
            x = pd.ExcelFile(path)
            e_prev, e_cur, i_prev, i_cur = _sheet1_totals(x)
            monthly[(y, m)] = {"export": e_cur, "import": i_cur}
            # each workbook also carries last year's same-period cumulative
            monthly.setdefault((y - 1, m), {"export": e_prev, "import": i_prev})
        except Exception as exc:
            log.warning("monthly totals failed for %s: %s", path.name, exc)
    series["trade_monthly"] = [
        {"year": y, "month": m,
         "export_cum_kusd": round(v["export"], 1) if v["export"] else None,
         "import_cum_kusd": round(v["import"], 1) if v["import"] else None}
        for (y, m), v in sorted(monthly.items())]
    log.info("Monthly cumulative trade points: %d", len(series["trade_monthly"]))

    # border prices (marginal unit values) per month
    # unit value of month m = (cumAmt_m - cumAmt_{m-1}) / (cumQty_m - cumQty_{m-1})
    cum = {}   # commodity -> {(y, m): (qty, amt)}
    for (y, m), path in sorted(files.items()):
        try:
            s8m = _read_sheet(pd.ExcelFile(path), "8")
            for key, (code, _lbl, _u) in EXPORT_COMMODITIES.items():
                d = _sheet8_commodity(s8m, code, "export")
                if not d:
                    continue
                cd = cum.setdefault(key, {})
                if d["qty_cur"] and d["amt_cur"]:
                    cd[(y, m)] = (d["qty_cur"], d["amt_cur"])
                if d["qty_prev"] and d["amt_prev"]:
                    cd.setdefault((y - 1, m), (d["qty_prev"], d["amt_prev"]))
        except Exception as exc:
            log.warning("border-price pass failed for %s: %s", path.name, exc)
    for key, cd in cum.items():
        pts = []
        for (y, m), (q, a) in sorted(cd.items()):
            q0, a0 = cd.get((y, m - 1), (0.0, 0.0)) if m > 1 else (0.0, 0.0)
            dq, da = q - q0, a - a0
            if dq and dq > 0:
                if key == "gold":       # thousand USD & kg -> USD/oz
                    price = da * 1000.0 / (dq * 32.1507)
                    unit = "USD/oz"
                else:                   # thousand USD & kg -> USD/tonne
                    price = da * 1e6 / dq
                    unit = "USD/t"
                marginal = (m == 1) or ((y, m - 1) in cd)
                pts.append({"year": y, "month": m, "price": round(price, 1),
                            "unit": unit,
                            "basis": "monthly" if marginal else
                                     f"cumulative avg (months 1-{m})"})
        if pts:
            series["border_prices"][key] = pts
    log.info("Border price series: %s",
             {k: len(v) for k, v in series["border_prices"].items()})
    return kpis, series


# NSO (json-stat2)
class JsonStat2:
    def __init__(self, data):
        self.value = data["value"]
        self.ids = data["id"]
        self.sizes = data["size"]
        self.dimension = data["dimension"]
        self.strides = [1] * len(self.sizes)
        for k in range(len(self.sizes) - 2, -1, -1):
            self.strides[k] = self.strides[k + 1] * self.sizes[k + 1]

    def labels(self, dim):
        cat = self.dimension[dim]["category"]
        idx, lab = cat.get("index", {}), cat.get("label", {})
        return {int(idx.get(c, c)): l for c, l in lab.items()}

    def cell(self, **by_dim):
        pos = []
        for dim in self.ids:
            key = by_dim[dim]
            if isinstance(key, int):
                pos.append(key)
            else:
                cat = self.dimension[dim]["category"]
                idx = cat.get("index", {})
                if key in idx:
                    pos.append(int(idx[key]))
                else:
                    lab = cat.get("label", {})
                    code = next((c for c, l in lab.items() if l == key), None)
                    if code is None:
                        raise KeyError(f"{key!r} not in {dim!r}")
                    pos.append(int(idx.get(code, code)))
        return self.value[sum(p * s for p, s in zip(pos, self.strides))]


def extract_nso():
    """Real-GDP KPIs + slide-4 cumulative growth + slide-5 contributions."""
    log.info("===== NSO NATIONAL ACCOUNTS =====")
    kpis, out = {}, {"quarterly": [], "cumulative_growth": {},
                     "contributions": {}, "sectors": []}
    try:
        js = JsonStat2(json.loads(NSO_JSON.read_text(encoding="utf-8")))
    except Exception as exc:
        log.error("NSO raw file unusable: %s", exc)
        return kpis, out

    # quarterly real GDP by sector -> {(year, q): {sector_pos: value}}
    sector_labels = js.labels(NSO_DIM_SECTOR)
    out["sectors"] = [sector_labels[i] for i in sorted(sector_labels)]
    data = {}
    for pos, lbl in js.labels(NSO_DIM_PERIOD).items():
        m = re.match(r"(\d{4})-(\d)$", str(lbl).strip())
        if not m:
            continue
        y, q = int(m.group(1)), int(m.group(2))
        row = {}
        for spos in sector_labels:
            try:
                v = js.cell(**{NSO_DIM_INDICATOR: NSO_REAL_GDP_2015,
                               NSO_DIM_SECTOR: spos, NSO_DIM_PERIOD: pos})
            except Exception:
                v = None
            if v is not None:
                row[spos] = float(v)
        if row.get(NSO_TOTAL_SECTOR) is not None:
            data[(y, q)] = row

    if not data:
        log.error("No usable NSO periods decoded.")
        return kpis, out

    quarters = sorted(data)
    latest_y, latest_q = quarters[-1]
    out["quarterly"] = [{"year": y, "quarter": q,
                         "real_gdp_2015p": round(data[(y, q)][NSO_TOTAL_SECTOR], 1)}
                        for (y, q) in quarters]

    # latest-quarter YoY (headline)
    prev = data.get((latest_y - 1, latest_q))
    if prev:
        yoy = (data[(latest_y, latest_q)][NSO_TOTAL_SECTOR]
               / prev[NSO_TOTAL_SECTOR] - 1) * 100
        kpis["gdp_growth"] = round(yoy, 2)
        kpis["gdp_growth_period"] = f"{latest_y}-Q{latest_q}"
        log.info("Real GDP YoY %s-Q%s: %.2f%%", latest_y, latest_q, yoy)

    # slide 4: cumulative growth I-Q for the last 3 years
    cum_labels = {1: "I-III", 2: "I-VI", 3: "I-IX", 4: "I-XII"}
    for y in range(latest_y - 2, latest_y + 1):
        row = {}
        for q_end in range(1, 5):
            need = [(y, q) for q in range(1, q_end + 1)]
            base = [(y - 1, q) for q in range(1, q_end + 1)]
            if all(k in data for k in need + base):
                cur = sum(data[k][NSO_TOTAL_SECTOR] for k in need)
                prv = sum(data[k][NSO_TOTAL_SECTOR] for k in base)
                row[cum_labels[q_end]] = round((cur / prv - 1) * 100, 2)
        if row:
            out["cumulative_growth"][str(y)] = row

    # slide 5: sector contributions -- prior year FULL YEAR vs current year
    # latest cumulative period (this is how the deck's 2025 vs 2026 bars work)
    for y, q_end in ((latest_y - 1, 4), (latest_y, latest_q)):
        need = [(y, q) for q in range(1, q_end + 1)]
        base = [(y - 1, q) for q in range(1, q_end + 1)]
        if not all(k in data for k in need + base):
            continue
        total_base = sum(data[k][NSO_TOTAL_SECTOR] for k in base)
        contrib = {}
        for spos, sname in sector_labels.items():
            if spos == NSO_TOTAL_SECTOR:
                continue
            try:
                cur = sum(data[k].get(spos, 0.0) for k in need)
                prv = sum(data[k].get(spos, 0.0) for k in base)
                contrib[sname] = round((cur - prv) / total_base * 100, 2)
            except Exception:
                continue
        out["contributions"][str(y)] = {
            "period": {1: "I-III", 2: "I-VI", 3: "I-IX", 4: "I-XII"}[q_end],
            **contrib}
    log.info("Cumulative growth years: %s | contribution years: %s",
             list(out["cumulative_growth"]), list(out["contributions"]))
    return kpis, out


# MONGOLBANK STATISTICAL BULLETIN (replaces the retired MED PDF, 2026-07)
# Calibrated against the real bulletin workbook (id 5115, May 2026). Sheets
# are long time series: rows are periods ('2025 12', then bare '01'..'05' in
# the current year), columns carry bilingual multi-row headers. Columns are
# resolved BY HEADER TEXT, never by fixed index, so layout shifts survive.
MIEG_JSON = RAW_DATA_DIR / "nso_mieg_raw.json"


def extract_mieg():
    """Monthly Indicator of Economic Growth (NSO DT_NSO_0500_001V5):
    cumulative YoY growth per month + sector contributions, monthly."""
    log.info("===== NSO MIEG (monthly economic growth) =====")
    kpis, out = {}, {"monthly": [], "contributions": {}}
    if not MIEG_JSON.exists():
        log.warning("MIEG raw file not cached -- run raw_ingestor.py "
                    "(table DT_NSO_0500_001V5). Monthly growth stays "
                    "flagged.")
        return kpis, out
    try:
        js = JsonStat2(json.loads(MIEG_JSON.read_text(encoding="utf-8")))
    except Exception as exc:
        log.warning("MIEG raw file unusable: %s", exc)
        return kpis, out

    ind_dim, sec_dim, per_dim = js.ids[0], js.ids[1], js.ids[-1]

    def pick(dim, fragments, fallback, what=""):
        """Positional index of the first category matching any fragment.

        Fragment lists carry the wording seen live plus older/English
        variants, and the chosen label is logged so that a future relabel
        shows up in the run log instead of silently selecting the fallback.
        """
        for pos, lbl in js.labels(dim).items():
            l = str(lbl).lower()
            if any(f in l for f in fragments):
                log.info("MIEG %s -> %r", what, str(lbl).strip()[:70])
                return pos
        if fallback is None:
            log.warning("MIEG %s: no category matched %s -- skipped.",
                        what, fragments)
        else:
            log.warning("MIEG %s: no category matched %s -- falling back to "
                        "index %s (%r). Check the table labels.", what,
                        fragments, fallback,
                        str(js.labels(dim).get(fallback, "?")).strip()[:70])
        return fallback

    # Live labels (2026-07): the growth indicator is 'Эдийн засгийн өсөлт,
    # бууралт, хувиар /өмнөх оны мөн үетэй харьцуулахад/' and the sector
    # split is '...өөрчлөлтөд салбаруудын ОРОЛЦОО, нэгж хувиар' -- not the
    # 'хувь нэмэр' wording used elsewhere on 1212.mn. The third indicator
    # ('Улирлын нөлөөллийг арилгасан ДНБ') is month-on-month and must not be
    # confused with either.
    growth_ind = pick(ind_dim,
                      ["эдийн засгийн өсөлт", "өсөлт, бууралт",
                       "сарын индикатор", "monthly indicator",
                       "economic growth"], 0, "growth indicator")
    contrib_ind = pick(ind_dim,
                       ["оролцоо", "хувь нэмэр", "contribution"],
                       None, "sector contribution indicator")
    total_sec = pick(sec_dim, ["бүгд", "нийт", "днб", "gdp", "total"], 0,
                     "total sector")

    periods = []
    for pos, lbl in js.labels(per_dim).items():
        m = re.match(r"^(\d{4})[-\s](\d{1,2})$", str(lbl).strip())
        if m and 1 <= int(m.group(2)) <= 12:
            y, mth = int(m.group(1)), int(m.group(2))
            # never look past the month being reported (see note below)
            if (y, mth) <= (TARGET_YEAR, TARGET_MONTH):
                periods.append((y, mth, pos))
    periods.sort()

    for y, mth, pos in periods:
        try:
            v = js.cell(**{ind_dim: growth_ind, sec_dim: total_sec,
                           per_dim: pos})
        except Exception:
            v = None
        if v is not None and -50 <= float(v) <= 100:
            out["monthly"].append({"period": f"{y}-{mth:02d}",
                                   "value": round(float(v), 2)})
    # A report for month M must never quote a later month. NSO keeps
    # publishing, so by the time May is rebuilt the table may already hold
    # June and July; taking the newest row would put a later month's growth
    # on a May slide. Everything after the target period is dropped here and
    # the fact is logged, so the series and the headline always stop at M.
    target_tag = f"{TARGET_YEAR}-{TARGET_MONTH:02d}"
    future = [d for d in out["monthly"] if d["period"] > target_tag]
    if future:
        out["monthly"] = [d for d in out["monthly"] if d["period"] <= target_tag]
        log.info("MIEG: ignoring %d period(s) after %s (newest published is "
                 "%s) -- the report stops at its own month.",
                 len(future), target_tag, future[-1]["period"])
    if out["monthly"]:
        last = out["monthly"][-1]
        kpis["mieg_growth_pct"] = last["value"]
        kpis["_mieg_period"] = last["period"]
        log.info("MIEG cumulative growth %s = %s%% (%d monthly points)",
                 last["period"], last["value"], len(out["monthly"]))
        if last["period"] < target_tag:
            kpis["_mieg_period_note"] = (
                f"WARNING: MIEG newest period is {last['period']}, target "
                f"is {target_tag}")
            log.warning(kpis["_mieg_period_note"])
    else:
        log.warning("MIEG: no monthly periods decoded -- check the table "
                    "structure in raw_files/nso_mieg_raw.json")

    if contrib_ind is not None and periods:
        y, mth, pos = periods[-1]
        for spos, sname in js.labels(sec_dim).items():
            if spos == total_sec:
                continue
            try:
                v = js.cell(**{ind_dim: contrib_ind, sec_dim: spos,
                               per_dim: pos})
            except Exception:
                v = None
            if v is not None:
                out["contributions"][str(sname).strip()] = round(float(v), 2)
        out["contributions_period"] = f"{y}-{mth:02d}"
        if out["contributions"]:
            # The sector contributions are additive by construction, so their
            # sum must reproduce the headline growth for the same month.
            tot = sum(out["contributions"].values())
            head = kpis.get("mieg_growth_pct")
            log.info("MIEG contributions %s: %d sectors, sum %.2f pp "
                     "(headline %s%%)", out["contributions_period"],
                     len(out["contributions"]), tot, head)
            if head is not None and abs(tot - head) > 0.5:
                log.warning("MIEG contributions sum %.2f pp does not match "
                            "headline %.2f%% -- sector split may be "
                            "mismapped.", tot, head)
        else:
            log.warning("MIEG: contribution indicator found but no sector "
                        "values decoded.")
    return kpis, out


BULLETIN_XLSX = RAW_DATA_DIR / "mongolbank_bulletin.xlsx"


def _bul_sheet(xls, fragment):
    for s in xls.sheet_names:
        if fragment.lower() in str(s).lower():
            return pd.read_excel(xls, sheet_name=s, header=None)
    return None


def _bul_col(df, keywords, hdr_rows=range(5, 15)):
    """First column whose concatenated header text contains ALL keywords."""
    for c in range(df.shape[1]):
        text = " ".join(str(df.iat[r, c]) for r in hdr_rows
                        if r < len(df) and isinstance(df.iat[r, c], str))
        text = text.lower()
        if all(k in text for k in keywords):
            return c
    return None


def _bul_lending_rate_cols(df, hdr_rows=range(0, 16)):
    """
    The four tugrik lending-rate columns on the bulletin's 'Rate' sheet.

    The sheet has a 'Зээлийн хүү / Loan rate' block split into loans ISSUED
    in the month and loans OUTSTANDING, and each of those gives the tugrik
    rate twice -- at market rates, then again with subsidised programme and
    project loans included. Four columns, all captioned 'Төгрөгийн', so a
    keyword match alone cannot separate them. They are located by walking the
    block: the sub-block caption says issued or outstanding, and within it the
    second tugrik column is the one carrying the parenthetical about
    below-market programme loans.
    """
    def text(c):
        return " ".join(str(df.iat[r, c]) for r in hdr_rows
                        if r < len(df) and isinstance(df.iat[r, c], str)).lower()

    cols, block = {}, None
    for c in range(df.shape[1]):
        t = text(c)
        if "олгосон зээлийн" in t or "lending rates (issued)" in t:
            block = "new"
        elif "үлдэгдэлд жигнэсэн" in t or "lending rates (outstanding)" in t:
            block = "stock"
        elif "зээлийн хүү" in t and "loan rate" in t and block is None:
            continue
        if block is None or "төгрөгийн" not in t:
            continue
        incl = "доогуур хүүгээр" in t or "хөтөлбөрийн зээлийг" in t
        key = f"loan_rate_{block}_mnt" + ("_incl" if incl else "")
        cols.setdefault(key, c)
    if len(cols) < 2:
        log.warning("bulletin: found %d of the 4 tugrik lending-rate columns "
                    "on the Rate sheet (%s) -- the layout may have changed.",
                    len(cols), ", ".join(cols) or "none")
    return cols


def _bul_last_row(df):
    """(row, year, month) of the newest period row. Explicit 'YYYY MM' labels
    set the year; bare 'MM' labels continue it (rolling over past December)."""
    year = month = last = None
    for r in range(len(df)):
        v = str(df.iat[r, 0]).strip()
        m = re.match(r"^(\d{4})\s+(\d{1,2})$", v)
        if m:
            year, month, last = int(m.group(1)), int(m.group(2)), r
        elif re.match(r"^\d{1,2}$", v) and year is not None:
            mm = int(v)
            if month is not None and mm < month:
                year += 1
            month, last = mm, r
    return last, year, month


def extract_bulletin():
    """
    Mongolbank statistical bulletin -> headline inflation (national, annual),
    POLICY RATE, and budget execution (revenue, expenditure, balance).
    Values are read from the newest period row of each sheet; every figure
    was verified against the retired MED PDF for May 2026 on first run.
    """
    log.info("===== MONGOLBANK STATISTICAL BULLETIN =====")
    kpis, notes = {}, {}
    if not BULLETIN_XLSX.exists():
        log.warning("Bulletin xlsx not cached -- run raw_ingestor.py "
                    "(report 'bulletin', stat.mongolbank.mn/bulletin). "
                    "Inflation/budget cells stay flagged.")
        return kpis, notes
    try:
        xls = pd.ExcelFile(BULLETIN_XLSX)

        # inflation: 'Inflation Nat', column 'Annual changes'
        df = _bul_sheet(xls, "inflation nat")
        if df is not None:
            c = _bul_col(df, ["annual", "changes"])
            r, y, mth = _bul_last_row(df)
            v = _f(df.iat[r, c]) if (c is not None and r is not None) else None
            if v is not None and -10 <= v <= 60:
                kpis["inflation"] = round(v, 1)
                kpis["_bulletin_period"] = f"{y}-{mth:02d}"
                if (y, mth) != (TARGET_YEAR, TARGET_MONTH):
                    kpis["_bulletin_period_note"] = (
                        f"WARNING: bulletin data is for {y}-{mth:02d}, "
                        f"target is {TARGET_YEAR}-{TARGET_MONTH:02d} (the "
                        "newer issue was not retrievable)")
                    log.warning(kpis["_bulletin_period_note"])
                log.info("bulletin inflation (national, YoY) = %s%% (%s-%02d)",
                         kpis["inflation"], y, mth)

        # policy rate: 'Rate' sheet, column 'Policy rate'
        df = _bul_sheet(xls, "rate")
        if df is not None:
            c = _bul_col(df, ["policy rate"])
            r, y, mth = _bul_last_row(df)
            v = _f(df.iat[r, c]) if (c is not None and r is not None) else None
            if v is not None and 0 < v <= 30:
                kpis["policy_rate"] = round(v, 2)
                kpis["_policy_rate_source"] = "MB bulletin, Rate sheet"
                log.info("bulletin policy rate = %s%% (%s-%02d)", v, y, mth)

        # budget: 'Budget' sheet (million MNT -> billion)
        df = _bul_sheet(xls, "budget")
        if df is not None:
            r, y, mth = _bul_last_row(df)
            cols = {
                "budget_revenue_collected_bln": ["нийт орлого", "дүн"],
                "budget_expenditure_bln":       ["зарлага", "дүн"],
                "budget_balance_bln_mnt":       ["зөрүү"],
            }
            for key, kws in cols.items():
                c = _bul_col(df, kws)
                v = _f(df.iat[r, c]) if (c is not None and r is not None)                     else None
                if v is not None:
                    kpis[key] = round(v / 1000.0, 1)   # mln -> bln MNT
            log.info("bulletin budget (bln MNT): revenue %s | expenditure %s "
                     "| balance %s (%s-%02d)",
                     kpis.get("budget_revenue_collected_bln"),
                     kpis.get("budget_expenditure_bln"),
                     kpis.get("budget_balance_bln_mnt"), y, mth)

        # monthly SERIES for deck slides 6 & 8 (last 16 months)
        def col_series(df, col, n=16):
            """[(year, month, value)] for the newest n period rows."""
            out, year, month = [], None, None
            for r in range(len(df)):
                v = str(df.iat[r, 0]).strip()
                m2 = re.match(r"^(\d{4})\s+(\d{1,2})$", v)
                if m2:
                    year, month = int(m2.group(1)), int(m2.group(2))
                elif re.match(r"^\d{1,2}$", v) and year is not None:
                    mm = int(v)
                    if month is not None and mm < month:
                        year += 1
                    month = mm
                else:
                    continue
                val = _f(df.iat[r, col])
                if val is not None:
                    out.append((year, month, val))
            return out[-n:]

        bser = {}
        specs = [
            ("inflation_nat", "inflation nat", ["annual", "changes"], 1),
            ("inflation_ub",  "inflation ub",  ["annual", "changes"], 1),
            ("policy_rate",   "rate",          ["policy rate"], 2),
            ("deposit_rate",  "rate",          ["шинэ", "төгрөгийн"], 2),
            ("usd_mnt_eop",   "exchange rate", ["usd", "эцэст"], 2),
        ]
        for key, sheet_frag, kws, nd in specs:
            df2 = _bul_sheet(xls, sheet_frag)
            c2 = _bul_col(df2, kws) if df2 is not None else None
            if c2 is not None:
                pts = col_series(df2, c2)
                if pts:
                    bser[key] = [{"period": f"{y}-{m:02d}",
                                  "value": round(v, nd)} for y, m, v in pts]

        # LENDING RATES -- there are two of each, and the difference is the
        # point, not an error. Mongolbank publishes the weighted average
        # tugrik lending rate twice: once for loans at market rates, and once
        # "зах зээлийн хүүнээс доогуур хүүгээр олгосон төсөл, хөтөлбөрийн
        # зээлийг оруулснаар" -- with subsidised programme and project loans
        # folded in. Those are cheaper, so the second figure sits below the
        # first, and the gap is how much the subsidised programmes pull the
        # cost of borrowing down. The deck should show both.
        df2 = _bul_sheet(xls, "rate")
        if df2 is not None:
            for key, cidx in _bul_lending_rate_cols(df2).items():
                pts = col_series(df2, cidx)
                if pts:
                    bser[key] = [{"period": f"{y}-{m:02d}",
                                  "value": round(v, 2)} for y, m, v in pts]
            pair = (bser.get("loan_rate_new_mnt"),
                    bser.get("loan_rate_new_mnt_incl"))
            if all(pair) and pair[0][-1]["period"] == pair[1][-1]["period"]:
                a, b = pair[0][-1]["value"], pair[1][-1]["value"]
                kpis["loan_rate_new_mnt_pct"] = a
                kpis["loan_rate_new_mnt_incl_pct"] = b
                kpis["loan_rate_subsidy_gap_pp"] = round(a - b, 2)
                log.info("bulletin lending rate (%s): market %.2f%% vs "
                         "%.2f%% including subsidised programme loans "
                         "-> the programmes pull it down by %.2f pp.",
                         pair[0][-1]["period"], a, b, a - b)
                if b > a:
                    log.warning("lending rates: the 'including programme "
                                "loans' rate (%.2f%%) is ABOVE the market "
                                "rate (%.2f%%) -- the two columns may have "
                                "swapped in the bulletin.", b, a)
        notes["series"] = bser
        if bser:
            log.info("bulletin series: %s",
                     {k: len(v) for k, v in bser.items()})

        # exchange rate KPI (slide 3: USD/MNT)
        pts = bser.get("usd_mnt_eop", [])
        if pts:
            kpis["usd_mnt_rate"] = pts[-1]["value"]
            prev = next((p["value"] for p in pts
                         if p["period"] ==
                         f"{int(pts[-1]['period'][:4]) - 1}"
                         f"{pts[-1]['period'][4:]}"), None)
            if prev:
                kpis["usd_mnt_yoy_pct"] = round(
                    (pts[-1]["value"] / prev - 1) * 100, 1)
            log.info("bulletin USD/MNT eop = %s (YoY %s%%)",
                     kpis["usd_mnt_rate"], kpis.get("usd_mnt_yoy_pct"))

        # depository-corporation loans (slide 16: growth by borrower)
        df2 = _bul_sheet(xls, "dc loan")
        if df2 is not None:
            for key, kws in (("dc_loans_total", ["total loans", "amount"]),
                             ("dc_loans_individuals", ["иргэд"])):
                c2 = _bul_col(df2, kws)
                if c2 is None:
                    continue
                pts = col_series(df2, c2)
                if not pts:
                    continue
                y, m, v = pts[-1]
                kpis[f"{key}_tln_mnt"] = round(v / 1e6, 2)   # mln -> tln
                prev = next((vv for (yy, mm, vv) in pts
                             if (yy, mm) == (y - 1, m)), None)
                if prev:
                    kpis[f"{key}_yoy_pct"] = round((v / prev - 1) * 100, 1)
            log.info("bulletin DC loans: total %s tln (%s%%) | "
                     "individuals %s tln (%s%%)",
                     kpis.get("dc_loans_total_tln_mnt"),
                     kpis.get("dc_loans_total_yoy_pct"),
                     kpis.get("dc_loans_individuals_tln_mnt"),
                     kpis.get("dc_loans_individuals_yoy_pct"))

        # CPI component breakdown: category columns sum to the general
        # index, so each category's pp contribution to headline inflation
        # = (level_now - level_year_ago) / general_year_ago * 100
        df = _bul_sheet(xls, "inflation nat")
        if df is not None:
            gen_c = _bul_col(df, ["general cpi"])
            rowmap, yr, mo = {}, None, None
            for r2 in range(len(df)):
                v2 = str(df.iat[r2, 0]).strip()
                m2 = re.match(r"^(\d{4})\s+(\d{1,2})$", v2)
                if m2:
                    yr, mo = int(m2.group(1)), int(m2.group(2))
                    rowmap[(yr, mo)] = r2
                elif re.match(r"^\d{1,2}$", v2) and yr is not None:
                    mm2 = int(v2)
                    if mo is not None and mm2 < mo:
                        yr += 1
                    mo = mm2
                    rowmap[(yr, mo)] = r2
            if gen_c and rowmap:
                # The same decomposition, for any month in the sheet. The
                # deck's table shows this year beside last year, and the
                # bulletin carries every month's history -- so the prior year
                # is a second call, not a second download.
                def _cpi_at(y, m):
                    rt, rp = rowmap.get((y, m)), rowmap.get((y - 1, m))
                    if rt is None or rp is None:
                        return []
                    gp = _f(df.iat[rp, gen_c])
                    if not gp:
                        return []
                    out = []
                    for c2 in range(1, gen_c):
                        hdr = " ".join(
                            str(df.iat[hr, c2]).strip()
                            for hr in range(6, 15)
                            if hr < len(df)
                            and isinstance(df.iat[hr, c2], str))
                        hdr = re.sub(r"\s+", " ", hdr).strip()
                        if not hdr or "эцэст" in hdr.lower() \
                                or "end-of-period" in hdr.lower():
                            continue
                        vt, vp = _f(df.iat[rt, c2]), _f(df.iat[rp, c2])
                        if vt is None or vp in (None, 0):
                            continue
                        out.append({
                            "name": hdr[:80],
                            "level": round(vt, 2),
                            "pp": round((vt - vp) / gp * 100, 2),
                            "yoy_pct": round((vt / vp - 1) * 100, 1)})
                    return out

                (ly, lm) = max(rowmap)
                comps = _cpi_at(ly, lm)
                if comps:
                    comps.sort(key=lambda d: -d["pp"])
                    notes["cpi_components"] = comps
                    notes["cpi_components_period"] = f"{ly}-{lm:02d}"
                    total_pp = sum(d["pp"] for d in comps)
                    log.info("CPI components: %d categories, pp sum %.2f "
                             "(headline %s)", len(comps), total_pp,
                             kpis.get("inflation"))
                    # the same month a year earlier, keyed by category name so
                    # the two years line up even if the sheet reorders columns
                    prev = _cpi_at(ly - 1, lm)
                    if prev:
                        notes["cpi_components_prev"] = prev
                        notes["cpi_components_prev_period"] = \
                            f"{ly - 1}-{lm:02d}"
                        log.info("CPI components %d-%02d (prior year): %d "
                                 "categories, pp sum %.2f", ly - 1, lm,
                                 len(prev), sum(d["pp"] for d in prev))
                    else:
                        log.info("CPI components: no prior-year comparison "
                                 "(the bulletin does not reach %d-%02d).",
                                 ly - 2, lm)
                    if kpis.get("inflation") is not None and \
                            abs(total_pp - kpis["inflation"]) > 0.5:
                        log.warning("CPI component pp sum deviates from "
                                    "headline by more than 0.5pp -- check "
                                    "category columns.")
        kpis["_bulletin_source"] = BULLETIN_XLSX.name
        if "inflation" in kpis:
            kpis["inflation_target"] = INFLATION_TARGET_PCT
        missing = [k for k in ("inflation", "policy_rate",
                               "budget_balance_bln_mnt") if k not in kpis]
        if missing:
            log.warning("bulletin: could not resolve %s -- header layout "
                        "may have changed; check the sheet headers.",
                        ", ".join(missing))
    except Exception as exc:
        log.warning("bulletin parsing failed: %s", exc)
    return kpis, notes


# MONGOLBANK CARDS (all swept sections)
def _parse_card_value(text):
    m = re.search(r"(-?\d[\d,]*(?:\.\d+)?)", str(text))
    return float(m.group(1).replace(",", "")) if m else None


def extract_mongolbank():
    log.info("===== MONGOLBANK CARDS =====")
    kpis, cards = {}, []
    for f in sorted(RAW_DATA_DIR.glob("mongolbank_*.json")):
        try:
            payload = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        for card in payload.get("result", []) or []:
            name = (card.get("name") or "").strip()
            data = (card.get("data") or "").strip()
            if not name or not data:
                continue
            entry = {"name": name, "data": data,
                     "date": card.get("data_date"),
                     "value": _parse_card_value(data),
                     "section_file": f.name}
            cards.append(entry)
            lname = name.lower()
            # the rates card carries two figures: loan and deposit rate
            if "хүү" in lname and "хугацаа" in lname:
                m = re.search(r"зээлийн хүү\s*([\d.]+)", data)
                if m and "loan_rate" not in kpis:
                    kpis["loan_rate"] = float(m.group(1))
                    kpis["_loan_rate_date"] = entry["date"]
                m = re.search(r"хадгаламжийн хүү\s*([\d.]+)", data)
                if m and "deposit_rate" not in kpis:
                    kpis["deposit_rate"] = float(m.group(1))
                    kpis["_deposit_rate_date"] = entry["date"]
                continue
            for frag, key in MONGOLBANK_CARD_MAP:
                if frag in lname and key not in kpis and entry["value"] is not None:
                    kpis[key] = entry["value"]
                    kpis[f"_{key}_date"] = entry["date"]
                    log.info("%-30s -> %s = %s (%s)", name, key,
                             entry["value"], entry["date"])
                    break
    if "policy_rate" not in kpis:
        log.warning("Policy rate not found in any Mongolbank card. "
                    "(Add its section to the ingestor sweep, or set manually.)")
    log.info("Cards indexed: %d", len(cards))
    return kpis, cards


# MONGOLBANK CONSOLIDATED BANK BALANCE SHEET (xlsx, full monthly columns)
# Rows are matched by LABEL (exact, stripped), never by row number; the row
# codes (а16, п74, ...) shown in comments are for human orientation only.
BS_ROW_LABELS = {
    "bs_total_assets":   "НИЙТ АКТИВ",              # а82
    "bs_total_loans":    "ДОТООДЫН ЗЭЭЛ",           # а16 (а15 is '/цэвэр/')
    "bs_cb_securities":  "ТӨВ БАНКНЫ ҮНЭТ ЦААС",    # а06
    "bs_past_due_loans": "Анхаарал хандуулах зээл", # а29 (special mention)
    "bs_npl":            "Чанаргүй зээл",           # а40 (first = aggregate)
    "bs_equity":         "ӨӨРИЙН ХӨРӨНГӨ",          # п90 (total; п85 is
                                                    # current-year profit)
    "bs_total_liabilities": "НИЙТ ПАССИВ",           # п78
    "bs_deposits_current": "Аж ахуйн нэгж, иргэдийн харилцах",   # п01
    "bs_deposits_savings": "Аж ахуйн нэгж, иргэдийн хадгаламж",  # п12
}

# Two independent ways to find each line, because BOTH drift between
# releases: labels get re-cased (п01/п12 arrived UPPERCASED in the 2026-06
# file) and row codes get renumbered when rows are inserted (НИЙТ ПАССИВ was
# п74 in the 2026-05 file, п78 in the 2026-06 one). The normalised label is
# therefore the primary key -- it survives re-casing -- and the code below is
# a last-resort fallback for the release where a label itself is reworded.
# The accounting identity at the end of the parse is what proves the match.
BS_ROW_CODES = {
    "bs_total_assets":     "а82",
    "bs_total_loans":      "а16",
    "bs_cb_securities":    "а06",
    "bs_past_due_loans":   "а29",
    "bs_npl":              "а40",
    "bs_equity":           "п90",
    "bs_total_liabilities": "п78",
    "bs_deposits_current": "п01",
    "bs_deposits_savings": "п12",
}


def _bs_norm(s) -> str:
    """Case/space-insensitive label key (Cyrillic-safe via casefold)."""
    return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip().casefold()


# DEPOSIT TREE (liability side of the balance sheet)
# The portal's 'Банкны салбарын тайлан тэнцэл' page lets the reader expand
# current accounts and deposits by currency, tenor and holder. The same tree
# is in this workbook, one row per node. Depth is NOT indicated by
# indentation here and the row codes renumber between releases, so the level
# is derived from the wording -- an ALL-CAPS line opens a block, 'хугацаа...'
# is a tenor tier, a currency line sits under it, and the four holder names
# are always leaves. Every node is then checked against the sum of its own
# children, which is what proves the tree was read correctly.
BS_DEPOSIT_BLOCKS = ["аж ахуйн нэгж, иргэдийн харилцах",
                     "аж ахуйн нэгж, иргэдийн хадгаламж"]
BS_HOLDERS = {"улсын байгууллага", "хувийн байгууллага", "иргэд",
              "бусад байгууллага"}


# The portal's own balance-sheet export (stat.mongolbank.mn/finance) carries
# the SAME tree but one column per month, where the downloadable report only
# has the few date columns Mongolbank chose to include. When that export is
# present it is preferred for the deposit block; the report remains the
# fallback so the sheet is never empty.
# Preference order matters: anything the ingestor fetched this run is newer
# than a file someone downloaded by hand weeks ago, so the manual drop-in is
# the LAST resort rather than the first. It also keeps the log honest -- a
# run that says 'deposits_manual.xlsx' really did fail to fetch.
DEPOSIT_EXPORT_XLSX = "deposits_export.xlsx"        # fetched, workbook body
DEPOSIT_EXPORT_JSON = "deposits_export.json"        # fetched, JSON body
DEPOSIT_MANUAL_XLSX = "deposits_manual.xlsx"        # hand-downloaded fallback
_DEPOSIT_PERIOD = re.compile(r"(20\d{2})[-/.](\d{1,2})")


def _deposit_json_tree(path):
    """
    Read the indicator API response into (columns, nodes).

    Response shape (captured live 2026-07):
        {"success": true, "controller": "indicator_data",
         "result": {"report": [ {...one row per indicator...} ],
                    "reporttype": 1, "cycleid": 3, "interval": 3}}
    Each row carries NAME_T (caption), LEVEL_T (depth in the tree),
    PARENT_ID_T, and one key per period spelled "'YYYY-MM#3'" -- quotes and
    cycle suffix included -- whose value is a thousands-separated STRING.
    Periods are matched by pattern rather than by exact key so the quoting
    or the cycle suffix can change without breaking the read.
    """
    js = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(js, dict) and js.get("success") is False:
        raise ValueError("API reported success=false")
    res = js.get("result") if isinstance(js, dict) else None
    rows = res.get("report") if isinstance(res, dict) else None
    if not isinstance(rows, list) or not rows:
        raise ValueError("result.report missing or empty")

    def num(v):
        if v is None:
            return None
        s = str(v).replace(",", "").replace("\xa0", "").strip()
        if not s or s in {"-", "--"}:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    # period key -> 'YYYY-MM', preserving the order the API returned them
    period_keys = {}
    for row in rows:
        for k in row:
            if k in period_keys:
                continue
            m = _DEPOSIT_PERIOD.fullmatch(str(k).strip("'\" ").split("#")[0])
            if m:
                period_keys[k] = f"{m.group(1)}-{int(m.group(2)):02d}"
    if not period_keys:
        raise ValueError("no period columns found in result.report")
    ordered = sorted(period_keys, key=lambda k: period_keys[k])
    cols = [period_keys[k] for k in ordered]

    levels = [int(r.get("LEVEL_T", 0) or 0) for r in rows]
    base = min(levels) if levels else 0
    nodes = []
    for row, lvl in zip(rows, levels):
        label = str(row.get("NAME_T") or "").strip()
        if not label:
            continue
        vals = [num(row.get(k)) for k in ordered]
        if not any(v is not None for v in vals):
            continue
        nodes.append({"level": max(0, lvl - base),
                      "label": re.sub(r"\s+", " ", label),
                      "values": vals,
                      "code": str(row.get("ID_T", "") or "")})
    if not nodes:
        raise ValueError("result.report carried no usable rows")
    return cols, nodes


def _deposit_export_tree(path):
    """
    Parse a portal balance-sheet export into (columns, nodes).

    Depth comes from the leading non-breaking spaces the export uses, so an
    extra tier upstream widens the tree instead of flattening it.
    """
    df = pd.read_excel(path, sheet_name=0, header=None)
    hdr, cols = None, {}
    for r in range(min(6, len(df))):
        found = {}
        for c in range(1, df.shape[1]):
            m = _DEPOSIT_PERIOD.search(str(df.iat[r, c]))
            if m:
                found[c] = f"{m.group(1)}-{m.group(2)}"
        if len(found) >= 2:
            hdr, cols = r, found
            break
    if hdr is None:
        raise ValueError("no monthly header row found")
    ordered = sorted(cols)

    depths = sorted({_bop_indent(str(df.iat[r, 0]))
                     for r in range(hdr + 1, len(df))
                     if isinstance(df.iat[r, 0], str) and df.iat[r, 0].strip()})
    level_of = {d: i for i, d in enumerate(depths)}

    nodes = []
    for r in range(hdr + 1, len(df)):
        raw = df.iat[r, 0]
        if not isinstance(raw, str) or not raw.strip():
            continue
        vals = [_f(df.iat[r, c]) for c in ordered]
        if not any(v is not None for v in vals):
            continue
        nodes.append({"level": level_of.get(_bop_indent(raw), 0),
                      "label": re.sub(r"\s+", " ", raw.strip()),
                      "values": vals, "code": ""})
    return [cols[c] for c in ordered], nodes


def _bs_deposit_tree(df, ordered):
    """[{level, label, values, code}] for the deposit blocks, in file order."""
    def vals(r):
        out = [_f(df.iat[r, c]) for c in ordered]
        return out if any(v is not None for v in out) else None

    nodes, block, seen_tenor, last_parent = [], None, False, 0
    for r in range(len(df)):
        raw = df.iat[r, 1]
        if not isinstance(raw, str) or not raw.strip():
            continue
        lab = re.sub(r"\s+", " ", raw.strip())
        low = lab.casefold()
        alpha = [ch for ch in lab if ch.isalpha()]
        is_caps = bool(alpha) and all(ch.isupper() for ch in alpha)
        # A block opens on its caption in ANY casing: the 2026-05 release
        # writes these two headings in sentence case and the 2026-06 one in
        # capitals, so casing cannot be the trigger.
        if low in BS_DEPOSIT_BLOCKS:
            block, seen_tenor, last_parent = low, False, 0
            v = vals(r)
            if v:
                nodes.append({"level": 0, "label": lab, "values": v,
                              "code": str(df.iat[r, 0]).strip()})
            continue
        if block is None:
            continue
        # A block closes at the next section: either a capitalised heading or
        # any caption that is not one of the tree's own words. Relying on
        # capitals alone would run past the end in the sentence-case release.
        if is_caps or not (low in BS_HOLDERS or low.startswith("хугацаа")
                           or low.startswith("төгрөгийн")
                           or low.startswith("гадаад валютын")
                           or low.startswith("валютын")):
            block = None
            continue
        v = vals(r)
        if v is None:
            continue
        if low in BS_HOLDERS:
            lvl = last_parent + 1
        elif low.startswith("хугацаа"):
            lvl, seen_tenor, last_parent = 1, True, 1
        else:                       # currency line
            lvl = 2 if seen_tenor else 1
            last_parent = lvl
        nodes.append({"level": lvl, "label": lab, "values": v,
                      "code": str(df.iat[r, 0]).strip()})
    return nodes


def _deposit_tree_plausible(nodes) -> str:
    """
    '' if the parsed rows really are the deposit tree, else the reason.

    The sum check alone is not enough to accept a source: a tree with no
    parent/child pairs passes it vacuously, so a truncated or wrongly shaped
    response could slip through looking valid. This asserts the shape first.
    """
    if len(nodes) < 10:
        return f"only {len(nodes)} rows"
    tops = [n for n in nodes if n.get("level") == 0]
    if len(tops) < 2:
        return f"{len(tops)} top-level block(s), expected at least 2"
    joined = " ".join(n.get("label", "").casefold() for n in tops)
    if "харилцах" not in joined or "хадгаламж" not in joined:
        return "top-level blocks are not харилцах / хадгаламж"
    levels = {n.get("level") for n in nodes}
    if len(levels) < 3:
        return f"only {len(levels)} tier(s) of hierarchy"
    return ""


def _check_deposit_tree(nodes, col=-1):
    """Each node must equal the sum of its immediate children."""
    bad = 0
    for i, n in enumerate(nodes):
        kids = []
        for m in nodes[i + 1:]:
            if m["level"] <= n["level"]:
                break
            if m["level"] == n["level"] + 1:
                kids.append(m)
        if not kids:
            continue
        parent = n["values"][col]
        tot = sum((k["values"][col] or 0) for k in kids)
        if parent is None:
            continue
        if abs(tot - parent) > max(1.0, abs(parent) * 1e-4):
            bad += 1
            log.warning("deposit tree: %r = %.1f but its %d components sum "
                        "to %.1f -- the hierarchy may have changed.",
                        n["label"][:40], parent, len(kids), tot)
    return bad


def _bs_date_columns(df):
    """{col_index: 'YYYY.MM.DD'} from the header rows (date-labelled cols)."""
    pat = re.compile(r"^(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})")
    for r in range(min(10, len(df))):
        cols = {}
        for c in range(2, df.shape[1]):
            v = df.iat[r, c]
            s = str(v).strip()
            if pat.match(s):
                cols[c] = s
            elif hasattr(v, "strftime"):
                cols[c] = v.strftime("%Y.%m.%d")
        if len(cols) >= 2:
            return cols
    return {}


def extract_banking_balance_sheet():
    """
    Parse mongolbank_banking_balance_sheet.xlsx (million MNT, one column per
    month-end) into slide-16 metrics: latest level, level at the first column
    of the report, and the full monthly series per metric. Deposits =
    current accounts + savings (п01 + п12).
    """
    log.info("===== MONGOLBANK BANK BALANCE SHEET =====")
    kpis, series = {}, {"columns": [], "rows": {}}
    path = RAW_DATA_DIR / "mongolbank_banking_balance_sheet.xlsx"
    if not path.exists():
        log.warning("Balance sheet xlsx not cached -- run raw_ingestor.py.")
        return kpis, series
    try:
        df = pd.read_excel(path, sheet_name=0, header=None)
        datecols = _bs_date_columns(df)
        if not datecols:
            raise ValueError("no date-labelled columns found")
        ordered = sorted(datecols)          # column order == chronological
        # The report often already contains a month later than the one being
        # reported. Every balance-sheet KPI is read from the LAST column, so
        # leaving those in would put the following month's assets, loans and
        # NPL on this month's sheet.
        tag = f"{TARGET_YEAR}-{TARGET_MONTH:02d}"
        keep = [c for c in ordered
                if str(datecols[c]).replace(".", "-")[:7] <= tag]
        if keep and len(keep) < len(ordered):
            dropped = [datecols[c] for c in ordered if c not in keep]
            log.info("balance sheet: ignoring %d column(s) after %s (%s).",
                     len(dropped), tag, ", ".join(dropped))
            ordered = keep
        elif not keep:
            log.warning("balance sheet: every column is later than %s -- "
                        "using the earliest available.", tag)
            ordered = ordered[:1]
        series["columns"] = [datecols[c] for c in ordered]

        def _row_values(r):
            vals = [_f(df.iat[r, c]) for c in ordered]
            return vals if any(x is not None for x in vals) else None

        def find_row(key, label):
            """Normalised label first (survives re-casing); code as fallback.

            Rows whose 'values' are the date header (e.g. the ПАССИВ and
            ӨӨРИЙН ХӨРӨНГӨ section banners repeat the label) yield no numbers
            and are skipped by _row_values, so the numeric line always wins.
            """
            target = _bs_norm(label)
            for r in range(len(df)):
                if _bs_norm(df.iat[r, 1]) == target:
                    vals = _row_values(r)
                    if vals is not None:
                        return vals, "label"
            code = _bs_norm(BS_ROW_CODES.get(key, ""))
            if code:
                for r in range(len(df)):
                    if _bs_norm(df.iat[r, 0]) == code:
                        vals = _row_values(r)
                        if vals is not None:
                            return vals, "code"
            return None, ""

        found = {}
        for key, label in BS_ROW_LABELS.items():
            vals, how = find_row(key, label)
            if vals is None:
                log.warning("balance-sheet row %r (%s) not found by code %r "
                            "or label", label, key, BS_ROW_CODES.get(key, "-"))
                continue
            if how == "code":
                log.info("balance-sheet %s matched by code %r (label %r "
                         "reworded upstream)", key, BS_ROW_CODES.get(key), label)
            found[key] = vals
            series["rows"][key] = {"label": label, "values": vals}
        # deposits = current + savings
        if "bs_deposits_current" in found and "bs_deposits_savings" in found:
            dep = [(a or 0) + (b or 0) for a, b in
                   zip(found["bs_deposits_current"],
                       found["bs_deposits_savings"])]
            found["bs_deposits"] = dep
            series["rows"]["bs_deposits"] = {
                "label": "Харилцах + хадгаламж (ААН, иргэд)", "values": dep}

        # Self-check: assets must equal liabilities + equity. Catches a row
        # mapped to the wrong line (e.g. current-year profit read as equity).
        a = (found.get("bs_total_assets") or [None])[-1]
        li = (found.get("bs_total_liabilities") or [None])[-1]
        eq = (found.get("bs_equity") or [None])[-1]
        if None not in (a, li, eq):
            gap = a - (li + eq)
            if abs(gap) > max(1.0, abs(a) * 1e-6):
                log.warning("balance-sheet identity FAILED: assets %.1f != "
                            "liabilities %.1f + equity %.1f (gap %.1f) -- a row "
                            "label/code may have moved.", a, li, eq, gap)
            else:
                log.info("balance-sheet identity OK (assets = liabilities + "
                         "equity, gap %.2f)", gap)
        else:
            log.warning("balance-sheet identity not checked (a row is missing)")

        first_lbl, last_lbl = series["columns"][0], series["columns"][-1]
        for key in ("bs_total_assets", "bs_total_loans", "bs_cb_securities",
                    "bs_past_due_loans", "bs_npl", "bs_equity",
                    "bs_total_liabilities", "bs_deposits"):
            vals = found.get(key)
            if not vals:
                continue
            kpis[f"{key}_mln_mnt"] = round(vals[-1], 1)
            kpis[f"{key}_first_mln_mnt"] = round(vals[0], 1)
        kpis["_bs_first_column"] = first_lbl
        kpis["_bs_last_column"] = last_lbl
        kpis["_bs_change_basis"] = (
            f"change measured {first_lbl} -> {last_lbl} (earliest column in "
            "the report; not necessarily 31 Dec)")
        log.info("Balance sheet parsed: %d metrics, columns %s..%s",
                 len(series["rows"]), first_lbl, last_lbl)

        # deposit tree (portal: Нийт пассив -> харилцах / хадгаламж).
        # The portal export wins when available because it carries a column
        # per month; the report's own columns are the fallback.
        tree, dep_cols, src = [], list(series["columns"]), "report"
        # candidates in order of preference: fetched/dropped workbook, then
        # the indicator-API JSON, then the report itself
        cands = [(RAW_DATA_DIR / DEPOSIT_EXPORT_XLSX, _deposit_export_tree),
                 (RAW_DATA_DIR / DEPOSIT_EXPORT_JSON, _deposit_json_tree),
                 (RAW_DATA_DIR / DEPOSIT_MANUAL_XLSX, _deposit_export_tree)]
        for path, reader in cands:
            if not path.exists():
                continue
            try:
                c, t = reader(path)
            except Exception as exc:
                log.warning("deposit source %s unusable (%s) -- trying the "
                            "next one.", path.name, exc)
                continue
            # a source must look like the deposit tree AND reproduce its
            # arithmetic before it is trusted
            why = _deposit_tree_plausible(t)
            if why:
                log.warning("deposit source %s does not look like the deposit "
                            "tree (%s) -- discarded, trying the next one.",
                            path.name, why)
                continue
            if _check_deposit_tree(t):
                log.warning("deposit source %s failed its parent/child sum "
                            "check -- discarded, trying the next one.",
                            path.name)
                continue
            # A deposit source can reach past the month being reported --
            # the fetched range ends at the target, but a hand-downloaded
            # export or a later report release may not. Trim so a May
            # workbook can never display a June column.
            tag = f"{TARGET_YEAR}-{TARGET_MONTH:02d}"
            keep = [i for i, p in enumerate(c)
                    if str(p).replace(".", "-")[:7] <= tag]
            if keep and len(keep) < len(c):
                log.info("deposit source %s: dropping %d column(s) after %s.",
                         path.name, len(c) - len(keep), tag)
                c = [c[i] for i in keep]
                for n in t:
                    n["values"] = [n["values"][i] for i in keep]
            dep_cols, tree, src = c, t, path.name
            break
        if not tree:
            tree, dep_cols, src = (_bs_deposit_tree(df, ordered),
                                   list(series["columns"]), "report")
        if tree:
            series["deposit_tree"] = tree
            series["deposit_columns"] = dep_cols
            bad = _check_deposit_tree(tree)
            tops = [n for n in tree if n["level"] == 0]
            total = sum((n["values"][-1] or 0) for n in tops)
            kpis["bs_deposits_total_mln_mnt"] = round(total, 1)
            kpis["_deposit_source"] = src
            log.info("Deposit tree: %d rows, %d top blocks, %d period "
                     "column(s) from %s; total %.0f mln MNT at %s; %s",
                     len(tree), len(tops), len(dep_cols), src, total,
                     dep_cols[-1] if dep_cols else "?",
                     "all parent/child sums agree" if not bad
                     else f"{bad} sum mismatch(es)")
        else:
            log.warning("Deposit tree not found -- the ПАССИВ block headings "
                        "may have been reworded.")
    except Exception as exc:
        log.warning("balance-sheet parsing failed: %s", exc)
    return kpis, series


# MONGOLBANK BALANCE OF PAYMENTS xlsx (when bop_summary/bop_detailed cached)
# Account labels as they appear in the export, matched EXACTLY after
# stripping the roman-numeral prefix ('I. УРСГАЛ ДАНС' -> 'УРСГАЛ ДАНС').
# Exact matching is essential: 'Санхүүгийн дансны тэнцэл (...)' is a
# different (subtotal) row and must NOT be caught by САНХҮҮГИЙН ДАНС.
BOP_ROW_LABELS = {
    "bop_current_account":   "УРСГАЛ ДАНС",
    "bop_capital_account":   "ХӨРӨНГИЙН ДАНС",
    "bop_financial_account": "САНХҮҮГИЙН ДАНС",
    "bop_errors_omissions":  "АЛДАА БОЛОН ОРХИГДУУЛГА",
    "bop_overall_balance":   "ТӨЛБӨРИЙН ТЭНЦЛИЙН НИЙТ ДҮН",
    "bop_reserve_assets":    "НӨӨЦ ХӨРӨНГӨ",
}
_BOP_PREFIX = re.compile(r"^[IVX]+\.\s*")
_BOP_PERIOD = re.compile(r"(20\d{2})-(\d{2})")


# Mongolbank's export portal puts several dashboards side by side and every
# one of them downloads as a nondescript .xlsx. The balance sheet in
# particular is easy to grab by mistake -- its Mongolian name, 'Банкны
# салбарын ТАЙЛАН ТЭНЦЭЛ', contains the same word as 'ТӨЛБӨРИЙН ТЭНЦЭЛ'.
# When the file is not the BoP, say which one it actually is and where the
# right one lives, rather than 'no account rows matched', which reads like
# the file is missing.
_BOP_IMPOSTERS = [
    (("нийт актив", "банкны нөөц", "индикатор нэр", "төв банкны үнэт цаас"),
     "Банкны салбарын тайлан тэнцэл (the banking sector BALANCE SHEET)",
     "that one is fetched automatically — you do not need to download it"),
    (("зээлийн үлдэгдэл", "хугацаа хэтэрсэн", "чанаргүй зээл",
      "олгосон зээл"),
     "Банкны салбарын зээл (the bank LOAN report)",
     "that one is fetched automatically too"),
    (("хэрэглээний үнийн индекс", "инфляц"),
     "a consumer price index export",
     "inflation comes from the Mongolbank bulletin, also automatic"),
]


def _bop_wrong_file(path, df):
    """Name the file the operator actually saved, and what to do about it."""
    text = " ".join(str(v).casefold() for v in df.iloc[:40, 0].tolist()
                    if isinstance(v, str))
    text += " " + " ".join(str(v).casefold() for v in df.iloc[:3].values.ravel()
                           if isinstance(v, str))
    for words, what, aside in _BOP_IMPOSTERS:
        if sum(w in text for w in words) >= 2:
            log.error(
                "%s is NOT the balance of payments — it is %s. %s.",
                path.name, what, aside.capitalize())
            break
    else:
        log.error("%s does not look like the balance of payments: none of "
                  "the account rows (УРСГАЛ ДАНС, ХӨРӨНГИЙН ДАНС, "
                  "САНХҮҮГИЙН ДАНС, НӨӨЦ ХӨРӨНГӨ) are in its first column.",
                  path.name)
    log.error("  What to download: stat.mongolbank.mn/external -> "
              "'Төлбөрийн тэнцэл' -> export to Excel -> save over "
              "raw_files/bop_manual.xlsx.")
    log.error("  How to tell you have the right one: its first column reads "
              "I. УРСГАЛ ДАНС, II. ХӨРӨНГИЙН ДАНС, III. САНХҮҮГИЙН ДАНС.")
    log.error("  The run continues. Everything else is unaffected; the BoP "
              "rows stay blank and Data_Vintage marks them MISSING, so "
              "nothing wrong reaches a slide.")


def extract_bop_detail():
    """
    Parse the manually downloaded BoP export (raw_files/bop_manual.xlsx,
    'Төлбөрийн тэнцэл (хураангуй)'): header row carries monthly columns
    ('2023-01 Эцсийн**', ...); account rows are matched exactly by label
    after stripping the roman-numeral prefix. Produces:
      * monthly series per account (mln USD),
      * annual sums (full years + current-year YTD -> deck slide 12 chart),
      * KPIs = current-year YTD per account.
    """
    log.info("===== MONGOLBANK BALANCE OF PAYMENTS (manual file) =====")
    kpis, series = {}, {"columns": [], "rows": {}, "annual": {}}
    paths = [RAW_DATA_DIR / "bop_manual.xlsx",              # human-in-the-loop
             RAW_DATA_DIR / "mongolbank_bop_summary.xlsx",  # legacy names
             RAW_DATA_DIR / "mongolbank_bop_detailed.xlsx"]
    path = next((p for p in paths if p.exists()), None)
    if path is None:
        log.warning("BoP file not present -- download the Төлбөрийн тэнцэл "
                    "Excel from stat.mongolbank.mn/external and save it as "
                    "raw_files/bop_manual.xlsx (the ingestor prompts for "
                    "this). Slide 12 detail stays flagged.")
        return kpis, series
    try:
        df = pd.read_excel(path, sheet_name=0, header=None)
        # header row = the one with >=3 'YYYY-MM' cells
        hdr, datecols = None, {}
        for r in range(min(6, len(df))):
            cols = {}
            for c in range(1, df.shape[1]):
                m = _BOP_PERIOD.search(str(df.iat[r, c]))
                if m:
                    cols[c] = (int(m.group(1)), int(m.group(2)))
            if len(cols) >= 3:
                hdr, datecols = r, cols
                break
        if hdr is None:
            raise ValueError("no monthly header row found")
        ordered = sorted(datecols)
        series["columns"] = [f"{y}-{m:02d}" for c in ordered
                             for (y, m) in [datecols[c]]]

        found = {}
        for r in range(hdr + 1, len(df)):
            v = df.iat[r, 0]
            if not isinstance(v, str):
                continue
            label = _BOP_PREFIX.sub("", v.strip()).upper()
            for key, want in BOP_ROW_LABELS.items():
                if label == want and key not in found:
                    found[key] = [_f(df.iat[r, c]) for c in ordered]
        for key, vals in found.items():
            series["rows"][key] = {"label": BOP_ROW_LABELS[key],
                                   "values": vals}
            annual = {}
            for (c, val) in zip(ordered, vals):
                if val is None:
                    continue
                y = datecols[c][0]
                annual[y] = annual.get(y, 0.0) + val
            months_last = [datecols[c][1] for c in ordered
                           if datecols[c][0] == max(annual) and
                           vals[ordered.index(c)] is not None]
            series["annual"][key] = {str(y): round(s, 1)
                                     for y, s in sorted(annual.items())}
            latest_year = max(annual)
            kpis[key] = round(annual[latest_year], 1)
            kpis[f"_{key}_period"] = (
                f"{latest_year} YTD (through month "
                f"{max(months_last) if months_last else '?'}) , mln USD")
        if not found:
            _bop_wrong_file(path, df)
        else:
            log.info("BoP accounts (latest-year YTD): %s",
                     {k: v for k, v in kpis.items()
                      if not k.startswith("_")})

        _bop_hierarchy(df, hdr, ordered, datecols, series, kpis)
    except Exception as exc:
        log.warning("BoP parsing failed for %s: %s", path.name, exc)
    return kpis, series


# Mongolbank encodes the BoP tree with leading non-breaking spaces: a section
# header sits at depth 0 and each level adds a fixed indent. Reading the depth
# instead of the numbering survives renumbering and reworded captions.
_BOP_SECTION = re.compile(r"^([IVX]+)\.\s*(.+)$")


def _bop_indent(s: str) -> int:
    return len(s) - len(s.lstrip("\xa0 \t"))


def _bop_hierarchy(df, hdr, ordered, datecols, series, kpis):
    """
    Section + immediate-children view of the BoP, cumulative for the target
    month against the same months of the previous year.

    Depth is taken from the indentation actually present in the file rather
    than assumed, so an extra tier upstream cannot silently flatten the tree.
    """
    depths = sorted({_bop_indent(str(df.iat[r, 0]))
                     for r in range(hdr + 1, len(df))
                     if isinstance(df.iat[r, 0], str) and str(df.iat[r, 0]).strip()})
    if len(depths) < 2:
        log.warning("BoP detail: no indentation tiers found -- section "
                    "breakdown skipped.")
        return
    top, child = depths[0], depths[1]

    cur_cols = [c for c in ordered
                if datecols[c][0] == TARGET_YEAR and datecols[c][1] <= TARGET_MONTH]
    prv_cols = [c for c in ordered
                if datecols[c][0] == TARGET_YEAR - 1 and datecols[c][1] <= TARGET_MONTH]
    if not cur_cols:
        log.warning("BoP detail: no %d columns at or before month %d -- the "
                    "export may predate the target period.",
                    TARGET_YEAR, TARGET_MONTH)
        return

    # Both years must cover the SAME span, and the span is what the export
    # actually holds -- not what was asked for. A file reaching only May,
    # summed for a June report, would otherwise print five months of data
    # under an "I-6" heading.
    span = min(max(datecols[c][1] for c in cur_cols),
               max((datecols[c][1] for c in prv_cols), default=12))
    cur_cols = [c for c in cur_cols if datecols[c][1] <= span]
    prv_cols = [c for c in prv_cols if datecols[c][1] <= span]
    lag_note = ""
    if span < TARGET_MONTH:
        lag_note = (f"BoP export reaches month {span} only; these are "
                    f"cumulative I-{span} figures, not I-{TARGET_MONTH}.")
        log.warning("BoP: %s", lag_note)

    def ytd(r, cols):
        vals = [_f(df.iat[r, c]) for c in cols]
        vals = [v for v in vals if v is not None]
        return round(sum(vals), 2) if vals else None

    grand = depths[2] if len(depths) > 2 else None
    numbered = re.compile(r'^"?\d+\.')

    detail, section, last_child_plain = [], None, False
    for r in range(hdr + 1, len(df)):
        raw = df.iat[r, 0]
        if not isinstance(raw, str) or not raw.strip():
            continue
        # the export quotes some captions ('"2. Багцын ..."')
        indent, text = _bop_indent(raw), raw.strip().strip('"').strip()
        m = _BOP_SECTION.match(text)
        if indent == top and m:
            section = f"{m.group(1)}. {m.group(2)}"
            last_child_plain = False
            detail.append({"level": 0, "section": section, "label": section,
                           "cur": ytd(r, cur_cols), "prev": ytd(r, prv_cols)})
        elif indent == top:
            continue                      # balance lines between sections
        elif indent == child and section is not None:
            # An unnumbered child is a subtotal ('Бараа ба үйлчилгээний
            # худалдаа'); the numbered rows beneath it carry the components
            # the reader actually wants, so those get pulled up one tier.
            last_child_plain = not numbered.match(text)
            detail.append({"level": 1, "section": section, "label": text[:70],
                           "cur": ytd(r, cur_cols), "prev": ytd(r, prv_cols)})
        elif (indent == grand and section is not None and last_child_plain
              and numbered.match(text)):
            detail.append({"level": 2, "section": section, "label": text[:70],
                           "cur": ytd(r, cur_cols), "prev": ytd(r, prv_cols)})

    if not detail:
        log.warning("BoP detail: hierarchy parsed but no rows produced.")
        return
    series["detail"] = detail
    series["detail_period"] = {
        "cur": f"{TARGET_YEAR} I-{span}",
        "prev": f"{TARGET_YEAR - 1} I-{span}",
        "months": span}
    if lag_note:
        series["detail_lag_note"] = lag_note
    kpis["_bop_period"] = f"{TARGET_YEAR}-{span:02d}"
    nsec = sum(1 for d in detail if d["level"] == 0)
    log.info("BoP detail: %d sections, %d components (%s vs %s)",
             nsec, len(detail) - nsec, series["detail_period"]["cur"],
             series["detail_period"]["prev"])

    # Self-check 1: each section must reconcile to its own components. Most
    # sections add up plainly, but the reserves section follows the standard
    # financing convention where exceptional items are DEDUCTED from the
    # reserve movement (verified month by month against the source:
    # V = ГВУН - ОУВС-гийн зээл - Тусгай санхүүжилт). Both conventions are
    # tried and the one that reconciles is named in the log, so a genuine
    # break is still visible.
    for sec in [d for d in detail if d["level"] == 0]:
        kids = [d for d in detail
                if d["level"] == 1 and d["section"] == sec["section"]]
        if not kids or sec["cur"] is None:
            continue
        tol = max(0.5, abs(sec["cur"]) * 0.01)
        additive = sum(k["cur"] or 0 for k in kids)
        financing = (kids[0]["cur"] or 0) - sum(k["cur"] or 0 for k in kids[1:])
        if abs(additive - sec["cur"]) <= tol:
            continue
        if abs(financing - sec["cur"]) <= tol:
            log.info("BoP %s reconciles on the financing convention "
                     "(first component net of the rest).", sec["section"][:34])
            continue
        log.warning("BoP %s: components sum %.1f (or %.1f net) != section "
                    "%.1f -- a tier may have been added upstream.",
                    sec["section"][:28], additive, financing, sec["cur"])

    # Self-check 2: the BoP identity. Mongolbank reports reserves as their own
    # section rather than inside the financial account, so the balancing
    # identity is CA + KA - FA + E&O = overall balance (= reserve movement),
    # not zero.
    by = {}
    for d in detail:
        if d["level"] != 0:
            continue
        u = d["label"].upper()
        for tag, need in (("ca", "УРСГАЛ"), ("ka", "ХӨРӨНГИЙН"),
                          ("fa", "САНХҮҮГИЙН"), ("eo", "АЛДАА"),
                          ("res", "НӨӨЦ")):
            if need in u and tag not in by:
                by[tag] = d["cur"]
    if all(by.get(t) is not None for t in ("ca", "ka", "fa", "eo", "res")):
        lhs = by["ca"] + by["ka"] - by["fa"] + by["eo"]
        resid = lhs - by["res"]
        scale = max(1.0, abs(by["res"]))
        if abs(resid) > scale * 0.02:
            log.warning("BoP identity off: CA %.1f + KA %.1f - FA %.1f + EO "
                        "%.1f = %.1f, but reserves %.1f (residual %.1f).",
                        by["ca"], by["ka"], by["fa"], by["eo"], lhs,
                        by["res"], resid)
        else:
            log.info("BoP identity OK (CA + KA - FA + E&O = %.2f = reserves "
                     "%.2f, residual %.2f)", lhs, by["res"], resid)
        kpis["_bop_identity_residual"] = round(resid, 2)


# MONGOLBANK BANK LOAN REPORT (monthly flow statement, by sector)
# Sheet per borrower type; within each, a grand-total block, then one block
# per economic sector (letter in column A). Column B carries the quality
# class (1..6) or maturity bucket (а/б/в), column C the caption.
LOAN_SHEETS = [
    ("Total",      "Нийт (Total)"),
    ("Private",    "Хувийн байгууллага (Private companies)"),
    ("Public",     "Улсын байгууллага (State organisations)"),
    ("OFC",        "Бусад санхүүгийн байгууллага (Other financial corp.)"),
    ("Individual", "Иргэд (Individuals)"),
    ("Other",      "Бусад (Other)"),
]
LOAN_QUALITY = [("1", "Хэвийн зээл (Performing)"),
                ("2", "Анхаарал хандуулах зээл (Special mention)"),
                ("3", "Хэвийн бус зээл (Substandard)"),
                ("4", "Эргэлзээтэй зээл (Doubtful)"),
                ("5", "Муу зээл (Loss)"),
                ("6", "Сан (Provisions)")]
LOAN_NPL_CLASSES = ("3", "4", "5")      # чанаргүй зээл = 3 + 4 + 5
LOAN_TOTAL_ROW = "ЗЭЭЛИЙН БҮГД ДҮН"
# Column positions are read from the 'Баганын дугаар' row at run time; these
# are the report's own column numbers, which are stable across months.
LOAN_COLNO = {"opening": 1, "disbursed": 2, "repaid": 4, "writeoff": 6,
              "fx_dr": 7, "fx_cr": 8, "other_dr": 9, "other_cr": 10,
              "closing": 11, "closing_mnt": 12,
              "closing_project": 13, "closing_project_mnt": 14,
              "borrowers": 15,
              "rate_new_mnt": 23, "rate_new_fx": 24,
              "maturity_stock_mnt": 25,
              "rate_stock_mnt": 27, "rate_stock_fx": 28}

# The stat.mongolbank.mn dashboard ('Банкны салбарын зээл') exposes four
# indicators, each also 'үүнээс: төгрөгийн зээл'. They are the same
# statistics this report carries -- its own 'Заавар LOAN' sheet states
# "Хугацаа хэтэрсэн (анхаарал хандуулах) болон чанаргүй (хэвийн бус,
# эргэлзээтэй, муу)" -- so the dashboard's wording maps onto the quality
# classes as below and the workbook reports them under the dashboard's names.
# Verified against a portal export for 2025-06..2026-06: every indicator and
# all 22 sectors agree to the cent for the overlapping month.
#   (key, label, quality classes to sum, column when it is its own column)
LOAN_INDICATORS = [
    ("loan_balance", "Зээлийн үлдэгдэл (Loan balance)",
     None, ("closing", "closing_mnt")),
    ("loan_overdue", "Хугацаа хэтэрсэн зээлийн үлдэгдэл (Past due)",
     ("2",), None),
    ("loan_npl", "Чанаргүй зээлийн үлдэгдэл (Non-performing)",
     LOAN_NPL_CLASSES, None),
    ("loan_project", "Төслөөр хэрэгжсэн зээлийн үлдэгдэл (Project loans)",
     None, ("closing_project", "closing_project_mnt")),
]
_LOAN_PERIOD = re.compile(r"(20\d{2})\s*ОНЫ\s*(\d{1,2})\s*ДУГААР\s*САР",
                          re.IGNORECASE)


# Each flow is found by its own caption first. Column NUMBERS are only a
# fallback: if Mongolbank inserts a column the numbering shifts underneath
# every later flow, which silently misreads the bridge (seen on the 2026-06
# report). Captions move far less often, and when one is reworded the bridge
# check fails loudly rather than quietly reporting a wrong decomposition.
LOAN_COLTEXT = {
    "opening":   ["эхний үлдэгдэл"],
    "disbursed": ["олгосон зээл"],
    "repaid":    ["төлөгдсөн зээл"],
    "writeoff":  ["сангаас хаагдсан"],
    "fx":        ["ханшийн тэгшитгэл"],
    "other":     ["бусад гүйлгээ"],
    "closing":   ["эцсийн үлдэгдэл"],
}


def _loan_colmaps(df):
    """
    Candidate column maps, best guess first: [by caption, by column number].

    Captions sit on a header row and span merged sub-columns ('нийт' /
    'төгрөгөөр', 'дебет' / 'кредит'); the caption marks the first of them,
    which is the total column. The caller picks whichever candidate makes the
    growth bridge reconcile, so an inserted column upstream cannot silently
    shift every later flow.
    """
    nums, numrow = {}, None
    for r in range(min(24, len(df))):
        found = {}
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            try:
                n = int(float(str(v).strip()))
            except (TypeError, ValueError):
                continue
            if 1 <= n <= 40 and n not in found:
                found[n] = c
        if len(found) >= 20:
            nums, numrow = found, r
            break
    positional = {k: nums[n] for k, n in LOAN_COLNO.items() if n in nums}
    if not positional:
        return []

    # caption scan: header band only, data columns only, and the caption must
    # START with the phrase -- the sheet title also contains 'олгосон зээл'
    caps = {}
    for r in range(0, numrow if numrow is not None else 18):
        for c in range(3, df.shape[1]):
            v = df.iat[r, c]
            if not isinstance(v, str):
                continue
            t = re.sub(r"\s+", " ", v).strip().casefold()
            if len(t) > 44:
                continue
            for key, frags in LOAN_COLTEXT.items():
                if key not in caps and any(t.startswith(f) for f in frags):
                    caps[key] = c
    by_caption = dict(positional)
    for key in ("opening", "disbursed", "repaid", "writeoff", "closing"):
        if key in caps:
            by_caption[key] = caps[key]
    if "fx" in caps:
        by_caption["fx_dr"], by_caption["fx_cr"] = caps["fx"], caps["fx"] + 1
    if "other" in caps:
        by_caption["other_dr"] = caps["other"]
        by_caption["other_cr"] = caps["other"] + 1

    out = [by_caption]
    if positional != by_caption:
        out.append(positional)
    return out


_LOAN_COLMAP_WARNED = []


def _loan_bridge_residual(df, cm):
    """Signed bridge error for the grand-total row, or None if unusable."""
    t = _parse_loan_sheet(df, cm).get("total") or {}
    op, cl = t.get("opening"), t.get("closing")
    if None in (op, cl):
        return None
    calc = (op + (t.get("disbursed") or 0) - (t.get("repaid") or 0)
            - (t.get("writeoff") or 0)
            + ((t.get("fx_dr") or 0) - (t.get("fx_cr") or 0))
            + ((t.get("other_dr") or 0) - (t.get("other_cr") or 0)))
    return calc - cl


def _loan_colmap(df):
    """The candidate map whose bridge reconciles; the first one otherwise."""
    cands = _loan_colmaps(df)
    if not cands:
        return {}
    for i, cm in enumerate(cands):
        resid = _loan_bridge_residual(df, cm)
        if resid is not None and abs(resid) <= 1.0:
            if i:
                log.info("loan columns: caption match did not reconcile; "
                         "using the report's column numbers instead.")
            return cm
    # warn once per run, not once per parsed workbook (13 history files)
    if not _LOAN_COLMAP_WARNED:
        _LOAN_COLMAP_WARNED.append(True)
        log.info("loan columns: neither mapping closes the bridge, so the "
                 "report's own flows do not sum to its closing balance. See "
                 "the bridge warning below for where the gap sits; the "
                 "closing balance, sectors and quality classes are read "
                 "directly and are unaffected.")
    return cands[0]


def _loan_period(df):
    for r in range(min(12, len(df))):
        for c in range(min(6, df.shape[1])):
            m = _LOAN_PERIOD.search(str(df.iat[r, c]))
            if m:
                return int(m.group(1)), int(m.group(2))
    return None


def _parse_loan_sheet(df, cm):
    """One borrower sheet -> {total, quality, sectors, sublines}."""
    def val(r, key):
        c = cm.get(key)
        return _f(df.iat[r, c]) if c is not None else None

    def block(r):
        return {k: val(r, k) for k in cm}

    out = {"total": None, "quality": {}, "sectors": [], "sublines": []}
    cur = None
    for r in range(len(df)):
        letter = str(df.iat[r, 0]).strip() if pd.notna(df.iat[r, 0]) else ""
        code = str(df.iat[r, 1]).strip() if pd.notna(df.iat[r, 1]) else ""
        raw = str(df.iat[r, 2]) if pd.notna(df.iat[r, 2]) else ""
        cap = raw.strip()
        if not cap:
            continue
        if cap.upper().startswith(LOAN_TOTAL_ROW):
            out["total"] = block(r)
            cur = "TOTAL"
            continue
        if len(letter) == 1 and letter.isalpha():
            entry = {"code": letter, "label": re.sub(r"\s+", " ", cap)[:70],
                     **block(r)}
            # Sector headings and their 'үүнээс' detail lines share the same
            # letter in column A, so the letter alone cannot separate them.
            # A real sector caption is unindented and written in capitals
            # ('УУЛ УУРХАЙ, ОЛБОРЛОЛТ'); detail lines are indented and in
            # sentence case ('б. Газар тариалан'), and only the first of them
            # carries the 'Үүнээс:' prefix -- so indentation and casing are
            # what separate them, never the prefix alone. Counting one detail
            # line as a sector silently inflates the sector total.
            alpha = [ch for ch in cap if ch.isalpha()]
            caps_ratio = (sum(ch.isupper() for ch in alpha) / len(alpha)
                          if alpha else 0)
            is_detail = (raw[:1].isspace() or "Үүнээс" in cap
                         or caps_ratio < 0.8)
            if is_detail:
                out["sublines"].append(entry)
                cur = None          # detail-line quality rows are not summed
            else:
                out["sectors"].append(entry)
                cur = "SECTOR"
            continue
        if cur == "TOTAL" and code in dict(LOAN_QUALITY):
            out["quality"][code] = block(r)
    return out


def extract_bank_loans():
    """
    Monthly bank loan report: how the loan book grew and what drove it.

    The report is a FLOW statement -- opening balance, loans disbursed,
    repaid, written off, FX revaluation, closing balance -- so the growth
    decomposes exactly rather than being inferred from two stock readings.
    Backfilled months (mb_loans_YYYY_MM.xlsx) turn it into a series; the
    single latest report is used on its own when no history is cached.
    """
    log.info("===== MONGOLBANK BANK LOAN REPORT =====")
    kpis, out = {}, {"bridge": {}, "sectors": [], "sublines": [],
                     "quality": [], "borrowers": [], "rates": {},
                     "monthly": [], "indicators": []}

    files = {}
    for p in sorted(RAW_DATA_DIR.glob("mb_loans_*.xlsx")):
        m = re.match(r"mb_loans_(\d{4})_(\d{2})\.xlsx$", p.name)
        if m:
            files[(int(m.group(1)), int(m.group(2)))] = p
    legacy = RAW_DATA_DIR / "mongolbank_bank_loan_report.xlsx"
    if not files and legacy.exists():
        files[(TARGET_YEAR, TARGET_MONTH)] = legacy
    if not files:
        log.warning("Loan report not cached -- run raw_ingestor.py. Loan "
                    "detail cells stay flagged.")
        return kpis, out
    # never let a rebuild for month M include months published after it
    files = {k: v for k, v in files.items() if k <= (TARGET_YEAR, TARGET_MONTH)}
    if not files:
        log.warning("Loan reports cached, but none at or before %d-%02d.",
                    TARGET_YEAR, TARGET_MONTH)
        return kpis, out

    parsed = {}
    for period in sorted(files):
        path = files[period]
        try:
            df = pd.read_excel(path, sheet_name="Total", header=None)
        except Exception as exc:
            log.warning("loan report %s unreadable (%s) -- skipped.",
                        path.name, exc)
            continue
        cm = _loan_colmap(df)
        if not cm or "closing" not in cm:
            log.warning("loan report %s: column header row not found -- "
                        "skipped.", path.name)
            continue
        stated = _loan_period(df)
        # For mb_loans_YYYY_MM.xlsx the filename period came from the portal's
        # own month listing, so it is authoritative and is used as the key:
        # keying on the in-file title instead would silently collapse two
        # months into one if the portal ever served a mislabelled file. The
        # legacy single-file download has no period in its name, so there the
        # title is all there is.
        if path is legacy:
            key = stated or period
        else:
            key = period
            if stated and stated != period:
                log.warning("loan report %s is labelled %d-%02d inside the "
                            "file; keeping the listed period %d-%02d.",
                            path.name, stated[0], stated[1], *period)
        if key in parsed:
            log.warning("loan report %s duplicates period %d-%02d -- keeping "
                        "the first.", path.name, *key)
            continue
        parsed[key] = (path, df, cm)

    if not parsed:
        log.warning("No loan report could be parsed.")
        return kpis, out

    for period in sorted(parsed):
        ppath, df, cm = parsed[period]
        t = _parse_loan_sheet(df, cm)["total"]
        if not t:
            continue
        row = {"period": f"{period[0]}-{period[1]:02d}",
               "opening": t.get("opening"), "closing": t.get("closing"),
               "disbursed": t.get("disbursed"), "repaid": t.get("repaid")}
        if None not in (row["opening"], row["closing"]) and row["opening"]:
            row["growth_pct"] = round((row["closing"] / row["opening"] - 1)
                                      * 100, 2)
        # Borrower-type balances for every cached month, not just the newest.
        # The deck's bank slide plots how fast INDIVIDUALS are borrowing
        # against businesses, and reading that off a single month is not
        # possible -- it needs the series, which is why that chart used to be
        # carried over from the previous deck and quietly went stale.
        row["by_borrower"] = {}
        for sheet, _lbl in LOAN_SHEETS:
            try:
                bdf = pd.read_excel(ppath, sheet_name=sheet, header=None)
                bcm = _loan_colmap(bdf)
                bt = _parse_loan_sheet(bdf, bcm)["total"] if bcm else None
            except Exception:
                bt = None
            if bt and bt.get("closing") is not None:
                row["by_borrower"][sheet] = round(bt["closing"], 1)
        out["monthly"].append(row)

    # Year-on-year growth per borrower type, once the whole history is read.
    # Thirteen cached months give exactly one YoY point per type; more months
    # give more. Anything shorter is left absent rather than approximated.
    by_period = {r["period"]: r.get("by_borrower", {}) for r in out["monthly"]}
    for r in out["monthly"]:
        y, m = int(r["period"][:4]), int(r["period"][5:])
        base = by_period.get(f"{y - 1}-{m:02d}")
        if not base:
            continue
        r["yoy_by_borrower"] = {
            s: round((v / base[s] - 1) * 100, 2)
            for s, v in r.get("by_borrower", {}).items()
            if base.get(s)}
    yoy_rows = [r for r in out["monthly"] if r.get("yoy_by_borrower")]
    if yoy_rows:
        last = yoy_rows[-1]
        log.info("loan YoY by borrower (%s): %s", last["period"],
                 " | ".join(f"{s} {v:+.1f}%"
                            for s, v in last["yoy_by_borrower"].items()))
    else:
        log.info("loan history: %d month(s) cached -- not enough for a "
                 "year-on-year comparison by borrower type (needs 13).",
                 len(out["monthly"]))

    latest = max(parsed)
    path, df, cm = parsed[latest]
    log.info("Loan report %d-%02d (%s); %d month(s) of history cached.",
             latest[0], latest[1], path.name, len(parsed))
    kpis["_loan_period"] = f"{latest[0]}-{latest[1]:02d}"
    if latest < (TARGET_YEAR, TARGET_MONTH):
        kpis["_loan_period_note"] = (
            f"WARNING: newest loan report is {kpis['_loan_period']}, target "
            f"is {TARGET_YEAR}-{TARGET_MONTH:02d}")
        log.warning(kpis["_loan_period_note"])

    tot = _parse_loan_sheet(df, cm)
    t = tot["total"] or {}
    op, cl = t.get("opening"), t.get("closing")

    # 1. Growth bridge -- must reconcile to the reported closing balance.
    fx = (t.get("fx_dr") or 0) - (t.get("fx_cr") or 0)
    oth = (t.get("other_dr") or 0) - (t.get("other_cr") or 0)
    out["bridge"] = {
        "opening": op, "disbursed": t.get("disbursed"),
        "repaid": t.get("repaid"), "writeoff": t.get("writeoff"),
        "fx": round(fx, 2), "other": round(oth, 2), "closing": cl}
    if None not in (op, cl):
        calc = (op + (t.get("disbursed") or 0) - (t.get("repaid") or 0)
                - (t.get("writeoff") or 0) + fx + oth)
        resid = calc - cl
        out["bridge"]["residual"] = round(resid, 2)
        out["bridge"]["change"] = round(cl - op, 1)
        out["bridge"]["change_pct"] = (round((cl / op - 1) * 100, 2)
                                       if op else None)
        kpis["loan_total_mln_mnt"] = round(cl, 1)
        kpis["loan_growth_mom_pct"] = out["bridge"]["change_pct"]
        if abs(resid) > max(1.0, abs(cl) * 1e-6):
            # Attribute the gap. If it concentrates in a few sectors while
            # the rest reconcile exactly, the flows in the SOURCE do not add
            # up to its own closing balance -- a publication inconsistency,
            # not a parsing fault. Naming the sectors lets that be checked or
            # queried with Mongolbank instead of guessed at.
            offenders = []
            for s in tot["sectors"]:
                so, sc = s.get("opening"), s.get("closing")
                if None in (so, sc):
                    continue
                sr = (so + (s.get("disbursed") or 0) - (s.get("repaid") or 0)
                      - (s.get("writeoff") or 0)
                      + ((s.get("fx_dr") or 0) - (s.get("fx_cr") or 0))
                      + ((s.get("other_dr") or 0) - (s.get("other_cr") or 0))
                      - sc)
                if abs(sr) > 1.0:
                    offenders.append((abs(sr), s["code"], s["label"][:34], sr))
            offenders.sort(reverse=True)
            log.warning("loan bridge does NOT reconcile: flows give %.1f but "
                        "the report states %.1f (gap %.1f = %.3f%% of the "
                        "book).", calc, cl, resid, abs(resid) / cl * 100)
            if offenders:
                log.warning("  the gap sits in %d of %d sectors, the rest "
                            "reconcile exactly -- this is an inconsistency in "
                            "the published report, not in the extraction:",
                            len(offenders), len(tot["sectors"]))
                for _, code, label, sr in offenders[:5]:
                    log.warning("    %s %-34s %+.1f", code, label, sr)
            out["bridge"]["unreconciled_sectors"] = [
                {"code": c, "label": l, "gap": round(sr, 1)}
                for _, c, l, sr in offenders]
        else:
            log.info("loan bridge OK: %.0f + %.0f - %.0f (+fx %.0f) = %.0f "
                     "(residual %.2f); month change %+.2f%%",
                     op, t.get("disbursed") or 0, t.get("repaid") or 0, fx,
                     cl, resid, out["bridge"]["change_pct"] or 0)

    # year-on-year from the backfilled history
    yoy_key = (latest[0] - 1, latest[1])
    if yoy_key in parsed:
        _, ydf, ycm = parsed[yoy_key]
        yt = _parse_loan_sheet(ydf, ycm)["total"] or {}
        if yt.get("closing") and cl:
            kpis["loan_growth_yoy_pct"] = round((cl / yt["closing"] - 1) * 100,
                                                2)
            out["yoy_base"] = {"period": f"{yoy_key[0]}-{yoy_key[1]:02d}",
                               "closing": yt["closing"]}
            log.info("loan YoY: %.1f -> %.1f = %+.2f%% (vs %d-%02d)",
                     yt["closing"], cl, kpis["loan_growth_yoy_pct"], *yoy_key)
    else:
        log.info("loan YoY not available yet (needs %d-%02d; backfill runs "
                 "on the next live ingest).", *yoy_key)

    # 2. Sectors, with each one's contribution to total growth in pp
    ybase = (out.get("yoy_base") or {}).get("period")
    ysec = {}
    if yoy_key in parsed:
        _, ydf, ycm = parsed[yoy_key]
        for s in _parse_loan_sheet(ydf, ycm)["sectors"]:
            ysec[s["code"]] = s.get("closing")
    for s in tot["sectors"]:
        so, sc = s.get("opening"), s.get("closing")
        d = {"code": s["code"], "label": s["label"],
             "opening": so, "closing": sc,
             "closing_mnt": s.get("closing_mnt"),
             "change": round(sc - so, 1) if None not in (so, sc) else None,
             "yoy_base": ysec.get(s["code"])}
        if None not in (so, sc) and op:
            d["pp"] = round((sc - so) / op * 100, 3)
        if sc and ysec.get(s["code"]):
            d["yoy_pct"] = round((sc / ysec[s["code"]] - 1) * 100, 1)
        out["sectors"].append(d)
    out["sectors"].sort(key=lambda d: -(d.get("change") or 0))
    for s in tot["sublines"]:
        so, sc = s.get("opening"), s.get("closing")
        out["sublines"].append({
            "code": s["code"], "label": s["label"], "opening": so,
            "closing": sc,
            "change": round(sc - so, 1) if None not in (so, sc) else None})

    if out["sectors"] and cl:
        ssum = sum(s["closing"] or 0 for s in out["sectors"])
        if abs(ssum - cl) > max(1.0, abs(cl) * 1e-4):
            log.warning("loan sectors sum %.1f != grand total %.1f (gap "
                        "%.1f) -- a sector row may be unmatched.",
                        ssum, cl, ssum - cl)
        else:
            log.info("loan sectors reconcile to the grand total (%d sectors, "
                     "%.0f)", len(out["sectors"]), ssum)
        pp = sum(s.get("pp") or 0 for s in out["sectors"])
        chg = out["bridge"].get("change_pct")
        if chg is not None and abs(pp - chg) > 0.05:
            log.warning("loan sector contributions sum %.2f pp but the book "
                        "grew %.2f%% -- the sector split is inconsistent with "
                        "the bridge.", pp, chg)
        else:
            log.info("loan growth contributions sum %.2f pp = month change "
                     "%.2f%%", pp, chg or 0)

    # 3. Loan quality and the NPL ratio
    npl = 0.0
    for code, label in LOAN_QUALITY:
        blk = tot["quality"].get(code) or {}
        c = blk.get("closing")
        out["quality"].append({"code": code, "label": label, "closing": c,
                               "opening": blk.get("opening"),
                               "closing_mnt": blk.get("closing_mnt")})
        if code in LOAN_NPL_CLASSES and c:
            npl += c

    # 3b. The dashboard's own three indicators, with the MNT split, so the
    # workbook can be cross-checked line for line against the portal page.
    for key, label, classes, cols in LOAN_INDICATORS:
        if cols is not None:
            tv, mv = t.get(cols[0]), t.get(cols[1])
        else:
            blocks = [tot["quality"].get(c) or {} for c in classes]
            tv = sum(b.get("closing") or 0 for b in blocks) or None
            mv = sum(b.get("closing_mnt") or 0 for b in blocks) or None
        out["indicators"].append({
            "key": key, "label": label, "total": tv, "mnt": mv,
            "share_pct": round(tv / cl * 100, 2) if (tv and cl) else None})
        if tv is not None:
            kpis[f"{key}_mln_mnt"] = round(tv, 1)
    ind = {d["key"]: d for d in out["indicators"]}
    if ind.get("loan_overdue", {}).get("total") and cl:
        kpis["loan_overdue_ratio_pct"] = ind["loan_overdue"]["share_pct"]
    log.info("loan indicators (portal wording): balance %.0f (MNT %.0f) | "
             "past due %.0f | non-performing %.0f | project %.0f",
             ind["loan_balance"]["total"] or 0,
             ind["loan_balance"]["mnt"] or 0,
             ind["loan_overdue"]["total"] or 0,
             ind["loan_npl"]["total"] or 0,
             (ind.get("loan_project") or {}).get("total") or 0)
    # the MNT part can never exceed the total it is drawn from
    for d in out["indicators"]:
        if d["total"] and d["mnt"] and d["mnt"] > d["total"] * 1.0001:
            log.warning("loan indicator %s: MNT part %.1f exceeds the total "
                        "%.1f -- the 'of which' column may have moved.",
                        d["key"], d["mnt"], d["total"])
    if npl and cl:
        kpis["loan_npl_mln_mnt"] = round(npl, 1)
        kpis["loan_npl_ratio_pct"] = round(npl / cl * 100, 2)
        out["npl"] = {"amount": round(npl, 1),
                      "ratio_pct": kpis["loan_npl_ratio_pct"]}
        log.info("loan quality: NPL (classes 3+4+5) %.0f = %.2f%% of the "
                 "book", npl, kpis["loan_npl_ratio_pct"])

    # 4. Borrower types (one sheet each)
    for sheet, label in LOAN_SHEETS:
        try:
            bdf = pd.read_excel(path, sheet_name=sheet, header=None)
        except Exception:
            continue
        bcm = _loan_colmap(bdf)
        if not bcm:
            continue
        bt = _parse_loan_sheet(bdf, bcm)["total"] or {}
        bo, bc = bt.get("opening"), bt.get("closing")
        out["borrowers"].append({
            "sheet": sheet, "label": label, "opening": bo, "closing": bc,
            "change": round(bc - bo, 1) if None not in (bo, bc) else None,
            "borrowers": bt.get("borrowers")})

    # 5. Weighted average rates on the total book
    out["rates"] = {k: t.get(k) for k in
                    ("rate_new_mnt", "rate_new_fx", "rate_stock_mnt",
                     "rate_stock_fx", "maturity_stock_mnt")
                    if t.get(k) is not None}
    if out["rates"].get("rate_stock_mnt"):
        kpis["loan_rate_stock_mnt_pct"] = round(out["rates"]
                                                ["rate_stock_mnt"], 2)
    return kpis, out


# REAL SECTOR (deck slides 16-21) — trade, services, construction, industry
#
# The method is the one used in the analyst's own workbooks, reproduced here
# exactly (verified against 'trade may 2025.xlsx' to the last decimal):
#
#   1. cumulative      cum_t = cum_{t-1} + level_t, restarting each January
#                      (tables already published 'өссөн дүнгээр' skip this)
#   2. headline growth g = cum_t / cum_{t-L} - 1        L = 12 months / 4 qtrs
#   3. contribution    c_i = (cum_i,t - cum_i,t-L)
#                            / (cum_tot,t - cum_tot,t-L) * g
#
# The contributions are additive by construction, so they must sum to g --
# which is the self-check at the end. On the deck these become a stacked bar
# of contributions with the headline growth as a line over the top.
REAL_SECTOR_TABLES = {
    "trade": {
        "file": "nso_trade_sales.json",
        "label": "Худалдааны салбарын нийт борлуулалт",
        "dim": "Бүс", "period": "Сар", "freq": "M",
        "cumulative_source": False,      # published as a monthly level
        "total": ("улсын дүн", "бүгд", "нийт"),
    },
    "hotel": {
        "file": "nso_hotel_income.json",
        "label": "Зочид буудлын салбарын орлого",
        "dim": "Бүс", "period": "Сар", "freq": "M",
        "cumulative_source": True,       # already өссөн дүнгээр
        "total": ("улсын дүн", "бүгд", "нийт"),
    },
    "food": {
        "file": "nso_food_income.json",
        "label": "Нийтийн хоолны салбарын орлого",
        "dim": "Бүс", "period": "Сар", "freq": "M",
        "cumulative_source": True,
        "total": ("улсын дүн", "бүгд", "нийт"),
    },
    "construction": {
        "file": "nso_construction.json",
        "label": "Барилга угсралт, их засварын ажил",
        "dim": "Үзүүлэлт", "period": "Улирал", "freq": "Q",
        "cumulative_source": True,       # quarterly, өссөн дүн
        "total": ("барилга угсралт, их засварын ажил", "бүгд", "нийт"),
        # This table is flat -- every caption sits at indent 0 even though
        # 'Орон сууцны бус' and 'Инженерийн' each have their own sub-types
        # below them. So the four headline types are named explicitly; they
        # sum to the total (verified against the published data), and they
        # are exactly the four series on the analyst's slide-21 chart.
        "components": ("орон сууцны барилга", "орон сууцны бус барилга",
                       "инженерийн барилга, байгууламж", "их засварын ажил"),
    },
    "industry": {
        "file": "nso_industry_sales.json",
        "label": "Аж үйлдвэрийн салбарын бүтээгдэхүүний борлуулалт",
        "dim": "Дэд салбар", "period": "Сар", "freq": "M",
        "cumulative_source": False,
        "total": ("нийт дүн", "бүгд", "нийт", "аж үйлдвэр"),
        # Hierarchical by INDENTATION: the four major groups sit at indent 8
        # and sum to 'Нийт дүн'; sub-sectors are indented deeper. Taking the
        # shallowest tier keeps the contribution split additive.
        "level_from_indent": True,
    },
    "transport": {
        "file": "nso_transport.json",
        "label": "Тээврийн салбарын нийт орлого",
        "dim": "Үзүүлэлт", "period": "Улирал", "freq": "Q",
        "cumulative_source": False,
        "unit": "тэрбум төг",
        # This table is not one hierarchy but FIVE independent blocks --
        # freight tonnes, freight turnover, passengers, passenger turnover
        # and revenue -- each followed by the same four modes. The indent-0
        # rows are different measures and do not sum to anything, so the
        # tier rule cannot apply. Name the block we want; its components are
        # the indented rows that follow it, up to the next block heading.
        "total": ("тээврийн нийт орлого, тэрбум төгрөг",),
        "block_children": True,
        # The four modes do not add up to the sector total: warehousing and
        # postal services belong to it too, and the API returns them empty.
        # The difference is shown as its own component so the decomposition
        # still closes and nothing is quietly dropped.
        "residual_label": "Бусад (агуулах, шуудан, туслах үйл ажиллагаа)",
    },
}
# the five statistical regions the analyst's workbooks use, plus the capital
REAL_SECTOR_REGIONS = {"баруун бүс", "хангайн бүс", "төвийн бүс",
                       "зүүн бүс", "улаанбаатар"}
_PERIOD_M = re.compile(r"^(\d{4})[-\s.](\d{1,2})$")
_PERIOD_Q = re.compile(r"^(\d{4})\s*[-\s]?\s*([IVX]{1,4})$", re.IGNORECASE)
_ROMAN = {"i": 1, "ii": 2, "iii": 3, "iv": 4}


def _rs_period(label, freq):
    """'2026-05' -> (2026,5); '2026-I' -> (2026,1). None if unparseable."""
    s = str(label).strip()
    m = _PERIOD_M.match(s)
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    m = _PERIOD_Q.match(s)
    if m and m.group(2).lower() in _ROMAN:
        return int(m.group(1)), _ROMAN[m.group(2).lower()]
    if freq == "Q":
        m = re.match(r"^(\d{4})\D+([1-4])$", s)
        if m:
            return int(m.group(1)), int(m.group(2))
    return None


def _rs_cumulate(pairs):
    """
    [(period, value)] -> running year-to-date, restarting each year.

    A missing period carries the running total forward rather than breaking
    the series. NSO publishes some components later than others (water
    transport lags the rest), and dropping the period outright desynchronises
    a component from its total, so the contributions stop adding up. Carrying
    forward states the honest thing: nothing was added in that period.
    """
    out, run, yr, seen = [], 0.0, None, False
    for (y, p), v in pairs:
        if y != yr:
            run, yr = 0.0, y
        if v is not None:
            run += float(v)
            seen = True
        out.append(((y, p), round(run, 4) if seen else None))
    return out


def _extract_one_real_sector(key, cfg):
    """{'label','unit','periods','total','components','growth',...} or None."""
    path = RAW_DATA_DIR / cfg["file"]
    if not path.exists():
        return None
    try:
        js = JsonStat2(json.loads(path.read_text(encoding="utf-8")))
    except Exception as exc:
        log.warning("real sector[%s]: %s unusable (%s)", key, cfg["file"], exc)
        return None

    dim_name = next((d for d in js.ids if d == cfg["dim"]), None)
    per_name = next((d for d in js.ids if d == cfg["period"]), None)
    if dim_name is None or per_name is None:
        log.warning("real sector[%s]: expected dimensions %r/%r, got %s",
                    key, cfg["dim"], cfg["period"], js.ids)
        return None

    # periods in chronological order, clamped to the reported month
    periods = []
    for pos, lbl in js.labels(per_name).items():
        pr = _rs_period(lbl, cfg["freq"])
        if not pr:
            continue
        if cfg["freq"] == "M" and pr > (TARGET_YEAR, TARGET_MONTH):
            continue
        if cfg["freq"] == "Q" and (pr[0] > TARGET_YEAR or
                                   (pr[0] == TARGET_YEAR and
                                    pr[1] > (TARGET_MONTH + 2) // 3)):
            continue
        periods.append((pr, pos))
    periods.sort()
    if len(periods) < 2:
        log.warning("real sector[%s]: fewer than two usable periods.", key)
        return None

    # a third dimension (region on the construction table) is collapsed to
    # its national-total member so the component split stays one-dimensional
    fixed = {}
    for other in js.ids:
        if other in (dim_name, per_name):
            continue
        pick = next((p for p, l in js.labels(other).items()
                     if str(l).strip().casefold() in
                     ("улсын дүн", "бүгд", "нийт", "нийт дүн")), 0)
        fixed[other] = pick

    def series_for(dpos):
        vals = []
        for pr, ppos in periods:
            try:
                v = js.cell(**{dim_name: dpos, per_name: ppos, **fixed})
            except Exception:
                v = None
            vals.append((pr, None if v is None else float(v)))
        return vals

    members = js.labels(dim_name)
    tot_pos = next((p for p, l in members.items()
                    if str(l).strip().casefold() in cfg["total"]), None)
    if tot_pos is None:
        tot_pos = min(members)
        log.info("real sector[%s]: no explicit total row; using %r.",
                 key, str(members[tot_pos]).strip()[:40])

    def prepare(dpos):
        s = series_for(dpos)
        return s if cfg["cumulative_source"] else _rs_cumulate(s)

    cum_tot = prepare(tot_pos)
    lag = 12 if cfg["freq"] == "M" else 4
    if len(cum_tot) <= lag:
        log.warning("real sector[%s]: only %d periods, need more than %d for "
                    "a year-on-year comparison.", key, len(cum_tot), lag)
        return None

    i = len(cum_tot) - 1
    now, base = cum_tot[i][1], cum_tot[i - lag][1]
    if not now or not base:
        log.warning("real sector[%s]: latest or year-ago total missing.", key)
        return None
    growth = now / base - 1
    denom = now - base

    # Which members are the components? These tables nest, so taking every
    # member would sum parents together with their children.
    # These tables mix aggregation levels, and each one mixes them
    # DIFFERENTLY: in the regional tables 'Улсын дүн' and the five regions
    # all sit at indent 0 with the aimags at indent 4, while the industry
    # table puts its four groups BELOW the total. No fixed rule covers both,
    # so every candidate tier is tested against the arithmetic and the
    # coarsest one that reconciles wins. A tier that does not reproduce the
    # total's own change is not a valid decomposition of it.
    def change_of(dpos):
        cs = prepare(dpos)
        a, b = cs[i][1], cs[i - lag][1]
        return None if (a is None or b is None) else a - b

    wanted, how = None, ""
    if cfg.get("block_children"):
        # the indented rows directly under the chosen heading, stopping at
        # the next row at the heading's own depth
        depth = {p: len(str(l)) - len(str(l).lstrip(" "))
                 for p, l in members.items() if str(l).strip()}
        # NB: not `base` -- that name already holds the year-ago total, and
        # shadowing it here silently published an indent depth as the prior
        # year's level.
        head_depth = depth.get(tot_pos, 0)
        wanted = set()
        for p in sorted(members):
            if p <= tot_pos:
                continue
            d = depth.get(p, head_depth)
            if d <= head_depth:
                break
            wanted.add(p)
        how = f"rows under {str(members[tot_pos]).strip()[:34]!r}"
    elif cfg.get("components"):
        want = set(cfg["components"])
        wanted = {p for p, l in members.items()
                  if str(l).strip().casefold() in want}
        how = "named in config"
    else:
        depth = {}
        for p, l in members.items():
            s = str(l)
            if s.strip():
                depth[p] = len(s) - len(s.lstrip(" "))
        for d in sorted(set(depth.values())):
            tier = {p for p, dd in depth.items() if dd == d and p != tot_pos}
            if not tier:
                continue
            tot_change = sum(x for x in (change_of(p) for p in tier)
                             if x is not None)
            if denom and abs(tot_change / denom - 1) <= 0.01:
                wanted, how = tier, f"indent {d}, reconciles"
                break
    if wanted:
        log.info("real sector[%s]: %d of %d members selected (%s).",
                 key, len(wanted), len(members), how)
    else:
        # Not an error on its own: when the API already returned a single
        # clean level (six members: the total and the five regions) there is
        # no deeper tier to choose and every non-total member is a component.
        # Whether that actually reconciles is asserted at the end.
        log.info("real sector[%s]: no tier selected from %d members; using "
                 "every non-total member.", key, len(members))

    comps = []
    sel_pos = {}                 # label -> dimension position, for the series
    for dpos, lbl in sorted(members.items()):
        if dpos == tot_pos:
            continue
        if wanted is not None and dpos not in wanted:
            continue
        if not str(lbl).strip():
            continue
        cs = prepare(dpos)
        a, b = cs[i][1], cs[i - lag][1]
        if a is None or b is None:
            continue
        contrib = (a - b) / denom * growth if denom else None
        sel_pos[re.sub(r"\s+", " ", str(lbl)).strip()] = dpos
        comps.append({
            "label": re.sub(r"\s+", " ", str(lbl)).strip(),
            "cur": round(a, 2), "prev": round(b, 2),
            "change": round(a - b, 2),
            "contribution_pct": (round(contrib * 100, 6)
                                 if contrib is not None else None),
            "yoy_pct": round((a / b - 1) * 100, 2) if b else None,
        })
    # If the chosen components are a genuine but PARTIAL decomposition --
    # the transport modes exclude warehousing and postal, which the source
    # publishes as empty -- close the gap with an explicit residual rather
    # than leaving the column not adding up.
    if cfg.get("residual_label") and comps and denom:
        gap = denom - sum(c["change"] for c in comps
                          if c.get("change") is not None)
        if abs(gap) > abs(denom) * 1e-6:
            # Carry LEVELS too, not just the change. Without them the sheet's
            # total row sums the components' levels to less than the table's
            # own total, and its YoY cell then reports a growth rate that
            # disagrees with the headline -- correct arithmetic on the wrong
            # base, which is the kind of thing that reaches a slide.
            c_lv = sum(c["cur"] for c in comps if c.get("cur") is not None)
            p_lv = sum(c["prev"] for c in comps if c.get("prev") is not None)
            comps.append({
                "label": cfg["residual_label"],
                "cur": round(now - c_lv, 2), "prev": round(base - p_lv, 2),
                "change": round(gap, 2),
                "contribution_pct": round(gap / denom * growth * 100, 6),
                "yoy_pct": (round(((now - c_lv) / (base - p_lv) - 1) * 100, 2)
                            if abs(base - p_lv) > 1e-9 else None)})
            log.info("real sector[%s]: added a residual of %s %s (%.1f%% of "
                     "the change) for components the source leaves empty.",
                     key, format(gap, ",.1f"), cfg.get("unit", ""),
                     gap / denom * 100)

    # HIERARCHY GUARD. These tables mix levels: 'Бүс' carries the five
    # regions AND the 21 aimags inside them, and the industry table nests
    # sub-sectors under group headings. Summing across levels double-counts,
    # which shows up as contributions coming to a clean multiple of the
    # headline. If the components sum to about 2x (or 3x) the total change,
    # keep only the coarsest level that reconciles.
    def _sum_change(rows):
        return sum(x["change"] for x in rows if x.get("change") is not None)

    if denom and comps:
        ratio = _sum_change(comps) / denom
        if abs(ratio - 1) > 0.02:
            keep = [c for c in comps
                    if abs(_sum_change([c])) <= abs(denom) * 1.0001]
            # try the members the analyst's own workbooks use: the five
            # regions, identified by name, before falling back to a warning
            named = [c for c in comps
                     if c["label"].casefold() in REAL_SECTOR_REGIONS]
            if named and abs(_sum_change(named) / denom - 1) <= 0.02:
                log.info("real sector[%s]: source mixes aggregation levels "
                         "(members summed to %.2fx the total change); using "
                         "the %d region rows only.", key, ratio, len(named))
                comps = named
            else:
                log.warning("real sector[%s]: components sum to %.2fx the "
                            "total change -- the table mixes aggregation "
                            "levels and no clean subset was found. The "
                            "contribution column is NOT reliable; the "
                            "headline and levels are.", key, ratio)

    comps.sort(key=lambda d: -(d.get("contribution_pct") or 0))

    per_tag = ("%d-%02d" % periods[i][0] if cfg["freq"] == "M"
               else "%d-Q%d" % periods[i][0])
    out = {
        "key": key, "label": cfg["label"], "freq": cfg["freq"],
        "period": per_tag,
        "unit": cfg.get("unit", "сая төг"),
        "total_cur": round(now, 2), "total_prev": round(base, 2),
        "growth_pct": round(growth * 100, 4),
        "components": comps,
        "monthly": [{"period": ("%d-%02d" % pr if cfg["freq"] == "M"
                                else "%d-Q%d" % pr),
                     "cumulative": v,
                     "value": None}
                    for pr, v in cum_tot],
    }
    # single-period value alongside the cumulative (his charts plot this for
    # the hotel and food sectors: cum_t - cum_{t-1}, reset each January)
    prev_v, prev_y = None, None
    for row, (pr, v) in zip(out["monthly"], cum_tot):
        if v is None:
            prev_v = None
            continue
        row["value"] = round(v if pr[0] != prev_y else v - (prev_v or 0), 2)
        prev_v, prev_y = v, pr[0]

    # ---- the analyst's chart needs the WHOLE history, not just this month.
    # His charts are a stacked bar of every component's contribution over
    # time with the headline growth drawn as a line across the top, so the
    # same three-line method is applied at every period, not only the last.
    # Values are fractions (0.0952 = 9.52%) so a 0.0% cell format displays
    # them, which is how his own workbooks are set up.
    labels = [c["label"] for c in comps]
    cum_by = {lbl: prepare(dp) for lbl, dp in sel_pos.items()
              if lbl in labels}
    chart, dropped, resid_lbl = [], [], cfg.get("residual_label")
    for t in range(lag, len(cum_tot)):
        a, b = cum_tot[t][1], cum_tot[t - lag][1]
        if not a or not b:
            continue
        g, dn = a / b - 1, a - b
        if not dn:
            continue
        row = {"period": ("%d-%02d" % periods[t][0] if cfg["freq"] == "M"
                          else "%d-Q%d" % periods[t][0]),
               "growth": round(g, 6), "contributions": {}}
        named = 0.0
        for lbl in labels:
            cs = cum_by.get(lbl)
            if cs is None:
                continue
            ca, cb = cs[t][1], cs[t - lag][1]
            if ca is None or cb is None:
                continue
            row["contributions"][lbl] = round((ca - cb) / dn * g, 6)
            named += ca - cb
        # the same labelled residual as the table, computed at every period
        if resid_lbl and resid_lbl in labels and abs(dn - named) > abs(dn) * 1e-9:
            row["contributions"][resid_lbl] = round((dn - named) / dn * g, 6)
        # A period whose bars would not stack up to the line is dropped, not
        # drawn. Early history is often incomplete -- some components only
        # start reporting later -- and a chart that silently fails to add up
        # is worse than one that starts a few years in.
        if row["contributions"] and abs(
                sum(row["contributions"].values()) - g) <= max(1e-4, abs(g) * 1e-3):
            chart.append(row)
        else:
            dropped.append(row["period"])
    if dropped:
        log.info("real sector[%s]: %d early period(s) left out of the chart "
                 "series because their components do not sum to the headline "
                 "(%s%s) -- incomplete history, not a parsing fault.",
                 key, len(dropped), ", ".join(dropped[:4]),
                 ", ..." if len(dropped) > 4 else "")
    out["chart_series"] = chart
    return out


# INDUSTRY, IN DETAIL -- the four themed charts on the analyst's own slides.
#
# The headline industry table splits into four major groups, and that is what
# `Real_Sector` shows. His deck goes a level deeper and gives sub-sectors
# four charts of their own: mining, food processing, light manufacturing, and
# chemicals and metals. Each series is that sub-sector's contribution to
# TOTAL industry growth, so the four charts read against one number.
#
# The groups are found from the table's own indentation, never by row number:
# NSO adds and reorders sub-sectors between releases.
INDUSTRY_GROUPS = [
    ("mining", "Уул уурхайн дэд салбарууд (Mining sub-sectors)",
     ("уул уурхай, олборлолт",)),
    ("food",   "Хүнсний үйлдвэрлэл (Food processing)",
     ("хүнсний бүтээгдэхүүний үйлдвэрлэл",)),
    ("light",  "Хөнгөн үйлдвэр (Light manufacturing)", None),   # see below
    ("chem",   "Хими, металл (Chemicals and metals)",
     ("кокс болон газрын түүхий тос",)),
]


def _children_of(members, heading_pos):
    """Rows indented deeper than a heading, up to the next row at its level."""
    depth = {p: len(str(l)) - len(str(l).lstrip(" "))
             for p, l in members.items() if str(l).strip()}
    d0 = depth.get(heading_pos)
    if d0 is None:
        return []
    kids = []
    for p in sorted(members):
        if p <= heading_pos:
            continue
        d = depth.get(p)
        if d is None:
            continue
        if d <= d0:
            break
        kids.append(p)
    return kids


def extract_industry_detail():
    """
    Industry sub-sectors, grouped into the analyst's four themed charts.

    Returns per-group component lists AND the full per-period contribution
    series, so the charts are a time series like the rest, not one month.
    """
    log.info("===== INDUSTRY SUB-SECTORS (deck slides 16-17) =====")
    cfg = REAL_SECTOR_TABLES["industry"]
    path = RAW_DATA_DIR / cfg["file"]
    out = {"groups": [], "period": None, "growth_pct": None,
           "unit": cfg.get("unit", "сая төг")}
    if not path.exists():
        log.info("industry detail: %s not cached.", cfg["file"])
        return out
    try:
        js = JsonStat2(json.loads(path.read_text(encoding="utf-8")))
    except Exception as exc:
        log.warning("industry detail: file unusable (%s).", exc)
        return out

    dim = next((d for d in js.ids if d.casefold() == cfg["dim"].casefold()),
               js.ids[0])
    per_dim = next((d for d in js.ids if d != dim), js.ids[-1])
    members = js.labels(dim)
    pers = sorted(((_rs_period(l, cfg["freq"]), p)
                   for p, l in js.labels(per_dim).items()
                   if _rs_period(l, cfg["freq"])), key=lambda x: x[0])
    pers = [(k, p) for k, p in pers
            if k <= (TARGET_YEAR, TARGET_MONTH)]        # never past the month
    if len(pers) <= 12:
        log.warning("industry detail: only %d period(s).", len(pers))
        return out

    def cum(dpos):
        vals = []
        for k, pp in pers:
            try:
                v = js.cell(**{dim: dpos, per_dim: pp})
            except Exception:
                v = None
            vals.append((k, None if v is None else float(v)))
        return _rs_cumulate(vals) if not cfg["cumulative_source"] else vals

    tot_pos = next((p for p, l in members.items()
                    if str(l).strip().casefold() in cfg["total"]), None)
    if tot_pos is None:
        log.warning("industry detail: no total row.")
        return out
    ctot = cum(tot_pos)
    lag = 12

    def head(words):
        return next((p for p, l in members.items()
                     if any(str(l).strip().casefold().startswith(w)
                            for w in words)), None)

    manu = head(("боловсруулах үйлдвэрлэл",))
    picks = {}
    for key, _label, words in INDUSTRY_GROUPS:
        if words:
            h = head(words)
            picks[key] = _children_of(members, h) if h is not None else []
        else:
            # light manufacturing: the direct children of 'Боловсруулах' that
            # are not themselves headings -- i.e. have no children of their
            # own. That excludes food processing and the coke/chemicals block,
            # which get their own charts, without naming either.
            direct = _children_of(members, manu) if manu is not None else []
            depth = {p: len(str(members[p])) - len(str(members[p]).lstrip(" "))
                     for p in direct}
            d0 = min(depth.values()) if depth else 0
            picks[key] = [p for p in direct if depth.get(p) == d0
                          and not _children_of(members, p)]

    cache = {}
    for key, label, _w in INDUSTRY_GROUPS:
        pos = picks.get(key) or []
        if not pos:
            log.warning("industry detail[%s]: no sub-sectors resolved.", key)
            continue
        for p in pos:
            cache.setdefault(p, cum(p))
        rows, comps = [], None
        for t in range(lag, len(pers)):
            a, b = ctot[t][1], ctot[t - lag][1]
            if not a or not b:
                continue
            g, dn = a / b - 1, a - b
            if not dn:
                continue
            contrib = {}
            for p in pos:
                cs = cache[p]
                ca, cb = cs[t][1], cs[t - lag][1]
                if ca is None or cb is None:
                    continue
                contrib[re.sub(r"\s+", " ", str(members[p])).strip()] = \
                    round((ca - cb) / dn * g, 6)
            if not contrib:
                continue
            rows.append({"period": "%d-%02d" % pers[t][0],
                         "contributions": contrib,
                         "subtotal": round(sum(contrib.values()), 6),
                         "growth": round(g, 6)})
        if not rows:
            continue
        comps = list(rows[-1]["contributions"])
        out["groups"].append({
            "key": key, "label": label, "components": comps,
            "series": rows,
            "latest": rows[-1]["contributions"],
            "subtotal_pct": round(rows[-1]["subtotal"] * 100, 4)})
        log.info("industry detail[%s]: %d sub-sector(s), %d period(s), "
                 "group contributes %+.2f pp of the %+.2f%% headline.",
                 key, len(comps), len(rows), rows[-1]["subtotal"] * 100,
                 rows[-1]["growth"] * 100)

    if out["groups"]:
        last = out["groups"][0]["series"][-1]
        out["period"] = last["period"]
        out["growth_pct"] = round(last["growth"] * 100, 4)
        # The four groups together should account for the whole of industry.
        # They will not always -- a sub-sector can be published empty -- so
        # the shortfall is stated rather than hidden.
        tot = sum(g["subtotal_pct"] for g in out["groups"])
        gap = out["growth_pct"] - tot
        log.info("industry detail: four groups sum to %+.2f pp of the "
                 "%+.2f%% headline (unattributed %+.2f pp).",
                 tot, out["growth_pct"], gap)
        if abs(gap) > max(0.5, abs(out["growth_pct"]) * 0.02):
            log.warning("industry detail: %+.2f pp of industry growth is not "
                        "covered by the four sub-sector groups -- a "
                        "sub-sector may be missing from the table.", gap)
    return out


def extract_real_sector():
    """Trade, hotel, food, construction, industry (and transport when added)."""
    log.info("===== NSO REAL SECTOR (trade, services, construction, "
             "industry) =====")
    out = {}
    for key, cfg in REAL_SECTOR_TABLES.items():
        if not (RAW_DATA_DIR / cfg["file"]).exists():
            if key == "transport":
                continue          # table id not captured yet, by design
            log.warning("real sector[%s]: %s not cached -- run "
                        "raw_ingestor.py.", key, cfg["file"])
            continue
        try:
            d = _extract_one_real_sector(key, cfg)
        except Exception as exc:
            log.warning("real sector[%s] failed: %s", key, exc)
            continue
        if not d:
            continue
        out[key] = d
        # The two published levels must themselves reproduce the published
        # growth rate. They are computed at different points in the routine,
        # so this catches a level being overwritten between them -- which is
        # exactly how an indent depth once reached the sheet as a prior-year
        # figure.
        if d["total_prev"]:
            lg = (d["total_cur"] / d["total_prev"] - 1) * 100
            if abs(lg - d["growth_pct"]) > 0.01:
                log.warning("real sector[%s]: levels %s -> %s imply %+.2f%% "
                            "but the headline says %+.2f%% -- the levels and "
                            "the growth rate disagree.", key,
                            format(d["total_prev"], ",.1f"),
                            format(d["total_cur"], ",.1f"), lg,
                            d["growth_pct"])
        else:
            log.warning("real sector[%s]: prior-year total is missing or "
                        "zero; the growth rate cannot be checked against "
                        "the levels.", key)
        # contributions are additive by construction: they MUST sum to the
        # headline. A mismatch means a component was missed or the total row
        # was misidentified.
        s = sum(c["contribution_pct"] or 0 for c in d["components"])
        ok = abs(s - d["growth_pct"]) <= max(0.05, abs(d["growth_pct"]) * 1e-3)
        log.info("%-13s %s: total %s (%+.2f%% YoY), %d components, "
                 "contributions %+.2f pp %s", key, d["period"],
                 format(d["total_cur"], ",.0f"), d["growth_pct"],
                 len(d["components"]), s,
                 "= headline" if ok else "!= headline")
        if not ok:
            log.warning("real sector[%s]: contributions sum %.2f but headline "
                        "is %.2f -- a component may be missing or the total "
                        "row misidentified.", key, s, d["growth_pct"])
    if not out:
        log.warning("No real-sector tables available.")
    return out


# IMF WORLD ECONOMIC OUTLOOK — the benchmark and the forward view
#
# Every other source in this workbook MEASURES the present from Mongolian
# primary data. WEO is different: it is the IMF's annual forecast, published
# twice a year. Listing its indicators as more numbers would waste it; the
# useful thing is the GAP between what we measure and what the IMF expects,
# plus the multi-year path.
#
# Comparisons are made only where the basis genuinely matches, and the basis
# is written on every row. WEO's fiscal balance and current account are
# percent of GDP while ours are cumulative levels in tugrik and dollars, so
# those are shown as projections only -- forcing a comparison there would be
# worse than not making one.
WEO_SERIES = [
    ("NGDP_RPCH",   "Бодит ДНБ-ий өсөлт (Real GDP growth)", "%", "gdp_growth",
     "IMF: full calendar year · ours: NSO cumulative to the reported quarter"),
    ("PCPIEPCH",    "Инфляц, оны эцэс (CPI, end of period)", "%", "inflation",
     "IMF: December on December · ours: bulletin year-on-year for the "
     "reported month"),
    ("PCPIPCH",     "Инфляц, жилийн дундаж (CPI, period average)", "%", None,
     "annual average — not comparable with a single month's reading"),
    ("GGXCNL_NGDP", "Төсвийн тэнцэл (Fiscal balance, % of GDP)", "% ДНБ", None,
     "IMF states this as % of GDP; the Budget sheet is cumulative bln MNT"),
    ("GGXWDG_NGDP", "Улсын өр (Government gross debt, % of GDP)", "% ДНБ",
     None, "projection only"),
    ("BCA_NGDPD",   "Урсгал дансны тэнцэл (Current account, % of GDP)",
     "% ДНБ", None,
     "IMF states this as % of GDP; External_Sector is cumulative mln USD"),
    ("NGDPD",       "ДНБ (GDP, тэрбум ам.доллар)", "тэрбум $", None,
     "projection only"),
    ("LUR",         "Ажилгүйдлийн түвшин (Unemployment rate)", "%", None,
     "projection only"),
    ("NID_NGDP",    "Хөрөнгө оруулалт (Gross capital formation, % of GDP)",
     "% ДНБ", None, "projection only"),
    ("NGSD_NGDP",   "Хуримтлал (Gross national savings, % of GDP)", "% ДНБ",
     None, "projection only"),
]
WEO_DIVERGENCE_PP = 1.5      # flag an actual this far from the forecast


# WEEKLY PRICES -- NSO's seven-day price survey for Ulaanbaatar.
# The deck has always carried a meat-and-fuel price slide; until now the two
# figures on it were typed in by hand from a presentation PDF. This is the
# same information at source, weekly, and it retires those yellow cells.
#
# Products are picked by KEYWORD, never by row position or an exact caption:
# NSO renames and reorders this list between releases, and the survey covers
# 31 items of which the deck wants a handful.
WEEKLY_PRICE_GROUPS = [
    ("meat",  "Мах (meat)",
     ("хонин", "хонь", "ямаан", "үхрийн", "гахайн", "адууны", "мах")),
    ("fuel",  "Шатахуун (fuel)",
     ("аи-", "аи ", "а-80", "а-92", "а-95", "бензин", "дизел", "дизель")),
    ("flour", "Гурил, талх (flour, bread)", ("гурил", "талх", "будаа")),
    ("other", "Бусад хүнс (other food)", ()),        # everything left over
]
WEEKLY_PRICE_LOOKBACK = 130          # ~2.5 years of weeks, enough for YoY
# NSO publishes this survey as more than one table. They cover the same
# thing but do not update together, so all cached tables are read and the
# freshest one wins. Add another line here if a third appears.
WEEKLY_PRICE_SOURCES = [
    ("nso_weekly_prices.json",       "capital city (DT_NSO_0600_001V4)"),
    ("nso_weekly_prices_aimag.json", "by aimag (DT_NSO_0300_010V5)"),
]


def _week_key(label):
    """
    A sortable (year, month, day) from whatever NSO writes in 'Хугацаа'.

    Seen so far: ISO '2026-08-03'. Older exports use '2026 оны 8 сарын 3' and
    '2026/08/03'. Anything with four digits then two more numbers is accepted;
    a label that yields nothing is skipped rather than guessed at.
    """
    s = str(label).strip()
    m = re.match(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        y, mo, d = (int(x) for x in m.groups())
        if 1 <= mo <= 12 and 1 <= d <= 31:
            return (y, mo, d)
    return None


def extract_weekly_prices():
    """
    Weekly Ulaanbaatar prices: every product's series, grouped for the deck.

    Two things the deck needs and the monthly CPI cannot give: the price of
    meat and fuel WITHIN the reported month, and the change since the start
    of the year. Both are computed here so no one has to read them off a
    chart.
    """
    log.info("===== WEEKLY PRICES (NSO 7-day survey) =====")
    out = {"products": [], "groups": [], "weeks": [], "as_of": None,
           "clamped": False, "source_file": None, "month_end_as_of": None,
           "considered": []}

    # NSO publishes this survey as more than one table and they do not
    # always update together: in August 2026 the capital-city table stalled
    # in July while the by-aimag table was current. Read every cached table
    # and use whichever reaches the newest week, rather than trusting one.
    best = None
    for fname, what in WEEKLY_PRICE_SOURCES:
        p = RAW_DATA_DIR / fname
        if not p.exists():
            continue
        try:
            js = JsonStat2(json.loads(p.read_text(encoding="utf-8")))
        except Exception as exc:
            log.warning("weekly prices: %s unusable (%s).", fname, exc)
            continue
        dim_p = next((d for d in js.ids if "бүтээгдэхүүн" in d.casefold()),
                     js.ids[0])
        dim_t = next((d for d in js.ids if "хугацаа" in d.casefold()),
                     js.ids[-1])
        dated = sorted(((_week_key(l), pos)
                        for pos, l in js.labels(dim_t).items()
                        if _week_key(l)), key=lambda x: x[0])
        if not dated:
            log.warning("weekly prices: %s has no parseable week labels.",
                        fname)
            continue
        newest = "%04d-%02d-%02d" % dated[-1][0]
        out["considered"].append(f"{what} -> {newest} ({len(dated)} weeks)")
        if best is None or dated[-1][0] > best[3][-1][0]:
            best = (fname, what, js, dated, dim_p, dim_t)
    if best is None:
        log.warning("weekly prices: no usable table cached -- run "
                    "raw_ingestor.py.")
        return {}, out
    fname, what, js, dated, dim_p, dim_t = best
    out["source_file"] = fname
    for line in out["considered"]:
        log.info("weekly prices: %s", line)
    if len(out["considered"]) > 1:
        log.info("weekly prices: using %s, the freshest of %d table(s).",
                 fname, len(out["considered"]))

    # Some of these tables carry a region dimension as well. Take the
    # national or capital row -- the deck talks about Ulaanbaatar prices --
    # and pin it, so the cell lookup below stays two-dimensional.
    fixed = {}
    for d in js.ids:
        if d in (dim_p, dim_t):
            continue
        labs = js.labels(d)
        pick = next((pos for pos, l in labs.items()
                     if any(w in str(l).casefold()
                            for w in ("улсын дүн", "улаанбаатар", "нийслэл",
                                      "бүгд", "нийт"))), min(labs))
        fixed[d] = pick
        log.info("weekly prices: %s has a %r dimension; using %r.",
                 fname, d, str(labs[pick]).strip())

    # NO FORWARD CLAMP HERE, and that is deliberate. Every other source is
    # cut off at the reported month because it measures that month. This one
    # does not: it is a LEADING indicator, published weekly, and the whole
    # point of the price slide is where meat and fuel are heading before the
    # monthly CPI catches up. Cutting it at the month end froze the slide on
    # stale prices. The series therefore runs to the newest published week
    # and carries that week's own date everywhere it appears; the value at
    # the reported month end is kept alongside it for continuity.
    keep = dated[-WEEKLY_PRICE_LOOKBACK:]
    if len(keep) < 8:
        log.warning("weekly prices: only %d week(s) available.", len(keep))
        return {}, out

    out["weeks"] = ["%04d-%02d-%02d" % k for k, _ in keep]
    out["as_of"] = out["weeks"][-1]
    month_end_i = max((i for i, (k, _) in enumerate(keep)
                       if (k[0], k[1]) <= (TARGET_YEAR, TARGET_MONTH)),
                      default=None)
    if month_end_i is not None:
        out["month_end_as_of"] = out["weeks"][month_end_i]
    ahead = len(keep) - 1 - (month_end_i if month_end_i is not None else -1)
    if ahead > 0:
        log.info("weekly prices: %d week(s) run past the reported month; "
                 "kept, because this is a leading indicator. Latest week %s, "
                 "last week inside %d-%02d was %s.", ahead, out["as_of"],
                 TARGET_YEAR, TARGET_MONTH, out["month_end_as_of"])

    def val(ppos, tpos):
        try:
            v = js.cell(**{dim_p: ppos, dim_t: tpos, **fixed})
        except Exception:
            return None
        return None if v is None else float(v)

    prods = js.labels(dim_p)

    # first week of the reported year, for the year-to-date change
    ytd0 = next((i for i, (k, _) in enumerate(keep) if k[0] == TARGET_YEAR), 0)
    for ppos, lbl in sorted(prods.items()):
        name = re.sub(r"\s+", " ", str(lbl)).strip()
        if not name:
            continue
        vals = [val(ppos, tp) for _, tp in keep]
        if not any(v is not None for v in vals):
            continue
        cur = next((v for v in reversed(vals) if v is not None), None)
        yr0 = next((v for v in vals[ytd0:] if v is not None), None)
        y_ago = vals[-53] if len(vals) >= 53 else None
        # the same product at the last week inside the reported month, so a
        # slide can quote either and say which it is quoting
        me = None
        if month_end_i is not None:
            me = next((v for v in reversed(vals[:month_end_i + 1])
                       if v is not None), None)
        out["products"].append({
            "label": name, "values": vals, "current": cur,
            "month_end": me,
            "ytd_pct": (round((cur / yr0 - 1) * 100, 2)
                        if cur and yr0 else None),
            "yoy_pct": (round((cur / y_ago - 1) * 100, 2)
                        if cur and y_ago else None),
        })

    # group them the way the slide talks about them
    taken = set()
    for key, label, words in WEEKLY_PRICE_GROUPS:
        if words:
            members = [p for p in out["products"]
                       if any(w in p["label"].casefold() for w in words)
                       and p["label"] not in taken]
        else:
            members = [p for p in out["products"]
                       if p["label"] not in taken]
        taken.update(p["label"] for p in members)
        ytd = [p["ytd_pct"] for p in members if p["ytd_pct"] is not None]
        yoy = [p["yoy_pct"] for p in members if p["yoy_pct"] is not None]
        if members:
            out["groups"].append({
                "key": key, "label": label,
                "members": [p["label"] for p in members],
                "ytd_pct": round(sum(ytd) / len(ytd), 2) if ytd else None,
                "yoy_pct": round(sum(yoy) / len(yoy), 2) if yoy else None,
            })

    kpis = {}
    for g in out["groups"]:
        if g["key"] in ("meat", "fuel"):
            # these two replace the yellow cells the operator used to fill
            kpis[f"weekly_{g['key']}_ytd_pct"] = g["ytd_pct"]
            kpis[f"weekly_{g['key']}_yoy_pct"] = g["yoy_pct"]
    kpis["_weekly_prices_as_of"] = out["as_of"]
    log.info("weekly prices: %d product(s), %d week(s) to %s. %s",
             len(out["products"]), len(out["weeks"]), out["as_of"],
             " | ".join(f"{g['label']}: YTD {g['ytd_pct']:+.1f}%"
                        for g in out["groups"]
                        if g["ytd_pct"] is not None))
    return kpis, out


def extract_imf_weo(kpis):
    """IMF WEO projections for Mongolia, plus actual-vs-forecast where the
    basis matches. Returns {} when no WEO export is cached."""
    log.info("===== IMF WORLD ECONOMIC OUTLOOK =====")
    cand = sorted(RAW_DATA_DIR.glob("*WEO*.csv")) + \
        sorted(RAW_DATA_DIR.glob("imf_weo*.csv"))
    if not cand:
        log.info("No IMF WEO export cached (raw_files/*WEO*.csv) -- the "
                 "Forecast sheet stays empty. Download Mongolia's WEO data "
                 "from imf.org and drop the CSV into raw_files/.")
        return {}
    path = cand[0]
    try:
        import csv as _csv
        with path.open(encoding="utf-8-sig", newline="") as fh:
            rows = list(_csv.DictReader(fh))
    except Exception as exc:
        log.warning("IMF WEO %s unreadable (%s).", path.name, exc)
        return {}
    if not rows:
        log.warning("IMF WEO %s is empty.", path.name)
        return {}

    years = sorted(k for k in rows[0]
                   if re.fullmatch(r"(19|20)\d{2}", str(k).strip()))
    if not years:
        log.warning("IMF WEO %s: no year columns found.", path.name)
        return {}

    def code_of(r):
        # 'MNG.NGDP_RPCH.A' -> 'NGDP_RPCH'
        parts = str(r.get("SERIES_CODE", "")).split(".")
        return parts[1] if len(parts) > 2 else ""

    by_code = {code_of(r): r for r in rows if code_of(r)}
    country = (rows[0].get("COUNTRY") or "").strip()

    def num(x):
        try:
            return float(str(x).replace(",", "").strip())
        except (TypeError, ValueError):
            return None

    out = {"country": country, "source": path.name,
           "vintage": _weo_vintage(path), "years": years,
           "indicators": [], "comparison": []}

    for code, label, unit, kpi_key, basis in WEO_SERIES:
        r = by_code.get(code)
        if not r:
            log.info("IMF WEO: series %s not in the export -- skipped.", code)
            continue
        vals = {y: num(r.get(y)) for y in years}
        if not any(v is not None for v in vals.values()):
            continue
        out["indicators"].append({
            "code": code, "label": label, "unit": unit, "basis": basis,
            "values": vals})
        # actual vs forecast, only where the basis matches
        if kpi_key and kpis.get(kpi_key) is not None:
            f = vals.get(str(TARGET_YEAR))
            if f is not None:
                a = float(kpis[kpi_key])
                out["comparison"].append({
                    "label": label, "unit": unit, "basis": basis,
                    "forecast": round(f, 3), "actual": round(a, 3),
                    "gap": round(a - f, 3)})
    if not out["indicators"]:
        log.warning("IMF WEO: none of the expected series were found in %s.",
                    path.name)
        return {}

    log.info("IMF WEO (%s, %s): %d indicator(s), %s-%s.", country,
             out["vintage"], len(out["indicators"]), years[0], years[-1])
    for c in out["comparison"]:
        flag = abs(c["gap"]) >= WEO_DIVERGENCE_PP
        (log.warning if flag else log.info)(
            "  %s %d: IMF %+.2f%s vs actual %+.2f%s -> %+.2f pp%s",
            c["label"][:34], TARGET_YEAR, c["forecast"], c["unit"],
            c["actual"], c["unit"], c["gap"],
            "  <-- diverging" if flag else "")
    return out


def _weo_vintage(path):
    """Best available description of which WEO release the file is."""
    m = re.search(r"(20\d{2})-(\d{2})-(\d{2})", path.name)
    if m:
        y, mth = int(m.group(1)), int(m.group(2))
        # WEO is published in April and October
        rel = "April" if 4 <= mth <= 9 else ("October" if mth >= 10 else
                                             "October (prev. year)")
        yr = y if mth >= 4 else y - 1
        return f"{yr} {rel} vintage (exported {m.group(0)})"
    try:
        return "exported " + datetime.fromtimestamp(
            path.stat().st_mtime).strftime("%Y-%m-%d")
    except Exception:
        return "vintage unknown"


# DATA VINTAGE — which month each source actually carries
# Sources publish on different calendars, so a report for month M is always
# assembled from a mix: some sources are at M, others still at M-1. Leaving
# the reader to infer that from scattered footnotes is how a previous month's
# figure gets read as the current one. This states it once, per source.
def build_vintage(kpis, series):
    tgt = f"{TARGET_YEAR}-{TARGET_MONTH:02d}"
    rows = []

    def add(source, feeds, period, note=""):
        # A quarterly period must not be compared as a string: '2026-Q1'
        # sorts after '2026-06' because 'Q' > '0', which once labelled a
        # first-quarter table as AHEAD of a June report. Quarters are their
        # own status and are ranked by the month the quarter ends in.
        p = str(period) if period else ""
        qm = re.fullmatch(r"(\d{4})-Q([1-4])", p)
        if not period:
            status = "MISSING"
        elif qm:
            status = "QUARTERLY"
            if f"{qm.group(1)}-{int(qm.group(2)) * 3:02d}" > tgt:
                status = "AHEAD"
        elif p == tgt:
            status = "CURRENT"
        elif p < tgt:
            status = "LAGGED"
        else:
            status = "AHEAD"
        rows.append({"source": source, "feeds": feeds,
                     "period": str(period) if period else "n/a",
                     "status": status, "note": note})

    add("Customs (gaali.mn)", "Trade, Commodities, Border prices, "
        "Trade_Sections", tgt if kpis.get("total_export") is not None else None,
        "monthly bulletin for the reported month"
        if kpis.get("total_export") is not None
        else "target workbook not available")

    nso_q = kpis.get("gdp_growth_period")
    add("NSO national accounts", "GDP_Growth, GDP_Sectors", nso_q,
        "quarterly series — the latest quarter ending on or before the "
        "reported month")
    rows[-1]["status"] = "QUARTERLY" if nso_q else "MISSING"

    # NSO's own name for this, from the table itself: 'ЭДИЙН ЗАСГИЙН ӨСӨЛТ,
    # БУУРАЛТ, сараар, өссөн дүнгээр'. It is the monthly GDP estimate at
    # constant prices, methodology approved by NSO order A/150 of 25 Oct
    # 2024 -- not an index and not an acronym anyone in Mongolia would
    # recognise as 'МЭЗҮИ'. The deck should call it what NSO calls it.
    add("Эдийн засгийн сарын өсөлт (NSO monthly GDP growth)",
        "GDP_Growth — monthly growth and sector contributions",
        kpis.get("_mieg_period"),
        kpis.get("_mieg_period_note", "")
        or "official monthly GDP estimate, cumulative from January")

    add("Mongolbank bulletin", "Inflation, policy rate, budget, FX",
        kpis.get("_bulletin_period"),
        kpis.get("_bulletin_period_note", "")
        or f"issue {kpis.get('_bulletin_source', '')}")

    bs_col = str(kpis.get("_bs_last_column", "") or "")
    add("Bank balance sheet", "Banking (assets, loans, NPL, equity)",
        bs_col[:7].replace(".", "-") if bs_col else None,
        f"latest column in the report: {bs_col}" if bs_col else "")

    dep_cols = (series.get("banking_balance_sheet", {})
                .get("deposit_columns") or [])
    add("Deposit tree", "Banking (харилцах, хадгаламж)",
        str(dep_cols[-1]).replace(".", "-")[:7] if dep_cols else None,
        f"source: {kpis.get('_deposit_source', 'n/a')}, "
        f"{len(dep_cols)} period column(s)")

    add("Bank loan report", "Loan_Detail (bridge, sectors, NPL, rates)",
        kpis.get("_loan_period"),
        kpis.get("_loan_period_note", "") or "monthly flow statement")

    add("Balance of payments", "External_Sector (BoP detail)",
        kpis.get("_bop_period"),
        (series.get("bop", {}) or {}).get("detail_lag_note", "")
        or "manual export, cumulative")

    for key, d in (series.get("real_sector") or {}).items():
        add(f"NSO {d.get('label','')[:26]}", "Real_Sector — " + key,
            d.get("period"),
            f"{d.get('freq','M')}, {len(d.get('components',[]))} components")

    wk = series.get("weekly_prices") or {}
    if wk.get("as_of"):
        me = wk.get("month_end_as_of")
        add("NSO 7 хоногийн үнийн мэдээ (weekly prices)",
            "Weekly_Prices — meat and fuel, slide 9", wk["as_of"],
            f"{len(wk.get('products', []))} products, "
            f"{len(wk.get('weeks', []))} weeks, from "
            f"{wk.get('source_file', '?')}"
            + (f" · LEADING INDICATOR: runs past the reported month; last "
               f"week inside it was {me}"
               if me and me != wk["as_of"] else ""))
        # A weekly date never equals the reported month string, so the
        # generic comparison would call it MISSING or AHEAD. It is neither:
        # it is the freshest week that still falls inside the month.
        rows[-1]["status"] = "WEEKLY"

    weo = series.get("imf_weo") or {}
    if weo:
        add("IMF WEO", "Forecast — projections and actual-vs-forecast",
            None, f"{weo.get('vintage','')} · ANNUAL projections, not "
                  f"measurements · {weo.get('source','')}")
        rows[-1]["status"] = "FORECAST"
        rows[-1]["period"] = weo.get("vintage", "n/a")

    card_dates = sorted({str(v) for k, v in kpis.items()
                         if k.endswith("_date") and v})
    add("Mongolbank cards", "External_Sector (reserves, FDI, debt)",
        None, "each card carries its own 'as of' date: "
        + ", ".join(card_dates[:4]) if card_dates else "")
    rows[-1]["status"] = "PER-CARD"
    rows[-1]["period"] = "see card dates"

    lagged = [r for r in rows if r["status"] == "LAGGED"]
    if lagged:
        log.warning("DATA VINTAGE: %d source(s) lag %s -> %s", len(lagged),
                    tgt, "; ".join(f"{r['source']} at {r['period']}"
                                   for r in lagged))
    else:
        log.info("DATA VINTAGE: every dated source is at %s", tgt)
    return {"target": tgt, "rows": rows}


# COVERAGE (slide-by-slide status for the Excel 'Coverage' sheet)
def build_coverage(kpis, series):
    ce = series.get("customs", {}).get("commodity_exports", {})
    gdp = series.get("nso", {})
    have = {
        "gdp": bool(gdp.get("cumulative_growth")),
        "sectors": bool(gdp.get("contributions")),
        "trade": kpis.get("total_export") is not None,
        "commod": bool(ce),
        "prices": bool(series.get("customs", {}).get("border_prices")),
        "med": kpis.get("inflation") is not None,
        "mb": kpis.get("reserves_mln_usd_mb") is not None,
        "policy": kpis.get("policy_rate") is not None,
        "monthly": len(series.get("customs", {}).get("trade_monthly", [])) > 2,
    }
    def s(ok, src, note=""):
        return {"status": "AUTO" if ok else "MISSING", "source": src, "note": note}
    rows = [
        (1,  "Title / date", s(True, "pipeline", "report_period string")),
        (2,  "Geopolitics", {"status": "MANUAL", "source": "-",
                             "note": "editorial; no data feed"}),
        (3,  "Macro forecast snapshot", s(have["med"] and have["gdp"], "Bulletin+NSO+MB",
             "policy rate " + ("auto" if have["policy"] else "manual"))),
        (4,  "Real GDP cumulative growth", s(have["gdp"], "NSO")),
        (5,  "Sector contributions", s(have["sectors"], "NSO",
             "NSO table has 10 aggregated sectors")),
        (6,  "Inflation vs target", s(have["med"], "MB bulletin",
             "headline + 16-month national/UB CPI series, all auto")),
        (7,  "Meat & fuel prices", {"status": "MANUAL",
             "source": "NSO CPI PDF",
             "note": "by decision (2026-07): heavy PDF parsing not worth "
                     "automating; enter values by hand"}),
        (8,  "Inflation -> interest rates", s(have["policy"], "MB bulletin",
             "policy & deposit rate monthly series + CPI, all auto")),
        (9,  "Trade balance & commodities", s(have["trade"] and have["commod"],
             "Customs")),
        (10, "Export volumes", s(have["commod"], "Customs")),
        (11, "Border prices", s(have["prices"], "Customs",
             "unit values; monthly once history cached" if not have["monthly"]
             else "monthly marginal unit values")),
        (12, "Balance of payments", s(have["mb"], "Mongolbank",
             "CA/KA/FA parsed from bop_manual.xlsx"
             if kpis.get("bop_current_account") is not None
             else "cards auto; CA/KA/FA from the manual BoP file "
             "(raw_files/bop_manual.xlsx -- ingestor prompts for it)")),
        (13, "State budget", s(have["med"], "MB bulletin")),
        (14, "Household income/expense", {"status": "MISSING", "source": "NSO",
             "note": "needs Household Socio-Economic Survey table id"}),
        (15, "Household debt", {"status": "MISSING", "source": "NSO",
             "note": "same survey table"}),
        (16, "Banking sector", s(kpis.get("bs_total_assets_mln_mnt")
                                 is not None, "Mongolbank",
             "balance-sheet xlsx parsed: assets, loans, CB securities, "
             "deposits, past-due, NPL, equity + monthly series")),
        (17, "Moody's slides (17-22)", {"status": "MANUAL", "source": "Moody's",
             "note": "proprietary images; never automatable"}),
    ]
    return [{"slide": n, "content": c, **d} for n, c, d in rows]


# MAIN
def main():
    log.info("########## MODULE 2 v2: METRICS + SERIES BUILD ##########")
    files = customs_files()
    trade_kpis, customs_series = extract_customs(files)
    nso_kpis, nso_series = extract_nso()
    med_kpis, med_notes = extract_bulletin()
    mb_kpis, mb_cards = extract_mongolbank()
    mieg_kpis, mieg_series = extract_mieg()
    bs_kpis, bs_series = extract_banking_balance_sheet()
    bop_kpis, bop_series = extract_bop_detail()
    loan_kpis, loan_series = extract_bank_loans()
    real_sector = extract_real_sector()
    industry_detail = extract_industry_detail()
    wk_kpis, wk_series = extract_weekly_prices()
    weo = extract_imf_weo({**trade_kpis, **nso_kpis, **mieg_kpis, **med_kpis,
                           **mb_kpis, **bs_kpis, **bop_kpis, **loan_kpis})

    period = {
        "target_year": TARGET_YEAR, "target_month": TARGET_MONTH,
        "report_period": f"{TARGET_YEAR} оны {TARGET_MONTH}-р сар",
        "report_date": f"{TARGET_YEAR} оны {TARGET_MONTH}-р сарын байдлаар",
        "report_period_en": f"{TARGET_YEAR}-{TARGET_MONTH:02d}",
    }

    kpis = {"_comment": ("Auto-generated by data_processor.py v2. Money in "
                         "THOUSAND USD unless suffixed (_mln_usd, _bln_mnt, "
                         "_tln); rates/growth in %.")}
    kpis.update(period)
    for chunk in (trade_kpis, nso_kpis, mieg_kpis, med_kpis, mb_kpis,
                  bs_kpis, bop_kpis, loan_kpis, wk_kpis):
        kpis.update(chunk)
    if "policy_rate" not in kpis:
        kpis["_policy_rate_source"] = "NOT FOUND - fill manually"
    elif not kpis.get("_policy_rate_source"):
        kpis["_policy_rate_source"] = "Mongolbank card"

    series = {"period": period, "customs": customs_series, "nso": nso_series,
              "bulletin_notes": med_notes, "mieg": mieg_series,
              "bulletin_cpi": {"components": med_notes.get("cpi_components",
                                                           []),
                               "period": med_notes.get(
                                   "cpi_components_period"),
                               "prev": med_notes.get("cpi_components_prev",
                                                     []),
                               "prev_period": med_notes.get(
                                   "cpi_components_prev_period")},
              "bulletin_series": med_notes.get("series", {}), "mongolbank_cards": mb_cards,
              "banking_balance_sheet": bs_series, "bop": bop_series,
              "bank_loans": loan_series, "real_sector": real_sector,
              "industry_detail": industry_detail,
              "weekly_prices": wk_series, "imf_weo": weo}
    series["vintage"] = build_vintage(kpis, series)
    series["coverage"] = build_coverage(kpis, series)

    METRICS_JSON.write_text(json.dumps(kpis, ensure_ascii=False, indent=2),
                            encoding="utf-8")
    SERIES_JSON.write_text(json.dumps(series, ensure_ascii=False, indent=2),
                           encoding="utf-8")
    log.info("Wrote %s (%d keys) and %s.", METRICS_JSON.name,
             len([k for k in kpis if not k.startswith("_")]), SERIES_JSON.name)
    log.info("########## DONE ##########")


if __name__ == "__main__":
    main()
