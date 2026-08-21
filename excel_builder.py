#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Builds Macro_Metrics_<year>_<month>.xlsx from the processed JSONs:
11 sheets, formula-driven, bilingual labels.
See the Data Dictionary document for the sheet-by-sheet layout.

Usage: python excel_builder.py
"""

import json
import logging
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
METRICS_JSON = BASE_DIR / "processed_metrics.json"
SERIES_JSON = BASE_DIR / "processed_series.json"

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s | %(levelname)-7s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("excel_builder")

# styles
FONT = "Arial"
F_TITLE = Font(name=FONT, size=13, bold=True, color="1F3864")
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_BODY = Font(name=FONT, size=10)
F_LINK = Font(name=FONT, size=10, color="008000")   # green = cross-sheet link
F_NOTE = Font(name=FONT, size=9, italic=True, color="808080")
FILL_HDR = PatternFill("solid", start_color="1F3864")
FILL_ALT = PatternFill("solid", start_color="EDF2F9")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
# number formats: thousands separators, parentheses negatives, dash zeros,
# applied via .number_format ONLY (never string conversion in Python)
NUM = "#,##0.0;(#,##0.0);\"-\""
NUM2 = "#,##0.00;(#,##0.00);\"-\""
INT = "#,##0;(#,##0);\"-\""
PCT = "0.0%;(0.0%);\"-\""
MULT = "0.0\"x\""

CUM_PERIODS = {1: "I-III", 2: "I-VI", 3: "I-IX", 4: "I-XII"}


def _title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = F_NOTE
    ws.freeze_panes = "A4"
    return 4   # first content row


# Every sheet states the vintage of its OWN data. The Data_Vintage sheet
# gives the whole picture, but a reader who opens straight into Inflation or
# Banking should not have to go looking: the month is on the sheet in front
# of them, green when it matches the reported month and red when it lags.
_VINTAGE = {"target": "", "by_source": {}}


def _set_vintage(series):
    v = series.get("vintage", {}) or {}
    _VINTAGE["target"] = v.get("target", "")
    _VINTAGE["by_source"] = {d.get("source"): d for d in v.get("rows", [])}


def _stamp(ws, r, *sources):
    """Write the 'data as of' banner for this sheet's own source(s)."""
    parts, lagged, missing = [], False, False
    for s in sources:
        d = _VINTAGE["by_source"].get(s)
        if not d:
            continue
        parts.append(f"{s} — {d.get('period')} ({d.get('status')})")
        if d.get("status") == "LAGGED":
            lagged = True
        if d.get("status") == "MISSING":
            missing = True
    if not parts:
        return r
    tgt = _VINTAGE.get("target", "")
    txt = ("МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  " + "   |   ".join(parts)
           + f"      [тайлант үе / reported month: {tgt}]")
    if lagged:
        txt += "   ← ЭНЭ ХУУДСЫН ТОО ӨМНӨХ САРЫНХ (previous month's data)"
    c = _cell(ws, r, 1, txt, bold=True)
    c.fill = PatternFill("solid",
                         start_color="FFC7CE" if lagged else
                         ("FFEB9C" if missing else "C6EFCE"))
    return r + 2


def _header_row(ws, row, headers, widths=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return row + 1


def _cell(ws, row, col, value, fmt=None, bold=False, alt=False, link=False):
    # Any division can meet an empty denominator when a source has not
    # published yet, and a sheet full of #DIV/0! is worse than a blank one.
    # Wrapping every division formula keeps the workbook presentable in the
    # worst case; the self-checks that matter are asserted in the run log,
    # not in the cells, so nothing diagnostic is being hidden here.
    if isinstance(value, str) and value.startswith("=") and "/" in value:
        value = f'=IFERROR({value[1:]},"")'
    c = ws.cell(row=row, column=col, value=value)
    c.font = F_LINK if link else (F_BOLD if bold else F_BODY)
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if alt:
        c.fill = FILL_ALT
    return c


def pct_static(x):
    """Percent number -> Excel fraction (11.2 -> 0.112) for PCT-formatted
    cells that hold a value rather than a formula."""
    return x / 100.0 if x is not None else None


COMMODITY_EN = {
    "coal": "Coal", "copper": "Copper concentrate", "gold": "Gold, unwrought",
    "iron": "Iron ore & concentrate", "wool_cashmere": "Wool & cashmere",
    "petroleum": "Petroleum products", "cars": "Passenger cars",
    "electricity": "Electricity", "trucks": "Trucks",
    "machinery": "Machinery & equipment", "base_metals": "Base metals",
    "food": "Food products",
}


def _commodity_label(key, d):
    en = COMMODITY_EN.get(key)
    label = d.get("label", key)
    return f"{label} ({en})" if en else str(label)


SECTOR_EN = {
    "Хөдөө аж ахуй": "Agriculture, forestry, fishing & hunting",
    "Уул уурхай": "Mining & quarrying",
    "Боловсруулах": "Manufacturing",
    "Цахилгаан": "Electricity, gas, steam & air conditioning",
    "Барилга": "Construction",
    "Бөөний болон жижиглэн": "Wholesale & retail trade; vehicle repair",
    "Тээвэр": "Transportation & storage",
    "Мэдээлэл": "Information & communication",
    "Үйлчилгээний бусад": "Other service activities",
    "Бүтээгдэхүүний цэвэр татвар": "Net taxes on products",
}


def _sector_label(name):
    """'Уул уурхай, олборлолт' -> 'Уул уурхай, олборлолт (Mining & quarrying)'."""
    for frag, en in SECTOR_EN.items():
        if str(name).startswith(frag):
            return f"{name} ({en})"
    return str(name)


# sheets
def sheet_gdp(wb, series):
    """
    Quarterly REAL GDP levels are the only hardcoded inputs; the YoY column
    and the whole cumulative-growth matrix are Excel formulas over them.
    Returns cell refs the Summary sheet links to.
    """
    nso = series.get("nso", {})
    ws = wb.create_sheet("GDP_Growth")
    r = _title(ws, "Real GDP (2015 prices) — levels & growth",
               "Deck slide 4. Source datapoints: NSO National Accounts "
               "(DT_NSO_0500_004V1). All growth figures are formulas over "
               "the quarterly levels below.")
    r = _stamp(ws, r, "NSO national accounts", "NSO MIEG")

    ws.cell(row=r, column=1, value="Улирлын бодит ДНБ (Quarterly real GDP levels, "
            "mln MNT, 2015 prices)").font = F_BOLD
    r += 1
    r = _header_row(ws, r, ["Он (Year)", "Улирал (Quarter)",
                            "Бодит ДНБ (Real GDP, 2015p)",
                            "Жилийн өсөлт (YoY growth)"], [10, 10, 20, 12])
    q = nso.get("quarterly", [])
    qrow = {}                         # (year, quarter) -> sheet row
    for d in q:
        qrow[(d["year"], d["quarter"])] = r
        _cell(ws, r, 1, d["year"])
        _cell(ws, r, 2, d["quarter"])
        _cell(ws, r, 3, d["real_gdp_2015p"], INT)   # SOURCE datapoint
        prev = qrow.get((d["year"] - 1, d["quarter"]))
        if prev:
            _cell(ws, r, 4, f"=C{r}/C{prev}-1", PCT)
        r += 1
    r += 2

    ws.cell(row=r, column=1, value="Өссөн дүнгээрх өсөлт (Cumulative growth by period "
            "= SUM of quarters vs same quarters a year earlier)").font = F_BOLD
    r += 1
    r = _header_row(ws, r, ["Он (Year)"] + list(CUM_PERIODS.values()),
                    [10, 12, 12, 12, 12])
    years = sorted({y for (y, _q) in qrow})
    latest_cell = None
    show_years = [y for y in years
                  if str(y) in nso.get("cumulative_growth", {})] or years[-3:]
    for y in show_years:
        _cell(ws, r, 1, y, bold=True)
        for j, (q_end, _lbl) in enumerate(CUM_PERIODS.items(), start=2):
            need = [(y, qq) for qq in range(1, q_end + 1)]
            base = [(y - 1, qq) for qq in range(1, q_end + 1)]
            if all(k in qrow for k in need + base):
                c1, c2 = qrow[need[0]], qrow[need[-1]]
                p1, p2 = qrow[base[0]], qrow[base[-1]]
                _cell(ws, r, j,
                      f"=SUM(C{c1}:C{c2})/SUM(C{p1}:C{p2})-1", PCT)
                latest_cell = f"{get_column_letter(j)}{r}"
        r += 1

    mieg = series.get("mieg", {}) or {}
    if mieg.get("monthly"):
        r += 2
        ws.cell(row=r, column=1, value="Эдийн засгийн өсөлтийн сарын "
                "индикатор (MIEG, cumulative YoY, NSO DT_NSO_0500_001V5)"
                ).font = F_BOLD
        r += 1
        r = _header_row(ws, r, ["Сар (Month)", "Өсөлт (Growth, YoY cum.)"],
                        [12, 18])
        for i, d in enumerate(mieg["monthly"]):
            alt = i % 2 == 1
            _cell(ws, r, 1, d["period"], alt=alt)
            _cell(ws, r, 2, d["value"] / 100.0, PCT, alt=alt)
            r += 1
    if mieg.get("contributions"):
        r += 2
        ws.cell(row=r, column=1, value="Салбаруудын хувь нэмэр (MIEG sector "
                f"contributions, pp, {mieg.get('contributions_period', '')})"
                ).font = F_BOLD
        r += 1
        r = _header_row(ws, r, ["Салбар (Sector)", "Хувь нэмэр (pp)"],
                        [40, 14])
        first_c = r
        for i, (sname, v) in enumerate(mieg["contributions"].items()):
            alt = i % 2 == 1
            _cell(ws, r, 1, sname, alt=alt)
            _cell(ws, r, 2, v, NUM2, alt=alt)
            r += 1
        _cell(ws, r, 1, "Нийт (Total, pp)", bold=True)
        _cell(ws, r, 2, f"=SUM(B{first_c}:B{r - 1})", NUM2, bold=True)
    return {"gdp_growth_cell": latest_cell}


def sheet_sectors(wb, series):
    contrib = series.get("nso", {}).get("contributions", {})
    ws = wb.create_sheet("GDP_Sectors")
    r = _title(ws, "Sector contributions to real GDP growth (pp)",
               "Deck slide 5. Contribution = Δsector real value added / "
               "prior-period total GDP (computed upstream from NSO levels). "
               "Total row is a SUM formula and must equal GDP growth.")
    r = _stamp(ws, r, "NSO national accounts")
    years = sorted(contrib)
    if not years:
        _cell(ws, r + 1, 1, "No NSO contribution data — run the pipeline.")
        return
    hdr = ["Салбар (Sector)"] + [f"{y} ({contrib[y].get('period', '')})"
                                  for y in years]
    r = _header_row(ws, r, hdr, [46] + [16] * len(years))
    sectors = [s for s in contrib[years[-1]] if s != "period"]
    first = r
    for i, s in enumerate(sectors):
        alt = i % 2 == 1
        _cell(ws, r, 1, _sector_label(s), alt=alt)
        for j, y in enumerate(years, start=2):
            _cell(ws, r, j, contrib[y].get(s), NUM2, alt=alt)
        r += 1
    _cell(ws, r, 1, "Нийт (Total = GDP growth, pp)", bold=True)
    for j, _y in enumerate(years, start=2):
        col = get_column_letter(j)
        _cell(ws, r, j, f"=SUM({col}{first}:{col}{r - 1})", NUM2, bold=True)


def sheet_inflation(wb, k, series):
    ws = wb.create_sheet("Inflation")
    r = _title(ws, "Inflation", "Deck slides 6-7. Headline from the Mongolbank statistical bulletin; the "
               "12-month CPI series requires the NSO CPI table id (see "
               "raw_ingestor.py NSO_TABLES). Values stored as fractions, "
               "shown as % via number format.")
    r = _stamp(ws, r, "Mongolbank bulletin")
    r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Утга (Value)", "Note"],
                    [36, 12, 46])
    first = r
    infl = k.get("inflation")
    tgt = k.get("inflation_target")
    med_src = k.get("_bulletin_period_note") or (
        "Mongolbank bulletin "
        f"({k.get('_bulletin_source', 'not cached -- run ingestor')}, "
        f"period {k.get('_bulletin_period', 'n/a')})")
    rows = [
        ("Инфляц (Headline inflation, YoY)",
         infl / 100.0 if infl is not None else None, PCT, med_src),
        ("Монголбанкны зорилт (Mongolbank target)",
         tgt / 100.0 if tgt is not None else None, PCT, "official target"),
        ("Зорилтоос зөрүү (Gap vs target)",
         f"=B{first}-B{first + 1}" if infl is not None else None, PCT,
         "formula: headline - target; positive = above target"),
        ("Махны хувь нэмэр (Meat contribution)",
         (k.get("meat_contribution_pp") / 100.0
          if k.get("meat_contribution_pp") is not None else None), PCT,
         "pp of headline, when stated in the source"),
        ("Бодлогын хүү (Policy rate)",
         (k.get("policy_rate") / 100.0
          if k.get("policy_rate") is not None else None), PCT,
         k.get("_policy_rate_source", "")),
        ("Махны үнийн өсөлт (Meat price growth, YoY)", None, PCT,
         "MANUAL entry (slide 7) -- CPI PDF automation dropped by design; "
         "type the value from the NSO CPI presentation here"),
        ("Шатахууны үнийн өсөлт (Fuel price growth, YoY)", None, PCT,
         "MANUAL entry (slide 7) -- CPI PDF automation dropped by design; "
         "type the value from the NSO CPI presentation here"),
    ]
    for name, val, fmt, note in rows:
        _cell(ws, r, 1, name, bold=True)
        c = _cell(ws, r, 2, val, fmt)
        if val is None:
            if "MANUAL" in note:      # yellow = cell awaiting manual input
                c.fill = PatternFill("solid", start_color="FFFF00")
            else:
                c.value = "n/a"
                c.number_format = "General"
        _cell(ws, r, 3, note)
        r += 1

    # monthly series from the bulletin (deck slides 6 & 8)
    bser = series.get("bulletin_series", {}) or {}
    if bser.get("inflation_nat"):
        r += 2
        ws.cell(row=r, column=1, value="Сарын цуваа (Monthly series, "
                "Mongolbank bulletin) — slides 6 & 8").font = F_BOLD
        r += 1
        # Two lending rates, not one. Mongolbank publishes the weighted
        # average tugrik lending rate at market rates and again with
        # subsidised programme and project loans folded in; the second sits
        # below the first and the gap is how far the programmes pull the cost
        # of borrowing down. Both belong on the rate slide.
        r = _header_row(ws, r, ["Сар (Month)",
                                "Инфляц, улс (CPI YoY, national)",
                                "Инфляц, УБ (CPI YoY, UB)",
                                "Зорилт (Target)",
                                "Бодлогын хүү (Policy rate)",
                                "Хадгаламжийн хүү (Deposit rate, new MNT)",
                                "Шинэ зээлийн хүү, зах зээлийн "
                                "(New loan rate, market)",
                                "Шинэ зээлийн хүү, хөтөлбөрийн зээл "
                                "оруулснаар (incl. subsidised)",
                                "Зөрүү (Gap, pp)",
                                "Ам.долларын ханш (USD/MNT, eop)"],
                        [10, 15, 15, 10, 13, 15, 17, 20, 12, 14])
        first_s = r
        by_period = {}
        for key in ("inflation_nat", "inflation_ub", "policy_rate",
                    "deposit_rate", "loan_rate_new_mnt",
                    "loan_rate_new_mnt_incl", "usd_mnt_eop"):
            for p in bser.get(key, []):
                by_period.setdefault(p["period"], {})[key] = p["value"]
        tgt = (k.get("inflation_target") or 8.0) / 100.0
        for i, per in enumerate(sorted(by_period)):
            d = by_period[per]
            alt = i % 2 == 1
            _cell(ws, r, 1, per, alt=alt)
            for j, key in enumerate(("inflation_nat", "inflation_ub"),
                                    start=2):
                v = d.get(key)
                _cell(ws, r, j, v / 100.0 if v is not None else None,
                      PCT, alt=alt)
            _cell(ws, r, 4, tgt, PCT, alt=alt)
            for j, key in enumerate(("policy_rate", "deposit_rate",
                                     "loan_rate_new_mnt",
                                     "loan_rate_new_mnt_incl"), start=5):
                v = d.get(key)
                _cell(ws, r, j, v / 100.0 if v is not None else None,
                      PCT, alt=alt)
            _cell(ws, r, 9, f"=IFERROR((G{r}-H{r})*100,\"\")", NUM2, alt=alt)
            _cell(ws, r, 10, d.get("usd_mnt_eop"), NUM, alt=alt)
            r += 1
        last_s = r - 1

        # Slide 10's chart, drawn here so it does not have to be rebuilt:
        # inflation against the target and the three policy/market rates.
        if last_s > first_s + 2:
            ln = LineChart()
            for col in (2, 4, 5, 6, 7, 8):
                ln.add_data(Reference(ws, min_col=col, min_row=first_s,
                                      max_row=last_s), titles_from_data=True)
            ln.set_categories(Reference(ws, min_col=1, min_row=first_s + 1,
                                        max_row=last_s))
            for i, s in enumerate(ln.series):
                s.marker = Marker(symbol="none")
                s.smooth = False
                s.graphicalProperties.line = LineProperties(
                    solidFill=CHART_COLORS[i % len(CHART_COLORS)], w=20000)
            ln.title = "Инфляц ба хүүгийн түвшин (inflation and rates)"
            ln.style = None
            ln.y_axis.numFmt = "0.0%"
            ln.y_axis.majorGridlines = None
            ln.x_axis.majorGridlines = None
            ln.x_axis.tickLblSkip = ln.x_axis.tickMarkSkip = 2
            ln.legend.position = "b"
            ln.legend.overlay = False
            ln.height, ln.width = 9.5, 26
            ws.add_chart(ln, f"L{first_s}")

    cpi = series.get("bulletin_cpi", {}) or {}
    comps = cpi.get("components", [])
    if comps:
        r += 2
        ws.cell(row=r, column=1, value="Инфляцад нөлөөлж буй бүлгүүд (CPI "
                "component contributions to headline, "
                f"{cpi.get('period', '')})").font = F_BOLD
        r += 1
        # This year beside last year, which is how the deck's table is laid
        # out. The prior year comes from the same bulletin sheet, matched by
        # category NAME so the two columns stay aligned even if the sheet
        # reorders itself between releases.
        prev = {d["name"]: d for d in (cpi.get("prev") or [])}
        pp_lbl = cpi.get("period", "")
        pv_lbl = cpi.get("prev_period", "")
        r = _header_row(ws, r, ["Бүлэг (Category)",
                                f"Хувь нэмэр, {pp_lbl} (pp)",
                                f"Эзлэх хувь, {pp_lbl} (share)",
                                f"Хувь нэмэр, {pv_lbl or 'n/a'} (pp)",
                                f"Эзлэх хувь, {pv_lbl or 'n/a'} (share)",
                                "Бүлгийн өсөлт (Category YoY)",
                                "Индекс (Level)"],
                        [46, 17, 17, 17, 17, 16, 12])
        first_c = r
        tot_now = sum(d["pp"] for d in comps) or 1
        tot_prv = sum(d["pp"] for d in prev.values()) or None
        for i, d in enumerate(comps):
            alt = i % 2 == 1
            pv = prev.get(d["name"])
            _cell(ws, r, 1, d["name"], alt=alt)
            _cell(ws, r, 2, d["pp"], NUM2, alt=alt)
            _cell(ws, r, 3, d["pp"] / tot_now, PCT, alt=alt)
            _cell(ws, r, 4, pv["pp"] if pv else None, NUM2, alt=alt)
            _cell(ws, r, 5, (pv["pp"] / tot_prv) if (pv and tot_prv) else None,
                  PCT, alt=alt)
            _cell(ws, r, 6, d["yoy_pct"] / 100.0, PCT, alt=alt)
            _cell(ws, r, 7, d["level"], NUM, alt=alt)
            r += 1
        _cell(ws, r, 1, "Улсын инфляц (Total, pp — must match headline)",
              bold=True)
        _cell(ws, r, 2, f"=SUM(B{first_c}:B{r - 1})", NUM2, bold=True)
        _cell(ws, r, 3, f"=SUM(C{first_c}:C{r - 1})", PCT, bold=True)
        _cell(ws, r, 4, f"=SUM(D{first_c}:D{r - 1})", NUM2, bold=True)
        _cell(ws, r, 5, f"=SUM(E{first_c}:E{r - 1})", PCT, bold=True)
        r += 1
        if not prev:
            _cell(ws, r, 1, "Өмнөх оны задаргаа алга — бюллетень тухайн "
                            "сарыг хамрахгүй байна (no prior-year breakdown: "
                            "the bulletin does not reach that month).")
            r += 1


def sheet_trade(wb, k, series):
    """
    Source datapoints: cumulative export/import for prev & current year
    (Customs sheet 1) and the monthly history. Everything else — YoY, trade
    balance, balance multiple, mln-USD conversions, monthly balances — is a
    formula. Returns cell refs for the Summary sheet.
    """
    ws = wb.create_sheet("Trade")
    r = _title(ws, "Foreign trade", "Deck slide 9. Source: Customs gaali.mn "
               "monthly bulletin, sheet 1. Money: thousand USD unless noted. "
               "YoY/balance columns are formulas.")
    r = _stamp(ws, r, "Customs (gaali.mn)")
    r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Өмнөх он (Prev year, cum.)",
                            "Тайлант он (Current year, cum.)",
                            "Өсөлт (YoY)", "сая $ (mln USD, cur.)"],
                    [24, 18, 18, 10, 14])
    exp_row, imp_row = r, r + 1
    for name, prev, cur, row in (
            ("Экспорт (Export)", k.get("total_export_prev"), k.get("total_export"),
             exp_row),
            ("Импорт (Import)", k.get("total_import_prev"), k.get("total_import"),
             imp_row)):
        _cell(ws, row, 1, name, bold=True)
        _cell(ws, row, 2, prev, NUM)          # SOURCE datapoint
        _cell(ws, row, 3, cur, NUM)           # SOURCE datapoint
        _cell(ws, row, 4, f"=C{row}/B{row}-1", PCT)
        _cell(ws, row, 5, f"=C{row}/1000", NUM)
    bal_row = imp_row + 1
    _cell(ws, bal_row, 1, "Худалдааны тэнцэл (Trade balance)", bold=True)
    _cell(ws, bal_row, 2, f"=B{exp_row}-B{imp_row}", NUM)
    _cell(ws, bal_row, 3, f"=C{exp_row}-C{imp_row}", NUM)
    _cell(ws, bal_row, 4, f"=C{bal_row}/B{bal_row}", MULT)
    _cell(ws, bal_row, 5, f"=C{bal_row}/1000", NUM)
    r = bal_row + 3

    ws.cell(row=r, column=1,
            value="Сар бүрийн өссөн худалдаа (Monthly cumulative trade, "
                  "mln USD) — grows as monthly workbooks are cached").font = \
        F_BOLD
    r += 1
    r = _header_row(ws, r, ["Он (Year)", "Сар (Month)",
                            "Экспорт, өссөн (Export cum., mln $)",
                            "Импорт, өссөн (Import cum., mln $)",
                            "Тэнцэл (Balance, mln $)",
                            "Экспортын өсөлт (Export YoY)",
                            "Импортын өсөлт (Import YoY)",
                            "Нийт эргэлт (Turnover, mln $)",
                            "Эргэлтийн өсөлт (Turnover YoY)"],
                    [10, 10, 18, 18, 16, 12, 12, 18, 14])
    # YoY formulas are keyed on a (year, month) row map, so every newly
    # appended month automatically gets its own YoY calculation.
    rowmap = {}
    for d in series.get("customs", {}).get("trade_monthly", []):
        rowmap[(d["year"], d["month"])] = r
        _cell(ws, r, 1, d["year"])
        _cell(ws, r, 2, d["month"])
        e = d.get("export_cum_kusd")
        i = d.get("import_cum_kusd")
        _cell(ws, r, 3, e / 1000.0 if e else None, NUM)   # SOURCE datapoint
        _cell(ws, r, 4, i / 1000.0 if i else None, NUM)   # SOURCE datapoint
        _cell(ws, r, 5, f"=C{r}-D{r}", NUM)
        # Turnover — export + import. The deck plots it and it is not in any
        # source as its own line, so a previous deck left that chart empty.
        _cell(ws, r, 8, f"=C{r}+D{r}", NUM)
        prev = rowmap.get((d["year"] - 1, d["month"]))
        if prev:
            _cell(ws, r, 6, f"=C{r}/C{prev}-1", PCT)
            _cell(ws, r, 7, f"=D{r}/D{prev}-1", PCT)
            _cell(ws, r, 9, f"=H{r}/H{prev}-1", PCT)
        r += 1
    return {"exp_row": exp_row, "imp_row": imp_row, "bal_row": bal_row}


def sheet_commodities(wb, series):
    """
    Source datapoints: prev/cur quantities and amounts per commodity.
    YoY changes, mln-USD and deck-unit volume conversions are formulas.
    Returns row map for Summary volume links.
    """
    ce = series.get("customs", {}).get("commodity_exports", {})
    ci = series.get("customs", {}).get("commodity_imports", {})
    ws = wb.create_sheet("Commodities")
    r = _title(ws, "Commodity exports & imports",
               "Deck slides 9-10. Source: Customs sheet 8 / 8.2 / 3, matched "
               "by HS code. Quantities in kg; amounts in thousand USD; every "
               "derived column is a formula.")
    r = _stamp(ws, r, "Customs (gaali.mn)")
    r = _header_row(ws, r, ["Экспортын бараа (Export commodity)", "HS",
                            "Тоо, өмнөх (Qty prev, kg)",
                            "Тоо, тайлант (Qty cur, kg)",
                            "Тооны өсөлт (Qty YoY)",
                            "Дүн, өмнөх (Amt prev, k$)",
                            "Дүн, тайлант (Amt cur, k$)",
                            "Дүнгийн өсөлт (Amt YoY)", "сая $ (mln $)",
                            "Хэмжээ (Volume, deck unit)", "Нэгж (Unit)"],
                    [26, 8, 16, 16, 10, 15, 15, 10, 14, 15, 10])
    order = ["coal", "copper", "gold", "iron", "wool_cashmere"]
    conv = {"coal": ("=D{r}/1E9", "mln t"),
            "copper": ("=D{r}/1E6", "thous. t"),
            "gold": ("=D{r}", "kg"),
            "iron": ("=D{r}/1E9", "mln t")}
    exp_rows = {}
    for i, key in enumerate([k_ for k_ in order if k_ in ce]):
        d = ce[key]
        alt = i % 2 == 1
        exp_rows[key] = r
        _cell(ws, r, 1, _commodity_label(key, d), bold=True, alt=alt)
        _cell(ws, r, 2, d.get("hs_code"), alt=alt)
        _cell(ws, r, 3, d.get("qty_prev"), INT, alt=alt)   # SOURCE
        _cell(ws, r, 4, d.get("qty_cur"), INT, alt=alt)    # SOURCE
        _cell(ws, r, 5, f"=D{r}/C{r}-1" if d.get("qty_prev") else None,
              PCT, alt=alt)
        _cell(ws, r, 6, d.get("amt_prev"), NUM, alt=alt)   # SOURCE
        _cell(ws, r, 7, d.get("amt_cur"), NUM, alt=alt)    # SOURCE
        _cell(ws, r, 8, f"=G{r}/F{r}-1" if d.get("amt_prev") else None,
              PCT, alt=alt)
        _cell(ws, r, 9, f"=G{r}/1000", NUM, alt=alt)
        if key in conv and d.get("qty_cur") is not None:
            f, unit = conv[key]
            _cell(ws, r, 10, f.format(r=r), NUM2, alt=alt)
            _cell(ws, r, 11, unit, alt=alt)
        r += 1
    r += 2

    ws.cell(row=r, column=1, value="Импортын бараа (Import commodities)").font = F_BOLD
    r += 1
    r = _header_row(ws, r, ["Импортын бараа (Import commodity)", "HS/section",
                            "Дүн, өмнөх (Amt prev, k$)",
                            "Дүн, тайлант (Amt cur, k$)",
                            "Дүнгийн өсөлт (Amt YoY)", "сая $ (mln $)"],
                    [26, 12, 16, 16, 10, 14])
    iorder = ["petroleum", "cars", "electricity", "trucks", "machinery",
              "base_metals", "food"]
    for i, key in enumerate([k_ for k_ in iorder if k_ in ci]):
        d = ci[key]
        alt = i % 2 == 1
        _cell(ws, r, 1, _commodity_label(key, d), bold=True, alt=alt)
        _cell(ws, r, 2, d.get("hs_code"), alt=alt)
        _cell(ws, r, 3, d.get("amt_prev"), NUM, alt=alt)   # SOURCE
        _cell(ws, r, 4, d.get("amt_cur"), NUM, alt=alt)    # SOURCE
        _cell(ws, r, 5, f"=D{r}/C{r}-1" if d.get("amt_prev") else None,
              PCT, alt=alt)
        _cell(ws, r, 6, f"=D{r}/1000", NUM, alt=alt)
        r += 1
    return {"exp_rows": exp_rows}


def sheet_trade_sections(wb, series):
    """
    Export and import as two INDEPENDENT tables.

    Mongolia's export and import structures barely overlap -- exports are
    dominated by mineral products while imports spread across machinery,
    vehicles and fuel -- so a shared row order forces one side into an order
    that means nothing for it. Each block is therefore ranked on its own
    values, largest first, with its own YoY, share and total.
    """
    secs = series.get("customs", {}).get("trade_sections", [])
    ws = wb.create_sheet("Trade_Sections")
    r = _title(ws, "Худалдаа, хэсгээр (Trade by HS section)",
               "Full sectional detail behind slide 9. Source: Customs "
               "sheet 3, thousand USD, cumulative both years. Export and "
               "import are separate tables, each ranked by its own current-"
               "year value; YoY, share and totals are formulas.")
    r = _stamp(ws, r, "Customs (gaali.mn)")
    if not secs:
        _cell(ws, r, 1, "No sectional data -- run the pipeline.")
        return

    def block(r, title, prev_key, cur_key, note):
        rows = [d for d in secs if d.get(cur_key) is not None
                or d.get(prev_key) is not None]
        rows.sort(key=lambda d: -(d.get(cur_key) or 0))
        _cell(ws, r, 1, title, bold=True)
        r += 1
        _cell(ws, r, 1, note)
        r += 1
        r = _header_row(ws, r, ["Эрэмбэ (Rank)", "Хэсэг (HS section)",
                                "Өмнөх он (Previous, k$)",
                                "Тайлант он (Current, k$)",
                                "Өсөлт (YoY)", "Хувь (Share)"],
                        [9, 52, 17, 17, 11, 11])
        first = r
        last = r + len(rows) - 1
        tot = last + 1
        for i, d in enumerate(rows):
            alt = i % 2 == 1
            label = d["mn"] + (f" ({d['en']})" if d.get("en") else "")
            _cell(ws, r, 1, i + 1, alt=alt)
            _cell(ws, r, 2, label, alt=alt)
            _cell(ws, r, 3, d.get(prev_key), NUM, alt=alt)
            _cell(ws, r, 4, d.get(cur_key), NUM, alt=alt)
            _cell(ws, r, 5, f"=D{r}/C{r}-1" if d.get(prev_key) else None,
                  PCT, alt=alt)
            _cell(ws, r, 6, f"=D{r}/D${tot}" if d.get(cur_key) else None,
                  PCT, alt=alt)
            r += 1
        _cell(ws, tot, 2, "Нийт (Total)", bold=True)
        _cell(ws, tot, 3, f"=SUM(C{first}:C{last})", NUM, bold=True)
        _cell(ws, tot, 4, f"=SUM(D{first}:D{last})", NUM, bold=True)
        _cell(ws, tot, 5, f"=D{tot}/C{tot}-1", PCT, bold=True)
        _cell(ws, tot, 6, f"=SUM(F{first}:F{last})", PCT, bold=True)
        return tot + 2

    r = block(r, "ЭКСПОРТ, хэсгээр (EXPORT by HS section)",
              "exp_prev", "exp_cur",
              "Ranked by current-year export value. Total must equal the "
              "headline export figure on the Trade sheet.")
    block(r, "ИМПОРТ, хэсгээр (IMPORT by HS section)",
          "imp_prev", "imp_cur",
          "Ranked independently by current-year import value -- the order "
          "differs from the export table by design.")


def sheet_border_prices(wb, series):
    bp = series.get("customs", {}).get("border_prices", {})
    ws = wb.create_sheet("Border_Prices")
    r = _title(ws, "Border prices (export unit values)",
               "Deck slide 11. Unit value = Δcumulative amount / Δcumulative "
               "quantity per month, from Customs sheet 8. Series fills in "
               "as monthly workbooks are cached by the ingestor.")
    r = _stamp(ws, r, "Customs (gaali.mn)")
    r = _header_row(ws, r, ["Бараа (Commodity)", "Он (Year)", "Сар (Month)",
                            "Үнэ (Price)", "Нэгж (Unit)", "Өсөлт (YoY)",
                            "Basis"], [22, 8, 8, 12, 10, 10, 30])
    labels = {"coal": "Нүүрс (Coal)",
              "copper": "Зэсийн баяжмал (Copper concentrate)",
              "gold": "Алт (Gold)", "iron": "Төмрийн хүдэр (Iron ore)"}
    i = 0
    for key in ["coal", "copper", "gold", "iron"]:
        rowmap = {}          # (year, month) -> row, per commodity: the YoY
        for d in bp.get(key, []):   # formula extends to every new month
            alt = i % 2 == 1
            rowmap[(d["year"], d["month"])] = r
            _cell(ws, r, 1, labels.get(key, key), alt=alt)
            _cell(ws, r, 2, d["year"], alt=alt)
            _cell(ws, r, 3, d["month"], alt=alt)
            _cell(ws, r, 4, d["price"], NUM, alt=alt)
            _cell(ws, r, 5, d["unit"], alt=alt)
            prev = rowmap.get((d["year"] - 1, d["month"]))
            _cell(ws, r, 6, f"=D{r}/D{prev}-1" if prev else None, PCT,
                  alt=alt)
            _cell(ws, r, 7, d["basis"], alt=alt)
            r += 1
            i += 1


def sheet_external(wb, series):
    cards = series.get("mongolbank_cards", [])
    ws = wb.create_sheet("External_Sector")
    r = _title(ws, "Mongolbank statistics cards",
               "Deck slide 12 (BoP, reserves, external debt, FDI) and slide "
               "16 inputs where exposed. Every card returned by the section "
               "sweep is listed.")
    r = _stamp(ws, r, "Mongolbank cards", "Balance of payments")
    r = _header_row(ws, r, ["Карт (Card)", "Утга (Reported value, raw text)",
                            "Тоон утга (Parsed value)", "Огноо (As of)",
                            "Section file"],
                    [32, 52, 14, 18, 26])
    for i, c in enumerate(cards):
        alt = i % 2 == 1
        _cell(ws, r, 1, c["name"], bold=True, alt=alt)
        _cell(ws, r, 2, c["data"], alt=alt)
        _cell(ws, r, 3, c.get("value"), NUM, alt=alt)
        _cell(ws, r, 4, c.get("date"), alt=alt)
        _cell(ws, r, 5, c.get("section_file"), alt=alt)
        r += 1
    if not cards:
        _cell(ws, r, 1, "No Mongolbank cards cached — run raw_ingestor.py")
        r += 1

    # Balance of payments accounts (slide 12 chart & detail)
    r += 2
    ws.cell(row=r, column=1, value="Balance of payments — annual sums "
            "(mln USD; latest year is YTD). Source: bop_manual.xlsx").font = \
        F_BOLD
    r += 1
    bop = series.get("bop", {}) or {}
    annual = bop.get("annual", {})
    order = ["bop_current_account", "bop_capital_account",
             "bop_financial_account", "bop_errors_omissions",
             "bop_overall_balance", "bop_reserve_assets"]
    labels = {
        "bop_current_account": "Урсгал данс (Current account)",
        "bop_capital_account": "Хөрөнгийн данс (Capital account)",
        "bop_financial_account": "Санхүүгийн данс (Financial account)",
        "bop_errors_omissions": "Алдаа болон орхигдуулга (Errors & omissions)",
        "bop_overall_balance": "Төлбөрийн тэнцлийн нийт дүн (Overall balance)",
        "bop_reserve_assets": "Нөөц хөрөнгө (Reserve assets)",
    }
    if annual:
        years = sorted({y for d in annual.values() for y in d})
        r = _header_row(ws, r, ["Данс (Account)"] + years, [40] + [12] * len(years))
        for i, key in enumerate([k_ for k_ in order if k_ in annual]):
            alt = i % 2 == 1
            _cell(ws, r, 1, labels.get(key, key), bold=True, alt=alt)
            for j, y in enumerate(years, start=2):
                _cell(ws, r, j, annual[key].get(y), NUM, alt=alt)
            r += 1
    else:
        r = _header_row(ws, r, ["Данс (Account)", "Утга (Value)", "Note"],
                        [40, 12, 40])
        for key in order[:3]:
            _cell(ws, r, 1, labels[key], bold=True)
            _cell(ws, r, 2, "n/a")
            _cell(ws, r, 3, "provide raw_files/bop_manual.xlsx (the ingestor "
                            "prompts for it during each run)")
            r += 1

    # Section-by-section breakdown of each BoP account (the '4 бүлгийг
    # задалж харуулах' request): every section with its own components,
    # cumulative for the target month against the same months last year.
    detail = bop.get("detail", [])
    per = bop.get("detail_period", {})
    if detail:
        r += 2
        ws.cell(row=r, column=1,
                value="Төлбөрийн тэнцэл, бүлгээр (Balance of payments by "
                      "section and component, mln USD)").font = F_BOLD
        r += 1
        _cell(ws, r, 1, "Each section is followed by its own components. "
                        "Indented rows sit one level deeper and are already "
                        "included in the line above them, so only the "
                        "section rows should be added together.")
        r += 1
        cur_lbl = per.get("cur", "current")
        prev_lbl = per.get("prev", "previous")
        r = _header_row(ws, r, ["Бүлэг / бүрэлдэхүүн (Section / component)",
                                f"{prev_lbl} (mln USD)",
                                f"{cur_lbl} (mln USD)",
                                "Өөрчлөлт (Change)"],
                        [58, 18, 18, 16])
        for i, d in enumerate(detail):
            lvl = d.get("level", 0)
            indent = {0: "", 1: "    ", 2: "        "}.get(lvl, "")
            alt = i % 2 == 1
            _cell(ws, r, 1, indent + d.get("label", ""),
                  bold=(lvl == 0), alt=alt)
            _cell(ws, r, 2, d.get("prev"), NUM, alt=alt)
            _cell(ws, r, 3, d.get("cur"), NUM, alt=alt)
            _cell(ws, r, 4,
                  f"=C{r}-B{r}" if d.get("cur") is not None
                  and d.get("prev") is not None else None,
                  NUM, bold=(lvl == 0), alt=alt)
            r += 1


def sheet_banking(wb, k, series):
    """
    Slide 16: consolidated bank balance sheet. Levels are source datapoints
    (million MNT); change-since-first-column and trillion-MNT conversions
    are formulas. Below the KPI block, the full monthly series per metric.
    Returns cell refs for the Summary sheet.
    """
    bs = series.get("banking_balance_sheet", {})
    ws = wb.create_sheet("Banking")
    basis = k.get("_bs_change_basis", "")
    r = _title(ws, "Banking sector — consolidated bank balance sheet",
               "Deck slide 16. Source: Mongolbank 'Банкуудын нэгдсэн тайлан "
               f"тэнцэл' xlsx (million MNT). {basis}")
    r = _stamp(ws, r, "Bank balance sheet", "Deposit tree")
    rows_meta = [
        ("Нийт актив (Total assets)", "bs_total_assets"),
        ("Дотоодын нийт зээл (Total loans, domestic)", "bs_total_loans"),
        ("Төв банкны үнэт цаас (Central bank securities)", "bs_cb_securities"),
        ("Харилцах, хадгаламж (Current accounts + deposits)", "bs_deposits"),
        ("Анхаарал хандуулах зээл (Past-due / special mention)", "bs_past_due_loans"),
        ("Чанаргүй зээл (Non-performing loans)", "bs_npl"),
        ("Өөрийн хөрөнгө (Equity)", "bs_equity"),
    ]
    cols = bs.get("columns", [])
    first_lbl = k.get("_bs_first_column", "first")
    last_lbl = k.get("_bs_last_column", "latest")
    r = _header_row(ws, r, ["Үзүүлэлт (Indicator)",
                            f"Түвшин {first_lbl} (Level, mln ₮)",
                            f"Түвшин {last_lbl} (Level, mln ₮)",
                            "Өөрчлөлт (Change since first col)",
                            "их наяд ₮ (Latest, tln ₮)"],
                    [32, 20, 20, 14, 12])
    refs = {}
    wrote_any = False
    for i, (label, key) in enumerate(rows_meta):
        first = k.get(f"{key}_first_mln_mnt")
        last = k.get(f"{key}_mln_mnt")
        alt = i % 2 == 1
        _cell(ws, r, 1, label, bold=True, alt=alt)
        _cell(ws, r, 2, first, NUM, alt=alt)            # SOURCE datapoint
        _cell(ws, r, 3, last, NUM, alt=alt)             # SOURCE datapoint
        _cell(ws, r, 4, f"=C{r}/B{r}-1" if first and last is not None
              else None, PCT, alt=alt)
        _cell(ws, r, 5, f"=C{r}/1E6" if last is not None else None,
              NUM2, alt=alt)
        refs[key] = r
        wrote_any = wrote_any or last is not None
        r += 1
    if not wrote_any:
        _cell(ws, r, 1, "Balance-sheet xlsx not parsed -- run "
                        "raw_ingestor.py (report id 5098).")
        return refs

    r += 2
    ws.cell(row=r, column=1,
            value="Сар бүрийн түвшин (Monthly levels, mln ₮) per balance-sheet metric").font = \
        F_BOLD
    r += 1
    r = _header_row(ws, r, ["Үзүүлэлт (Metric)"] + cols,
                    [32] + [16] * len(cols))
    for i, (label, key) in enumerate(rows_meta):
        vals = (bs.get("rows", {}).get(key) or {}).get("values")
        if not vals:
            continue
        alt = i % 2 == 1
        _cell(ws, r, 1, label, bold=True, alt=alt)
        for j, v in enumerate(vals, start=2):
            _cell(ws, r, j, v, NUM, alt=alt)
        r += 1

    # DEPOSITS — the liability tree, as on the portal's balance-sheet page
    tree = bs.get("deposit_tree", [])
    if tree:
        dcols = bs.get("deposit_columns") or cols
        r += 2
        ws.cell(row=r, column=1,
                value="ХАРИЛЦАХ, ХАДГАЛАМЖ (deposits and current accounts, "
                      "mln ₮) — Нийт пассив breakdown").font = F_BOLD
        r += 1
        _cell(ws, r, 1, "Each block splits by currency, then by tenor for "
                        "savings, then by holder. Indented rows are already "
                        "included in the line above them, so only the two "
                        "top-level blocks add to total deposits. Source: "
                        + str(k.get("_deposit_source", "report")))
        r += 1
        last_col = dcols[-1] if dcols else ""
        prev_col = dcols[-2] if len(dcols) > 1 else ""
        r = _header_row(ws, r,
                        ["Үзүүлэлт (Item)"] + dcols
                        + [f"Өөрчлөлт ({prev_col}→{last_col})", "Эзлэх (Share)"],
                        [46] + [14] * len(dcols) + [18, 11])
        ncol = len(dcols)
        top_rows = []
        first = r
        for i, n in enumerate(tree):
            lvl = n.get("level", 0)
            indent = {0: "", 1: "    ", 2: "        ", 3: "            "}.get(
                lvl, "")
            alt = i % 2 == 1
            _cell(ws, r, 1, indent + n.get("label", ""), bold=(lvl == 0),
                  alt=alt)
            for j, v in enumerate(n.get("values", []), start=2):
                _cell(ws, r, j, v, NUM, alt=alt)
            cl = get_column_letter(1 + ncol)          # last data column
            pl = get_column_letter(ncol)              # previous data column
            if ncol > 1:
                _cell(ws, r, 2 + ncol, f"={cl}{r}-{pl}{r}", NUM,
                      bold=(lvl == 0), alt=alt)
            if lvl == 0:
                top_rows.append(r)
            r += 1
        tot = r
        cl = get_column_letter(1 + ncol)
        _cell(ws, tot, 1, "Нийт харилцах, хадгаламж (Total deposits)",
              bold=True)
        for j in range(2, 2 + ncol):
            col = get_column_letter(j)
            _cell(ws, tot, j,
                  "=" + "+".join(f"{col}{tr}" for tr in top_rows), NUM,
                  bold=True)
        # share of the total, for the two blocks and every row beneath them
        for rr in range(first, tot):
            _cell(ws, rr, 3 + ncol, f"={cl}{rr}/{cl}${tot}", PCT)
    return refs


def sheet_budget(wb, k, commod_refs):
    ws = wb.create_sheet("Budget")
    r = _title(ws, "State budget", "Deck slide 13. Source: Mongolbank statistical bulletin. "
               "Coal actual links to the Commodities sheet.")
    r = _stamp(ws, r, "Mongolbank bulletin")
    r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Утга (Value)", "Нэгж (Unit)",
                    "Note"], [34, 14, 12, 40])
    coal_row = commod_refs.get("exp_rows", {}).get("coal")
    med_src = ("Mongolbank bulletin "
               f"({k.get('_bulletin_source', 'not cached -- run ingestor')})")
    rows = [
        ("Төсвийн тэнцэл (Budget balance, cumulative)", k.get("budget_balance_bln_mnt"),
         NUM, "bln MNT", med_src, False),
        ("Төсвийн орлого, гүйцэтгэл (Revenue collected, cumulative)",
         k.get("budget_revenue_collected_bln"), NUM, "bln MNT", med_src,
         False),
        ("Төсвийн зарлага, гүйцэтгэл (Expenditure, cumulative)",
         k.get("budget_expenditure_bln"), NUM, "bln MNT", med_src, False),
        ("Төсвийн орлого, төлөвлөгөө (Revenue plan, annual)", k.get("budget_revenue_plan_tln"), NUM,
         "tln MNT", "MANUAL: annual plan figure, set once a year", False),
        ("Нүүрсний экспортын төлөвлөгөө (Coal export plan, annual)", k.get("coal_export_plan_mt"), NUM,
         "mln t", "MANUAL: annual plan figure, set once a year", False),
        ("Нүүрсний экспорт, гүйцэтгэл (Coal export actual)",
         f"=Commodities!D{coal_row}/1E9" if coal_row else None, NUM2,
         "mln t", k.get("report_date", ""), True),
    ]
    for name, val, fmt, unit, note, link in rows:
        _cell(ws, r, 1, name, bold=True)
        c = _cell(ws, r, 2, val, fmt, link=link)
        if val is None:
            if "төлөвлөгөө" in name:   # annual plan: manual yellow input
                c.fill = PatternFill("solid", start_color="FFFF00")
            else:
                c.value = "n/a"
                c.number_format = "General"
        _cell(ws, r, 3, unit)
        _cell(ws, r, 4, note)
        r += 1


def sheet_loans(wb, k, series):
    """
    Bank loan detail: how the loan book grew and what drove it.

    The source is a monthly FLOW report, so the growth is decomposed exactly
    (opening + disbursed - repaid +/- revaluation = closing) rather than
    inferred from two stock readings. Sections: growth bridge, economic
    sectors ranked by contribution, loan quality with the NPL ratio,
    borrower types, weighted average rates, and the monthly series.
    """
    d = series.get("bank_loans", {}) or {}
    ws = wb.create_sheet("Loan_Detail")
    per = k.get("_loan_period", "n/a")
    r = _title(ws, "Банкны салбарын зээл, дэлгэрэнгүй (Bank loans in detail)",
               f"Mongolbank monthly loan report, {per} (million MNT). The "
               "report is a flow statement, so the change in the loan book "
               "is decomposed rather than estimated. Growth, shares and "
               "totals are formulas.")
    r = _stamp(ws, r, "Bank loan report")
    note = k.get("_loan_period_note")
    if note:
        _cell(ws, r, 1, note, bold=True)
        r += 1
    if not d.get("bridge"):
        _cell(ws, r, 1, "No loan report cached — run raw_ingestor.py.")
        return

    # 0. THE PORTAL'S OWN INDICATORS
    # Named exactly as on the stat.mongolbank.mn 'Банкны салбарын зээл' page
    # so the workbook can be checked line for line against it.
    inds = d.get("indicators", [])
    if inds:
        _cell(ws, r, 1, "0. ҮЗҮҮЛЭЛТҮҮД (indicators as published on "
                        "stat.mongolbank.mn — Банкны салбарын зээл)",
              bold=True)
        r += 1
        _cell(ws, r, 1, "Same figures as the portal's indicator list. Per "
                        "the report's own Заавар sheet, 'хугацаа хэтэрсэн' is "
                        "the special-mention class and 'чанаргүй' is "
                        "substandard + doubtful + loss.")
        r += 1
        r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Нийт (Total)",
                                "Үүнээс: төгрөгийн зээл (of which MNT)",
                                "Валютын хувь (FX share)",
                                "Нийт зээлд эзлэх (Share of book)"],
                        [52, 18, 26, 14, 16])
        for i, ind in enumerate(inds):
            alt = i % 2 == 1
            _cell(ws, r, 1, ind.get("label"), bold=True, alt=alt)
            _cell(ws, r, 2, ind.get("total"), NUM, alt=alt)
            _cell(ws, r, 3, ind.get("mnt"), NUM, alt=alt)
            _cell(ws, r, 4, f"=1-C{r}/B{r}" if ind.get("mnt") else None,
                  PCT, alt=alt)
            _cell(ws, r, 5, f"=B{r}/B${r - i}" if ind.get("total") else None,
                  PCT, alt=alt)
            r += 1
        r += 1

    # 1. GROWTH BRIDGE
    b = d["bridge"]
    _cell(ws, r, 1, "1. ЗЭЭЛИЙН ӨСӨЛТИЙН ЗАДАРГАА (how the loan book grew)",
          bold=True)
    r += 1
    r = _header_row(ws, r, ["Үзүүлэлт (Flow)", "Дүн (mln MNT)"], [52, 20])
    first = r
    steps = [("Эхний үлдэгдэл (Opening balance)", b.get("opening")),
             ("+ Олгосон зээл (Loans disbursed)", b.get("disbursed")),
             ("− Төлөгдсөн зээл (Loans repaid)",
              -(b.get("repaid") or 0) if b.get("repaid") is not None else None),
             ("− Сангаас хаагдсан (Written off)",
              -(b.get("writeoff") or 0) if b.get("writeoff") is not None
              else None),
             ("± Ханшийн тэгшитгэл (FX revaluation)", b.get("fx")),
             ("± Бусад гүйлгээ (Other movements)", b.get("other"))]
    # If the report carries a flow this bridge does not model, show the gap
    # as its own line rather than leaving the column not adding up. The
    # closing balance below is the REPORTED one either way, so no figure is
    # invented -- the difference is named and visible.
    resid = b.get("residual")
    if resid is not None and abs(resid) > 1.0:
        steps.append(("± Тайлбарлагдаагүй хөдөлгөөн (unexplained movement — "
                      "a flow this bridge does not model)", -resid))
    for i, (lbl, v) in enumerate(steps):
        _cell(ws, r, 1, lbl, alt=i % 2 == 1)
        _cell(ws, r, 2, v, NUM, alt=i % 2 == 1)
        r += 1
    last = r - 1
    # no leading '=' in the caption: Excel would read it as a formula
    _cell(ws, r, 1, "Эцсийн үлдэгдэл (Closing balance, computed)",
          bold=True)
    _cell(ws, r, 2, f"=SUM(B{first}:B{last})", NUM, bold=True)
    computed = r
    r += 1
    _cell(ws, r, 1, "Тайлант эцсийн үлдэгдэл (Closing, as reported)")
    _cell(ws, r, 2, b.get("closing"), NUM)
    reported = r
    r += 1
    _cell(ws, r, 1, "Зөрүү (Residual — must be 0)")
    _cell(ws, r, 2, f"=B{computed}-B{reported}", NUM)
    r += 1
    _cell(ws, r, 1, "Өөрчлөлт (Net change)", bold=True)
    _cell(ws, r, 2, f"=B{reported}-B{first}", NUM, bold=True)
    r += 1
    _cell(ws, r, 1, "Өсөлт, сараар (Growth, month on month)", bold=True)
    _cell(ws, r, 2, f"=B{reported}/B{first}-1", PCT, bold=True)
    r += 1
    if k.get("loan_growth_yoy_pct") is not None:
        base = (d.get("yoy_base") or {})
        _cell(ws, r, 1, f"Өсөлт, жилээр (Growth YoY vs "
                        f"{base.get('period', '')})", bold=True)
        _cell(ws, r, 2, pct_static(k["loan_growth_yoy_pct"]), PCT, bold=True)
        r += 1
    r += 1

    # 2. SECTORS
    secs = d.get("sectors", [])
    if secs:
        _cell(ws, r, 1, "2. САЛБАРААР (by economic sector — what drove it)",
              bold=True)
        r += 1
        _cell(ws, r, 1, "Ranked by contribution to the month's growth. The "
                        "pp column sums to the total growth rate above.")
        r += 1
        cols = ["Код", "Салбар (Sector)", "Эхний үлдэгдэл (Opening)",
                "Эцсийн үлдэгдэл (Closing)", "Өөрчлөлт (Change)",
                "Хувь нэмэр (pp)", "Эзлэх (Share)",
                "Үүнээс: төгрөгийн (of which MNT)"]
        widths = [7, 50, 17, 17, 15, 12, 11, 22]
        has_yoy = any(s.get("yoy_pct") is not None for s in secs)
        if has_yoy:
            cols.append("Жилийн өсөлт (YoY)")
            widths.append(13)
        r = _header_row(ws, r, cols, widths)
        sfirst = r
        slast = r + len(secs) - 1
        stot = slast + 1
        for i, s in enumerate(secs):
            alt = i % 2 == 1
            _cell(ws, r, 1, s.get("code"), alt=alt)
            _cell(ws, r, 2, s.get("label"), alt=alt)
            _cell(ws, r, 3, s.get("opening"), NUM, alt=alt)
            _cell(ws, r, 4, s.get("closing"), NUM, alt=alt)
            _cell(ws, r, 5, f"=D{r}-C{r}", NUM, alt=alt)
            _cell(ws, r, 6, f"=(D{r}-C{r})/C${stot}*100"
                  if s.get("opening") is not None else None, NUM2, alt=alt)
            _cell(ws, r, 7, f"=D{r}/D${stot}", PCT, alt=alt)
            _cell(ws, r, 8, s.get("closing_mnt"), NUM, alt=alt)
            if has_yoy:
                _cell(ws, r, 9, pct_static(s.get("yoy_pct")), PCT, alt=alt)
            r += 1
        _cell(ws, stot, 2, "Нийт (Total)", bold=True)
        _cell(ws, stot, 3, f"=SUM(C{sfirst}:C{slast})", NUM, bold=True)
        _cell(ws, stot, 4, f"=SUM(D{sfirst}:D{slast})", NUM, bold=True)
        _cell(ws, stot, 5, f"=SUM(E{sfirst}:E{slast})", NUM, bold=True)
        _cell(ws, stot, 6, f"=SUM(F{sfirst}:F{slast})", NUM2, bold=True)
        _cell(ws, stot, 7, f"=SUM(G{sfirst}:G{slast})", PCT, bold=True)
        _cell(ws, stot, 8, f"=SUM(H{sfirst}:H{slast})", NUM, bold=True)
        r = stot + 2

    # 2b. named detail lines (salary, pension, herder loans ...)
    subs = d.get("sublines", [])
    if subs:
        _cell(ws, r, 1, "2b. Нэрлэсэн зээлүүд (named loan products, already "
                        "included in their sector above)", bold=True)
        r += 1
        r = _header_row(ws, r, ["Код", "Зээл (Product)",
                                "Эхний үлдэгдэл (Opening)",
                                "Эцсийн үлдэгдэл (Closing)",
                                "Өөрчлөлт (Change)"],
                        [7, 50, 17, 17, 15])
        for i, s in enumerate(subs):
            alt = i % 2 == 1
            _cell(ws, r, 1, s.get("code"), alt=alt)
            _cell(ws, r, 2, s.get("label"), alt=alt)
            _cell(ws, r, 3, s.get("opening"), NUM, alt=alt)
            _cell(ws, r, 4, s.get("closing"), NUM, alt=alt)
            _cell(ws, r, 5, f"=D{r}-C{r}", NUM, alt=alt)
            r += 1
        r += 1

    # 3. LOAN QUALITY / NPL
    qual = [q for q in d.get("quality", []) if q.get("closing") is not None]
    if qual:
        _cell(ws, r, 1, "3. ЧАНАРААР (loan quality and NPL)", bold=True)
        r += 1
        r = _header_row(ws, r, ["Ангилал (Class)", "Эхний үлдэгдэл (Opening)",
                                "Эцсийн үлдэгдэл (Closing)",
                                "Өөрчлөлт (Change)", "Эзлэх (Share)"],
                        [46, 17, 17, 15, 11])
        qrows = {}
        book = d["bridge"].get("closing")
        for i, q in enumerate(qual):
            alt = i % 2 == 1
            _cell(ws, r, 1, q.get("label"), alt=alt)
            _cell(ws, r, 2, q.get("opening"), NUM, alt=alt)
            _cell(ws, r, 3, q.get("closing"), NUM, alt=alt)
            _cell(ws, r, 4, f"=C{r}-B{r}", NUM, alt=alt)
            _cell(ws, r, 5, f"=C{r}/{book}" if book else None, PCT, alt=alt)
            qrows[q.get("code")] = r
            r += 1
        npl_cells = [f"C{qrows[c]}" for c in ("3", "4", "5") if c in qrows]
        if npl_cells:
            _cell(ws, r, 1, "Чанаргүй зээл (NPL = substandard + doubtful + "
                            "loss)", bold=True)
            _cell(ws, r, 3, "=" + "+".join(npl_cells), NUM, bold=True)
            if book:
                _cell(ws, r, 5, f"=C{r}/{book}", PCT, bold=True)
            r += 1
        r += 1

    # 4. BORROWER TYPE
    bor = [b_ for b_ in d.get("borrowers", [])
           if b_.get("closing") is not None and b_.get("sheet") != "Total"]
    if bor:
        _cell(ws, r, 1, "4. ЗЭЭЛДЭГЧЭЭР (by borrower type)", bold=True)
        r += 1
        r = _header_row(ws, r, ["Зээлдэгч (Borrower)",
                                "Эхний үлдэгдэл (Opening)",
                                "Эцсийн үлдэгдэл (Closing)",
                                "Өөрчлөлт (Change)", "Өсөлт (Growth)",
                                "Эзлэх (Share)", "Зээлдэгчийн тоо (Count)"],
                        [46, 17, 17, 15, 11, 11, 17])
        bfirst = r
        for i, b_ in enumerate(bor):
            alt = i % 2 == 1
            _cell(ws, r, 1, b_.get("label"), alt=alt)
            _cell(ws, r, 2, b_.get("opening"), NUM, alt=alt)
            _cell(ws, r, 3, b_.get("closing"), NUM, alt=alt)
            _cell(ws, r, 4, f"=C{r}-B{r}", NUM, alt=alt)
            _cell(ws, r, 5, f"=C{r}/B{r}-1" if b_.get("opening") else None,
                  PCT, alt=alt)
            _cell(ws, r, 6, f"=C{r}/C${bfirst + len(bor)}", PCT, alt=alt)
            _cell(ws, r, 7, b_.get("borrowers"), NUM, alt=alt)
            r += 1
        _cell(ws, r, 1, "Нийт (Total)", bold=True)
        _cell(ws, r, 2, f"=SUM(B{bfirst}:B{r - 1})", NUM, bold=True)
        _cell(ws, r, 3, f"=SUM(C{bfirst}:C{r - 1})", NUM, bold=True)
        _cell(ws, r, 4, f"=SUM(D{bfirst}:D{r - 1})", NUM, bold=True)
        _cell(ws, r, 7, f"=SUM(G{bfirst}:G{r - 1})", NUM, bold=True)
        r += 2

    # 5. RATES
    rates = d.get("rates", {})
    if rates:
        _cell(ws, r, 1, "5. ХҮҮ, ХУГАЦАА (weighted average, total book)",
              bold=True)
        r += 1
        r = _header_row(ws, r, ["Үзүүлэлт (Measure)", "Төгрөг (MNT)",
                                "Гадаад валют (FX)"], [46, 18, 18])
        _cell(ws, r, 1, "Тайлант сард олгосон зээлийн хүү (New lending rate)")
        _cell(ws, r, 2, pct_static(rates.get("rate_new_mnt")), PCT)
        _cell(ws, r, 3, pct_static(rates.get("rate_new_fx")), PCT)
        r += 1
        _cell(ws, r, 1, "Үлдэгдэл зээлийн хүү (Rate on outstanding stock)")
        _cell(ws, r, 2, pct_static(rates.get("rate_stock_mnt")), PCT)
        _cell(ws, r, 3, pct_static(rates.get("rate_stock_fx")), PCT)
        r += 1
        if rates.get("maturity_stock_mnt") is not None:
            _cell(ws, r, 1, "Үлдэгдэл зээлийн хугацаа, сар (Maturity, months)")
            _cell(ws, r, 2, round(rates["maturity_stock_mnt"], 1), NUM2)
            r += 1
        r += 1

    # 6. MONTHLY SERIES
    monthly = d.get("monthly", [])
    if len(monthly) > 1:
        _cell(ws, r, 1, "6. САРААР (monthly series — self-extending)",
              bold=True)
        r += 1
        r = _header_row(ws, r, ["Сар (Month)", "Эхний үлдэгдэл (Opening)",
                                "Эцсийн үлдэгдэл (Closing)",
                                "Олгосон (Disbursed)", "Төлөгдсөн (Repaid)",
                                "Өсөлт (MoM)"],
                        [12, 17, 17, 17, 17, 11])
        for i, m in enumerate(monthly):
            alt = i % 2 == 1
            _cell(ws, r, 1, m.get("period"), alt=alt)
            _cell(ws, r, 2, m.get("opening"), NUM, alt=alt)
            _cell(ws, r, 3, m.get("closing"), NUM, alt=alt)
            _cell(ws, r, 4, m.get("disbursed"), NUM, alt=alt)
            _cell(ws, r, 5, m.get("repaid"), NUM, alt=alt)
            _cell(ws, r, 6, f"=C{r}/B{r}-1" if m.get("opening") else None,
                  PCT, alt=alt)
            r += 1
    elif monthly:
        _cell(ws, r, 1, "6. САРААР — only one month cached so far; the "
                        "13-month history is downloaded on the next live "
                        "ingest and this block then fills automatically.")


# ---------------------------------------------------------------- charts
# The analyst draws these by hand in his own workbooks every month. Two
# patterns, taken from `trade may 2025.xlsx` and `hotels and restaurant may
# 2025.xlsx` and reproduced here so the deck can lift them directly:
#
#   A. contribution chart — components stacked as columns, the sector's own
#      growth drawn as a line across the top. Bars touch (overlap 100), no
#      gridlines, legend underneath, axis as a percentage.
#   B. level chart — the single-period level as columns against the
#      cumulative year-on-year growth as a line on a second axis.
#
# His first series is purple and the rest run through the Office accents;
# the headline line is black. Values are written as FRACTIONS so a 0.0%
# axis format displays them, exactly as his sheets do.
CHART_COLORS = ["7030A0", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
                "4472C4", "ED7D31", "264478", "9E480E", "636363"]
CHART_LINE = "000000"


def _style_axes(ch, numfmt):
    ch.y_axis.numFmt = numfmt
    ch.y_axis.majorGridlines = None
    ch.x_axis.majorGridlines = None
    ch.y_axis.majorTickMark = "out"
    ch.x_axis.majorTickMark = "out"
    ch.x_axis.delete = False
    ch.y_axis.delete = False


def _contribution_chart(ws, title, hdr_row, first, last, ncomp, growth_col,
                        freq):
    """Pattern A — stacked component contributions with the headline on top."""
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 0
    bar.add_data(Reference(ws, min_col=2, max_col=1 + ncomp,
                           min_row=hdr_row, max_row=last),
                 titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=first, max_row=last)
    bar.set_categories(cats)
    for i, s in enumerate(bar.series):
        s.graphicalProperties.solidFill = CHART_COLORS[i % len(CHART_COLORS)]
        s.graphicalProperties.line.noFill = True

    ln = LineChart()
    ln.add_data(Reference(ws, min_col=growth_col, min_row=hdr_row,
                          max_row=last), titles_from_data=True)
    ln.set_categories(cats)
    s = ln.series[0]
    s.graphicalProperties.line = LineProperties(solidFill=CHART_LINE, w=22000)
    s.marker = Marker(symbol="none")
    s.smooth = False

    bar += ln
    bar.title = title
    bar.style = None
    _style_axes(bar, "0.0%")
    bar.x_axis.tickLblSkip = 6 if freq == "M" else 4
    bar.x_axis.tickMarkSkip = 6 if freq == "M" else 4
    bar.legend.position = "b"
    bar.legend.overlay = False
    bar.height, bar.width = 9.5, 26
    return bar


def _level_chart(ws, title, hdr_row, first, last, level_col, growth_col,
                 freq, unit):
    """Pattern B — the period's level as columns, YoY growth on a 2nd axis."""
    bar = BarChart()
    bar.type, bar.grouping, bar.gapWidth = "col", "clustered", 50
    bar.add_data(Reference(ws, min_col=level_col, min_row=hdr_row,
                           max_row=last), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    bar.series[0].graphicalProperties.solidFill = "BFBFBF"
    bar.series[0].graphicalProperties.line.solidFill = "808080"
    bar.y_axis.numFmt = "#,##0"
    bar.y_axis.title = unit
    bar.y_axis.majorGridlines = None
    bar.x_axis.majorGridlines = None
    bar.x_axis.tickLblSkip = 6 if freq == "M" else 4
    bar.x_axis.tickMarkSkip = 6 if freq == "M" else 4

    ln = LineChart()
    ln.add_data(Reference(ws, min_col=growth_col, min_row=hdr_row,
                          max_row=last), titles_from_data=True)
    s = ln.series[0]
    s.graphicalProperties.line = LineProperties(solidFill=CHART_LINE, w=22000)
    s.marker = Marker(symbol="none")
    s.smooth = False
    ln.y_axis.axId = 200
    ln.y_axis.numFmt = "0.0%"
    ln.y_axis.title = "Жилийн өсөлт"
    ln.y_axis.majorGridlines = None
    ln.y_axis.crosses = "max"

    bar += ln
    bar.title = title
    bar.style = None
    bar.legend.position = "b"
    bar.legend.overlay = False
    bar.height, bar.width = 9.5, 26
    return bar


# How much history each chart shows. Three years, not five: the 2021-22
# reopening put a 195% year-on-year spike in the hotel series, and any window
# containing it rescales the axis to 250% and flattens everything recent into
# a smear. The full history is in the JSON if a longer chart is ever wanted.
CHART_WINDOW = {"M": 36, "Q": 16}


def _industry_detail_blocks(ws, series, r):
    """
    The four themed industry charts from the analyst's own deck: mining, food
    processing, light manufacturing, and chemicals and metals. Each stacks
    its sub-sectors' contributions to TOTAL industry growth, with the group's
    subtotal as the line -- so the four read against one number and can be
    compared with each other.
    """
    det = series.get("industry_detail") or {}
    if not det.get("groups"):
        return r
    _cell(ws, r, 1, "Аж үйлдвэр, дэд салбараар — 4 бүлэг. Хувь нэмэр нь "
                    "АЖ ҮЙЛДВЭРИЙН НИЙТ өсөлтөд эзлэх хувь (each series is "
                    "its contribution to TOTAL industry growth, so the four "
                    f"charts share one denominator: {det.get('growth_pct')}%)",
          bold=True)
    r += 2
    for g in det["groups"]:
        rows = (g.get("series") or [])[-CHART_WINDOW["M"]:]
        labels = [l for l in g.get("components", [])
                  if any(l in x["contributions"] for x in rows)]
        if len(rows) < 4 or not labels:
            continue
        _cell(ws, r, 1, f"{g.get('label','')} — {det.get('period','')}",
              bold=True)
        r += 1
        hdr = r
        r = _header_row(ws, r, ["Үе (Period)"] + labels
                        + ["Бүлгийн дүн (Group subtotal)"],
                        [14] + [17] * len(labels) + [22])
        first = r
        gcol = 2 + len(labels)
        for i, x in enumerate(rows):
            alt = i % 2 == 1
            _cell(ws, r, 1, x["period"], alt=alt)
            for j, lbl in enumerate(labels):
                _cell(ws, r, 2 + j, x["contributions"].get(lbl), PCT, alt=alt)
            _cell(ws, r, gcol,
                  f"=SUM(B{r}:{get_column_letter(gcol - 1)}{r})", PCT, alt=alt)
            r += 1
        ws.add_chart(_contribution_chart(
            ws, f"{g.get('label','')} — аж үйлдвэрийн өсөлтөд оруулсан "
                f"хувь нэмэр, {det.get('period','')}",
            hdr, first, r - 1, len(labels), gcol, "M"),
            f"{get_column_letter(gcol + 3)}{hdr}")
        r += 21
    return r


def sheet_sector_charts(wb, k, series):
    """
    Deck-ready charts for slides 16-21, drawn the way the analyst draws them.

    The data block under each chart is live: the contributions are the same
    figures as Real_Sector, and every block carries a SUM row that must equal
    the headline growth, so a chart that stops adding up is visible in the
    sheet rather than only on the slide.
    """
    rs = series.get("real_sector", {}) or {}
    ws = wb.create_sheet("Sector_Charts")
    r = _title(ws, "Эдийн засгийн голлох салбаруудын гүйцэтгэл "
                   "(Performance of the main economic sectors)",
               "Deck slides 16-21. Each chart stacks the components' "
               "contributions with the sector's own growth as a line, in the "
               "house style. Values are fractions shown as percentages; the "
               "contributions in each row sum to that row's growth. Copy a "
               "chart straight into the deck, or read the block beneath it. "
               "Each chart shows the last three years — change CHART_WINDOW "
               "in excel_builder.py for a longer view.")
    if not rs:
        _cell(ws, r, 1, "No real-sector tables cached — run raw_ingestor.py.")
        return

    tgt = _VINTAGE.get("target", "")
    pers = sorted({d.get("period") for d in rs.values() if d.get("period")})
    c = _cell(ws, r, 1, "МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  " + ", ".join(pers)
              + f"      [тайлант үе / reported month: {tgt}]", bold=True)
    c.fill = PatternFill("solid", start_color="C6EFCE")
    r += 2

    order = ["trade", "hotel", "food", "construction", "industry", "transport"]
    for key in [x for x in order if x in rs] + [x for x in rs if x not in order]:
        d = rs[key]
        rows = (d.get("chart_series") or [])[-CHART_WINDOW.get(
            d.get("freq"), 60):]
        if len(rows) < 4:
            continue
        labels = [c["label"] for c in d.get("components", [])]
        labels = [l for l in labels
                  if any(l in x["contributions"] for x in rows)]
        if not labels:
            continue

        _cell(ws, r, 1, f"{d.get('label','')} — {d.get('period','')}",
              bold=True)
        r += 1
        hdr = r
        r = _header_row(ws, r, ["Үе (Period)"] + labels
                        + ["Салбарын өсөлт (Growth)", "Нийлбэр (check)"],
                        [14] + [17] * len(labels) + [22, 16])
        first = r
        gcol = 2 + len(labels)
        for i, x in enumerate(rows):
            alt = i % 2 == 1
            _cell(ws, r, 1, x["period"], alt=alt)
            for j, lbl in enumerate(labels):
                _cell(ws, r, 2 + j, x["contributions"].get(lbl), PCT, alt=alt)
            _cell(ws, r, gcol, x["growth"], PCT, alt=alt)
            # rounded to 5dp so ordinary floating-point dust reads as zero;
            # a genuine break is orders of magnitude larger than that
            _cell(ws, r, gcol + 1,
                  f"=ROUND(SUM(B{r}:{get_column_letter(gcol - 1)}{r})"
                  f"-{get_column_letter(gcol)}{r},5)", PCT, alt=alt)
            r += 1
        last = r - 1
        _cell(ws, r, 1, "Шалгалт: сүүлийн багана бүх мөрөнд 0 байх ёстой "
                        "(the check column must read 0 on every row — the "
                        "bars stack up to the line)")
        r += 1

        ws.add_chart(_contribution_chart(
            ws, f"{d.get('label','')} — хувь нэмэр, {d.get('period','')}",
            hdr, first, last, len(labels), gcol, d.get("freq")),
            f"{get_column_letter(gcol + 3)}{hdr}")

        # INDUSTRY, IN DETAIL. His deck gives industry four charts of its own,
        # one per theme, each showing sub-sectors' contributions to TOTAL
        # industry growth with the group's own subtotal as the line. Those
        # go straight after the industry headline block.
        if key == "industry":
            r = _industry_detail_blocks(ws, series, r)

        # the level chart he draws for the hotel and food sectors
        if key in ("hotel", "food"):
            monthly = [m for m in d.get("monthly", [])
                       if m.get("value") is not None][-CHART_WINDOW.get(
                           d.get("freq"), 60):]
            if len(monthly) > 4:
                lh = r
                r = _header_row(ws, r, ["Үе (Period)",
                                        f"Тухайн үе ({d.get('unit','')})",
                                        "Өссөн дүнгийн жилийн өсөлт"],
                                [14, 26, 26])
                lf = r
                growth = {x["period"]: x["growth"] for x in d["chart_series"]}
                for i, m in enumerate(monthly):
                    alt = i % 2 == 1
                    _cell(ws, r, 1, m.get("period"), alt=alt)
                    _cell(ws, r, 2, m.get("value"), NUM, alt=alt)
                    _cell(ws, r, 3, growth.get(m.get("period")), PCT, alt=alt)
                    r += 1
                ws.add_chart(_level_chart(
                    ws, f"{d.get('label','')} — түвшин ба өсөлт",
                    lh, lf, r - 1, 2, 3, d.get("freq"), d.get("unit", "")),
                    f"E{lh}")
        r += 22


def sheet_weekly_prices(wb, k, series):
    """
    NSO's seven-day price survey for Ulaanbaatar — deck slide 9.

    This used to be two yellow cells copied off a presentation PDF. It is now
    the survey itself: every product, every week, with the change since the
    start of the year and against a year ago, and a chart of the two the deck
    talks about.

    Unlike every other source, this one is NOT cut off at the reported month.
    It is a leading indicator — the point of the slide is where meat and fuel
    are heading before the monthly CPI catches up — so it runs to the newest
    published week and carries that week's own date. The value at the last
    week inside the reported month is shown beside it, so a slide can quote
    either and say which.
    """
    wp = series.get("weekly_prices") or {}
    ws = wb.create_sheet("Weekly_Prices")
    r = _title(ws, "7 хоногийн үнийн мэдээ, Улаанбаатар "
                   "(Weekly prices, Ulaanbaatar)",
               "Deck slide 9. Source: NSO, Хүнсний гол нэрийн барааны 7 "
               "хоногийн үнийн мэдээ. Prices in tugrik. YTD compares the "
               "latest week with the first week of the reported year; YoY "
               "with the same week last year.")
    if not wp.get("products"):
        _cell(ws, r, 1, "No weekly price table cached — run raw_ingestor.py.")
        return
    tgt = _VINTAGE.get("target", "")
    me = wp.get("month_end_as_of")
    ahead = bool(me and me != wp.get("as_of"))
    c = _cell(ws, r, 1, "МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  "
              f"{wp.get('as_of','')} (7 хоног бүр / weekly)      "
              f"[тайлант үе / reported month: {tgt}]"
              + (f"   ← ТЭРГҮҮЛЭХ ҮЗҮҮЛЭЛТ: тайлант сараас хойшхи 7 хоногийг "
                 f"хассангүй. Тайлант сарын сүүлийн 7 хоног: {me} "
                 f"(leading indicator — weeks after the reported month are "
                 f"kept; last week inside it was {me})" if ahead else ""),
              bold=True)
    c.fill = PatternFill("solid", start_color="FFEB9C" if ahead else "C6EFCE")
    r += 1
    _cell(ws, r, 1, f"Эх файл (source table): {wp.get('source_file','?')}"
          + (("   ·   " + "   ·   ".join(wp.get("considered", [])))
             if len(wp.get("considered", [])) > 1 else ""))
    r += 2

    if wp.get("groups"):
        _cell(ws, r, 1, "Бүлгээр (by group — simple average of the members' "
                        "changes)", bold=True)
        r += 1
        r = _header_row(ws, r, ["Бүлэг (Group)", "Оны эхнээс (YTD)",
                                "Жилийн өмнөхөөс (YoY)", "Бүрэлдэхүүн"],
                        [26, 16, 18, 74])
        for i, g in enumerate(wp["groups"]):
            alt = i % 2 == 1
            _cell(ws, r, 1, g["label"], alt=alt)
            _cell(ws, r, 2, pct_static(g.get("ytd_pct")), PCT, alt=alt)
            _cell(ws, r, 3, pct_static(g.get("yoy_pct")), PCT, alt=alt)
            _cell(ws, r, 4, ", ".join(g.get("members", []))[:230], alt=alt)
            r += 1
        r += 1

    _cell(ws, r, 1, "Бүтээгдэхүүнээр (by product)", bold=True)
    r += 1
    r = _header_row(ws, r, ["Бүтээгдэхүүн (Product)",
                            f"Сүүлийн үнэ (Latest, {wp.get('as_of','')})",
                            f"Тайлант сарын эцэст ({me or 'n/a'})",
                            "Оны эхнээс (YTD)", "Жилийн өмнөхөөс (YoY)"],
                    [46, 20, 20, 14, 16])
    for i, p in enumerate(wp["products"]):
        alt = i % 2 == 1
        _cell(ws, r, 1, p["label"], alt=alt)
        _cell(ws, r, 2, p.get("current"), NUM, alt=alt)
        _cell(ws, r, 3, p.get("month_end"), NUM, alt=alt)
        _cell(ws, r, 4, pct_static(p.get("ytd_pct")), PCT, alt=alt)
        _cell(ws, r, 5, pct_static(p.get("yoy_pct")), PCT, alt=alt)
        r += 1
    r += 2

    # The weekly series themselves, and the slide's chart on top of them.
    weeks = wp.get("weeks", [])
    show = [p for p in wp["products"]
            if any(w in p["label"].casefold()
                   for w in ("үхрийн", "хонин", "хонь", "аи-", "дизел"))]
    if weeks and show:
        _cell(ws, r, 1, "Долоо хоногийн цуваа — махны болон шатахууны үнэ "
                        "(weekly series, meat and fuel)", bold=True)
        r += 1
        hdr = r
        r = _header_row(ws, r, ["7 хоног (Week)"] + [p["label"] for p in show],
                        [14] + [20] * len(show))
        first = r
        for i, wk in enumerate(weeks):
            alt = i % 2 == 1
            _cell(ws, r, 1, wk, alt=alt)
            for j, p in enumerate(show):
                _cell(ws, r, 2 + j, p["values"][i], NUM, alt=alt)
            r += 1
        ln = LineChart()
        ln.add_data(Reference(ws, min_col=2, max_col=1 + len(show),
                              min_row=hdr, max_row=r - 1),
                    titles_from_data=True)
        ln.set_categories(Reference(ws, min_col=1, min_row=first,
                                    max_row=r - 1))
        for i, s in enumerate(ln.series):
            s.marker = Marker(symbol="none")
            s.smooth = False
            s.graphicalProperties.line = LineProperties(
                solidFill=CHART_COLORS[i % len(CHART_COLORS)], w=20000)
        ln.title = f"7 хоногийн үнэ, ₮ — {wp.get('as_of','')}"
        ln.style = None
        ln.y_axis.numFmt = "#,##0"
        ln.y_axis.majorGridlines = None
        ln.x_axis.majorGridlines = None
        ln.x_axis.tickLblSkip = ln.x_axis.tickMarkSkip = 13
        ln.legend.position = "b"
        ln.legend.overlay = False
        ln.height, ln.width = 9.5, 26
        ws.add_chart(ln, f"{get_column_letter(len(show) + 4)}{hdr}")


def sheet_real_sector(wb, k, series):
    """
    Trade, services, construction and industry — deck slides 16-21.

    Laid out the way the analyst's own workbooks are: the cumulative total
    and its year-on-year growth, then each component's CONTRIBUTION to that
    growth in percentage points. Contributions are additive by construction,
    so the total row is a SUM formula that must equal the headline growth --
    visible proof in the sheet, not just in the log.

    On the deck these become a stacked bar of contributions with the
    headline growth as a line; hotel and food are plotted as the
    single-period value with the year-on-year growth as the line.
    """
    rs = series.get("real_sector", {}) or {}
    ws = wb.create_sheet("Real_Sector")
    r = _title(ws, "Бодит салбарууд (Real sector — trade, services, "
                   "construction, industry)",
               "Deck slides 16-21. Source: NSO 1212.mn. Cumulative from the "
               "start of the year (өссөн дүнгээр); growth is year on year on "
               "that cumulative; each component's contribution is its share "
               "of the year-on-year change, so the contributions sum to the "
               "headline growth. Money in сая төг.")
    if not rs:
        _cell(ws, r, 1, "No real-sector tables cached — run raw_ingestor.py.")
        return

    order = ["trade", "hotel", "food", "construction", "industry", "transport"]
    keys = [x for x in order if x in rs] + [x for x in rs if x not in order]

    # Sheet-level banner. Each table below also carries its own month, because
    # NSO publishes these six on different calendars -- but a reader opening
    # straight into this sheet should see the spread without scrolling.
    tgt = _VINTAGE.get("target", "")
    pers = sorted({rs[x].get("period") for x in keys if rs[x].get("period")})

    def _as_month(p):
        # '2026-Q1' must rank as March, not sort after '2026-06' on the 'Q'.
        m = re.fullmatch(r"(\d{4})-Q([1-4])", str(p))
        return f"{m.group(1)}-{int(m.group(2)) * 3:02d}" if m else str(p)

    lag = [p for p in pers if tgt and _as_month(p) < tgt]
    txt = ("МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  " + ", ".join(pers)
           + f"      [тайлант үе / reported month: {tgt}]")
    if lag:
        txt += ("   ← ЗАРИМ ХҮСНЭГТ ӨМНӨХ САРЫНХ (some tables lag; "
                "each table's own month is on its header line)")
    c = _cell(ws, r, 1, txt, bold=True)
    c.fill = PatternFill("solid", start_color="FFC7CE" if lag else "C6EFCE")
    r += 2

    for key in keys:
        d = rs[key]
        _cell(ws, r, 1, f"{d.get('label','')} — {d.get('period','')}",
              bold=True)
        r += 1
        _cell(ws, r, 1, ("Улирлаар, өссөн дүнгээр" if d.get("freq") == "Q"
                         else "Сараар, өссөн дүнгээр")
              + f"   ·   нийт {format(d.get('total_cur') or 0, ',.1f')} сая төг"
              + f"   ·   өмнөх оны мөн үе {format(d.get('total_prev') or 0, ',.1f')}")
        r += 1
        r = _header_row(ws, r, ["Бүрэлдэхүүн (Component)",
                                "Өмнөх он (Prev, cum.)",
                                "Тайлант он (Cur, cum.)",
                                "Өөрчлөлт (Change)",
                                "Хувь нэмэр (Contribution, pp)",
                                "Жилийн өсөлт (YoY)"],
                        [46, 18, 18, 16, 22, 14])
        first = r
        comps = d.get("components", [])
        for i, c in enumerate(comps):
            alt = i % 2 == 1
            _cell(ws, r, 1, c.get("label"), alt=alt)
            _cell(ws, r, 2, c.get("prev"), NUM, alt=alt)
            _cell(ws, r, 3, c.get("cur"), NUM, alt=alt)
            _cell(ws, r, 4, f"=C{r}-B{r}", NUM, alt=alt)
            _cell(ws, r, 5, c.get("contribution_pct"), NUM2, alt=alt)
            _cell(ws, r, 6, pct_static(c.get("yoy_pct")), PCT, alt=alt)
            r += 1
        last = r - 1
        if comps:
            _cell(ws, r, 1, "Нийт (Total — must equal the headline growth)",
                  bold=True)
            _cell(ws, r, 2, f"=SUM(B{first}:B{last})", NUM, bold=True)
            _cell(ws, r, 3, f"=SUM(C{first}:C{last})", NUM, bold=True)
            _cell(ws, r, 4, f"=C{r}-B{r}", NUM, bold=True)
            _cell(ws, r, 5, f"=SUM(E{first}:E{last})", NUM2, bold=True)
            _cell(ws, r, 6, f"=C{r}/B{r}-1", PCT, bold=True)
            r += 1
        _cell(ws, r, 1, "Салбарын өсөлт (Headline growth, YoY)", bold=True)
        _cell(ws, r, 5, d.get("growth_pct"), NUM2, bold=True)
        r += 2

        # the series behind the chart: cumulative, and the single-period value
        monthly = [m for m in d.get("monthly", [])
                   if m.get("cumulative") is not None]
        if len(monthly) > 1:
            _cell(ws, r, 1, "Цуваа (series for the chart)", bold=True)
            r += 1
            r = _header_row(ws, r, ["Үе (Period)", "Өссөн дүн (Cumulative)",
                                    "Тухайн үе (Single period)"],
                            [14, 20, 20])
            for i, m in enumerate(monthly):
                alt = i % 2 == 1
                _cell(ws, r, 1, m.get("period"), alt=alt)
                _cell(ws, r, 2, m.get("cumulative"), NUM, alt=alt)
                _cell(ws, r, 3, m.get("value"), NUM, alt=alt)
                r += 1
        r += 2


def sheet_forecast(wb, k, series):
    """
    IMF World Economic Outlook — the benchmark and the forward view.

    Everything else in this workbook measures the present from Mongolian
    primary sources. These are ANNUAL PROJECTIONS from a twice-yearly IMF
    publication, so the sheet says so at the top and on every block. The
    comparison table is deliberately short: only where the IMF's basis
    matches ours, with the basis written on the row.
    """
    w = series.get("imf_weo") or {}
    ws = wb.create_sheet("Forecast")
    r = _title(ws, "IMF World Economic Outlook — төсөөлөл (projections)",
               "ТӨСӨӨЛӨЛ, ХЭМЖИЛТ БИШ / PROJECTIONS, NOT MEASUREMENTS. "
               "Annual figures from the IMF's twice-yearly World Economic "
               "Outlook. Every other sheet in this workbook measures what "
               "happened; this one states what the IMF expects. Never "
               "present a figure from this sheet as an outturn.")
    if not w:
        _cell(ws, r, 1, "No IMF WEO export cached — drop the CSV from "
                        "imf.org into raw_files/ (any name containing 'WEO').")
        return
    c = _cell(ws, r, 1,
              "МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  "
              f"{w.get('vintage','')} (FORECAST — жилийн төсөөлөл, сарын "
              f"хэмжилт биш)      [тайлант үе / reported month: "
              f"{_VINTAGE.get('target','')}]", bold=True)
    c.fill = PatternFill("solid", start_color="FFEB9C")
    r += 1
    _cell(ws, r, 1, f"Улс: {w.get('country','')}   ·   {w.get('vintage','')}"
                    f"   ·   эх файл: {w.get('source','')}", bold=True)
    r += 2

    comp = w.get("comparison") or []
    if comp:
        _cell(ws, r, 1, f"1. ГҮЙЦЭТГЭЛ vs ТӨСӨӨЛӨЛ, {k.get('target_year','')} "
                        f"(actual against the IMF forecast)", bold=True)
        r += 1
        _cell(ws, r, 1, "Only indicators whose basis genuinely matches are "
                        "compared. The basis column says exactly what each "
                        "side measures — read it before quoting a gap.")
        r += 1
        r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "IMF төсөөлөл",
                                "Бодит гүйцэтгэл (actual)",
                                "Зөрүү (gap, pp)", "Суурь (basis)"],
                        [42, 16, 22, 16, 62])
        for i, c in enumerate(comp):
            alt = i % 2 == 1
            _cell(ws, r, 1, c.get("label"), bold=True, alt=alt)
            _cell(ws, r, 2, c.get("forecast"), NUM2, alt=alt)
            _cell(ws, r, 3, c.get("actual"), NUM2, alt=alt)
            cell = _cell(ws, r, 4, c.get("gap"), NUM2, bold=True, alt=alt)
            if abs(c.get("gap") or 0) >= 1.5:
                cell.fill = PatternFill("solid", start_color="FFEB9C")
            _cell(ws, r, 5, c.get("basis"), alt=alt)
            r += 1
        r += 2

    inds = w.get("indicators") or []
    years = w.get("years") or []
    if inds and years:
        _cell(ws, r, 1, "2. ТӨСӨӨЛӨЛ, оноор (IMF projections by year)",
              bold=True)
        r += 1
        _cell(ws, r, 1, "Years up to the last outturn are IMF estimates of "
                        "history; later years are projections.")
        r += 1
        r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Нэгж (Unit)"] + years,
                        [46, 12] + [11] * len(years))
        for i, d in enumerate(inds):
            alt = i % 2 == 1
            _cell(ws, r, 1, d.get("label"), bold=True, alt=alt)
            _cell(ws, r, 2, d.get("unit"), alt=alt)
            for j, y in enumerate(years, start=3):
                _cell(ws, r, j, (d.get("values") or {}).get(y), NUM2, alt=alt)
            r += 1
        r += 1
        _cell(ws, r, 1, "Тайлбар (basis notes)", bold=True)
        r += 1
        for d in inds:
            _cell(ws, r, 1, f"{d.get('label')} — {d.get('basis','')}")
            r += 1


def sheet_vintage(wb, series):
    """
    One table answering 'which month is each number from?'.

    Sources publish on different calendars, so a report is always a mix of
    vintages. Stating that per source, in one place, is what stops a reader
    taking a lagged figure for a current one.
    """
    v = series.get("vintage", {}) or {}
    rows = v.get("rows", [])
    ws = wb.create_sheet("Data_Vintage")
    tgt = v.get("target", "")
    r = _title(ws, f"Мэдээллийн он сар (Data vintage) — тайлант үе {tgt}",
               "Which month each source actually carries. CURRENT = the "
               "reported month. LAGGED = the source has not published the "
               "reported month yet, so the PREVIOUS month's figure is shown "
               "and the period column says which. Nothing is ever "
               "extrapolated or carried forward silently.")
    if not rows:
        _cell(ws, r, 1, "No vintage information — run the pipeline.")
        return
    r = _header_row(ws, r, ["Эх сурвалж (Source)", "Хаана орсон (Feeds)",
                            "Аль сарынх (Period of the data)",
                            "Төлөв (Status)", "Тайлбар (Note)"],
                    [30, 42, 22, 14, 62])
    fills = {"CURRENT": "C6EFCE", "LAGGED": "FFC7CE", "MISSING": "FFEB9C",
             "QUARTERLY": "DDEBF7", "PER-CARD": "F2F2F2", "AHEAD": "FFEB9C"}
    for i, d in enumerate(rows):
        alt = i % 2 == 1
        _cell(ws, r, 1, d.get("source"), bold=True, alt=alt)
        _cell(ws, r, 2, d.get("feeds"), alt=alt)
        _cell(ws, r, 3, d.get("period"), bold=True, alt=alt)
        c = _cell(ws, r, 4, d.get("status"), bold=True)
        col = fills.get(d.get("status"))
        if col:
            c.fill = PatternFill("solid", start_color=col)
        _cell(ws, r, 5, d.get("note"), alt=alt)
        r += 1
    r += 1
    lagged = [d for d in rows if d.get("status") == "LAGGED"]
    if lagged:
        _cell(ws, r, 1, "АНХААР (read before using): "
              + "; ".join(f"{d['source']} is {d['period']}, not {tgt}"
                          for d in lagged), bold=True)
    else:
        _cell(ws, r, 1, f"Бүх эх сурвалж {tgt} үеийнх "
                        f"(every dated source is at the reported month).",
              bold=True)


def sheet_coverage(wb, series):
    ws = wb.create_sheet("Coverage")
    r = _title(ws, "Deck coverage — what this workbook automates",
               "AUTO = filled by pipeline | MISSING = source not yet wired "
               "(see note) | MANUAL = can never be data-driven.")
    c = _cell(ws, r, 1, "МЭДЭЭЛЛИЙН ОН САР / DATA AS OF:  энэ хуудас "
                        "бүтэц, тоо биш (this sheet is structure, not data)."
                        f"      [тайлант үе / reported month: "
                        f"{_VINTAGE.get('target','')}]", bold=True)
    c.fill = PatternFill("solid", start_color="F2F2F2")
    r += 2
    r = _header_row(ws, r, ["Slide", "Content", "Status", "Source", "Note"],
                    [8, 34, 10, 16, 52])
    colors = {"AUTO": "C6EFCE", "MISSING": "FFEB9C", "MANUAL": "F2F2F2"}
    for d in series.get("coverage", []):
        _cell(ws, r, 1, d["slide"])
        _cell(ws, r, 2, d["content"])
        c = _cell(ws, r, 3, d["status"], bold=True)
        c.fill = PatternFill("solid",
                             start_color=colors.get(d["status"], "FFFFFF"))
        _cell(ws, r, 4, d["source"])
        _cell(ws, r, 5, d.get("note", ""))
        r += 1


def sheet_summary(wb, k, gdp_refs, trade_refs, commod_refs, bank_refs=None):
    """
    Built LAST (placed first): every figure that exists on a data sheet is a
    cross-sheet formula reference; only single-source KPIs with no home table
    (bulletin/Mongolbank point values) are entered here directly.
    """
    ws = wb.create_sheet("Summary", 0)
    r = _title(ws, f"Macroeconomic KPI Summary — {k.get('report_period', '')}",  # noqa: E501
               "Green figures are live links to the data sheets. Feeds deck "
               "slide 3 and headline text on slides 4, 6, 9, 12, 13.")
    # The Summary mixes sources, so it names every one of them and their
    # vintages rather than a single 'as of' month.
    _lag = [d for d in _VINTAGE["by_source"].values()
            if d.get("status") == "LAGGED"]
    _c = _cell(ws, r, 1,
               "МЭДЭЭЛЛИЙН ОН САР / DATA AS OF: энэ хуудас олон эх сурвалжийг "
               "нэгтгэнэ (this sheet mixes sources). "
               + (("ӨМНӨХ САРЫНХ / previous month: "
                   + "; ".join(f"{d['source']} = {d['period']}"
                               for d in _lag))
                  if _lag else
                  f"бүгд {_VINTAGE.get('target','')} үеийнх (all current)")
               + "  —  дэлгэрэнгүйг Data_Vintage хуудаснаас (see Data_Vintage)",
               bold=True)
    _c.fill = PatternFill("solid",
                          start_color="FFC7CE" if _lag else "C6EFCE")
    r += 2
    r = _header_row(ws, r, ["Үзүүлэлт (Indicator)", "Утга (Value)", "Нэгж (Unit)",
                            "Period / note", "Source", "Deck slide"],
                    [38, 14, 16, 30, 22, 10])
    e, i, b = (trade_refs.get("exp_row"), trade_refs.get("imp_row"),
               trade_refs.get("bal_row"))
    coal_r = commod_refs.get("exp_rows", {}).get("coal")
    copper_r = commod_refs.get("exp_rows", {}).get("copper")
    gold_r = commod_refs.get("exp_rows", {}).get("gold")
    gdp_cell = gdp_refs.get("gdp_growth_cell")

    # Bulletin figures must be stamped with the BULLETIN's own period, not
    # the month being reported. When Mongolbank has not yet published the
    # target month the pipeline serves the previous issue, and labelling that
    # "as of <target month>" would present one month's data as another's.
    bul_asof = (k.get("_bulletin_period_note")
                or f"Mongolbank bulletin, {k.get('report_date', '')}")
    bul_short = (k.get("_bulletin_period_note")
                 or f"bln MNT, {k.get('report_date', '')}")

    rows = [
        ("Бодит ДНБ-ий өсөлт (Real GDP growth, YoY cum.)",
         f"=GDP_Growth!{gdp_cell}" if gdp_cell else None, PCT,
         k.get("gdp_growth_period", ""), "NSO 1212.mn", "3, 4", True),
        ("Эдийн засгийн өсөлтийн сарын индикатор (MIEG, YoY cum.)",
         pct_static(k.get("mieg_growth_pct")), PCT,
         k.get("_mieg_period_note")
         or f"NSO MIEG, {k.get('_mieg_period', 'not cached')}",
         "NSO 1212.mn", "4", False),
        ("Инфляц (Inflation, headline)", pct_static(k.get("inflation")), PCT,
         bul_asof, "MB bulletin", "3, 6", False),
        ("Инфляцын зорилтот түвшин (Inflation target)", pct_static(k.get("inflation_target")), PCT,
         "Mongolbank official target", "Mongolbank", "6", False),
        ("Бодлогын хүү (Policy rate)", pct_static(k.get("policy_rate")), PCT,
         k.get("_policy_rate_source", ""), "Mongolbank", "3, 8", False),
        ("Зээлийн хүү (Bank loan rate, new loans)", pct_static(k.get("loan_rate")), PCT,
         k.get("_loan_rate_date", ""), "Mongolbank", "8", False),
        ("Хадгаламжийн хүү (Deposit rate, new deposits)", pct_static(k.get("deposit_rate")),
         PCT, k.get("_deposit_rate_date", ""), "Mongolbank", "8", False),
        ("Банкны салбарын зээл (Bank loans outstanding)", k.get("bank_loans_tln_mnt"), NUM,
         f"tln MNT, {k.get('_bank_loans_tln_mnt_date', '')}", "Mongolbank",
         "16", False),
        ("Ипотекийн зээл (Mortgage loans outstanding)", k.get("mortgage_loans_tln_mnt"), NUM,
         f"tln MNT, {k.get('_mortgage_loans_tln_mnt_date', '')}",
         "Mongolbank", "16", False),
        ("Нийт экспорт (Total export)", f"=Trade!C{e}" if e else None, NUM,
         f"thous. USD, {k.get('report_date', '')}", "Customs gaali.mn", "9",
         True),
        ("Нийт импорт (Total import)", f"=Trade!C{i}" if i else None, NUM,
         f"thous. USD, {k.get('report_date', '')}", "Customs gaali.mn", "9",
         True),
        ("Худалдааны тэнцэл (Trade balance)", f"=Trade!C{e}-Trade!C{i}" if e and i else None, NUM,
         "thous. USD; formula: export cell - import cell",
         "Customs gaali.mn", "9", True),
        ("Экспортын өсөлт (Export growth, YoY)",
         f"=Trade!C{e}/Trade!B{e}-1" if e else None, PCT,
         "formula: current / previous - 1", "Customs gaali.mn", "9", True),
        ("Импортын өсөлт (Import growth, YoY)",
         f"=Trade!C{i}/Trade!B{i}-1" if i else None, PCT,
         "formula: current / previous - 1", "Customs gaali.mn", "9", True),
        ("Худалдааны тэнцэл, өмнөх оны эсрэг (Trade balance vs prev yr)",
         f"=Trade!C{b}/Trade!B{b}" if b else None, MULT,
         "balance multiple", "Customs gaali.mn", "9", True),
        ("Нүүрсний экспортын хэмжээ (Coal export volume)",
         f"=Commodities!D{coal_r}/1E9" if coal_r else None, NUM2,
         "mln t", "Customs gaali.mn", "10, 13", True),
        ("Нүүрсний хэмжээний өсөлт (Coal volume growth, YoY)",
         f"=Commodities!D{coal_r}/Commodities!C{coal_r}-1" if coal_r else None,
         PCT, "formula over quantity cells", "Customs gaali.mn", "10", True),
        ("Зэсийн баяжмалын экспортын хэмжээ (Copper conc. export volume)",
         f"=Commodities!D{copper_r}/1E6" if copper_r else None, NUM2,
         "thous. t", "Customs gaali.mn", "10", True),
        ("Зэсийн хэмжээний өсөлт (Copper volume growth, YoY)",
         f"=Commodities!D{copper_r}/Commodities!C{copper_r}-1"
         if copper_r else None, PCT, "formula over quantity cells",
         "Customs gaali.mn", "10", True),
        ("Алтны экспортын хэмжээ (Gold export volume)",
         f"=Commodities!D{gold_r}" if gold_r else None, NUM,
         "kg", "Customs gaali.mn", "10", True),
        ("Алтны хэмжээний өсөлт (Gold volume growth, YoY)",
         f"=Commodities!D{gold_r}/Commodities!C{gold_r}-1"
         if gold_r else None, PCT, "formula over quantity cells",
         "Customs gaali.mn", "10", True),
        ("Гадаад валютын нөөц (FX reserves)",
         k.get("reserves_mln_usd_mb", k.get("reserves_mln_usd")), NUM,
         f"mln USD, {k.get('_reserves_mln_usd_mb_date', '')}",
         "Mongolbank", "3, 12", False),
        ("Төлбөрийн тэнцэл (Balance of payments)", k.get("balance_of_payments_mln_usd"), NUM,
         f"mln USD, {k.get('_balance_of_payments_mln_usd_date', '')}",
         "Mongolbank", "12", False),
        ("Нийт гадаад өр (External debt)", k.get("external_debt_mln_usd"), NUM,
         f"mln USD, {k.get('_external_debt_mln_usd_date', '')}",
         "Mongolbank", "12", False),
        ("ШХО-ын үлдэгдэл (FDI stock)", k.get("fdi_mln_usd"), NUM,
         f"mln USD, {k.get('_fdi_mln_usd_date', '')}", "Mongolbank", "12",
         False),
        ("Ам.долларын ханш (USD/MNT rate, eop)", k.get("usd_mnt_rate"),
         NUM, f"YoY {k.get('usd_mnt_yoy_pct', 'n/a')}%, MB bulletin "
         "Exchange rate sheet", "MB bulletin", "3", False),
        ("Банкны нийт зээлийн өсөлт (DC loans growth, YoY)",
         (k.get("dc_loans_total_yoy_pct") / 100.0
          if k.get("dc_loans_total_yoy_pct") is not None else None), PCT,
         f"loans outstanding {k.get('dc_loans_total_tln_mnt', 'n/a')} tln ₮",
         "MB bulletin", "16", False),
        ("Иргэдийн зээлийн өсөлт (Individuals' loans growth, YoY)",
         (k.get("dc_loans_individuals_yoy_pct") / 100.0
          if k.get("dc_loans_individuals_yoy_pct") is not None else None),
         PCT, f"individuals {k.get('dc_loans_individuals_tln_mnt', 'n/a')} "
         "tln ₮", "MB bulletin", "16", False),
        ("Банкны нийт актив (Bank total assets)",
         (f"=Banking!C{bank_refs['bs_total_assets']}/1E6"
          if bank_refs and k.get("bs_total_assets_mln_mnt") is not None
          else None), NUM2,
         f"tln MNT, as of {k.get('_bs_last_column', '')}", "Mongolbank",
         "16", True),
        ("Чанаргүй зээлийн өөрчлөлт (Bank NPL change since Jan)",
         (f"=Banking!C{bank_refs['bs_npl']}/Banking!B{bank_refs['bs_npl']}-1"
          if bank_refs and k.get("bs_npl_mln_mnt") is not None else None),
         PCT, k.get("_bs_change_basis", ""), "Mongolbank", "16", True),
        ("Төсвийн тэнцэл (Budget balance, cum.)", k.get("budget_balance_bln_mnt"), NUM,
         bul_short, "MB bulletin", "13", False),
        ("Төсвийн орлого, гүйцэтгэл (Budget revenue collected)", k.get("budget_revenue_collected_bln"),
         NUM, "bln MNT", "MB bulletin", "13", False),
    ]
    for idx, (name, val, fmt, note, src, slide, link) in enumerate(rows):
        alt = idx % 2 == 1
        _cell(ws, r, 1, name, bold=True, alt=alt)
        _cell(ws, r, 2, val if val is not None else "n/a",
              fmt if val is not None else None, alt=alt, link=link)
        unit_note = note
        _cell(ws, r, 3, fmt and {PCT: "%", NUM: "level", NUM2: "level",
                                 MULT: "x"}.get(fmt, ""), alt=alt)
        _cell(ws, r, 4, unit_note, alt=alt)
        _cell(ws, r, 5, src, alt=alt)
        _cell(ws, r, 6, slide, alt=alt)
        r += 1


# main
def _load_json(path):
    """Readable instruction instead of a traceback when run out of order."""
    if not path.exists():
        log.error("%s not found. Run 'python run_all.py' (or at least "
                  "'python data_processor.py <month> <year>') first.",
                  path.name)
        sys.exit(1)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.error("%s is corrupt (%s). Delete it and re-run "
                  "'python run_all.py --offline'.", path.name, exc)
        sys.exit(1)


def main():
    k = _load_json(METRICS_JSON)
    series = _load_json(SERIES_JSON)
    out = BASE_DIR / (f"Macro_Metrics_{k.get('target_year', 'X')}_"
                      f"{int(k.get('target_month', 0)):02d}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)                       # data sheets first ...
    _set_vintage(series)                       # per-sheet 'data as of' stamps
    gdp_refs = sheet_gdp(wb, series)
    sheet_sectors(wb, series)
    sheet_inflation(wb, k, series)
    trade_refs = sheet_trade(wb, k, series)
    commod_refs = sheet_commodities(wb, series)
    sheet_trade_sections(wb, series)
    sheet_border_prices(wb, series)
    sheet_external(wb, series)
    bank_refs = sheet_banking(wb, k, series)
    sheet_loans(wb, k, series)
    sheet_real_sector(wb, k, series)
    sheet_sector_charts(wb, k, series)
    sheet_weekly_prices(wb, k, series)
    sheet_forecast(wb, k, series)
    sheet_budget(wb, k, commod_refs)
    sheet_vintage(wb, series)
    sheet_coverage(wb, series)
    sheet_summary(wb, k, gdp_refs, trade_refs, commod_refs,
                  bank_refs)                                # ... linked last
    wb.move_sheet("Data_Vintage", offset=-(len(wb.sheetnames) - 2))
    wb.save(out)
    log.info("Wrote %s (%d sheets, formula-driven).", out.name,
             len(wb.sheetnames))
    return out


if __name__ == "__main__":
    main()
