#!/usr/bin/env python3
"""
deck_digest.py — turn the run's output into ONE compact markdown file for the
deck chat.

Why this exists: the workbook is the deliverable, but it is an expensive thing
for a model to read. Building the deck straight from the .xlsx means opening a
16-sheet binary, walking it cell by cell, and re-reading it for every number.
This writes every figure a slide could want as plain text, once, from the same
processed JSON the workbook itself is built from -- so the two agree by
construction and cannot drift.

    python deck_digest.py            # after run_all.py, in the same folder
    -> deck_digest.md

Attach deck_digest.md to the deck chat ALONGSIDE the workbook. The model reads
the digest; the workbook stays the citable source and the file you send on.

Only recent history is included by default (RECENT_MONTHS / RECENT_YEARS) --
a deck never plots 2015. Raise them if you need a longer chart.
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "deck_digest.md"
RECENT_MONTHS = 30          # monthly series: how many periods to carry
RECENT_YEARS = 6            # quarterly / annual series
L = []


def w(s=""):
    L.append(s)


def n(v, d=1):
    if v is None or v == "":
        return "—"
    if isinstance(v, str):
        return v
    if not isinstance(v, (int, float)):
        return str(v)[:60]
    return f"{v:,.{d}f}"


def sg(v, d=2):
    if not isinstance(v, (int, float)):
        return "—"
    return f"{v:+.{d}f}"


def table(headers, rows):
    rows = [r for r in rows if r is not None]
    if not rows:
        return
    w("| " + " | ".join(headers) + " |")
    w("|" + "|".join("---" for _ in headers) + "|")
    for r in rows:
        w("| " + " | ".join(str(x) for x in r) + " |")
    w()


def line(label, pairs, d=2):
    """A whole time series on one line — far cheaper than a table."""
    pairs = [(p, v) for p, v in pairs if v is not None]
    if not pairs:
        return
    w(f"- **{label}**: " + "; ".join(f"{p} {n(v, d)}" for p, v in pairs))


def main():
    mp, sp = HERE / "processed_metrics.json", HERE / "processed_series.json"
    if not (mp.exists() and sp.exists()):
        sys.exit("Run run_all.py first — processed_*.json not found here.")
    k = json.loads(mp.read_text(encoding="utf-8"))
    s = json.loads(sp.read_text(encoding="utf-8"))
    per = s.get("period", {})
    y, m = per.get("target_year"), per.get("target_month")
    tag = f"{y}_{m:02d}" if y and m else "?"

    w(f"# Deck digest — {per.get('report_period', tag)}")
    w()
    w(f"Every figure below comes from the same processed data as "
      f"`Macro_Metrics_{tag}.xlsx`, so the two agree exactly.")
    w()
    w("**Units, read this before charting anything.** Percentages in THIS "
      "FILE are written as percentages: `12.00` means 12%. The workbook "
      "stores the same figures as fractions (`0.12`) because its cells carry "
      "a `0.0%` format. If you put a number from this file into a chart "
      "whose axis is formatted as a percentage, **divide it by 100 first** — "
      "otherwise 12% is drawn as 1200%, which has happened. Contributions "
      "are percentage points and sum to their headline. Indented rows in the "
      "BoP and deposit tables are ALREADY INCLUDED in the line above — never "
      "sum a parent with its children.")
    w()

    # ---------------- vintage
    w("## 1. Data vintage — READ FIRST")
    w()
    w("Any figure from a LAGGED source must carry its own month on the slide.")
    w()
    table(["Source", "Feeds", "Period", "Status"],
          [[r.get("source"), r.get("feeds"), r.get("period"), r.get("status")]
           for r in (s.get("vintage") or {}).get("rows", [])])

    # ---------------- KPIs
    w("## 2. Headline KPIs")
    w()
    w("Every metric the pipeline computed, with the unit in the key. `_pct` "
      "and `_ratio` are already percentages; `_tln`/`_bln`/`_mln`/`kusd` name "
      "the money unit. Names are printed verbatim so you can cite them.")
    w()
    skip = {"target_year", "target_month"}
    table(["Metric", "Value"],
          [[a, n(v, 2 if abs(v) < 1000 else 0)]
           for a, v in sorted(k.items())
           if isinstance(v, (int, float)) and a not in skip])
    txt = {a: v for a, v in sorted(k.items()) if isinstance(v, str)}
    if txt:
        table(["Label", "Text"], [[a, v] for a, v in txt.items()])

    # ---------------- growth
    w("## 3. Growth")
    w()
    nso = s.get("nso", {})
    q = nso.get("quarterly", [])[-RECENT_YEARS * 4:]
    line("Real GDP level, 2015 prices (mln MNT)",
         [(f"{d['year']}-Q{d['quarter']}", d.get("real_gdp_2015p")) for d in q],
         0)
    for yr in sorted(nso.get("cumulative_growth", {})):
        cg = nso["cumulative_growth"][yr]
        line(f"Cumulative real GDP growth % — {yr}",
             list(cg.items()))
    cont = nso.get("contributions", {})
    newest = max(cont, default=None)
    if newest:
        c = dict(cont[newest])
        pl = c.pop("period", "")
        w()
        w(f"GDP sector contributions, {newest} {pl}, pp "
          "(these sum to GDP growth):")
        w()
        table(["Sector", "pp"], [[a, sg(v)] for a, v in c.items()])
    mi = s.get("mieg", {})
    line("MIEG monthly growth %",
         [(d.get("period"), d.get("value"))
          for d in mi.get("monthly", [])[-RECENT_MONTHS:]])
    if mi.get("contributions"):
        w()
        w(f"MIEG sector contributions, {mi.get('contributions_period','')}, "
          "pp (sum to the MIEG headline):")
        w()
        table(["Sector", "pp"],
              [[a, sg(v)] for a, v in mi["contributions"].items()])

    # ---------------- inflation
    w("## 4. Inflation, rates, FX")
    w()
    for key, lbl in (("inflation_nat", "Inflation national %"),
                     ("inflation_ub", "Inflation UB %"),
                     ("policy_rate", "Policy rate %"),
                     ("deposit_rate", "Deposit rate %"),
                     ("usd_mnt_eop", "USD/MNT eop")):
        ser = (s.get("bulletin_notes", {}).get("series", {}) or {}).get(key)
        if ser:
            line(lbl, [(d.get("period"), d.get("value")) for d in ser])
    cpi = s.get("bulletin_cpi", {})
    if cpi.get("components"):
        w()
        w(f"CPI components, {cpi.get('period','')} — the pp column sums to "
          "headline inflation:")
        w()
        table(["Category", "Weight", "YoY %", "pp"],
              [[c.get("name"), n(c.get("level"), 1), n(c.get("yoy_pct"), 1),
                sg(c.get("pp"))] for c in cpi["components"]])

    # ---------------- trade
    w("## 5. External trade")
    w()
    tm = s.get("customs", {}).get("trade_monthly", [])[-RECENT_MONTHS:]
    line("Export cumulative, thou USD",
         [(f"{d['year']}-{d['month']:02d}", d.get("export_cum_kusd"))
          for d in tm], 0)
    line("Import cumulative, thou USD",
         [(f"{d['year']}-{d['month']:02d}", d.get("import_cum_kusd"))
          for d in tm], 0)
    line("Balance cumulative, thou USD",
         [(f"{d['year']}-{d['month']:02d}",
           (d.get("export_cum_kusd") or 0) - (d.get("import_cum_kusd") or 0))
          for d in tm], 0)
    for side in ("commodity_exports", "commodity_imports"):
        d = s.get("customs", {}).get(side) or {}
        if d:
            w()
            w(f"{'Export' if 'exp' in side else 'Import'} commodities — "
              "cumulative, amount in thousand USD:")
            w()
            table(["Item", "Qty prior", "Qty now", "Qty YoY %",
                   "Amt prior", "Amt now", "Amt YoY %", "Volume unit"],
                  [[v.get("label"), n(v.get("qty_prev"), 0),
                    n(v.get("qty_cur"), 0), sg(v.get("qty_yoy_pct"), 1),
                    n(v.get("amt_prev"), 0), n(v.get("amt_cur"), 0),
                    sg(v.get("amt_yoy_pct"), 1), v.get("volume_unit", "")]
                   for v in d.values()])
    secs = s.get("customs", {}).get("trade_sections") or []
    for side, cur, prev in (("EXPORT", "exp_cur", "exp_prev"),
                            ("IMPORT", "imp_cur", "imp_prev")):
        rows = sorted(secs, key=lambda r: -(r.get(cur) or 0))
        tot = sum(r.get(cur) or 0 for r in rows)
        w(f"HS sections — {side}, ranked, thousand USD "
          f"(total {n(tot, 0)}, equals the headline):")
        w()
        table(["Section", "Prior year", "Current", "Share %", "YoY %"],
              [[r.get("mn"), n(r.get(prev), 0), n(r.get(cur), 0),
                n((r.get(cur) or 0) / tot * 100 if tot else None, 1),
                sg(((r.get(cur) or 0) / r[prev] - 1) * 100
                   if r.get(prev) else None, 1)] for r in rows])
    bp = s.get("customs", {}).get("border_prices") or {}
    if bp:
        w("Export unit values (marginal):")
        w()
        for name, ser in bp.items():
            u = ser[0].get("unit", "") if ser else ""
            line(f"{name} ({u})",
                 [(f"{d['year']}-{d['month']:02d}", d.get("price"))
                  for d in ser[-RECENT_MONTHS:]])
        w()

    # ---------------- external sector
    w("## 6. External sector")
    w()
    table(["Card", "Reported", "As of"],
          [[c.get("name"), c.get("data"), c.get("date")]
           for c in s.get("mongolbank_cards", [])])
    bop = s.get("bop", {})
    dp = bop.get("detail_period", {})
    if bop.get("detail"):
        w(f"Balance of payments, {dp.get('cur','')} against "
          f"{dp.get('prev','')}, mln USD. {bop.get('detail_lag_note','')}")
        w()
        table(["Item", "Lvl", "Current", "Prior year", "Change"],
              [[("··" * (r.get("level") or 0)) + str(r.get("label")),
                r.get("level"), n(r.get("cur")), n(r.get("prev")),
                sg((r.get("cur") or 0) - (r.get("prev") or 0), 1)]
               for r in bop["detail"]])
    ann = bop.get("annual") or {}
    if ann:
        w("BoP annual totals, mln USD:")
        w()
        for code, d in ann.items():
            vals = d.get("values") if isinstance(d, dict) else None
            if isinstance(vals, dict):
                line(d.get("label", code),
                     sorted(vals.items())[-RECENT_YEARS:], 1)
        w()

    # ---------------- banking
    w("## 7. Banking")
    w()
    bb = s.get("banking_balance_sheet", {})
    cols = bb.get("deposit_columns") or []
    if bb.get("deposit_tree"):
        w(f"Deposit tree, mln MNT — newest column {cols[-1] if cols else ''}. "
          "Indented rows are ALREADY INCLUDED above them:")
        w()
        prevcol = -13 if len(cols) >= 13 else 0
        table(["Item", "Lvl", "Latest", "Year ago", "YoY %"],
              [[("··" * (r.get("level") or 0)) + str(r.get("label")),
                r.get("level"), n((r.get("values") or [None])[-1], 0),
                n((r.get("values") or [None])[prevcol], 0),
                sg(((r["values"][-1] / r["values"][prevcol] - 1) * 100)
                   if r.get("values") and len(r["values"]) > 12
                   and r["values"][prevcol] else None, 1)]
               for r in bb["deposit_tree"]])
    bl = s.get("bank_loans", {})
    if bl.get("bridge"):
        w("Loan growth bridge, mln MNT. If an 'unexplained movement' or "
          "residual appears it is an inconsistency in Mongolbank's own "
          "published report — do NOT put it on a slide:")
        w()
        table(["Component", "Amount"],
              [[a, n(v, 1)] for a, v in bl["bridge"].items()])
    if bl.get("indicators"):
        table(["Indicator", "Total", "of which MNT", "Share %"],
              [[i.get("label"), n(i.get("total"), 1), n(i.get("mnt"), 1),
                n(i.get("share_pct"), 1)] for i in bl["indicators"]])
    if bl.get("sectors"):
        w("Loans by sector, mln MNT — the pp column sums to the month's "
          "growth rate:")
        w()
        table(["Sector", "Opening", "Closing", "Change", "pp", "YoY %"],
              [[r.get("label"), n(r.get("opening"), 0), n(r.get("closing"), 0),
                sg(r.get("change"), 1), sg(r.get("pp")),
                sg(r.get("yoy_pct"), 1)] for r in bl["sectors"]])
    for key, lbl in (("quality", "Loan quality classes"),
                     ("borrowers", "Borrower types"),
                     ("sublines", "Named loan products (already inside "
                                  "their sector — do not add)")):
        if bl.get(key):
            w(f"{lbl}, mln MNT:")
            w()
            table(["Item", "Opening", "Closing"],
                  [[r.get("label"), n(r.get("opening"), 0),
                    n(r.get("closing"), 0)] for r in bl[key]])
    if bl.get("rates"):
        table(["Rate", "%"], [[a, n(v, 2)] for a, v in bl["rates"].items()])
    if bl.get("monthly"):
        line("Loan book closing, mln MNT",
             [(d.get("period"), d.get("closing")) for d in bl["monthly"]], 0)
        line("Loan growth %",
             [(d.get("period"), d.get("growth_pct")) for d in bl["monthly"]])
        yb = [d for d in bl["monthly"] if d.get("yoy_by_borrower")]
        if yb:
            w()
            w("Year-on-year loan growth by borrower type (%) — this is the "
              "series the bank slide plots, and it is current:")
            w()
            types = list(yb[-1]["yoy_by_borrower"])
            table(["Period"] + types,
                  [[d["period"]] + [sg(d["yoy_by_borrower"].get(t))
                                    for t in types] for d in yb])
        w()

    # ---------------- real sector
    w("## 8. Real sector")
    w()
    for key, d in (s.get("real_sector") or {}).items():
        w(f"### {d.get('label')} — {d.get('period')} ({d.get('unit')})")
        w()
        w(f"Cumulative {n(d.get('total_cur'))} against "
          f"{n(d.get('total_prev'))} a year earlier = "
          f"**{sg(d.get('growth_pct'))}%**. The contributions below sum to "
          "that figure.")
        w()
        table(["Component", "Prior yr", "Current", "Change", "pp", "YoY %"],
              [[c.get("label"), n(c.get("prev")), n(c.get("cur")),
                sg(c.get("change"), 1), sg(c.get("contribution_pct")),
                sg(c.get("yoy_pct"), 1)] for c in d.get("components", [])])
        keep = RECENT_MONTHS if d.get("freq") == "M" else RECENT_YEARS * 4
        line("cumulative series",
             [(x.get("period"), x.get("cumulative"))
              for x in d.get("monthly", [])[-keep:]], 0)
        line("single-period series",
             [(x.get("period"), x.get("value"))
              for x in d.get("monthly", [])[-keep:]], 0)
        cs = (d.get("chart_series") or [])[-keep:]
        if cs:
            w()
            w("Contribution series behind the `Sector_Charts` chart — one row "
              "per period, contributions sum to that row's growth. Percent "
              "values:")
            w()
            labs = [c["label"] for c in d.get("components", [])]
            table(["Period"] + labs + ["Growth"],
                  [[x["period"]]
                   + [sg((x["contributions"].get(l) or 0) * 100)
                      for l in labs]
                   + [sg(x["growth"] * 100)] for x in cs])
        w()

    # ---------------- industry in detail
    det = s.get("industry_detail") or {}
    if det.get("groups"):
        w("## 8b. Industry sub-sectors — the four themed charts")
        w()
        w(f"Every figure is that sub-sector's contribution to TOTAL industry "
          f"growth ({det.get('growth_pct')}% in {det.get('period')}), so the "
          "four groups share one denominator and can be compared with each "
          "other. This is the split the analyst's own deck uses.")
        w()
        for g in det["groups"]:
            w(f"**{g.get('label')}** — group subtotal "
              f"{sg(g.get('subtotal_pct'))} pp")
            w()
            table(["Sub-sector", "pp"],
                  [[a, sg(v * 100)] for a, v in g.get("latest", {}).items()])
            rows = (g.get("series") or [])[-RECENT_MONTHS:]
            if rows:
                line("group subtotal over time (pp)",
                     [(x["period"], x["subtotal"] * 100) for x in rows])
                w()

    # ---------------- weekly prices
    wk = s.get("weekly_prices") or {}
    if wk.get("products"):
        w("## 8c. Weekly prices, Ulaanbaatar")
        w()
        me = wk.get("month_end_as_of")
        w(f"NSO seven-day price survey, latest week **{wk.get('as_of')}**, "
          f"from {wk.get('source_file', '?')}. Prices in tugrik.")
        w()
        if me and me != wk.get("as_of"):
            w(f"**This is a leading indicator and it runs past the reported "
              f"month on purpose** — the slide is about where meat and fuel "
              f"are heading before the monthly CPI catches up. The latest "
              f"week is **{wk['as_of']}**; the last week inside the reported "
              f"month was **{me}**. Whichever you quote, put its date on the "
              f"slide.")
            w()
        w()
        table(["Group", "YTD %", "YoY %", "Members"],
              [[g["label"], sg(g.get("ytd_pct")), sg(g.get("yoy_pct")),
                ", ".join(g.get("members", []))[:90]]
               for g in wk.get("groups", [])])
        table(["Product", "Latest ₮", "YTD %", "YoY %"],
              [[p["label"], n(p.get("current"), 0), sg(p.get("ytd_pct")),
                sg(p.get("yoy_pct"))] for p in wk["products"]])
        show = [p for p in wk["products"]
                if any(x in p["label"].casefold()
                       for x in ("үхрийн", "хонин", "хонь", "аи-", "дизел"))]
        for p in show[:6]:
            line(p["label"], list(zip(wk["weeks"], p["values"]))[-40:], 0)
        w()

    # ---------------- forecast
    weo = s.get("imf_weo") or {}
    if weo.get("indicators"):
        w("## 9. IMF WEO — PROJECTIONS, NOT MEASUREMENTS")
        w()
        w(f"{weo.get('country','')} · {weo.get('vintage','')} · "
          f"{weo.get('source','')}. Never present one of these as an "
          "outturn; label it 'IMF-ийн төсөөлөл'.")
        w()
        if weo.get("comparison"):
            w("Actual vs forecast — only where the basis genuinely matches:")
            w()
            table(["Indicator", "IMF", "Actual", "Gap pp", "Basis"],
                  [[c.get("label"), n(c.get("forecast"), 2),
                    n(c.get("actual"), 2), sg(c.get("gap")), c.get("basis")]
                   for c in weo["comparison"]])
        w("Projection path:")
        w()
        for i in weo["indicators"]:
            line(f"{i.get('label')} ({i.get('unit','')})",
                 sorted((i.get("values") or {}).items()), 2)
        w()

    # ---------------- budget + coverage
    w("## 10. Budget")
    w()
    for a in ("budget_revenue", "budget_expenditure", "budget_balance"):
        if k.get(a) is not None:
            w(f"- **{a}**: {n(k[a], 1)} bln MNT (cumulative)")
    w()
    w("## 11. Deck coverage")
    w()
    table(["Slide", "Content", "Status", "Source", "Note"],
          [[c.get("slide"), c.get("content"), c.get("status"),
            c.get("source"), (c.get("note") or "")[:60]]
           for c in s.get("coverage", [])])

    OUT.write_text("\n".join(L) + "\n", encoding="utf-8")
    print(f"wrote {OUT.name} "
          f"({OUT.stat().st_size / 1024:,.1f} KB, {len(L)} lines)")


if __name__ == "__main__":
    main()
