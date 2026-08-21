# AutoMEO — Monthly Deck Prompt

Run `python run_all.py`, complete the Balance of Payments download when
prompted, fill the two yellow cells on `Budget`, and run `python deck_digest.py`.
Open a new deck conversation and attach exactly:

1. last month's approved deck (the base file to edit in place);
2. `deck_digest.md`;
3. `Macro_Metrics_<year>_<month>.xlsx`.

The July 2026 deck is the canonical structure. Preserve its masters, layouts,
fonts, colours, chart style, slide order, and grouped objects. Do not create a
new presentation or use a different template. The deck has 21 slides:

| Slide | Disposition and source |
|---|---|
| 1 | Structural cover. Update the meeting date. |
| 2 | Manual geopolitics editorial. Cite every external claim in speaker notes. |
| 3 | Macro forecast snapshot, including the budget block. Source: `Summary`, `Forecast`, `Budget`. |
| 4 | Quarterly real GDP cumulative growth. Source: `GDP_Growth`. |
| 5 | GDP sector contributions. Source: `GDP_Sectors`. |
| 6 | NSO's monthly GDP estimate, “Эдийн засгийн сарын өсөлт”. Source: `GDP_Growth`; show its own vintage. |
| 7 | Inflation against target. Source: `Inflation`; include the CPI contribution breakdown when present. |
| 8 | Weekly meat and fuel prices. Source: `Weekly_Prices`; show the latest week and YTD change. |
| 9 | Inflation and interest rates. Source: `Inflation`; include policy, deposit, market new-loan, subsidised-loan rates, and their gap. |
| 10 | Trade balance and headline commodities. Source: `Trade`, `Trade_Sections`, `Commodities`. |
| 11 | Export volumes and year-on-year change. Source: `Commodities`. |
| 12 | Border prices: coal, copper, gold, and iron. Source: `Border_Prices`. |
| 13 | Balance of payments and reserves. Source: `External_Sector`; label the actual BoP span. |
| 14 | Banking-sector balance-sheet metrics. Source: `Banking`. |
| 15 | Loan growth structure. Source: `Loan_Detail`; include the bridge, sector contributions, and borrower growth where shown. |
| 16 | Structural divider for the sector section. |
| 17 | Industry I: mining and food processing. Source: `Sector_Charts`. |
| 18 | Industry II: light manufacturing and chemicals/metals. Source: `Sector_Charts`. |
| 19 | Trade by region. Source: `Sector_Charts`. |
| 20 | Services: hotel and food. Source: `Sector_Charts`. |
| 21 | Construction and transport. Source: `Sector_Charts`; state the quarter. |

## Non-negotiable rules

- Read `Data_Vintage` first. Any lagged source must carry its own month or
  quarter on the slide. Weekly prices may be later than the reported month,
  but must show their own latest week and remain labelled as a leading
  indicator.
- Every data-bearing number must trace to `deck_digest.md` or the workbook.
  Geopolitics claims must have cited web sources in speaker notes. Do not use
  remembered, inferred, or untraceable figures.
- Workbook percentages are fractions for Excel formatting. Digest percentages
  are written as percentage points. Convert exactly once when moving data into
  a chart.
- Never add a parent row to its children in deposits, BoP, loans, or sector
  contributions.
- Use NSO's wording “Эдийн засгийн сарын өсөлт”; do not call it “МЭЗҮИ”.
- IMF values are projections, not outturns. Label them “ОУВС-ийн төсөөлөл”.
- Keep contribution charts as time-series stacked columns with the headline as
  a line, touching bars, no gridlines, and the legend underneath. Preserve the
  July deck's colours and chart geometry.
- Do not leave stale text, dates, charts, or bitmap tables from the base deck.

## Build and QA

1. Inventory the attached July-style base deck and map every output slide to a
   source slide before editing. Wait for approval of the map.
2. Edit the base deck in place using its inherited text frames and charts.
   Preserve the Moody's/grouped-shape treatment if the approved base deck
   contains those objects in a future version.
3. Complete the slide-to-source table and an old-to-new mapping only if slides
   were added, removed, or merged.
4. Run `python check_deck.py <finished-deck>.pptx`. If it reports missing line
   markers, run `python check_deck.py <finished-deck>.pptx --fix`, use the
   repaired copy, and rerun the checker. Zero projector-breaking problems are
   required; explain any remaining notes.
5. Search the finished deck for prior-month strings, check meeting-date
   consistency, verify percentage axes, and inspect every slide in PowerPoint.

Add these facts to the prompt message:

```text
Meeting date:      <the real one>
Reported month:    <e.g. 2026 оны 7-р сар>
Geopolitics slide: <the angle, or "your call">
Yellow cells:      annual revenue plan <X> их наяд ₮, coal export plan <Y> сая тонн
```
