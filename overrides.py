#!/usr/bin/env python3
"""
overrides.py — hand-set the identifiers the pipeline normally discovers.

Every source here is reached through an id that the pipeline works out at run
time: Mongolbank's report file ids, the survey ids behind them, NSO's table
ids, the deposit indicator ids. That discovery is deliberately layered and
usually survives a redesign — but if a portal changes enough, it will stop,
and the person who has to fix it may not be the person who wrote this.

So every one of those ids can be pinned by hand, in `overrides.json`, without
touching a line of code:

    {
      "mongolbank": {
        "report_ids": { "2026-06": { "bulletin": 5217 } }
      }
    }

Run `python overrides.py` to check the file parses and see what it will do.
The companion operations manual explains how to capture each id from the browser.

Nothing here is required. With no file, or an empty one, the pipeline behaves
exactly as it does today; overrides only ever *add* certainty.
"""
import json
import logging
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OVERRIDES_PATH = BASE_DIR / "overrides.json"
log = logging.getLogger("overrides")

# Shape of the file, with every key optional. Keep this in step with the
# example file and the troubleshooting doc.
SCHEMA = {
    "mongolbank": {
        "report_ids":     "period 'YYYY-MM' -> {report name: numeric file id}",
        "survey_ids":     "report name -> surveyid of its download page",
        "static_fallback": "report name -> a known-good id, any month",
        "api":            "base URLs: main, sublist, data, indicator",
    },
    "nso": {
        "tables": "save name -> {id, sector, subsector, extra}",
    },
    "deposits": {
        "report_id": "the report's numeric id (string)",
        "parent_id": "the parent id (integer)",
        "indicator_ids": "list of indicator id strings, in tree order",
    },
    "customs": {
        "list_api": "the statistic-news list endpoint",
    },
}


def load(path=None):
    """Read overrides.json. A missing file is normal; a broken one is loud."""
    p = Path(path) if path else OVERRIDES_PATH
    if not p.exists():
        return {}
    try:
        raw = p.read_text(encoding="utf-8")
    except Exception as exc:
        log.error("overrides: %s could not be read (%s) — ignoring it.",
                  p.name, exc)
        return {}
    # Allow // comments so the file can explain itself to the next operator.
    lines = [ln for ln in raw.splitlines()
             if not ln.strip().startswith("//")]
    try:
        data = json.loads("\n".join(lines) or "{}")
    except json.JSONDecodeError as exc:
        log.error("overrides: %s is not valid JSON (%s at line %d). "
                  "IGNORING THE WHOLE FILE — fix the syntax and re-run, or "
                  "the ids you set will silently not apply.",
                  p.name, exc.msg, exc.lineno)
        return {}
    if not isinstance(data, dict):
        log.error("overrides: %s must contain a JSON object.", p.name)
        return {}
    unknown = [k for k in data if k not in SCHEMA and not k.startswith("_")]
    if unknown:
        log.warning("overrides: ignoring unknown section(s) %s. Known "
                    "sections: %s.", ", ".join(unknown), ", ".join(SCHEMA))
    return data


def section(data, *path, default=None):
    """Fetch a nested key, tolerating any level being absent."""
    cur = data
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


def describe(data):
    """Human-readable summary of what the file will change. Logged on every
    run so an override can never be silently in force."""
    out = []
    for per, ids in (section(data, "mongolbank", "report_ids") or {}).items():
        for name, fid in (ids or {}).items():
            out.append(f"Mongolbank {name} for {per} pinned to id {fid}")
    for name, sid in (section(data, "mongolbank", "survey_ids") or {}).items():
        out.append(f"Mongolbank {name} surveyid pinned to {sid}")
    for name, fid in (section(data, "mongolbank",
                              "static_fallback") or {}).items():
        out.append(f"Mongolbank {name} fallback id set to {fid}")
    for k, v in (section(data, "mongolbank", "api") or {}).items():
        out.append(f"Mongolbank {k} API URL replaced")
    for name, spec in (section(data, "nso", "tables") or {}).items():
        bits = ", ".join(f"{a}={b}" for a, b in (spec or {}).items())
        out.append(f"NSO table {name}: {bits}")
    dep = section(data, "deposits") or {}
    if dep:
        out.append("deposit endpoint ids overridden ("
                   + ", ".join(sorted(dep)) + ")")
    if section(data, "customs", "list_api"):
        out.append("Customs list endpoint replaced")
    return out


def apply_and_log(data=None):
    """Load, then state plainly what is pinned. Call once at ingestion start."""
    data = load() if data is None else data
    notes = describe(data)
    if notes:
        log.warning("OVERRIDES ACTIVE — %d setting(s) are pinned by hand in "
                    "%s, not discovered:", len(notes), OVERRIDES_PATH.name)
        for n in notes:
            log.warning("    %s", n)
        log.warning("    Remove them once the portal works again, so the "
                    "pipeline goes back to finding ids itself.")
    return data


EXAMPLE = """{
  // AutoMEO overrides — every key is optional, delete what you do not need.
  // Lines starting with // are ignored. See the companion operations manual.

  "mongolbank": {
    // Pin a specific file id for one month. Use this when the run logs
    // "no listing covers YYYY-MM" and the id probe does not find it either.
    // Capture the id from DevTools: open the report's page, click the Excel
    // download, look at the request to /api/survey/data and read `id`.
    "report_ids": {
      // "2026-06": { "bulletin": 5217, "banking_balance_sheet": 5219 }
    },

    // The surveyid of a report's download page, if auto-discovery fails.
    // DevTools: open the page, find the POST to /api/survey/sublist,
    // read `surveyid` from the query string.
    "survey_ids": {
      // "banking_balance_sheet": 12
    },

    // A known-good id for any month, tried last when everything else fails.
    "static_fallback": {
      // "bulletin": 5217
    },

    // Only if Mongolbank moves its API.
    "api": {
      // "main":      "https://stat.mongolbank.mn/api/report/main",
      // "sublist":   "https://stat.mongolbank.mn/api/survey/sublist",
      // "data":      "https://stat.mongolbank.mn/api/survey/data",
      // "indicator": "https://stat.mongolbank.mn/api/indicator/data"
    }
  },

  "nso": {
    // Replace a table id when 1212.mn renames or republishes one.
    // DevTools: open the table on 1212.mn, find the POST to
    // /api/table-view, and read id / sector / subsector from the URL.
    "tables": {
      // "nso_mieg_raw": {
      //   "id": "DT_NSO_0500_001V5.px",
      //   "sector": "Economy%2C%20environment",
      //   "subsector": "National%20Accounts",
      //   "extra": ""
      // }
    }
  },

  "deposits": {
    // The interactive balance-sheet endpoint on stat.mongolbank.mn/finance.
    // "report_id": "86",
    // "parent_id": 10,
    // "indicator_ids": ["60564", "60565"]
  },

  "customs": {
    // "list_api": "https://www.gaali.mn/shared-api/api/statistic-news"
  }
}
"""


def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(levelname)-8s| %(message)s")
    if not OVERRIDES_PATH.exists():
        print(f"No {OVERRIDES_PATH.name} — the pipeline will discover every "
              f"id itself, which is the normal state.")
        print(f"To start one:  python overrides.py --init")
        return
    data = load()
    notes = describe(data)
    if not notes:
        print(f"{OVERRIDES_PATH.name} parses, and pins nothing.")
        return
    print(f"{OVERRIDES_PATH.name} parses. It pins {len(notes)} setting(s):")
    for n in notes:
        print("  •", n)


if __name__ == "__main__":
    import sys
    if "--init" in sys.argv:
        if OVERRIDES_PATH.exists():
            print(f"{OVERRIDES_PATH.name} already exists — not overwriting.")
        else:
            OVERRIDES_PATH.write_text(EXAMPLE, encoding="utf-8")
            print(f"Wrote {OVERRIDES_PATH.name}. Everything in it is "
                  f"commented out; uncomment what you need.")
    else:
        main()
