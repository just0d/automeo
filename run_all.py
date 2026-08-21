#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AutoMEO pipeline orchestrator: ingest -> process -> Excel.

Usage: python run_all.py [month] [year] [--offline]
No arguments = auto-detect the newest published month.
See the Data Dictionary and Technical Overview documents for details.
"""

import json
import logging
import subprocess
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PY = sys.executable

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s | %(levelname)-7s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("run_all")


def run_step(script, extra_args=None) -> bool:
    cmd = [PY, str(BASE_DIR / script)] + (extra_args or [])
    log.info("=" * 60)
    log.info("STEP: %s %s", script, " ".join(extra_args or []))
    log.info("=" * 60)
    start = time.time()
    try:
        proc = subprocess.run(cmd, cwd=str(BASE_DIR))
    except Exception as exc:
        log.error("Could not launch %s: %s", script, exc)
        return False
    ok = proc.returncode == 0
    log.info("%s finished in %.1fs (exit code %d).",
             script, time.time() - start, proc.returncode)
    return ok


def check_processed_output(started_at: float, period_args) -> bool:
    """
    Confirm the processor actually rewrote its output for THIS period.

    A step that exits 0 without doing its work -- a truncated script, an
    interpreter that dies before main(), a permissions problem on the output
    -- would otherwise let the Excel builder rebuild from the previous run's
    JSON and hand over a complete-looking workbook full of last month's
    numbers. Checking freshness and period turns that into a hard stop.
    """
    out = BASE_DIR / "processed_metrics.json"
    if not out.exists():
        log.error("data_processor.py exited 0 but %s does not exist.",
                  out.name)
        return False
    if out.stat().st_mtime < started_at - 1:
        log.error("%s was not rewritten by this run (last modified %s). The "
                  "processing step did nothing -- refusing to build the Excel "
                  "from stale data.", out.name,
                  time.strftime("%Y-%m-%d %H:%M:%S",
                                time.localtime(out.stat().st_mtime)))
        return False
    if period_args:
        try:
            k = json.loads(out.read_text(encoding="utf-8"))
            want_m = int(period_args[0])
            want_y = (int(period_args[1]) if len(period_args) > 1
                      else int(k.get("target_year", 0)))
            if (int(k.get("target_month", 0)), int(k.get("target_year", 0))) \
                    != (want_m, want_y):
                log.error("%s is for %s-%s but this run asked for %s-%s.",
                          out.name, k.get("target_year"),
                          k.get("target_month"), want_y, want_m)
                return False
        except Exception as exc:
            log.error("%s could not be read back (%s).", out.name, exc)
            return False
    return True


def check_dependencies() -> bool:
    """Verify required packages are importable BEFORE running any step."""
    import importlib
    needed = {"requests": "requests", "pandas": "pandas",
              "openpyxl": "openpyxl"}
    missing = []
    for module, pkg in needed.items():
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(pkg)
    if missing:
        log.error("Missing required package(s): %s", ", ".join(missing))
        log.error("Install them into THIS Python with:")
        log.error("    %s -m pip install %s", PY, " ".join(missing))
        return False
    return True


def main() -> None:
    log.info("################ MACRO PIPELINE v2: RUN ALL ################")
    t0 = time.time()

    if not check_dependencies():
        log.error("Aborting: install the packages above, then re-run.")
        sys.exit(1)

    offline = "--offline" in sys.argv[1:]
    period_args = [a for a in sys.argv[1:] if a.strip().isdigit()][:2]

    if offline:
        log.info("OFFLINE mode: skipping ingestion, using cached raw_files/.")
    elif not run_step("raw_ingestor.py", period_args):
        log.warning("Ingestion had a hard failure; continuing with cached "
                    "raw files (see raw_files/manifest.json).")

    proc_started = time.time()
    if not run_step("data_processor.py", period_args):
        log.error("Pipeline halted: data_processor.py failed.")
        sys.exit(1)
    if not check_processed_output(proc_started, period_args):
        log.error("Pipeline halted: the processing step produced no fresh "
                  "output, so the workbook would have been stale.")
        sys.exit(1)

    if not run_step("excel_builder.py"):
        log.error("Pipeline halted: excel_builder.py failed.")
        sys.exit(1)

    log.info("################ PIPELINE COMPLETE in %.1fs ################",
             time.time() - t0)
    log.info("Deliverable: Macro_Metrics_<year>_<month>.xlsx in %s", BASE_DIR)


if __name__ == "__main__":
    main()
