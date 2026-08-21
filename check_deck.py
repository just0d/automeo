#!/usr/bin/env python3
"""
check_deck.py — catch the faults a converter will not show you.

    python check_deck.py Macroeconomic_Update_Jul_2026.pptx
    python check_deck.py deck.pptx --fix        # write deck_FIXED.pptx

Run this on the finished deck before sending it. LibreOffice, Preview and
every online viewer will render charts that PowerPoint leaves blank, so
"I converted it to PDF and it looked fine" is not a check.

What it looks for
-----------------
1. **Line charts with no `<c:marker>` on the `c:lineChart` element.**
   The element is optional in the schema and LibreOffice ignores its
   absence; PowerPoint draws the axes and nothing else. This is what emptied
   an entire slide of border-price charts in the July deck while every
   converter showed them correctly. `--fix` inserts it in its schema
   position and touches nothing else.
2. Series with no values, or mostly blank.
3. Charts whose series and categories are different lengths.
4. Charts missing their embedded-workbook relationship.
5. Percentage axes whose values look like they were never divided by 100
   (a 12% inflation series plotted as 1200%).
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

CHART = re.compile(r"^ppt/charts/chart(\d+)\.xml$")
TAG = re.compile(r"<(/?)c:(\w+)([^>]*?)(/?)>")


def _direct_children(body):
    """Tag names that are direct children of an already-unwrapped element."""
    out, depth = [], 0
    for m in TAG.finditer(body):
        close, tag, _attrs, selfc = m.groups()
        if depth == 0 and not close:
            out.append((tag, m.start()))
        if not close and not selfc:
            depth += 1
        elif close:
            depth -= 1
    return out


def _slide_of(z):
    """chart number -> slide number, from the slide relationships."""
    out = {}
    for name in z.namelist():
        m = re.match(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$", name)
        if not m:
            continue
        for ch in re.findall(r"chart(\d+)\.xml",
                             z.read(name).decode("utf8", "ignore")):
            out[int(ch)] = int(m.group(1))
    return out


def _needs_marker(xml):
    """(True, insertion offset) when a lineChart is missing its marker."""
    m = re.search(r"<c:lineChart>(.*?)</c:lineChart>", xml, re.S)
    if not m:
        return False, None
    kids = _direct_children(m.group(1))
    if any(t == "marker" for t, _ in kids):
        return False, None
    # schema order: ... ser*, dLbls, dropLines, hiLowLines, upDownBars,
    # marker, smooth, axId, axId
    at = next((off for t, off in kids if t in ("smooth", "axId")),
              len(m.group(1)))
    return True, m.start(1) + at


def _all_value_axes_are_percent(xml):
    """
    True only when EVERY value axis in the chart is percent-formatted.

    A combo chart with a level on one axis and growth on the other has both,
    and flagging the level series there would be noise -- a checker that
    cries wolf gets ignored, which is worse than not having one.
    """
    fmts = []
    for ax in re.findall(r"<c:valAx>(.*?)</c:valAx>", xml, re.S):
        m = re.search(r'<c:numFmt formatCode="([^"]*)"', ax)
        fmts.append(m.group(1) if m else "")
    return bool(fmts) and all("%" in f for f in fmts)


def check(path, fix=False):
    z = zipfile.ZipFile(path)
    slide = _slide_of(z)
    problems, notes, blank_marker = [], [], []
    n_line = n_chart = 0

    for name in sorted(z.namelist()):
        m = CHART.match(name)
        if not m:
            continue
        num = int(m.group(1))
        sl = slide.get(num, "?")
        xml = z.read(name).decode("utf8", "ignore")
        n_chart += 1

        if "<c:lineChart>" in xml:
            n_line += 1
            need, _ = _needs_marker(xml)
            if need:
                blank_marker.append(num)
                problems.append(
                    f"slide {sl}, chart{num}: line chart has no <c:marker> on "
                    "c:lineChart — PowerPoint will draw this BLANK")

        if "<c:externalData" not in xml:
            problems.append(f"slide {sl}, chart{num}: no embedded workbook "
                            "relationship")

        pct_only = _all_value_axes_are_percent(xml)
        for i, ser in enumerate(re.findall(r"<c:ser>(.*?)</c:ser>", xml, re.S)):
            nm = (re.findall(r"<c:v>([^<]*)</c:v>", ser) or [""])[0][:34]
            val = re.search(r"<c:val>(.*?)</c:val>", ser, re.S)
            cat = re.search(r"<c:cat>(.*?)</c:cat>", ser, re.S)
            vals = re.findall(r"<c:pt idx=\"\d+\"><c:v>([^<]*)</c:v>",
                              val.group(1)) if val else []
            ncat = len(re.findall(r"<c:pt idx=\"\d+\">", cat.group(1))) \
                if cat else 0
            if not vals:
                problems.append(f"slide {sl}, chart{num}, series {i} "
                                f"'{nm}': no values at all — this plots blank")
                continue
            if ncat and ncat != len(vals):
                notes.append(f"slide {sl}, chart{num}, series {i} "
                             f"'{nm}': {len(vals)} values against {ncat} "
                             "categories (deliberate gap, or a short series?)")
            nums = [float(v) for v in vals
                    if re.fullmatch(r"-?\d+(\.\d+)?([eE]-?\d+)?", v)]
            if pct_only and nums and max(abs(x) for x in nums) > 5:
                problems.append(
                    f"slide {sl}, chart{num}, series {i} '{nm}': every value "
                    f"axis is percent-formatted but the values reach "
                    f"{max(nums):.1f} — 12% will be drawn as 1200%. Divide "
                    "by 100.")

    print(f"{Path(path).name}: {n_chart} charts ({n_line} line charts)")
    if problems:
        print(f"\nWILL BE WRONG ON THE PROJECTOR — {len(problems)}:")
        for p in problems:
            print("  •", p)
    else:
        print("\nNothing that will break in PowerPoint.")
    if notes:
        print(f"\nWorth a look — {len(notes)}:")
        for p in notes:
            print("  ·", p)

    if fix and blank_marker:
        out = Path(path).with_name(Path(path).stem + "_FIXED.pptx")
        zin = zipfile.ZipFile(path)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zin.infolist():
                data = zin.read(item.filename)
                mm = CHART.match(item.filename)
                if mm and int(mm.group(1)) in blank_marker:
                    x = data.decode("utf8")
                    need, at = _needs_marker(x)
                    if need:
                        data = (x[:at] + '<c:marker val="1"/>'
                                + x[at:]).encode("utf8")
                zo.writestr(item, data)
        print(f"\nWrote {out.name} — {len(blank_marker)} chart(s) repaired, "
              "no series data touched.")
    elif fix:
        print("\nNothing to fix.")
    return 1 if problems else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    sys.exit(check(args[0], fix="--fix" in sys.argv))
