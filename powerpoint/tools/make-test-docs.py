# -*- coding: utf-8 -*-
"""
make-test-docs.py - answer keys and part-done presentations for testing the PowerPoint checks.

Run from the powerpoint/ folder:  python3 tools/make-test-docs.py
It writes tools/test-docs/*.pptx and tools/test-docs/expected.json. Each case
says which lesson file to read and, for each activity, whether each step should
pass (true), fail (false) or has no checks (null). check-tests.html compares
those with what the checker says.

The presentations come from the same builders as the starters (tools/pp_w*.py),
with some steps marked done. "all" means every checked step passes; "none"
means every checked step fails (a starter should never pass a step by itself).
"""
import json, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
ROOT = os.path.join(HERE, "..")
OUT = os.path.join(HERE, "test-docs")
os.makedirs(OUT, exist_ok=True)

import importlib
W = {n: importlib.import_module("pp_w%d" % n) for n in range(1, 6) if os.path.exists(os.path.join(HERE, "pp_w%d.py" % n))}


def steps_of(lesson, activity):
    """For each step of an activity in a lesson file: True if it has a check (or an answer), else False."""
    text = open(os.path.join(ROOT, lesson), encoding="utf8").read()
    m = re.search(r"^::: (?:drill|skillcheck|preflight) %s\n(.*?)^:::$" % re.escape(activity), text, re.M | re.S)
    if not m:
        raise KeyError(activity + " not in " + lesson)
    steps = re.split(r"^- (?:step|task):", m.group(1), flags=re.M)[1:]
    return [bool(re.search(r"^  check:", s, re.M)) for s in steps]


def expect(lesson, activity, spec):
    """'all' / 'none' / a list of True, False and None (one per step)."""
    checked = steps_of(lesson, activity)
    if spec == "all":
        return [True if c else None for c in checked]
    if spec == "none":
        return [False if c else None for c in checked]
    assert len(spec) == len(checked), (activity, len(spec), len(checked))
    return spec


CASES = []


def case(doc, lesson, acts, fileName=None):
    c = {"doc": doc, "lesson": lesson, "expect": {a: expect(lesson, a, s) for a, s in acts.items()}}
    if fileName:
        c["fileName"] = fileName
    CASES.append(c)


def t(name):
    return os.path.join(OUT, name)


# ---------------------------------------------------------------------------
# Week 1
# ---------------------------------------------------------------------------
def week1():
    w = W[1]
    L1, L2, L3, L4 = ("content/week-01/01-meet-powerpoint.md", "content/week-01/02-slides-sections.md",
                      "content/week-01/03-text-lists.md", "content/week-01/04-size-notes-print.md")
    SC1, P1 = "content/week-01/05-skill-check-sc1.md", "content/week-01/06-project-p1.md"
    # Lesson 1
    case("starters/week-01/P1-L1-welcome.pptx", L1, {"D1.01": [False], "D1.02": "none", "D1.03": [None]})
    case("starters/week-01/P1-L1-welcome.pptx", L1, {"D1.01": [True]}, fileName="Rivera_P1_Welcome.pptx")
    w.build_l1(t("P1-L1-solved.pptx"), {"outline", "notes", "sorter"})
    case("tools/test-docs/P1-L1-solved.pptx", L1, {"D1.02": "all"})
    w.build_l1(t("P1-L1-partial.pptx"), {"notes"})
    case("tools/test-docs/P1-L1-partial.pptx", L1, {"D1.02": [False, True, None, False]})
    # Lesson 2
    case("starters/week-01/P1-L2-orientation.pptx", L2, {"D1.04": "none", "D1.05": "none", "D1.06": "none"})
    every2 = {"newslides", "layout", "duplicate", "reuse", "outline", "sections", "movesection", "reorder", "hide"}
    w.build_l2(t("P1-L2-solved.pptx"), every2)
    case("tools/test-docs/P1-L2-solved.pptx", L2, {"D1.04": "all", "D1.05": "all", "D1.06": "all"})
    w.build_l2(t("P1-L2-partial.pptx"), {"newslides", "duplicate", "rename", "hide"})
    case("tools/test-docs/P1-L2-partial.pptx", L2, {"D1.04": [True, True, False, True], "D1.05": "none", "D1.06": [True, False, False, False, True]})
    w.build_l2(t("P1-L2-sections.pptx"), {"sections"})   # sections added, Gear not moved yet, Questions not last
    case("tools/test-docs/P1-L2-sections.pptx", L2, {"D1.06": [True, True, False, False, False]})
    # Lesson 3
    case("starters/week-01/P1-L3-text.pptx", L3, {"D1.07": "none", "D1.08": "none"})
    w.build_l3(t("P1-L3-solved.pptx"), {"titlefmt", "wordart", "columns", "bullets", "numbers"})
    case("tools/test-docs/P1-L3-solved.pptx", L3, {"D1.07": "all", "D1.08": "all"})
    w.build_l3(t("P1-L3-partial.pptx"), {"wordart", "numbers"})
    case("tools/test-docs/P1-L3-partial.pptx", L3, {"D1.07": [False, True, False], "D1.08": [False, False, True]})
    # Lesson 4
    case("starters/week-01/P1-L4-print.pptx", L4, {"D1.09": "none", "D1.10": "none", "D1.11": [False, False, None, None, None]})
    w.build_l4(t("P1-L4-solved.pptx"), {"size", "footer", "notes", "title"})
    case("tools/test-docs/P1-L4-solved.pptx", L4, {"D1.09": "all", "D1.10": "all", "D1.11": [True, True, None, None, None]})
    # Skill Check 1
    every_sc1 = {"size", "title", "newslide", "duplicate", "hide", "numbers_footer", "rename", "addsection", "movesection", "reorder", "italic", "columns", "numbers"}
    case("starters/week-01/P1-SC1-club.pptx", SC1, {"SC1": "none"})
    w.build_sc1(t("P1-SC1-solved.pptx"), every_sc1)
    case("tools/test-docs/P1-SC1-solved.pptx", SC1, {"SC1": "all"})
    w.build_sc1(t("P1-SC1-half.pptx"), {"size", "hide", "rename", "reorder", "numbers", "duplicate"})
    case("tools/test-docs/P1-SC1-half.pptx", SC1, {"SC1": [True, False, False, True, True, False, True, False, False, True, False, False, True]})
    # Project 1 preflight: the solved orientation deck has 13 slides, sections, no notes
    case("tools/test-docs/P1-L2-solved.pptx", P1, {"P1-check": [False, True, False, False, False, False]})


# ---------------------------------------------------------------------------
# Week 2
# ---------------------------------------------------------------------------
def week2():
    w = W[2]
    L1, L2, L3, L4 = ("content/week-02/01-themes-backgrounds.md", "content/week-02/02-slide-master.md",
                      "content/week-02/03-handouts-templates.md", "content/week-02/04-design-reads.md")
    SC2, P2 = "content/week-02/05-skill-check-sc2.md", "content/week-02/06-project-p2.md"
    # Lesson 1 (the other slides keep the theme background in the starter, so that step passes)
    case("starters/week-02/P2-L1-design.pptx", L1, {"D2.01": [False, None, False], "D2.02": [False, False, True]})
    w.build_l1(t("P2-L1-solved.pptx"), {"theme", "fonts", "picturebg", "solidbg"})
    case("tools/test-docs/P2-L1-solved.pptx", L1, {"D2.01": "all", "D2.02": "all"})
    w.build_l1(t("P2-L1-theme.pptx"), {"theme", "picturebg"})
    case("tools/test-docs/P2-L1-theme.pptx", L1, {"D2.01": [True, None, False], "D2.02": [True, False, True]})
    w.build_l1(t("P2-L1-applyall.pptx"), {"picturebg", "applyall"})     # the picture went to every slide (Apply to All)
    case("tools/test-docs/P2-L1-applyall.pptx", L1, {"D2.02": [True, False, False]})
    # Lesson 2
    case("starters/week-02/P2-L2-master.pptx", L2, {"D2.03": [False, False, False, None], "D2.04": "none"})
    w.build_l2(t("P2-L2-solved.pptx"), {"masterbg", "masterlogo", "titlefont", "newlayout", "editlayout", "uselayout"})
    case("tools/test-docs/P2-L2-solved.pptx", L2, {"D2.03": "all", "D2.04": "all"})
    w.build_l2(t("P2-L2-partial.pptx"), {"masterlogo", "newlayout"})
    case("tools/test-docs/P2-L2-partial.pptx", L2, {"D2.03": [False, True, False, None], "D2.04": [True, True, False, False]})
    # Lesson 3
    case("starters/week-02/P2-L3-handouts.pptx", L3, {"D2.05": [False, False, None], "D2.06": [False, None]})
    case("starters/week-02/P2-L3-handouts.pptx", L3, {"D2.06": [True, None]}, fileName="Rivera_Lakeshore.potx")
    case("starters/week-02/P2-L3-handouts.pptx", L3, {"D2.06": [False, None]}, fileName="Rivera_Lakeshore.pptx")
    w.build_l3(t("P2-L3-solved.pptx"), {"handout", "notesfooter"})
    case("tools/test-docs/P2-L3-solved.pptx", L3, {"D2.05": "all"})
    # Lesson 4
    case("starters/week-02/P2-L4-cluttered.pptx", L4, {"D2.07": "none"})
    w.build_l4(t("P2-L4-solved.pptx"), {"fonts", "contrast", "split"})
    case("tools/test-docs/P2-L4-solved.pptx", L4, {"D2.07": "all"})
    w.build_l4(t("P2-L4-partial.pptx"), {"contrast"})
    case("tools/test-docs/P2-L4-partial.pptx", L4, {"D2.07": [False, True, False]})
    # Skill Check 2
    every2 = {"masterbg", "titlefont", "masterlogo", "newlayout", "uselayout", "editlayout", "handout", "notesfooter", "picturebg", "solidbg"}
    case("starters/week-02/P2-SC2-kickoff.pptx", SC2, {"SC2": "none"})
    w.build_sc2(t("P2-SC2-solved.pptx"), every2)
    case("tools/test-docs/P2-SC2-solved.pptx", SC2, {"SC2": "all"})
    w.build_sc2(t("P2-SC2-half.pptx"), {"titlefont", "newlayout", "handout", "picturebg"})
    case("tools/test-docs/P2-SC2-half.pptx", SC2, {"SC2": [False, True, False, True, False, False, True, False, True, False]})
    # Project 2 preflight
    case("tools/test-docs/P2-L1-solved.pptx", P2, {"P2-check": [False, True, False, False, False]})
    case("tools/test-docs/P2-SC2-solved.pptx", P2, {"P2-check": [True, False, True, True, False]})


# ---------------------------------------------------------------------------
# Week 3
# ---------------------------------------------------------------------------
def week3():
    w = W[3]
    L1, L2, L3, L4 = ("content/week-03/01-pictures.md", "content/week-03/02-shapes-align.md",
                      "content/week-03/03-tables-charts.md", "content/week-03/04-smartart.md")
    SC3, P3 = "content/week-03/05-skill-check-sc3.md", "content/week-03/06-project-p3.md"
    # Lesson 1
    case("starters/week-03/P3-L1-pictures.pptx", L1, {"D3.01": "none", "D3.02": "none"})
    w.build_l1(t("P3-L1-solved.pptx"), {"insert", "resize", "crop", "style", "alt", "clipping"})
    case("tools/test-docs/P3-L1-solved.pptx", L1, {"D3.01": "all", "D3.02": "all"})
    w.build_l1(t("P3-L1-stretched.pptx"), {"insert", "stretch"})       # inserted, then pulled out of shape
    case("tools/test-docs/P3-L1-stretched.pptx", L1, {"D3.01": [True, False, False], "D3.02": [False, False, False]})
    w.build_l1(t("P3-L1-resized.pptx"), {"insert", "resize"})
    case("tools/test-docs/P3-L1-resized.pptx", L1, {"D3.01": [True, True, False]})
    # Lesson 2
    case("starters/week-03/P3-L2-shapes.pptx", L2, {"D3.03": "none", "D3.04": "none"})
    w.build_l2(t("P3-L2-solved.pptx"), {"callout", "size", "shapestyle", "textbox", "front", "align", "distribute", "group", "guides", "ink"})
    case("tools/test-docs/P3-L2-solved.pptx", L2, {"D3.03": "all", "D3.04": "all"})
    w.build_l2(t("P3-L2-partial.pptx"), {"callout", "align", "front"})   # aligned but not distributed; callout not sized
    case("tools/test-docs/P3-L2-partial.pptx", L2, {"D3.03": [True, False, False, False], "D3.04": [True, False, False, False, False]})
    # Lesson 3
    case("starters/week-03/P3-L3-data.pptx", L3, {"D3.05": "none", "D3.06": "none"})
    w.build_l3(t("P3-L3-solved.pptx"), {"table", "addrow", "delcol", "tablestyle", "chart", "charttype", "charttitle", "chartalt"})
    case("tools/test-docs/P3-L3-solved.pptx", L3, {"D3.05": "all", "D3.06": "all"})
    w.build_l3(t("P3-L3-partial.pptx"), {"table", "addrow", "chart"})
    case("tools/test-docs/P3-L3-partial.pptx", L3, {"D3.05": [True, True, False, False], "D3.06": [True, False, False, False]})
    # Lesson 4 (the starter has no graphics, so the alt-text step passes on its own; the last step has no check)
    case("starters/week-03/P3-L4-smartart.pptx", L4, {"D3.07": [False, False, False, True, None]})
    w.build_l4(t("P3-L4-solved.pptx"), {"smartart", "addstep", "layout", "colors", "smartalt"})
    case("tools/test-docs/P3-L4-solved.pptx", L4, {"D3.07": "all"})
    w.build_l4(t("P3-L4-partial.pptx"), {"smartart", "addstep"})
    case("tools/test-docs/P3-L4-partial.pptx", L4, {"D3.07": [True, True, False, False, None]})
    # Skill Check 3
    every3 = {"crop", "style", "alt", "shape", "size", "shapestyle", "front", "align", "distribute", "group", "table", "tablestyle", "chart", "charttitle", "smartart", "addshape"}
    case("starters/week-03/P3-SC3-report.pptx", SC3, {"SC3": "none"})
    w.build_sc3(t("P3-SC3-solved.pptx"), every3)
    case("tools/test-docs/P3-SC3-solved.pptx", SC3, {"SC3": "all"})
    w.build_sc3(t("P3-SC3-half.pptx"), {"crop", "alt", "shape", "group", "table", "chart", "smartart"})
    case("tools/test-docs/P3-SC3-half.pptx", SC3, {"SC3": [True, False, True, True, False, False, False, False, True, True, False, True, False, False]})
    # Project 3 preflight
    case("tools/test-docs/P3-SC3-solved.pptx", P3, {"P3-check": [True, True, True, True, True, False]})
    case("tools/test-docs/P3-L3-solved.pptx", P3, {"P3-check": [False, True, True, False, False, True]})


# ---------------------------------------------------------------------------
# Week 4
# ---------------------------------------------------------------------------
def week4():
    w = W[4]
    L1, L2, L3, L4 = ("content/week-04/01-transitions.md", "content/week-04/02-animations.md",
                      "content/week-04/03-media-3d.md", "content/week-04/04-links-zoom.md")
    SC4, P4 = "content/week-04/05-skill-check-sc4.md", "content/week-04/06-project-p4.md"
    # Lesson 1
    case("starters/week-04/P4-L1-transitions.pptx", L1, {"D4.01": "none", "D4.02": "none"})
    w.build_l1(t("P4-L1-solved.pptx"), {"fade", "3d", "option", "duration", "timing", "sound"})
    case("tools/test-docs/P4-L1-solved.pptx", L1, {"D4.01": "all", "D4.02": "all"})
    w.build_l1(t("P4-L1-partial.pptx"), {"fade", "3d"})
    case("tools/test-docs/P4-L1-partial.pptx", L1, {"D4.01": [True, True, False, False], "D4.02": "none"})
    w.build_l1(t("P4-L1-fadeonly.pptx"), {"fade"})                       # Fade on the title slide too: not 3D
    case("tools/test-docs/P4-L1-fadeonly.pptx", L1, {"D4.01": [True, False, False, False]})
    # Lesson 2
    case("starters/week-04/P4-L2-animations.pptx", L2, {"D4.03": "none", "D4.04": "none", "D4.05": "none"})
    w.build_l2(t("P4-L2-solved.pptx"), {"entrance", "emphasis", "exit", "after", "duration", "reorder", "path", "reverse"})
    case("tools/test-docs/P4-L2-solved.pptx", L2, {"D4.03": "all", "D4.04": "all", "D4.05": "all"})
    w.build_l2(t("P4-L2-partial.pptx"), {"entrance", "emphasis", "exit", "path"})   # effects in, timing not set, path still goes right
    case("tools/test-docs/P4-L2-partial.pptx", L2, {"D4.03": "all", "D4.04": [False, False, False], "D4.05": [True, False]})
    # Lesson 3
    case("starters/week-04/P4-L3-media.pptx", L3, {"D4.06": "none", "D4.07": "none", "D4.08": "none"})
    w.build_l3(t("P4-L3-solved.pptx"), {"video", "trim", "fade", "auto", "music", "recording", "model", "view", "turntable"})
    case("tools/test-docs/P4-L3-solved.pptx", L3, {"D4.06": "all", "D4.07": "all", "D4.08": "all"})
    w.build_l3(t("P4-L3-partial.pptx"), {"video", "auto", "model"})
    case("tools/test-docs/P4-L3-partial.pptx", L3, {"D4.06": [True, False, False, False], "D4.07": [False], "D4.08": [True, False, False]})
    # Lesson 4
    case("starters/week-04/P4-L4-links.pptx", L4, {"D4.09": "none", "D4.10": "none"})
    w.build_l4(t("P4-L4-solved.pptx"), {"weblink", "slidelink", "button", "summary", "section", "slidezoom", "return"})
    case("tools/test-docs/P4-L4-solved.pptx", L4, {"D4.09": "all", "D4.10": "all"})
    w.build_l4(t("P4-L4-partial.pptx"), {"weblink", "slidezoom"})          # slide zoom without Return to Zoom
    case("tools/test-docs/P4-L4-partial.pptx", L4, {"D4.09": [True, False, False], "D4.10": [False, False, False]})
    # Skill Check 4
    every4 = {"push", "pushoption", "after", "flyin", "listafter", "spin", "first", "path", "video", "auto", "trim", "weblink", "summary", "slidezoom"}
    case("starters/week-04/P4-SC4-kiosk.pptx", SC4, {"SC4": "none"})
    w.build_sc4(t("P4-SC4-solved.pptx"), every4)
    case("tools/test-docs/P4-SC4-solved.pptx", SC4, {"SC4": "all"})
    w.build_sc4(t("P4-SC4-half.pptx"), {"push", "flyin", "spin", "video", "weblink", "summary"})
    case("tools/test-docs/P4-SC4-half.pptx", SC4, {"SC4": [True, False, False, True, False, True, False, False, False, False, True, True, False]})
    # Project 4 preflight
    case("tools/test-docs/P4-SC4-solved.pptx", P4, {"P4-check": [True, True, True, True, True, False]})   # the video has no alt text
    case("tools/test-docs/P4-L1-solved.pptx", P4, {"P4-check": [False, False, True, False, False, True]})


# ---------------------------------------------------------------------------
# Week 5
# ---------------------------------------------------------------------------
def week5():
    w = W[5]
    L1, L2, L3, L4 = ("content/week-05/01-comments-coauthor.md", "content/week-05/02-compare-combine.md",
                      "content/week-05/03-deliver.md", "content/week-05/04-inspect-share.md")
    SC5, P5, CAP = "content/week-05/05-skill-check-sc5.md", "content/week-05/06-project-p5.md", "content/week-05/07-capstone.md"
    # Lesson 1 (the Rain policy slide is in the starter, so that check alone passes; its step also needs the partner's slide)
    case("starters/week-05/P5-L1-review.pptx", L1, {"D5.01": "none", "D5.02": [None, False, False]})
    w.build_l1(t("P5-L1-solved.pptx"), {"reply", "resolve", "delete", "add", "partner"})
    case("tools/test-docs/P5-L1-solved.pptx", L1, {"D5.01": "all", "D5.02": [None, True, True]})
    w.build_l1(t("P5-L1-partial.pptx"), {"reply", "partner", "deleted"})       # replied, not resolved; slide deleted, not restored
    case("tools/test-docs/P5-L1-partial.pptx", L1, {"D5.01": [True, False, False, False], "D5.02": [None, True, False]})
    # Lesson 2: your copy, the reviewer's copy, and the combined result
    case("starters/week-05/P5-L2-compare.pptx", L2, {"D5.03": [None, False, True]})
    case("starters/week-05/P5-L2-reviewer.pptx", L2, {"D5.03": [None, True, False]})
    w.build_l2(t("P5-L2-solved.pptx"), {"title", "price"})
    case("tools/test-docs/P5-L2-solved.pptx", L2, {"D5.03": [None, True, True]})
    w.build_l2(t("P5-L2-titleonly.pptx"), {"title"})
    case("tools/test-docs/P5-L2-titleonly.pptx", L2, {"D5.03": [None, False, True]})
    # Lesson 3
    case("starters/week-05/P5-L3-show.pptx", L3, {"D5.04": "none", "D5.05": "none"})
    w.build_l3(t("P5-L3-solved.pptx"), {"custom", "kiosk", "loop", "rehearse", "recording"})
    case("tools/test-docs/P5-L3-solved.pptx", L3, {"D5.04": "all", "D5.05": "all"})
    w.build_l3(t("P5-L3-partial.pptx"), {"loop", "rehearse"})              # loops, but not a kiosk show; timings ignored
    case("tools/test-docs/P5-L3-partial.pptx", L3, {"D5.04": [False, False], "D5.05": [True, False, None]})
    # Lesson 4
    case("starters/week-05/P5-L4-final.pptx", L4, {"D5.06": "none", "D5.07": [None, None, None]})
    w.build_l4(t("P5-L4-solved.pptx"), {"inspect", "offslide", "alt", "slidetitle", "header", "fonts", "password", "final"})
    case("tools/test-docs/P5-L4-solved.pptx", L4, {"D5.06": "all"})
    w.build_l4(t("P5-L4-partial.pptx"), {"inspect", "alt", "fonts"})        # off-slide note still there; title and header missing
    case("tools/test-docs/P5-L4-partial.pptx", L4, {"D5.06": [False, False, True, False, False]})
    # Skill Check 5
    every5 = {"reply", "resolve", "delete", "add", "custom", "loop", "rehearse", "personal", "offslide", "alt", "fonts"}
    case("starters/week-05/P5-SC5-team.pptx", SC5, {"SC5": "none"})
    w.build_sc5(t("P5-SC5-solved.pptx"), every5)
    case("tools/test-docs/P5-SC5-solved.pptx", SC5, {"SC5": "all"})
    w.build_sc5(t("P5-SC5-half.pptx"), {"reply", "delete", "custom", "personal", "fonts"})
    case("tools/test-docs/P5-SC5-half.pptx", SC5, {"SC5": [True, False, True, False, True, False, False, True, False, False, True]})
    # Project 5 preflight: the solved inspect deck has 4 slides and only one note
    case("tools/test-docs/P5-L4-solved.pptx", P5, {"P5-check": [False, False, True, True, True]})
    # Capstone practice exam
    every_cap = {"size", "titlefont", "title", "newslide", "hide", "footer", "section", "numbers", "link", "style", "alt", "shape", "align",
                 "row", "tablestyle", "charttype", "smartart", "fade", "anim", "custom"}
    case("starters/week-05/P5-CAP-exam.pptx", CAP, {"CAP": "none"})
    w.build_cap(t("P5-CAP-solved.pptx"), every_cap)
    case("tools/test-docs/P5-CAP-solved.pptx", CAP, {"CAP": "all"})
    w.build_cap(t("P5-CAP-half.pptx"), {"size", "title", "hide", "numbers", "alt", "row", "fade", "custom", "newslide", "anim"})
    case("tools/test-docs/P5-CAP-half.pptx", CAP, {"CAP": [True, False, True, True, True, False, False, True, False, False,
                                                         True, False, False, True, False, False, False, True, True, True]})


if __name__ == "__main__":
    for n in range(1, 6):
        fn = globals().get("week%d" % n)
        if fn and n in W:
            fn()
    about = ("Expected checker results. For each presentation: which lesson file to read, and for each activity, the expected result "
             "of every step (true = pass, false = fail, null = no checks). check-tests.html compares these with what the checker says.")
    with open(os.path.join(OUT, "expected.json"), "w") as f:
        json.dump({"about": about, "cases": CASES}, f, indent=1)
    print(len(CASES), "cases")
