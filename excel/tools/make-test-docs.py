# -*- coding: utf-8 -*-
"""
make-test-docs.py - answer keys and part-done workbooks for testing the Excel checks.

Run from the excel/ folder:  python3 tools/make-test-docs.py
It writes tools/test-docs/*.xlsx and tools/test-docs/expected.json. Each case
says which lesson file to read and, for each activity, whether each step should
pass (true), fail (false) or has no checks (null). check-tests.html compares
those with what the checker says.

The workbooks come from the same builders as the starters (tools/xl_w*.py),
with some steps marked done. "all" means every checked step passes; "none"
means every checked step fails (a starter should never pass a step by itself).
"""
import json, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
ROOT = os.path.join(HERE, "..")
OUT = os.path.join(HERE, "test-docs")
os.makedirs(OUT, exist_ok=True)

import xl_w1, xl_w2, xl_w3, xl_w4, xl_w5


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
    L1, L2, L3, L4 = ("content/week-01/01-meet-excel.md", "content/week-01/02-enter-edit.md",
                      "content/week-01/03-format-cells.md", "content/week-01/04-page-print.md")
    SC1 = "content/week-01/05-skill-check-sc1.md"
    # Lesson 1
    case("starters/week-01/E1-L1-rentals.xlsx", L1, {"D1.01": [False], "D1.02": "none", "D1.04": "none", "D1.05": [None, False]})
    case("starters/week-01/E1-L1-rentals.xlsx", L1, {"D1.01": [True]}, fileName="Rivera_E1_Rentals.xlsx")
    xl_w1.build_l1(t("E1-L1-solved.xlsx"), {"rename", "addsheet", "tabcolor", "freeze", "pagelayout", "window", "replace"})
    case("tools/test-docs/E1-L1-solved.xlsx", L1, {"D1.02": "all", "D1.04": "all", "D1.05": [None, True]})
    xl_w1.build_l1(t("E1-L1-partial.xlsx"), {"rename", "tabcolor"})
    case("tools/test-docs/E1-L1-partial.xlsx", L1, {"D1.02": [True, False, True], "D1.04": "none"})
    # Lesson 2
    case("starters/week-01/E1-L2-log.xlsx", L2, {"D1.06": "none", "D1.07": "none"})
    xl_w1.build_l2(t("E1-L2-solved.xlsx"), {"months", "titlerow", "numbers", "columns", "deletecell", "values", "transpose"})
    case("tools/test-docs/E1-L2-solved.xlsx", L2, {"D1.06": "all", "D1.07": "all"})
    xl_w1.build_l2(t("E1-L2-partial.xlsx"), {"months", "titlerow", "numbers"})   # cell not deleted, columns not added
    case("tools/test-docs/E1-L2-partial.xlsx", L2, {"D1.06": [True, True, True, False, False], "D1.07": "none"})
    xl_w1.build_l2(t("E1-L2-nodelete.xlsx"), {"deletecell"})                    # deleted, but no title row yet
    case("tools/test-docs/E1-L2-nodelete.xlsx", L2, {"D1.06": [False, False, False, False, False]})
    # Lesson 3
    case("starters/week-01/E1-L3-report.xlsx", L3, {"D1.08": "none", "D1.09": "none", "D1.10": [False, False, False, False]})
    full3 = {"currency", "percent", "dates", "borderfill", "merge", "wrap", "rotate", "indent", "title", "total", "painter", "clear"}
    xl_w1.build_l3(t("E1-L3-solved.xlsx"), full3)
    case("tools/test-docs/E1-L3-solved.xlsx", L3, {"D1.08": "all", "D1.09": "all", "D1.10": "all"})
    xl_w1.build_l3(t("E1-L3-partial.xlsx"), {"currency", "merge", "wrap", "borderfill"})
    case("tools/test-docs/E1-L3-partial.xlsx", L3, {"D1.08": [True, False, False, True], "D1.09": [True, True, False, False], "D1.10": [False, False, False, False]})
    # Lesson 4
    case("starters/week-01/E1-L4-report.xlsx", L4, {"D1.11": [False, False, None], "D1.12": "none", "D1.13": [False, None]})
    xl_w1.build_l4(t("E1-L4-solved.xlsx"), {"widths", "heights", "landscape", "fit", "printarea", "footer", "header", "props"})
    case("tools/test-docs/E1-L4-solved.xlsx", L4, {"D1.11": [True, True, None], "D1.12": "all", "D1.13": [True, None]})
    # Skill check 1
    case("starters/week-01/E1-SC1-inventory.xlsx", SC1, {"SC1": "none"})
    xl_w1.build_sc1(t("E1-SC1-solved.xlsx"), {"replace", "fill", "supplier", "paste", "merge", "headers", "currency", "style", "width", "freeze", "landscape", "footer", "props"})
    case("tools/test-docs/E1-SC1-solved.xlsx", SC1, {"SC1": "all"})


# ---------------------------------------------------------------------------
# Week 2
# ---------------------------------------------------------------------------
def week2():
    L1, L2, L3, L4 = ("content/week-02/01-formulas-references.md", "content/week-02/02-everyday-functions.md",
                      "content/week-02/03-names-links.md", "content/week-02/04-text-functions.md")
    SC2 = "content/week-02/05-skill-check-sc2.md"
    case("starters/week-02/E2-L1-orders.xlsx", L1, {"D2.01": "none", "D2.02": "none", "D2.03": "none", "D2.04": "none"})
    xl_w2.build_l1(t("E2-L1-solved.xlsx"), {"lines", "subtotal", "total", "absolute", "mixed", "fixtotal", "days", "showformulas"})
    case("tools/test-docs/E2-L1-solved.xlsx", L1, {"D2.01": "all", "D2.02": "all", "D2.03": "all", "D2.04": "all"})
    xl_w2.build_l1(t("E2-L1-partial.xlsx"), {"lines", "relative", "days"})   # no $ signs; days typed but total still typed
    case("tools/test-docs/E2-L1-partial.xlsx", L1, {"D2.01": [True, False, False], "D2.02": [False], "D2.04": [False, False, False]})
    case("starters/week-02/E2-L2-functions.xlsx", L2, {"D2.05": "none", "D2.06": "none", "D2.07": "none"})
    xl_w2.build_l2(t("E2-L2-solved.xlsx"), {"rowsum", "sum", "average", "max", "min", "count", "counta", "countblank", "if"})
    case("tools/test-docs/E2-L2-solved.xlsx", L2, {"D2.05": "all", "D2.06": "all", "D2.07": "all"})
    xl_w2.build_l2(t("E2-L2-partial.xlsx"), {"rowsum", "sum", "counta"})
    case("tools/test-docs/E2-L2-partial.xlsx", L2, {"D2.05": [True, True, False], "D2.06": [False, True, False]})
    case("starters/week-02/E2-L3-invoice.xlsx", L3, {"D2.08": [False, False, None], "D2.09": [False, False, False]})
    xl_w2.build_l3(t("E2-L3-solved.xlsx"), {"name", "usename", "linksheet", "linkweb", "removebroken"})
    case("tools/test-docs/E2-L3-solved.xlsx", L3, {"D2.08": [True, True, None], "D2.09": "all"})
    case("starters/week-02/E2-L4-members.xlsx", L4, {"D2.10": "none", "D2.11": "none", "D2.12": "none"})
    xl_w2.build_l4(t("E2-L4-solved.xlsx"), {"left", "mid", "right", "upper", "len", "lower", "concat", "textjoin", "sequence", "randbetween"})
    case("tools/test-docs/E2-L4-solved.xlsx", L4, {"D2.10": "all", "D2.11": "all", "D2.12": "all"})
    case("starters/week-02/E2-SC2-orders.xlsx", SC2, {"SC2": "none"})
    xl_w2.build_sc2(t("E2-SC2-solved.xlsx"), {"amount", "discount", "name", "sum", "average", "counta", "if", "left", "upper", "concat", "sequence", "link", "showformulas"})
    case("tools/test-docs/E2-SC2-solved.xlsx", SC2, {"SC2": "all"})


# ---------------------------------------------------------------------------
# Week 3
# ---------------------------------------------------------------------------
def week3():
    L1, L2, L3, L4 = ("content/week-03/01-bring-data-in.md", "content/week-03/02-excel-tables.md",
                      "content/week-03/03-sort-filter.md", "content/week-03/04-visual-summaries.md")
    SC3 = "content/week-03/05-skill-check-sc3.md"
    case("starters/week-03/E3-L1-import.xlsx", L1, {"D3.01": "none"})
    xl_w3.build_l1(t("E3-L1-solved.xlsx"), {"import", "dedupe"})
    case("tools/test-docs/E3-L1-solved.xlsx", L1, {"D3.01": "all"})
    xl_w3.build_l1(t("E3-L1-partial.xlsx"), {"import"})            # imported, duplicates still there
    case("tools/test-docs/E3-L1-partial.xlsx", L1, {"D3.01": [True, False]})
    case("starters/week-03/E3-L1-web.xlsx", L1, {"D3.02": "none"})
    xl_w3.build_l1_web(t("E3-L1-web-solved.xlsx"), {"web"})
    case("tools/test-docs/E3-L1-web-solved.xlsx", L1, {"D3.02": "all"})
    case("starters/week-03/E3-L2-tables.xlsx", L2, {"D3.03": "none", "D3.04": "none", "D3.05": "none", "D3.06": "none"})
    xl_w3.build_l2(t("E3-L2-solved.xlsx"), {"table", "banded", "total", "deposit", "deleterow", "cost", "convert"})
    case("tools/test-docs/E3-L2-solved.xlsx", L2, {"D3.03": "all", "D3.04": "all", "D3.05": "all", "D3.06": "all"})
    xl_w3.build_l2(t("E3-L2-partial.xlsx"), {"table", "total"})      # table made, empty row still in it
    case("tools/test-docs/E3-L2-partial.xlsx", L2, {"D3.03": "all", "D3.04": [False, True, False, False], "D3.05": "none"})
    case("starters/week-03/E3-L3-sort.xlsx", L3, {"D3.07": "none", "D3.08": "none"})
    xl_w3.build_l3(t("E3-L3-solved.xlsx"), {"sort", "filter", "unique", "sortfn"})
    case("tools/test-docs/E3-L3-solved.xlsx", L3, {"D3.07": "all", "D3.08": "all"})
    xl_w3.build_l3(t("E3-L3-partial.xlsx"), {"sortone"})               # sorted by boat only
    case("tools/test-docs/E3-L3-partial.xlsx", L3, {"D3.07": [False, False]})
    case("starters/week-03/E3-L4-visual.xlsx", L4, {"D3.09": "none", "D3.10": "none"})
    xl_w3.build_l4(t("E3-L4-solved.xlsx"), {"databars", "highlight", "removeold", "sparklines"})
    case("tools/test-docs/E3-L4-solved.xlsx", L4, {"D3.09": "all", "D3.10": "all"})
    case("starters/week-03/E3-SC3-gear.xlsx", SC3, {"SC3": "none"})
    xl_w3.build_sc3(t("E3-SC3-solved.xlsx"), {"import", "table", "name", "style", "firstcol", "total", "value", "sort", "filter", "unique", "sortfn", "cf", "removecf", "sparklines", "old"})
    case("tools/test-docs/E3-SC3-solved.xlsx", SC3, {"SC3": "all"})


def week4():
    W = "content/week-04/"
    L1, L2, L3, L4, L5 = (W + "01-create-charts.md", W + "02-format-charts.md", W + "03-graphics.md",
                          W + "04-themes-templates.md", W + "05-what-if.md")
    SC4 = W + "06-skill-check-sc4.md"
    logo = os.path.join(ROOT, "starters/week-04/logo.png")
    # Lessons 1 and 2 share one chart workbook.
    case("starters/week-04/E4-L1-charts.xlsx", L1, {"D4.01": "none", "D4.02": "none", "D4.03": "none"})
    case("starters/week-04/E4-L1-charts.xlsx", L2, {"D4.04": "none", "D4.05": "none"})
    every = {"chart", "chartsheet", "addseries", "switch", "title", "axes", "labels", "legend", "layout", "style", "alt"}
    xl_w4.build_l1(t("E4-L1-solved.xlsx"), every)
    case("tools/test-docs/E4-L1-solved.xlsx", L1, {"D4.01": "all", "D4.02": "all", "D4.03": "all"})
    case("tools/test-docs/E4-L1-solved.xlsx", L2, {"D4.04": "all", "D4.05": "all"})
    xl_w4.build_l1(t("E4-L1-partial.xlsx"), {"chart"})                 # chart inserted, still on Season
    case("tools/test-docs/E4-L1-partial.xlsx", L1, {"D4.01": "all", "D4.02": "none", "D4.03": "none"})
    case("tools/test-docs/E4-L1-partial.xlsx", L2, {"D4.04": "none", "D4.05": "none"})
    xl_w4.build_l1(t("E4-L1-onsheet.xlsx"), every - {"chartsheet"})     # everything, but the chart stays on Season
    case("tools/test-docs/E4-L1-onsheet.xlsx", L1, {"D4.01": "all", "D4.02": [False], "D4.03": "all"})
    case("tools/test-docs/E4-L1-onsheet.xlsx", L2, {"D4.04": "all", "D4.05": "all"})
    xl_w4.build_l1(t("E4-L1-noswitch.xlsx"), {"chart", "addseries", "title"})
    case("tools/test-docs/E4-L1-noswitch.xlsx", L1, {"D4.01": "all", "D4.02": "none", "D4.03": [True, False]})
    # Lesson 3: graphics
    case("starters/week-04/E4-L3-graphics.xlsx", L3, {"D4.06": "none"})
    xl_w4.build_l3(t("E4-L3-solved.xlsx"), {"logo", "textbox", "icon", "alt"}, logo=logo)
    case("tools/test-docs/E4-L3-solved.xlsx", L3, {"D4.06": "all"})
    xl_w4.build_l3(t("E4-L3-noalt.xlsx"), {"logo", "textbox", "icon"}, logo=logo)
    case("tools/test-docs/E4-L3-noalt.xlsx", L3, {"D4.06": [True, True, True, None, False]})
    # Lesson 4: theme, cell style, template
    case("starters/week-04/E4-L4-planner.xlsx", L4, {"D4.07": "none"})
    xl_w4.build_l4(t("E4-L4-solved.xlsx"), {"theme", "style"})
    case("tools/test-docs/E4-L4-solved.xlsx", L4, {"D4.07": "all"}, fileName="Ghidiu_Planner.xltx")
    case("tools/test-docs/E4-L4-solved.xlsx", L4, {"D4.07": [True, True, True, False]}, fileName="Ghidiu_Planner.xlsx")
    # Lesson 5: what-if
    case("starters/week-04/E4-L5-whatif.xlsx", L5, {"D4.08": "none", "D4.09": "none", "D4.10": "none"})
    xl_w4.build_l5(t("E4-L5-solved.xlsx"), {"pmt", "goalseek", "scenarios", "summary", "datatable"})
    case("tools/test-docs/E4-L5-solved.xlsx", L5, {"D4.08": "all", "D4.09": "all", "D4.10": "all"})
    xl_w4.build_l5(t("E4-L5-pmt.xlsx"), {"pmt"})                       # PMT done, nothing else yet
    case("tools/test-docs/E4-L5-pmt.xlsx", L5, {"D4.08": [True, False], "D4.09": "none", "D4.10": "none"})
    xl_w4.build_l5(t("E4-L5-noreport.xlsx"), {"pmt", "scenarios"})     # scenarios, no summary
    case("tools/test-docs/E4-L5-noreport.xlsx", L5, {"D4.09": [True, False]})
    # Skill Check 4
    case("starters/week-04/E4-SC4-sales.xlsx", SC4, {"SC4": "none"})
    xl_w4.build_sc4(t("E4-SC4-solved.xlsx"), {"chart", "chartsheet", "addseries", "switch", "title", "layout", "style", "alt",
                                              "titlestyle", "goalseek", "scenarios"})
    case("tools/test-docs/E4-SC4-solved.xlsx", SC4, {"SC4": "all"})
    xl_w4.build_sc4(t("E4-SC4-partial.xlsx"), {"chart"})
    case("tools/test-docs/E4-SC4-partial.xlsx", SC4, {"SC4": [True] + [False] * 10})
    # Project 4 preflight: the solved what-if file has scenarios and a data table but no chart or theme
    case("tools/test-docs/E4-L5-solved.xlsx", W + "07-project-p4.md", {"P4-check": [False, True, True, False, False]})


def week5():
    W = "content/week-05/"
    L1, L2, L3, L4 = W + "01-comments-notes.md", W + "02-coauthor-combine.md", W + "03-inspect-protect.md", W + "04-web-export.md"
    SC5, P5, CAP = W + "05-skill-check-sc5.md", W + "06-project-p5.md", W + "07-capstone.md"
    # Lesson 1: comments and notes
    case("starters/week-05/E5-L1-review.xlsx", L1, {"D5.01": "none"})
    xl_w5.build_l1(t("E5-L1-solved.xlsx"), {"note", "comment", "reply", "resolve", "delnote"})
    case("tools/test-docs/E5-L1-solved.xlsx", L1, {"D5.01": "all"})
    xl_w5.build_l1(t("E5-L1-replied.xlsx"), {"reply"})                  # replied but not resolved
    case("tools/test-docs/E5-L1-replied.xlsx", L1, {"D5.01": [False, False, True, False, False]})
    # Lesson 2: co-authoring and combining
    case("starters/week-05/E5-L2-schedule.xlsx", L2, {"D5.02": "none"})
    xl_w5.build_l2_schedule(t("E5-L2-schedule-solved.xlsx"), {"comment"})
    case("tools/test-docs/E5-L2-schedule-solved.xlsx", L2, {"D5.02": "all"})
    case("starters/week-05/E5-L2-combine.xlsx", L2, {"D5.03": "none"})
    xl_w5.build_l2_combine(t("E5-L2-combine-solved.xlsx"), {"link", "consolidate"})
    case("tools/test-docs/E5-L2-combine-solved.xlsx", L2, {"D5.03": "all"})
    xl_w5.build_l2_combine(t("E5-L2-combine-link.xlsx"), {"link"})
    case("tools/test-docs/E5-L2-combine-link.xlsx", L2, {"D5.03": [True, False, None]})
    # Lesson 3: inspect, accessibility, protection
    case("starters/week-05/E5-L3-inspect.xlsx", L3, {"D5.04": "none"})
    xl_w5.build_l3_inspect(t("E5-L3-inspect-solved.xlsx"), {"comments", "hidden", "personal"})
    case("tools/test-docs/E5-L3-inspect-solved.xlsx", L3, {"D5.04": "all"})
    xl_w5.build_l3_inspect(t("E5-L3-inspect-partial.xlsx"), {"comments"})
    case("tools/test-docs/E5-L3-inspect-partial.xlsx", L3, {"D5.04": [True, False, False, None]})
    case("starters/week-05/E5-L3-access.xlsx", L3, {"D5.05": "none"})
    xl_w5.build_l3_access(t("E5-L3-access-solved.xlsx"), {"sheets", "table", "alt", "title"})
    case("tools/test-docs/E5-L3-access-solved.xlsx", L3, {"D5.05": "all"})
    case("starters/week-05/E5-L3-signup.xlsx", L3, {"D5.06": "none"})
    xl_w5.build_l3_protect(t("E5-L3-signup-solved.xlsx"), {"unlock", "list", "whole", "protect", "structure"})
    case("tools/test-docs/E5-L3-signup-solved.xlsx", L3, {"D5.06": "all"})
    xl_w5.build_l3_protect(t("E5-L3-signup-locked.xlsx"), {"list", "whole", "protect"})   # protected without unlocking
    case("tools/test-docs/E5-L3-signup-locked.xlsx", L3, {"D5.06": [False, True, True, True, False]})
    # Lesson 4: web page, PDF, CSV
    case("starters/week-05/E5-L4-schedule.xlsx", L4, {"D5.07": "none"})
    xl_w5.build_l4(t("E5-L4-solved.xlsx"), {"title"})
    case("tools/test-docs/E5-L4-solved.xlsx", L4, {"D5.07": "all"})
    # Skill Check 5
    every5 = {"reply", "resolve", "note", "delnote", "unlock", "list", "protect", "sheetname", "title", "table", "alt", "hidden", "personal", "structure"}
    case("starters/week-05/E5-SC5-roster.xlsx", SC5, {"SC5": "none"})
    xl_w5.build_sc5(t("E5-SC5-solved.xlsx"), every5)
    case("tools/test-docs/E5-SC5-solved.xlsx", SC5, {"SC5": "all"})
    # Project 5 preflight: the solved roster still has its comments; the inspected price list has no structure protection
    case("tools/test-docs/E5-SC5-solved.xlsx", P5, {"P5-check": [False, True, True, True, True]})
    case("tools/test-docs/E5-L3-inspect-solved.xlsx", P5, {"P5-check": [True, True, True, False, False]})
    # Capstone practice exam
    every_cap = {"freeze", "table", "style", "total", "sort", "currency", "cf", "name", "sum", "counta", "if", "tax", "unique",
                 "upper", "chart", "charttitle", "alt", "page", "footer", "title"}
    case("starters/week-05/E5-CAP-exam.xlsx", CAP, {"CAP": "none"})
    xl_w5.build_cap(t("E5-CAP-solved.xlsx"), every_cap)
    case("tools/test-docs/E5-CAP-solved.xlsx", CAP, {"CAP": "all"})
    xl_w5.build_cap(t("E5-CAP-half.xlsx"), {"freeze", "table", "style", "sum", "counta", "chart"})
    case("tools/test-docs/E5-CAP-half.xlsx", CAP, {"CAP": [True, True, True, False, False, False, False, False, True, True,
                                                         False, False, False, False, True, False, False, False, False, False]})


if __name__ == "__main__":
    week1()
    for w in (2, 3, 4, 5):
        fn = globals().get("week%d" % w)
        if fn:
            fn()
    about = ("Expected checker results. For each workbook: which lesson file to read, and for each activity, the expected result "
             "of every step (true = pass, false = fail, null = no checks). check-tests.html compares these with what the checker says.")
    with open(os.path.join(OUT, "expected.json"), "w") as f:
        json.dump({"about": about, "cases": CASES}, f, indent=1)
    print(len(CASES), "cases")
