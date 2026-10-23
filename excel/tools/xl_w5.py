# -*- coding: utf-8 -*-
"""
Week 5 practice workbooks: Collaborate, Protect and Share, plus the practice exam.

build_*(path, done) as in xl_w1.py: `done` is a set of step names, and the
starter is build_x(path) with nothing done. Everything is fictional
(Lakeshore Paddle Company, Paddle Day, its staff and volunteers).
"""
import os
import random
from xlsxkit import *

# ---------------------------------------------------------------------------
# Shared numbers: the Paddle Day budget the reviewers argue about
# ---------------------------------------------------------------------------
ITEMS = ["Boat rental", "Food", "Signs", "Prizes", "Staff shirts"]
ESTIMATE = [600, 450, 120, 200, 300]
ACTUAL = [640, 410, 135, 180, 300]

# Three reviewers each sent back their own copy of the estimate. Lena added
# a line the others don't have, so Consolidate has to match rows by label.
REVIEWS = {
    "Priya": [("Boat rental", 650), ("Food", 450), ("Signs", 150), ("Prizes", 200), ("Staff shirts", 320)],
    "Marcus": [("Boat rental", 600), ("Food", 500), ("Signs", 120), ("Prizes", 250), ("Staff shirts", 300)],
    "Lena": [("Boat rental", 620), ("Food", 480), ("Signs", 140), ("Prizes", 220), ("Staff shirts", 310), ("Parking", 90)],
}
PRIYA_TOTAL = sum(v for _, v in REVIEWS["Priya"])     # 1770


def _budget_sheet(wb, F, name="Budget", with_actual=True):
    money = F(num_format='"$"#,##0')
    ws = wb.add_worksheet(name)
    ws.write("A1", "Paddle Day budget", F(bold=True, font_size=14))
    head = ["Item", "Estimate"] + (["Actual", "Difference"] if with_actual else [])
    ws.write_row("A3", head, F(bold=True, bottom=1))
    for k, item in enumerate(ITEMS):
        r = 3 + k
        ws.write(r, 0, item)
        ws.write_number(r, 1, ESTIMATE[k], money)
        if with_actual:
            ws.write_number(r, 2, ACTUAL[k], money)
            ws.write_formula(r, 3, "=C%d-B%d" % (r + 1, r + 1), money, ACTUAL[k] - ESTIMATE[k])
    ws.write("A9", "Total", F(bold=True, top=1))
    ws.write_formula("B9", "=SUM(B4:B8)", F(bold=True, top=1, num_format='"$"#,##0'), sum(ESTIMATE))
    if with_actual:
        ws.write_formula("C9", "=SUM(C4:C8)", F(bold=True, top=1, num_format='"$"#,##0'), sum(ACTUAL))
        ws.write_formula("D9", "=C9-B9", F(bold=True, top=1, num_format='"$"#,##0'), sum(ACTUAL) - sum(ESTIMATE))
    ws.set_column("A:A", 14); ws.set_column("B:D", 11)
    return ws


# ---------------------------------------------------------------------------
# Lesson 1: Comments and notes (D5.01)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    ws = _budget_sheet(wb, F)
    # Priya's threaded comment. Excel also saves a "[Threaded comment]" note copy for older versions.
    ws.write_comment("B4", "[Threaded comment]\nYour version of Excel lets you read this threaded comment.\n\nComment:\n    Is the dock fee in this estimate?", {"author": "Priya Shah"})
    if "comment" in done:
        ws.write_comment("C5", "[Threaded comment]\nYour version of Excel lets you read this threaded comment.\n\nComment:\n    The caterer's receipt is in the shared folder.", {"author": "Learner"})
    if "note" in done:
        ws.write_comment("B6", "Includes two banners.", {"author": "Learner"})
    if "delnote" not in done:
        ws.write_comment("D9", "Old: check the totals with Sam before May.", {"author": "Dana Ruiz"})
    wb.close()
    threads = [{"ref": "B4", "author": "Priya Shah", "text": "Is the dock fee in this estimate?",
                "replies": [("Learner", "Yes, $40 of it is the dock fee.")] if "reply" in done else [],
                "done": "resolve" in done}]
    if "comment" in done:
        threads.append({"ref": "C5", "author": "Learner", "text": "The caterer's receipt is in the shared folder."})
    add_threaded_comments(path, 0, threads)


# ---------------------------------------------------------------------------
# Lesson 2: Co-authoring (D5.02) and combining reviewers' work (D5.03)
# ---------------------------------------------------------------------------
SCHEDULE = [("9:00", "Set up the dock", "Priya"), ("10:00", "Free demos open", "Marcus"), ("12:00", "Lunch", "Everyone"),
            ("1:00", "Paddle race", "Lena"), ("2:00", "Free demos close", "Marcus"), ("3:00", "Clean up", "Everyone")]


def _schedule_sheet(wb, F):
    ws = wb.add_worksheet("Schedule")
    ws.write("A1", "Paddle Day schedule", F(bold=True, font_size=14))
    ws.write_row("A3", ["Time", "What", "Who"], F(bold=True, bottom=1))
    for k, row in enumerate(SCHEDULE):
        ws.write_row(3 + k, 0, row)
    ws.set_column("A:A", 8); ws.set_column("B:B", 20); ws.set_column("C:C", 10)
    return ws


def build_l2_schedule(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    ws = _schedule_sheet(wb, F)
    if "comment" in done:
        ws.write_comment("C7", "[Threaded comment]\nYour version of Excel lets you read this threaded comment.\n\nComment:\n    @Partner can you run the race?", {"author": "Learner"})
    wb.close()
    if "comment" in done:
        add_threaded_comments(path, 0, [{"ref": "C7", "author": "Learner", "text": "@Partner can you run the race?"}])


def build_priya(path):
    """Review-Priya.xlsx: the copy Priya sent back (a separate file learners link to)."""
    wb = new_book(path, author="Lakeshore Paddle Company")
    F = Fmt(wb)
    money = F(num_format='"$"#,##0')
    ws = wb.add_worksheet("Budget")
    ws.write("A1", "Paddle Day budget - Priya's review", F(bold=True, font_size=14))
    ws.write_row("A3", ["Item", "Estimate"], F(bold=True, bottom=1))
    for k, (item, v) in enumerate(REVIEWS["Priya"]):
        ws.write(3 + k, 0, item); ws.write_number(3 + k, 1, v, money)
    ws.write("A9", "Total", F(bold=True, top=1))
    ws.write_formula("B9", "=SUM(B4:B8)", F(bold=True, top=1, num_format='"$"#,##0'), PRIYA_TOTAL)
    ws.set_column("A:A", 14)
    wb.close()


def build_l2_combine(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0')
    s = wb.add_worksheet("Summary")
    s.write("A1", "Combining the reviews", F(bold=True, font_size=14))
    s.write("A3", "Priya's total (linked)")
    if "link" in done:
        s.write_formula("B3", "=[1]Budget!$B$9", money, PRIYA_TOTAL)
    s.set_column("A:A", 22); s.set_column("B:B", 11)
    for who, rows in REVIEWS.items():
        ws = wb.add_worksheet(who)
        ws.write_row("A1", ["Item", "Estimate"], F(bold=True, bottom=1))
        for k, (item, v) in enumerate(rows):
            ws.write(1 + k, 0, item); ws.write_number(1 + k, 1, v, money)
        ws.set_column("A:A", 14)
    c = wb.add_worksheet("Combined")
    c.write("A1", "All three reviews, added together", F(bold=True, font_size=14))
    if "consolidate" in done:
        # What Data > Consolidate (Sum, Top row + Left column labels) writes at A3.
        c.write("B3", "Estimate")
        labels = []
        for rows in REVIEWS.values():
            for item, _ in rows:
                if item not in labels:
                    labels.append(item)
        for k, item in enumerate(labels):
            c.write(3 + k, 0, item)
            c.write_number(3 + k, 1, sum(v for rows in REVIEWS.values() for i, v in rows if i == item), money)
    c.set_column("A:A", 14)
    wb.close()
    if "link" in done:
        external_link(path, "Review-Priya.xlsx", "Budget", {"B9": PRIYA_TOTAL})
    if "consolidate" in done:
        refs = [(who, "$A$1:$B$%d" % (1 + len(rows))) for who, rows in REVIEWS.items()]
        add_consolidate(path, 4, refs, labels=True)


# ---------------------------------------------------------------------------
# Lesson 3: Inspect (D5.04), accessibility (D5.05), protection (D5.06)
# ---------------------------------------------------------------------------
PRICES = [("Solo kayak", 15, 35), ("Tandem kayak", 22, 50), ("Canoe", 20, 45), ("Paddleboard", 12, 30)]


def build_l3_inspect(path, done=frozenset()):
    wb = new_book(path, author="Dana Ruiz")
    F = Fmt(wb)
    money = F(num_format='"$"#,##0')
    ws = wb.add_worksheet("Prices")
    ws.write("A1", "Rental prices", F(bold=True, font_size=14))
    ws.write_row("A3", ["Boat", "Hour", "Half day"], F(bold=True, bottom=1))
    for k, (b, h, d) in enumerate(PRICES):
        ws.write(3 + k, 0, b); ws.write_number(3 + k, 1, h, money); ws.write_number(3 + k, 2, d, money)
    ws.set_column("A:A", 14)
    if "comments" not in done:
        ws.write_comment("C6", "Raise to $50 next year? - Dana", {"author": "Dana Ruiz"})
        ws.write_comment("B4", "[Threaded comment]\nYour version of Excel lets you read this threaded comment.\n\nComment:\n    Check against the Lakeview price list.", {"author": "Dana Ruiz"})
    if "hidden" not in done:
        old = wb.add_worksheet("Old prices")
        old.write_row("A1", ["Boat", "Hour"])
        old.write_row("A2", ["Solo kayak", 12])
        old.write("A4", "Owner's cost notes: do not share.")
        old.hide()
    ws.activate()
    wb.close()
    if "comments" not in done:
        add_threaded_comments(path, 0, [{"ref": "B4", "author": "Dana Ruiz", "text": "Check against the Lakeview price list."}])
    if "personal" in done:
        remove_personal_info(path)
    else:
        set_last_modified_by(path, "Dana Ruiz")


def build_l3_access(path, done=frozenset()):
    wb = new_book(path, title="Rental prices" if "title" in done else "")
    F = Fmt(wb)
    money = F(num_format='"$"#,##0')
    ws = wb.add_worksheet("Prices" if "sheets" in done else "Sheet1")
    ws.add_table(0, 0, len(PRICES), 2, {"name": "PriceList" if "table" in done else "Table1", "style": "Table Style Medium 2",
                                        "columns": [{"header": "Boat"}, {"header": "Hour", "format": money}, {"header": "Half day", "format": money}],
                                        "data": [list(p) for p in PRICES]})
    ws.set_column("A:A", 14); ws.set_column("B:C", 10)
    cs = wb.add_worksheet("Chart" if "sheets" in done else "Sheet2")
    ch = wb.add_chart({"type": "column"})
    sheet = "Prices" if "sheets" in done else "Sheet1"
    ch.add_series({"name": "=%s!$B$1" % sheet, "categories": "=%s!$A$2:$A$5" % sheet, "values": "=%s!$B$2:$B$5" % sheet})
    ch.set_title({"name": "Price per hour"})
    cs.insert_chart("A1", ch, {"description": "Tandem kayaks cost the most per hour ($22); paddleboards cost the least ($12)." if "alt" in done else ""})
    ws.activate()
    wb.close()


BOAT_CHOICES = ["Solo kayak", "Tandem kayak", "Canoe", "Paddleboard"]


def build_l3_protect(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00')
    entry = F(locked=False, bg_color="#FFF2CC") if "unlock" in done else F(bg_color="#FFF2CC")
    ws = wb.add_worksheet("Sign-up")
    ws.write("A1", "Paddle Day boat sign-up", F(bold=True, font_size=14))
    ws.write_row("A3", ["Name", "Boat", "Hours", "Cost"], F(bold=True, bottom=1))
    sample = [("Avery", "Canoe", 2), ("Jordan", "Solo kayak", 1)]
    for r in range(4, 14):
        name, boat, hours = sample[r - 4] if r - 4 < len(sample) else ("", "", None)
        ws.write(r - 1, 0, name, entry)
        ws.write(r - 1, 1, boat, entry)
        if hours is None:
            ws.write_blank(r - 1, 2, None, entry)
        else:
            ws.write_number(r - 1, 2, hours, entry)
        cost = ""
        if boat:
            cost = hours * dict((b, h) for b, h, _ in PRICES)[boat]
        ws.write_formula(r - 1, 3, '=IF(B%d="","",C%d*VLOOKUP(B%d,Rates!$A$2:$B$5,2,FALSE))' % (r, r, r), money, cost)
    ws.set_column("A:B", 14); ws.set_column("D:D", 10)
    if "list" in done:
        ws.data_validation("B4:B13", {"validate": "list", "source": BOAT_CHOICES})
    if "whole" in done:
        ws.data_validation("C4:C13", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 8,
                                      "error_message": "Type a whole number of hours from 1 to 8."})
    if "protect" in done:
        ws.protect()
    rates = wb.add_worksheet("Rates")
    rates.write_row("A1", ["Boat", "Hour"], F(bold=True))
    for k, (b, h, _) in enumerate(PRICES):
        rates.write(1 + k, 0, b); rates.write_number(1 + k, 1, h, money)
    rates.set_column("A:A", 14)
    ws.activate()
    wb.close()
    if "structure" in done:
        protect_structure(path, password=False)


# ---------------------------------------------------------------------------
# Lesson 4: Web page, PDF and CSV (D5.07)
# ---------------------------------------------------------------------------
def build_l4(path, done=frozenset()):
    wb = new_book(path, title="Paddle Day schedule" if "title" in done else "")
    F = Fmt(wb)
    _schedule_sheet(wb, F)
    wb.close()


# ---------------------------------------------------------------------------
# Skill Check 5: volunteer roster
# ---------------------------------------------------------------------------
VOLUNTEERS = [("Avery", "Dock", "M"), ("Blake", "Demos", "L"), ("Casey", "Food", "S"), ("Devon", "Race", "XL"),
              ("Emery", "Dock", "M"), ("Finley", "Signs", "L"), ("Harper", "Demos", "M"), ("Jules", "Food", "S")]


def build_sc5(path, done=frozenset()):
    wb = new_book(path, author="Dana Ruiz", title="Volunteer roster" if "title" in done else "")
    F = Fmt(wb)
    entry = F(locked=False) if "unlock" in done else None
    ws = wb.add_worksheet("Roster")
    ws.add_table(0, 0, len(VOLUNTEERS), 3, {"name": "Volunteers" if "table" in done else "Table1", "style": "Table Style Medium 2",
                                            "columns": [{"header": h} for h in ["Name", "Job", "Shirt", "Hours"]],
                                            "data": [[n, j, s, 3] for n, j, s in VOLUNTEERS]})
    for k, (n, j, s) in enumerate(VOLUNTEERS):
        ws.write(1 + k, 2, s, entry)
        ws.write_number(1 + k, 3, 3, entry)
    ws.set_column("A:B", 12)
    ws.write_comment("C2", "[Threaded comment]\nYour version of Excel lets you read this threaded comment.\n\nComment:\n    Do we have shirt sizes for everyone?", {"author": "Marcus Lee"})
    if "note" in done:
        ws.write_comment("D9", "Jules can only stay until noon.", {"author": "Learner"})
    if "delnote" not in done:
        ws.write_comment("B5", "Old: ask Casey about the grill.", {"author": "Dana Ruiz"})
    if "list" in done:
        ws.data_validation("C2:C9", {"validate": "list", "source": ["S", "M", "L", "XL"]})
    if "protect" in done:
        ws.protect()
    hrs = wb.add_worksheet("Hours")
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "=Roster!$D$1", "categories": "=Roster!$A$2:$A$9", "values": "=Roster!$D$2:$D$9"})
    ch.set_title({"name": "Volunteer hours"})
    hrs.insert_chart("A1", ch, {"description": "Every volunteer is signed up for 3 hours." if "alt" in done else ""})
    lists = wb.add_worksheet("Lists" if "sheetname" in done else "Sheet3")
    lists.write_column("A1", ["Jobs", "Dock", "Demos", "Food", "Race", "Signs"])
    if "hidden" not in done:
        d = wb.add_worksheet("Draft")
        d.write("A1", "Draft list - includes phone numbers (made up)")
        d.hide()
    ws.activate()
    wb.close()
    add_threaded_comments(path, 0, [{"ref": "C2", "author": "Marcus Lee", "text": "Do we have shirt sizes for everyone?",
                                     "replies": [("Learner", "Yes, they're all in column C.")] if "reply" in done else [],
                                     "done": "resolve" in done}])
    if "personal" in done:
        remove_personal_info(path)
    else:
        set_last_modified_by(path, "Dana Ruiz")
    if "structure" in done:
        protect_structure(path, password=False)


# ---------------------------------------------------------------------------
# Capstone: practice exam (all weeks)
# ---------------------------------------------------------------------------
RATE = {"Solo kayak": 15, "Tandem kayak": 22, "Canoe": 20, "Paddleboard": 12}


def exam_orders():
    """20 orders in date order scrambled, the same every time."""
    rnd = random.Random(5150)
    boats = ["Solo kayak"] * 6 + ["Tandem kayak"] * 4 + ["Canoe"] * 6 + ["Paddleboard"] * 4
    rnd.shuffle(boats)
    days = list(range(20))
    rnd.shuffle(days)
    import datetime
    start = datetime.date(2027, 5, 29)
    rows = []
    for i, b in enumerate(boats):
        d = start + datetime.timedelta(days=days[i] * 4)
        serial = (d - datetime.date(1899, 12, 30)).days
        rows.append([3001 + i, serial, b, rnd.choice([1, 2, 2, 3, 4, 5])])
    return rows


def build_cap(path, done=frozenset()):
    wb = new_book(path, title="Season report" if "title" in done else "")
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00') if "currency" in done else None
    date = F(num_format="m/d/yyyy")
    rows = exam_orders()
    if "sort" in done:
        rows.sort(key=lambda r: (r[2], r[1]))            # Boat A to Z, then Date oldest first
    head = ["Order ID", "Date", "Boat", "Hours", "Rate", "Amount"]
    ws = wb.add_worksheet("Orders")
    n = len(rows)
    if "table" in done:
        cols = [{"header": h} for h in head]
        cols[1]["format"] = date
        cols[5]["formula"] = "=[@Hours]*[@Rate]"
        if money:
            cols[5]["format"] = money
        if "total" in done:
            cols[0]["total_string"] = "Total"
            cols[5]["total_function"] = "sum"
        ws.add_table(0, 0, n + (1 if "total" in done else 0), 5,
                     {"name": "Orders", "style": "Table Style Medium 9" if "style" in done else "Table Style Medium 2",
                      "columns": cols, "data": [r + [RATE[r[2]]] for r in rows], "total_row": "total" in done})
        for k, r in enumerate(rows):
            ws.write_formula(1 + k, 5, "=Orders[[#This Row],[Hours]]*Orders[[#This Row],[Rate]]", money, r[3] * RATE[r[2]])
        if "total" in done:
            ws.write_formula(n + 1, 5, "=SUBTOTAL(109,Orders[Amount])", money, sum(r[3] * RATE[r[2]] for r in rows))
    else:
        ws.write_row(0, 0, head, F(bold=True, bottom=1))
        for k, r in enumerate(rows):
            ws.write(1 + k, 0, r[0]); ws.write_number(1 + k, 1, r[1], date); ws.write(1 + k, 2, r[2])
            ws.write_number(1 + k, 3, r[3]); ws.write_number(1 + k, 4, RATE[r[2]])
            ws.write_formula(1 + k, 5, "=D%d*E%d" % (k + 2, k + 2), money, r[3] * RATE[r[2]])
    if "freeze" in done:
        ws.freeze_panes(1, 0)
    if "cf" in done:
        ws.conditional_format("D2:D%d" % (n + 1), {"type": "cell", "criteria": ">", "value": 3, "format": F(bg_color="#FFC7CE")})
    if "page" in done:
        ws.set_landscape(); ws.fit_to_pages(1, 0)
    if "footer" in done:
        ws.set_footer("&CPage &P")
    ws.set_column("A:A", 9); ws.set_column("B:B", 11); ws.set_column("C:C", 14); ws.set_column("F:F", 10)

    total = sum(r[3] * RATE[r[2]] for r in rows)
    s = wb.add_worksheet("Summary")
    s.write("A1", "Season summary", F(bold=True, font_size=14))
    labels = ["Total amount", "Number of orders", "Goal", "Sales tax"]
    s.write_column("A3", labels)
    m2 = F(num_format='"$"#,##0.00')
    if "sum" in done:
        s.write_formula("B3", "=SUM(Orders[Amount])", m2, total)
    if "counta" in done:
        s.write_formula("B4", "=COUNTA(Orders[Order ID])", None, n)
    if "if" in done:
        s.write_formula("B5", '=IF(B3>=800,"Goal met","Keep going")', None, "Goal met" if total >= 800 else "Keep going")
    if "tax" in done:
        s.write_formula("B6", "=B3*TaxRate", m2, round(total * 0.08, 6))
    s.write("D2", "Boats", F(bold=True))
    boats = []
    for r in rows:
        if r[2] not in boats:
            boats.append(r[2])
    if "unique" in done:
        s.write_dynamic_array_formula("D3:D%d" % (2 + len(boats)), "=_xlfn.UNIQUE(Orders[Boat])", None, boats[0])
        for k, b in enumerate(boats[1:]):
            s.write(3 + k, 3, b)
    s.write_row("A10", ["Boat", "Rentals"], F(bold=True, bottom=1))
    for k, b in enumerate(RATE):
        s.write(10 + k, 0, b); s.write_number(10 + k, 1, sum(1 for r in rows if r[2] == b))
    s.set_column("A:A", 18); s.set_column("B:B", 12); s.set_column("D:D", 14)
    if "chart" in done:
        ch = wb.add_chart({"type": "column"})
        ch.add_series({"name": "=Summary!$B$10", "categories": "=Summary!$A$11:$A$14", "values": "=Summary!$B$11:$B$14"})
        ch.set_title({"name": "Rentals by boat"} if "charttitle" in done else {"none": True})
        ch.set_legend({"none": True})
        s.insert_chart("F2", ch, {"description": "Solo kayaks and canoes were rented most, 6 times each." if "alt" in done else ""})

    rt = wb.add_worksheet("Rates")
    rt.write_row("A1", ["Sales tax rate", ""], F(bold=True))
    rt.write("A2", "Rate"); rt.write_number("B2", 0.08, F(num_format="0%"))
    rt.write("A5", "lpc-2027-season")
    if "upper" in done:
        rt.write_formula("B5", "=UPPER(A5)", None, "LPC-2027-SEASON")
    rt.set_column("A:B", 16)
    if "name" in done:
        wb.define_name("TaxRate", "=Rates!$B$2")
    ws.activate()
    wb.close()


def build_all(base):
    d = os.path.join(base, "week-05")
    os.makedirs(d, exist_ok=True)
    build_l1(os.path.join(d, "E5-L1-review.xlsx"))
    build_l2_schedule(os.path.join(d, "E5-L2-schedule.xlsx"))
    build_l2_combine(os.path.join(d, "E5-L2-combine.xlsx"))
    build_priya(os.path.join(d, "Review-Priya.xlsx"))
    build_l3_inspect(os.path.join(d, "E5-L3-inspect.xlsx"))
    build_l3_access(os.path.join(d, "E5-L3-access.xlsx"))
    build_l3_protect(os.path.join(d, "E5-L3-signup.xlsx"))
    build_l4(os.path.join(d, "E5-L4-schedule.xlsx"))
    build_sc5(os.path.join(d, "E5-SC5-roster.xlsx"))
    build_cap(os.path.join(d, "E5-CAP-exam.xlsx"))
