# -*- coding: utf-8 -*-
"""
Week 1 practice workbooks: The Working Spreadsheet.

Each build_* function makes one workbook. `done` is a set of step names:
build(path) makes the starter a learner opens; build(path, done={"rename", ...})
makes the same workbook with those steps finished, for testing the checks.
Everything is fictional (Lakeshore Paddle Company, its rentals and its staff).
"""
import random
from xlsxkit import *

BOATS = ["Kayak (single)", "Kayak (tandem)", "Canoe", "Paddleboard"]


def june_rentals():
    """40 June rentals, the same every time (a fixed random seed)."""
    rnd = random.Random(2027)
    boats = ["Kayak (single)"] * 12 + ["Kayak (tandem)"] * 9 + ["Canoe"] * 10 + ["Paddleboard"] * 9
    rnd.shuffle(boats)
    rows = []
    for i, b in enumerate(boats):
        day = 1 + (i * 3) // 4
        rows.append([1001 + i, "2027-06-%02d" % day, b, rnd.choice([1, 2, 2, 3, 4]), rnd.choice(["Member", "Visitor", "Visitor"])])
    return rows


def excel_date(iso):
    """'2027-06-05' -> Excel's date serial number."""
    import datetime
    d = datetime.date.fromisoformat(iso)
    return (d - datetime.date(1899, 12, 30)).days


# ---------------------------------------------------------------------------
# Lesson 1: Meet Excel  (D1.01 save, D1.02 sheets, D1.04 views, D1.05 find)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    wb = new_book(path, title="")
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    date = F(num_format="m/d/yyyy")
    names = ("June" if "rename" in done else "Sheet1")
    ws = wb.add_worksheet(names)
    if "addsheet" in done:
        jul = wb.add_worksheet("July")
        jul.write("A1", "July rentals start here.")
    summary = wb.add_worksheet("Summary")
    ws.write_row("A1", ["Rental #", "Date", "Boat", "Hours", "Customer"], head)
    rows = june_rentals()
    for i, r in enumerate(rows):
        boat = "Solo kayak" if ("replace" in done and r[2] == "Kayak (single)") else r[2]
        ws.write(1 + i, 0, r[0])
        ws.write_number(1 + i, 1, excel_date(r[1]), date)
        ws.write(1 + i, 2, boat)
        ws.write(1 + i, 3, r[3])
        ws.write(1 + i, 4, r[4])
    ws.set_column("A:A", 10); ws.set_column("B:B", 11); ws.set_column("C:C", 16); ws.set_column("E:E", 11)
    if "tabcolor" in done:
        ws.set_tab_color("#0E7C86")
    if "freeze" in done:
        ws.freeze_panes(1, 1)
    summary.write("A1", "June summary", F(bold=True, font_size=14))
    summary.write_row("A3", ["Rentals", len(rows)])
    summary.write_row("A4", ["Hours", sum(r[3] for r in rows)])
    summary.write("A5", "Most rented")
    summary.write("B5", "Solo kayak" if "replace" in done else "Kayak (single)")
    summary.set_column("A:A", 14); summary.set_column("B:B", 16)
    if "pagelayout" in done:
        summary.set_page_view(1)
    ws.activate()
    wb.close()
    if "window" in done:
        add_windows(path, 2)


# ---------------------------------------------------------------------------
# Lesson 2: Enter and Edit Data  (D1.06 fill, insert, delete; D1.07 paste special)
# ---------------------------------------------------------------------------
LOG_BOATS = ["Canoe", "Solo kayak", "Paddleboard", "Tandem kayak", "Canoe", "Solo kayak", "Solo kayak", "Paddleboard",
             "Canoe", "Solo kayak", "Tandem kayak", "Paddleboard", "Canoe", "Solo kayak", "Paddleboard", "Canoe",
             "Tandem kayak", "Solo kayak", "Paddleboard", "Canoe"]
PRICES = [("Solo kayak", 35), ("Tandem kayak", 50), ("Canoe", 45), ("Paddleboard", 30), ("Life jacket", 0), ("Dry bag", 5)]


def build_l2(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    money = F(num_format='"$"#,##0.00')
    # --- Fleet: month names to fill ---
    fleet = wb.add_worksheet("Fleet")
    fleet.write("A1", "2027 Fleet Plan", F(bold=True, font_size=14))
    fleet.write("A3", "Boat", head)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for i, m in enumerate(months if "months" in done else months[:1]):
        fleet.write(2, 1 + i, m, head)
    for i, b in enumerate(["Solo kayak", "Tandem kayak", "Canoe", "Paddleboard"]):
        fleet.write(3 + i, 0, b)
    fleet.set_column("A:A", 14)
    # --- Log: numbers to fill, rows and columns to insert, a cell to delete ---
    log = wb.add_worksheet("Log")
    top = 1 if "titlerow" in done else 0          # the inserted title row pushes everything down
    if top:
        log.write("A1", "June Rental Log", F(bold=True, font_size=14))
    headers = ["Rental #", "Date", "Boat", "Hours"] + (["Deposit", "Paid?"] if "columns" in done else []) + ["Notes"]
    log.write_row(top, 0, headers, head)
    notes_col = len(headers) - 1
    boats = list(LOG_BOATS)
    if "deletecell" not in done:
        boats.insert(9, "Tandem kayak")           # the extra entry that pushes the rest of the column down
    for i in range(20):
        r = top + 1 + i
        if "numbers" in done or i < 2:
            log.write_number(r, 0, 1001 + i)
        log.write_number(r, 1, excel_date("2027-06-%02d" % (1 + i // 2)), F(num_format="m/d/yyyy"))
        log.write_number(r, 3, 1 + (i * 7) % 4)
        if i in (3, 11):
            log.write(r, notes_col, "Late return" if i == 3 else "Paid at dock")
    for i, b in enumerate(boats):
        log.write(top + 1 + i, 2, b)
    log.set_column("A:A", 10); log.set_column("B:B", 11); log.set_column("C:C", 14); log.set_column(notes_col, notes_col, 14)
    # --- Prices: paste special ---
    pr = wb.add_worksheet("Prices")
    pr.write_row("A1", ["Boat", "Daily rate", "With tax", "Quote"], head)
    for i, (b, p) in enumerate(PRICES):
        pr.write(1 + i, 0, b)
        pr.write_number(1 + i, 1, p, money)
        pr.write_formula(1 + i, 2, "=B%d*1.08" % (2 + i), money, round(p * 1.08, 2))
        if "values" in done:
            pr.write_number(1 + i, 3, round(p * 1.08, 2), money)
    if "transpose" in done:
        pr.write_row("A10", ["Boat"] + [b for b, p in PRICES], head)
        pr.write("A11", "Daily rate", head)
        for i, (b, p) in enumerate(PRICES):
            pr.write_number(10, 1 + i, p, money)
    pr.set_column("A:A", 14); pr.set_column("B:G", 12)
    fleet.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Lessons 3 and 4: Format Cells; Page Setup and Printing
# ---------------------------------------------------------------------------
REPORT = [  # label, rentals, revenue, share, first rental (or None for a category row)
    ("Kayaks", None, None, None, None),
    ("Solo kayak", 52, 1820, 0.31, "2027-06-01"),
    ("Tandem kayak", 31, 1550, 0.26, "2027-06-02"),
    ("Canoes", None, None, None, None),
    ("Canoe", 38, 1710, 0.29, "2027-06-01"),
    ("Paddleboards", None, None, None, None),
    ("Paddleboard", 28, 840, 0.14, "2027-06-03"),
]
HEADERS = ["Boat type", "Rentals", "Revenue", "Share of revenue", "First rental", "Notes about this boat type"]


def _report_sheet(wb, F, ws, done, formatted_month):
    """One month's report sheet. `formatted_month` makes the July copy plain or finished."""
    title_fmt = F(bold=True, font_size=14)
    if "merge" in done:
        ws.merge_range("A1:F1", "June Rental Report" if formatted_month == "June" else "July Rental Report", F(bold=True, font_size=14, align="center"))
    else:
        ws.write("A1", ("June" if formatted_month == "June" else "July") + " Rental Report", title_fmt)
    ws.write("A2", "Lakeshore Paddle Company")
    head_props = {"bold": True}
    if "wrap" in done:
        head_props.update(text_wrap=True, valign="top")
    if "borderfill" in done:
        head_props.update(bottom=2, bg_color="#DDEBF7")
    for i, h in enumerate(HEADERS):
        ws.write(2, i, h, F(**head_props))
    money = F(num_format='"$"#,##0.00') if "currency" in done else None
    pct = F(num_format="0%") if "percent" in done else None
    dfmt = F(num_format="m/d/yyyy") if "dates" in done else None
    for i, (label, n, rev, share, first) in enumerate(REPORT):
        r = 3 + i
        indent = F(indent=1) if ("indent" in done and n is not None) else None
        ws.write(r, 0, label, indent if n is not None else F(bold=True))
        if n is None:
            continue
        ws.write_number(r, 1, n)
        ws.write_number(r, 2, rev, money)
        ws.write_number(r, 3, share, pct)
        ws.write_number(r, 4, excel_date(first), dfmt)
    total_fmt = F(bold=True, top=1) if "total" not in done else None
    ws.write(10, 0, "Total", total_fmt)
    ws.write_number(10, 1, 149, total_fmt)
    ws.write_number(10, 2, 5920, money if money else total_fmt)
    ws.write_number(10, 3, 1, pct if pct else total_fmt)
    ws.write_blank(10, 4, None, total_fmt or F())   # empty cells in the total row still carry its look
    ws.write_blank(10, 5, None, total_fmt or F())
    # Week labels to angle
    ws.write("A13", "Week", F(bold=True))
    rot = F(bold=True, rotation=45) if "rotate" in done else F(bold=True)
    for i in range(4):
        ws.write(12, 1 + i, "Week %d" % (i + 1), rot)
        ws.write(13, 1 + i, [30, 41, 37, 41][i])
    ws.write("A14", "Rentals")


def build_l3(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    june = wb.add_worksheet("June")
    july = wb.add_worksheet("July")
    _report_sheet(wb, F, june, done, "June")
    # July has the same layout but only plain headers (for Format Painter).
    july_done = {d for d in done if d not in ("wrap", "borderfill")}
    if "painter" in done:
        july_done |= {"wrap", "borderfill"}
    _report_sheet(wb, F, july, july_done - {"merge", "currency", "percent", "dates", "indent", "rotate", "total", "clear"}, "July")
    # A notes area with leftover formatting to clear.
    messy = None if "clear" in done else F(bold=True, italic=True, font_color="#C00000", bg_color="#FFFF00", border=1)
    june.write("H3", "Notes", messy)
    june.write("H4", "Prices went up June 1.", messy)
    june.write("H5", "Two canoes in repair.", messy)
    june.write("H6", "Count again in July.", messy)
    for ws in (june, july):
        ws.set_column("A:A", 14); ws.set_column("B:E", 11); ws.set_column("F:F", 18)
    june.activate()
    wb.close()
    if "title" in done:
        cell_style(path, 0, ["A1"], "Title")
    if "total" in done:
        cell_style(path, 0, ["A11", "B11", "C11", "D11", "E11", "F11"], "Total")


def build_l4(path, done=frozenset()):
    """The formatted report, ready for column widths, grouping and page setup."""
    fmt_done = {"merge", "wrap", "borderfill", "currency", "percent", "dates", "indent", "rotate", "painter", "clear"}
    wb = new_book(path, title="June Rental Report" if "props" in done else "", keywords="rentals, June" if "props" in done else "")
    F = Fmt(wb)
    june = wb.add_worksheet("June")
    july = wb.add_worksheet("July")
    _report_sheet(wb, F, june, fmt_done, "June")
    _report_sheet(wb, F, july, fmt_done, "July")
    for ws in (june, july):
        ws.set_column("A:A", 20 if "widths" in done else 14)
        ws.set_column("B:E", 11); ws.set_column("F:F", 18)
        if "heights" in done:
            ws.set_row(2, 32)
    if "landscape" in done:
        june.set_landscape()
    if "fit" in done:
        june.fit_to_pages(1, 0)
    if "printarea" in done:
        june.print_area("A1:F11")
    if "footer" in done:
        june.set_footer("&L&A&RPage &P")
    if "header" in done:
        june.set_header("&CLakeshore Paddle Company")
    june.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Skill Check 1: the gear inventory
# ---------------------------------------------------------------------------
GEAR = [("Lifejacket, adult", "Safety", 39.0, 60), ("Lifejacket, youth", "Safety", 34.0, 25), ("Paddle, kayak", "Paddles", 89.0, 40),
        ("Paddle, canoe", "Paddles", 54.0, 30), ("Paddle, SUP", "Paddles", 79.0, 20), ("Dry bag, 10 L", "Gear", 18.5, 35),
        ("Dry bag, 20 L", "Gear", 24.0, 20), ("Whistle", "Safety", 3.5, 80), ("Bilge pump", "Safety", 32.0, 12),
        ("Spray skirt", "Gear", 65.0, 15), ("Kayak seat pad", "Gear", 22.0, 18), ("Lifejacket, child", "Safety", 29.0, 20),
        ("Rescue rope", "Safety", 27.0, 10), ("Canoe yoke pad", "Gear", 16.0, 8), ("Leash, SUP", "Gear", 19.0, 22),
        ("Fin, SUP", "Gear", 25.0, 14), ("First aid kit", "Safety", 45.0, 6)]


def build_sc1(path, done=frozenset()):
    wb = new_book(path, title="Gear Inventory" if "props" in done else "")
    F = Fmt(wb)
    inv = wb.add_worksheet("Inventory")
    old = wb.add_worksheet("Old Counts")
    supplier = "supplier" in done
    if "merge" in done:
        inv.merge_range("A1:G1" if supplier else "A1:F1", "Gear Inventory - Spring 2027", F(bold=True, font_size=14, align="center"))
    else:
        inv.write("A1", "Gear Inventory - Spring 2027", F(bold=True, font_size=14))
    headers = ["Item #", "Item", "Category"] + (["Supplier"] if supplier else []) + ["Cost", "On hand", "Reorder at"]
    hp = {"bold": True, "bottom": 1}
    if "headers" in done:
        hp.update(text_wrap=True, align="center")
    inv.write_row("A3", headers, F(**hp))
    money = F(num_format='"$"#,##0.00') if "currency" in done else None
    off = 1 if supplier else 0
    for i, (item, cat, cost, onhand) in enumerate(GEAR):
        r = 3 + i
        if "fill" in done or i < 2:
            inv.write(r, 0, "G-%d" % (101 + i))
        name = item.replace("Lifejacket", "Life jacket") if "replace" in done else item
        inv.write(r, 1, name)
        inv.write(r, 2, cat)
        inv.write_number(r, 3 + off, cost, money)
        inv.write_number(r, 4 + off, onhand)
        if "paste" in done:
            inv.write_number(r, 5 + off, round(onhand * 0.25))
    inv.set_column("A:A", 8)
    inv.set_column("B:B", 30 if "width" in done else 14)
    inv.set_column("C:G", 11)
    if "freeze" in done:
        inv.freeze_panes(3, 0)
    if "landscape" in done:
        inv.set_landscape(); inv.fit_to_pages(1, 0); inv.print_area("A1:G20" if supplier else "A1:F20")
    if "footer" in done:
        inv.set_footer("&CPage &P")
    old.write_row("A1", ["Item #", "On hand", "Reorder at"], F(bold=True))
    for i, (item, cat, cost, onhand) in enumerate(GEAR):
        old.write(1 + i, 0, "G-%d" % (101 + i))
        old.write_number(1 + i, 1, onhand)
        old.write_formula(1 + i, 2, "=ROUND(B%d*0.25,0)" % (2 + i), None, round(onhand * 0.25))
    inv.activate()
    wb.close()
    if "style" in done:
        cell_style(path, 0, ["A1"], "Title")


# ---------------------------------------------------------------------------
def build_all(base):
    """Write every Week 1 starter into base/week-01."""
    import os
    d = os.path.join(base, "week-01")
    os.makedirs(d, exist_ok=True)
    build_l1(os.path.join(d, "E1-L1-rentals.xlsx"))
    build_l2(os.path.join(d, "E1-L2-log.xlsx"))
    build_l3(os.path.join(d, "E1-L3-report.xlsx"))
    build_l4(os.path.join(d, "E1-L4-report.xlsx"))
    build_sc1(os.path.join(d, "E1-SC1-inventory.xlsx"))
