# -*- coding: utf-8 -*-
"""
Week 4 practice workbooks: Charts, Graphics and What-If.

build_*(path, done) as in xl_w1.py. Everything is fictional. Loan math uses
monthly payments: PMT(rate/12, years*12, -amount).
"""
import os, re, struct, zlib
from xlsxkit import *

SEASON = [("Solo kayak", [41, 58, 66, 49]), ("Tandem kayak", [22, 31, 37, 25]), ("Canoe", [30, 44, 51, 38]), ("Paddleboard", [18, 29, 34, 21])]
MONTHS = ["May", "Jun", "Jul", "Aug"]


def pmt(rate, n, pv):
    r = rate / 12
    return pv * r / (1 - (1 + r) ** -n)


def chart_alt(path, drawing_number, text):
    """Edit Alt Text on a chart (also works for a chart on its own chart sheet)."""
    def fn(parts):
        name = "xl/drawings/drawing%d.xml" % drawing_number
        d = parts[name].decode("utf8")
        d = re.sub(r'(<xdr:cNvPr [^>]*?)( descr="[^"]*")?(/?>)', lambda m: m.group(1) + ' descr="%s"' % text.replace('"', "&quot;") + m.group(3), d, count=1)
        parts[name] = d.encode("utf8")
    patch(path, fn)


def png(path, w=120, h=60, color=(14, 124, 134)):
    """A small solid-color PNG (the practice logo), made without any image library."""
    raw = b"".join(b"\x00" + bytes(color) * w for _ in range(h))
    def chunk(t, data):
        return struct.pack(">I", len(data)) + t + data + struct.pack(">I", zlib.crc32(t + data) & 0xFFFFFFFF)
    data = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(data)


def _season_sheet(wb, F, name="Season"):
    ws = wb.add_worksheet(name)
    ws.write("A1", "Rentals by month, 2027", F(bold=True, font_size=14))
    ws.write_row("A3", ["Boat"] + MONTHS, F(bold=True, bottom=1))
    for k, (b, nums) in enumerate(SEASON):
        ws.write(3 + k, 0, b)
        ws.write_row(3 + k, 1, nums)
    ws.set_column("A:A", 14)
    return ws


def _rentals_chart(wb, done, last_col):
    """The chart learners make in Lesson 1, in whatever state `done` says."""
    ch = wb.add_chart({"type": "column", "subtype": "clustered"})
    cols = "BCDE"[:last_col]
    if "switch" in done:   # series are the boats (rows); months along the axis
        for k, (b, _) in enumerate(SEASON):
            ch.add_series({"name": "=Season!$A$%d" % (4 + k), "categories": "=Season!$B$3:$%s$3" % cols[-1],
                           "values": "=Season!$B$%d:$%s$%d" % (4 + k, cols[-1], 4 + k),
                           "data_labels": {"value": True} if "labels" in done else None})
    else:                  # Excel's default here: series are the months (columns); boats along the axis
        for c in cols:
            ch.add_series({"name": "=Season!$%s$3" % c, "categories": "=Season!$A$4:$A$7", "values": "=Season!$%s$4:$%s$7" % (c, c),
                           "data_labels": {"value": True} if "labels" in done else None})
    if "title" in done:
        ch.set_title({"name": "Rentals by boat, May to August"})
    else:
        ch.set_title({"none": True})
    if "axes" in done:
        ch.set_x_axis({"name": "Month" if "switch" in done else "Boat"})
        ch.set_y_axis({"name": "Rentals"})
    ch.set_legend({"position": "bottom" if "legend" in done else "right"})
    if "layout" in done:
        ch.set_table()
    return ch


# ---------------------------------------------------------------------------
# Lesson 1 and 2: Create and format charts (D4.01 - D4.05)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    ws = _season_sheet(wb, F)
    chart_in = None
    if "chart" in done:
        ch = _rentals_chart(wb, done, 4 if "addseries" in done else 3)
        if "chartsheet" in done:
            cs = wb.add_chartsheet("Rentals Chart")
            cs.set_chart(ch)
            chart_in = "sheet"
        else:
            ws.insert_chart("G3", ch, {"description": "Solo kayaks were rented most every month, peaking in July." if "alt" in done else ""})
            chart_in = "season"
    ws.activate()
    wb.close()
    if chart_in:
        if "style" in done:
            chart_style(path, 1, 205)
        else:
            chart_style(path, 1, 201)
        if chart_in == "sheet" and "alt" in done:
            chart_alt(path, 1, "Solo kayaks were rented most every month, peaking in July.")


# ---------------------------------------------------------------------------
# Lesson 3: Graphics (D4.06)
# ---------------------------------------------------------------------------
def build_l3(path, done=frozenset(), logo=None):
    wb = new_book(path)
    F = Fmt(wb)
    ws = wb.add_worksheet("Flyer")
    ws.write("A1", "Paddle Day price list", F(bold=True, font_size=16))
    ws.write_row("A3", ["Boat", "Hour", "Half day"], F(bold=True, bottom=1))
    for k, (b, h, d) in enumerate([("Solo kayak", 15, 35), ("Tandem kayak", 22, 50), ("Canoe", 20, 45), ("Paddleboard", 12, 30)]):
        ws.write(3 + k, 0, b); ws.write_number(3 + k, 1, h, F(num_format='"$"#,##0')); ws.write_number(3 + k, 2, d, F(num_format='"$"#,##0'))
    ws.set_column("A:A", 14)
    alt = "alt" in done
    n_pics = 0
    if "logo" in done and logo:
        ws.insert_image("F2", logo, {"description": "Lakeshore Paddle Company logo" if alt else "", "object_position": 1})
        n_pics += 1
    if "textbox" in done:
        ws.insert_textbox("F8", "Free demos Saturday, 10 a.m. to 2 p.m.", {"width": 220, "height": 50, "fill": {"color": "#E2F0D9"}})
    if "icon" in done and logo:
        ws.insert_image("A10", logo, {"description": "Kayak icon" if alt else "", "x_scale": 0.4, "y_scale": 0.6})
        n_pics += 1
    wb.close()
    if "icon" in done and logo:
        make_icon(path, 1, n_pics - 1)


# ---------------------------------------------------------------------------
# Lesson 4: Themes, cell styles and templates (D4.07)
# ---------------------------------------------------------------------------
def build_l4(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    ws = wb.add_worksheet("Planner")
    ws.write("A1", "Trip planner", F(bold=True, font_size=16))
    ws.write_column("A3", ["Boats", "Hours", "Rate"], F(bold=True))
    ws.write_column("B3", [2, 3, 35])
    ws.write("A7", "Cost", F(bold=True)); ws.write_formula("B7", "=B3*B4*B5", F(num_format='"$"#,##0.00'), 210)
    ws.set_column("A:A", 12)
    wb.close()
    if "theme" in done:
        set_theme(path, "Retrospect", "Retrospect")
    if "style" in done:
        cell_style(path, 0, ["B3", "B4", "B5"], "Entry", builtin=False)


# ---------------------------------------------------------------------------
# Lesson 5: What-if analysis (D4.08 - D4.10)
# ---------------------------------------------------------------------------
RATES = [0.04, 0.05, 0.06, 0.07, 0.08]


def build_l5(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00')
    loan = wb.add_worksheet("Loan")
    loan.write("A1", "New boat loan", F(bold=True, font_size=14))
    price = 14931.39 if "goalseek" in done else 12000
    rows = [("Boat price", price, money), ("Down payment", 2000, money), ("Interest rate (yearly)", 0.06, F(num_format="0.0%")), ("Years", 5, None)]
    for k, (lab, v, f) in enumerate(rows):
        loan.write(2 + k, 0, lab); loan.write_number(2 + k, 1, v, f)
    loan.write("A8", "Monthly payment", F(bold=True))
    amount = price - 2000
    if "pmt" in done:
        loan.write_formula("B8", "=PMT(B5/12,B6*12,-(B3-B4))", money, round(pmt(0.06, 60, amount), 6))
    loan.write_row("D3", ["Rate", "Payment"], F(bold=True, bottom=1))
    if "datatable" in done:
        loan.write_formula("E4", "=B8", money, round(pmt(0.06, 60, amount), 6))
    for k, r in enumerate(RATES):
        loan.write_number(4 + k, 3, r, F(num_format="0%"))
        if "datatable" in done:
            loan.write_number(4 + k, 4, round(pmt(r, 60, amount), 6), money)
    loan.set_column("A:A", 22); loan.set_column("B:B", 12); loan.set_column("D:E", 11)
    season = wb.add_worksheet("Season")
    season.write("A1", "Paddle Day season plan", F(bold=True, font_size=14))
    srows = [("Price per rental", 40, money), ("Rentals per week", 60, None), ("Weeks", 14, None), ("Costs per week", 900, money)]
    for k, (lab, v, f) in enumerate(srows):
        season.write(2 + k, 0, lab); season.write_number(2 + k, 1, v, f)
    season.write("A8", "Season profit", F(bold=True))
    season.write_formula("B8", "=(B3*B4-B6)*B5", money, (40 * 60 - 900) * 14)
    season.set_column("A:A", 20); season.set_column("B:B", 12)
    if "summary" in done:
        ss = wb.add_worksheet("Scenario Summary")
        ss.write("B2", "Scenario Summary", F(bold=True, font_size=14))
        ss.write_row("C4", ["Current Values:", "Low", "Expected", "High"], F(bold=True))
        ss.write_row("B6", ["$B$4", 60, 35, 60, 90])
        ss.write_row("B8", ["$B$8", (40 * 60 - 900) * 14, (40 * 35 - 900) * 14, (40 * 60 - 900) * 14, (40 * 90 - 900) * 14])
    loan.activate()
    wb.close()
    if "scenarios" in done:
        add_scenarios(path, 1, [{"name": "Low", "inputs": {"B4": 35}}, {"name": "Expected", "inputs": {"B4": 60}}, {"name": "High", "inputs": {"B4": 90}}])
    if "datatable" in done:
        add_data_table(path, 0, "E5:E9", "B5")


# ---------------------------------------------------------------------------
# Skill Check 4: season sales
# ---------------------------------------------------------------------------
SALES = [("Rentals", [8200, 14100, 16800, 9300]), ("Lessons", [1500, 3200, 3900, 1800]), ("Gear", [900, 1700, 2100, 1100])]
QUARTERS = ["Spring", "Early summer", "Late summer", "Fall"]


def build_sc4(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0')
    ws = wb.add_worksheet("Sales")
    if "titlestyle" in done:
        ws.write("A1", "Sales by season", F(bold=True, font_size=15))
    else:
        ws.write("A1", "Sales by season")
    ws.write_row("A3", ["Source"] + QUARTERS, F(bold=True, bottom=1))
    for k, (src, nums) in enumerate(SALES):
        ws.write(3 + k, 0, src)
        for j, n in enumerate(nums):
            ws.write_number(3 + k, 1 + j, n, money)
    ws.set_column("A:A", 12); ws.set_column("B:E", 13)
    p = wb.add_worksheet("Price")
    p.write("A1", "Lesson price check", F(bold=True, font_size=14))
    price = 32.5 if "goalseek" in done else 25
    p.write_column("A3", ["Price per lesson", "Lessons sold", "Costs"])
    p.write_number("B3", price, money); p.write_number("B4", 120); p.write_number("B5", 1500, money)
    p.write("A7", "Profit", F(bold=True)); p.write_formula("B7", "=B3*B4-B5", money, price * 120 - 1500)
    p.set_column("A:A", 18)
    if "chart" in done:
        ch = wb.add_chart({"type": "column", "subtype": "clustered" if "stacked" not in done else "stacked"})
        n_series = 3 if "addseries" in done else 2
        if "switch" in done:
            for j, q in enumerate(QUARTERS):
                ch.add_series({"name": "=Sales!$%s$3" % "BCDE"[j], "categories": "=Sales!$A$4:$A$%d" % (3 + n_series),
                               "values": "=Sales!$%s$4:$%s$%d" % ("BCDE"[j], "BCDE"[j], 3 + n_series)})
        else:
            for k in range(n_series):
                ch.add_series({"name": "=Sales!$A$%d" % (4 + k), "categories": "=Sales!$B$3:$E$3", "values": "=Sales!$B$%d:$E$%d" % (4 + k, 4 + k),
                               "data_labels": {"value": True} if "labels" in done else None})
        ch.set_title({"name": "Sales by season"} if "title" in done else {"none": True})
        if "layout" in done:
            ch.set_table()
        cs = wb.add_chartsheet("Sales Chart") if "chartsheet" in done else None
        if cs:
            cs.set_chart(ch)
        else:
            ws.insert_chart("G3", ch)
    ws.activate()
    wb.close()
    if "chart" in done:
        chart_style(path, 1, 208 if "style" in done else 201)
        if "alt" in done:
            chart_alt(path, 1, "Rentals bring in most of the money every season, peaking in late summer.")
    if "titlestyle" in done:
        cell_style(path, 0, ["A1"], "Title")
    if "scenarios" in done:
        add_scenarios(path, 1, [{"name": "Busy", "inputs": {"B4": 160}}, {"name": "Slow", "inputs": {"B4": 80}}])


def build_all(base):
    d = os.path.join(base, "week-04")
    os.makedirs(d, exist_ok=True)
    build_l1(os.path.join(d, "E4-L1-charts.xlsx"))
    png(os.path.join(d, "logo.png"))
    build_l3(os.path.join(d, "E4-L3-graphics.xlsx"))
    build_l4(os.path.join(d, "E4-L4-planner.xlsx"))
    build_l5(os.path.join(d, "E4-L5-whatif.xlsx"))
    build_sc4(os.path.join(d, "E4-SC4-sales.xlsx"))
