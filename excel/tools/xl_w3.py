# -*- coding: utf-8 -*-
"""
Week 3 practice workbooks: Tables and Data.

build_*(path, done) as in xl_w1.py. The rental list here is the same 32
rentals in every workbook; sorting and filtering are done in Python so the
saved files look exactly as Excel would leave them. Everything is fictional.
"""
import csv, datetime, os, random
from xlsxkit import *

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
RATES = {"Solo kayak": 35, "Tandem kayak": 50, "Canoe": 45, "Paddleboard": 30}


def rentals():
    """32 rentals: [Rental #, date (ISO), Day, Boat, Hours, Rate, Customer]."""
    rnd = random.Random(31)
    out = []
    for i in range(32):
        d = datetime.date(2027, 7, 1) + datetime.timedelta(days=(i * 5) // 4)
        boat = rnd.choice(list(RATES))
        out.append([2001 + i, d.isoformat(), DAYS[d.weekday()], boat, rnd.choice([1, 2, 2, 3, 4, 5]), RATES[boat], rnd.choice(["Member", "Visitor", "Visitor"])])
    return out


def serial(iso):
    d = datetime.date.fromisoformat(iso)
    return (d - datetime.date(1899, 12, 30)).days


HEAD = ["Rental #", "Date", "Day", "Boat", "Hours", "Rate", "Customer"]


def write_csv(path, rows, dupes=True):
    """The rental export as a .csv file, with the 4 duplicate rows a real export often has."""
    rows = list(rows)
    if dupes:
        for k in (3, 10, 17, 25):
            rows.insert(k + 1, list(rows[k]))
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEAD)
        for r in rows:
            w.writerow([r[0], datetime.date.fromisoformat(r[1]).strftime("%m/%d/%Y"), r[2], r[3], r[4], r[5], r[6]])


# ---------------------------------------------------------------------------
# Lesson 1: Bring data in (D3.01 text import, D3.02 web import)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    a = wb.add_worksheet("About")
    a.write("A1", "July rentals", F(bold=True, font_size=14))
    a.write("A3", "Import rentals-july.csv into this workbook with Data > Get Data > From File > From Text/CSV.")
    if "import" in done:
        rows = rentals()
        ws = wb.add_worksheet("rentals-july")
        data = [[r[0], serial(r[1]), r[2], r[3], r[4], r[5], r[6]] for r in rows]
        if "dedupe" not in done:
            for k in (3, 10, 17, 25):
                data.insert(k + 1, list(data[k]))
        ws.add_table(0, 0, len(data), 6, {"name": "rentals_july", "style": "Table Style Medium 7",
                                           "columns": [{"header": h} for h in HEAD], "data": data})
        ws.set_column("B:B", 11, F(num_format="m/d/yyyy"))
    wb.close()
    if "import" in done:
        add_connection(path, "Query - rentals-july")


def build_l1_web(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    a = wb.add_worksheet("About")
    a.write("A1", "Finger Lakes data", F(bold=True, font_size=14))
    a.write("A3", "Import a table from a web page with Data > Get Data > From Other Sources > From Web.")
    if "web" in done:
        ws = wb.add_worksheet("Lakes")
        lakes = [["Lake", "Length (mi)", "Max depth (ft)"], ["Seneca", 38, 618], ["Cayuga", 38, 435], ["Keuka", 20, 186], ["Canandaigua", 16, 276]]
        ws.add_table(0, 0, 4, 2, {"name": "Lakes", "columns": [{"header": h} for h in lakes[0]], "data": lakes[1:]})
    wb.close()
    if "web" in done:
        add_connection(path, "Query - Lakes")


# ---------------------------------------------------------------------------
# Lesson 2: Excel tables (D3.03 - D3.06)
# ---------------------------------------------------------------------------
def build_l2(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    date = F(num_format="m/d/yyyy")
    money = F(num_format='"$"#,##0.00')
    ws = wb.add_worksheet("Rentals")
    rows = rentals()
    data = [[r[0], serial(r[1]), r[2], r[3], r[4], r[5], r[6]] for r in rows]
    blank_at = 16                                  # an empty row in the middle of the list
    if "deleterow" not in done:
        data.insert(blank_at, [None] * 7)
    cols = list(HEAD) + (["Deposit"] if "deposit" in done else []) + (["Cost"] if "cost" in done else [])
    if "table" in done:
        columns = [{"header": h} for h in HEAD]
        if "total" in done:
            columns[3]["total_string"] = "Total"
            columns[4]["total_function"] = "sum"
            columns[5]["total_function"] = "average"
        if "deposit" in done:
            columns.append({"header": "Deposit"})
        if "cost" in done:
            columns.append({"header": "Cost", "formula": "=[@Hours]*[@Rate]", "format": money})
        extra = [[r[4] * 5 if "deposit" in done else None] if "deposit" in done else [] for r in data]
        tdata = []
        for k, r in enumerate(data):
            row = list(r) + ([10 if r[0] else None] if "deposit" in done else [])
            tdata.append(row)
        opts = {"name": "Rentals", "style": "Table Style Medium 2", "columns": columns, "data": tdata,
                "total_row": "total" in done, "banded_columns": "banded" in done}
        last_row = len(data) + (1 if "total" in done else 0)
        ws.add_table(0, 0, last_row, len(cols) - 1, opts)
        if "cost" in done:
            ccol = len(cols) - 1
            for k, r in enumerate(data):
                if r[0]:
                    ws.write_formula(1 + k, ccol, "=Rentals[[#This Row],[Hours]]*Rentals[[#This Row],[Rate]]", money, r[4] * r[5])
        if "total" in done:
            n = len(data) + 1
            ws.write_formula(n, 4, "=SUBTOTAL(109,Rentals[Hours])", None, sum(r[4] for r in data if r[0]))
            ws.write_formula(n, 5, "=SUBTOTAL(101,Rentals[Rate])", None, sum(r[5] for r in data if r[0]) / len([r for r in data if r[0]]))
    else:
        ws.write_row(0, 0, HEAD, head)
        for k, r in enumerate(data):
            if r[0]:
                ws.write_row(1 + k, 0, r)
    ws.set_column("B:B", 11, date); ws.set_column("D:D", 13); ws.set_column("G:G", 10)
    # Prices: an old table to convert back to a range
    p = wb.add_worksheet("Prices")
    prices = [["Boat", "Old rate"], ["Solo kayak", 30], ["Tandem kayak", 45], ["Canoe", 40], ["Paddleboard", 25]]
    if "convert" in done:
        p.write_row(0, 0, prices[0], head)
        for k, r in enumerate(prices[1:]):
            p.write_row(1 + k, 0, r)
    else:
        p.add_table(0, 0, 4, 1, {"name": "OldPrices", "style": "Table Style Light 1", "columns": [{"header": h} for h in prices[0]], "data": prices[1:]})
    p.set_column("A:A", 14)
    ws.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Lesson 3: Sort and filter, SORT() and UNIQUE() (D3.07, D3.08)
# ---------------------------------------------------------------------------
def build_l3(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00')
    ws = wb.add_worksheet("Rentals")
    rows = rentals()
    data = [[r[0], serial(r[1]), r[2], r[3], r[4], r[5], r[6]] for r in rows]
    if "sort" in done:
        data.sort(key=lambda r: (r[3].lower(), -r[1]))       # Boat A to Z, then Date newest first
    elif "sortone" in done:
        data.sort(key=lambda r: r[3].lower())
    ws.add_table(0, 0, len(data), 6, {"name": "Rentals", "style": "Table Style Medium 2", "columns": [{"header": h} for h in HEAD], "data": data})
    if "filter" in done:
        for k, r in enumerate(data):
            if r[2] not in ("Sat", "Sun"):
                ws.set_row(1 + k, None, None, {"hidden": True})
    ws.set_column("B:B", 11, F(num_format="m/d/yyyy")); ws.set_column("D:D", 13)
    s = wb.add_worksheet("Summary")
    s.write_row("A1", ["Boats", "", "Hours, largest first"], F(bold=True))
    if "unique" in done:
        uniq = []
        for r in data:
            if r[3] not in uniq:
                uniq.append(r[3])
        s.write_dynamic_array_formula("A2:A%d" % (1 + len(uniq)), "=_xlfn.UNIQUE(Rentals[Boat])", None, uniq[0])
        for k, u in enumerate(uniq[1:]):
            s.write(2 + k, 0, u)
    if "sortfn" in done:
        hours = sorted((r[4] for r in data), reverse=True)
        s.write_dynamic_array_formula("C2:C33", "=_xlfn._xlws.SORT(Rentals[Hours],,-1)", None, hours[0])
        for k, h in enumerate(hours[1:]):
            s.write_number(2 + k, 2, h)
    s.set_column("A:A", 14); s.set_column("C:C", 20)
    ws.activate()
    wb.close()
    if "filter" in done:
        table_filter(path, 1, 2, ["Sat", "Sun"])


# ---------------------------------------------------------------------------
# Lesson 4: Conditional formatting and sparklines (D3.09, D3.10)
# ---------------------------------------------------------------------------
TREND = [("Solo kayak", [41, 58, 66, 49]), ("Tandem kayak", [22, 31, 37, 25]), ("Canoe", [30, 44, 51, 38]), ("Paddleboard", [18, 29, 34, 21])]


def build_l4(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00')
    ws = wb.add_worksheet("Rentals")
    rows = rentals()
    data = [[r[0], serial(r[1]), r[2], r[3], r[4], r[5], r[6], r[4] * r[5]] for r in rows]
    ws.add_table(0, 0, len(data), 7, {"name": "Rentals", "style": "Table Style Medium 2",
                                      "columns": [{"header": h} for h in HEAD] + [{"header": "Cost", "format": money}], "data": data})
    n = len(data) + 1
    if "databars" in done:
        ws.conditional_format("E2:E%d" % n, {"type": "data_bar", "bar_color": "#5B9BD5"})
    if "highlight" in done:
        ws.conditional_format("H2:H%d" % n, {"type": "cell", "criteria": ">", "value": 100, "format": F(bg_color="#FFC7CE", font_color="#9C0006")})
    if "removeold" not in done:
        ws.conditional_format("F2:F%d" % n, {"type": "3_color_scale"})
    ws.set_column("B:B", 11, F(num_format="m/d/yyyy")); ws.set_column("D:D", 13); ws.set_column("H:H", 10)
    t = wb.add_worksheet("Trend")
    t.write("A1", "Rentals by month", F(bold=True, font_size=14))
    t.write_row("A2", ["Boat", "May", "Jun", "Jul", "Aug", "Trend"], F(bold=True, bottom=1))
    for k, (b, nums) in enumerate(TREND):
        t.write(2 + k, 0, b)
        t.write_row(2 + k, 1, nums)
        if "sparklines" in done:
            t.add_sparkline(2 + k, 5, {"range": "Trend!B%d:E%d" % (3 + k, 3 + k), "high_point": True})
    t.set_column("A:A", 14); t.set_column("F:F", 16)
    ws.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Skill Check 3: gear deliveries
# ---------------------------------------------------------------------------
GEAR = [("G-101", "Life jacket, adult", "Safety", 60, 39.0), ("G-102", "Life jacket, youth", "Safety", 25, 34.0),
        ("G-103", "Paddle, kayak", "Paddles", 40, 89.0), ("G-104", "Paddle, canoe", "Paddles", 30, 54.0),
        ("G-105", "Paddle, SUP", "Paddles", 20, 79.0), ("G-106", "Dry bag, 10 L", "Gear", 35, 18.5),
        ("G-107", "Dry bag, 20 L", "Gear", 20, 24.0), ("G-108", "Whistle", "Safety", 80, 3.5),
        ("G-109", "Bilge pump", "Safety", 12, 32.0), ("G-110", "Spray skirt", "Gear", 15, 65.0),
        ("G-111", "Rescue rope", "Safety", 10, 27.0), ("G-112", "Leash, SUP", "Gear", 22, 19.0)]
DELIVERIES = [("D-01", "G-103", 10), ("D-02", "G-108", 40), ("D-03", "G-101", 20), ("D-04", "G-110", 5), ("D-05", "G-106", 12)]


def write_deliveries_csv(path):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Delivery", "Item #", "Quantity"])
        for r in DELIVERIES:
            w.writerow(r)


def build_sc3(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    money = F(num_format='"$"#,##0.00')
    ws = wb.add_worksheet("Gear")
    head = ["Item #", "Item", "Category", "On hand", "Cost"] + (["Value"] if "value" in done else [])
    data = [list(g) for g in GEAR]
    if "sort" in done:
        data.sort(key=lambda r: (r[2].lower(), -r[3]))      # Category A to Z, then On hand largest first
    if "table" in done:
        cols = [{"header": h} for h in head[:5]]
        cols[4]["format"] = money
        if "value" in done:
            cols.append({"header": "Value", "formula": "=[@[On hand]]*[@Cost]", "format": money})
        if "total" in done:
            cols[0]["total_string"] = "Total"
            cols[3]["total_function"] = "sum"
        ws.add_table(0, 0, len(data) + (1 if "total" in done else 0), len(head) - 1,
                     {"name": "Gear" if "name" in done else "Table1", "style": "Table Style Medium 4" if "style" in done else "Table Style Light 1",
                      "columns": cols, "data": data, "total_row": "total" in done, "first_column": "firstcol" in done})
        if "value" in done:
            for k, r in enumerate(data):
                ws.write_formula(1 + k, 5, "=Gear[[#This Row],[On hand]]*Gear[[#This Row],[Cost]]", money, r[3] * r[4])
        if "total" in done:
            ws.write_formula(len(data) + 1, 3, "=SUBTOTAL(109,Gear[On hand])", None, sum(r[3] for r in data))
    else:
        ws.write_row(0, 0, head, F(bold=True))
        for k, r in enumerate(data):
            ws.write_row(1 + k, 0, r)
    if "filter" in done:
        for k, r in enumerate(data):
            if r[2] != "Safety":
                ws.set_row(1 + k, None, None, {"hidden": True})
    n = len(data) + 1
    if "cf" in done:
        ws.conditional_format("D2:D%d" % n, {"type": "cell", "criteria": "<", "value": 15, "format": F(bg_color="#FFC7CE")})
    if "removecf" not in done:
        ws.conditional_format("E2:E%d" % n, {"type": "data_bar"})
    ws.set_column("B:B", 20); ws.set_column("E:F", 10)
    s = wb.add_worksheet("Lists")
    s.write_row("A1", ["Categories", "", "Costs, highest first"], F(bold=True))
    if "unique" in done:
        cats = []
        for r in data:
            if r[2] not in cats:
                cats.append(r[2])
        s.write_dynamic_array_formula("A2:A%d" % (1 + len(cats)), "=_xlfn.UNIQUE(Gear[Category])", None, cats[0])
        for k, c in enumerate(cats[1:]):
            s.write(2 + k, 0, c)
    if "sortfn" in done:
        costs = sorted((r[4] for r in data), reverse=True)
        s.write_dynamic_array_formula("C2:C13", "=_xlfn._xlws.SORT(Gear[Cost],,-1)", None, costs[0])
        for k, c in enumerate(costs[1:]):
            s.write_number(2 + k, 2, c)
    tr = wb.add_worksheet("Trend")
    tr.write_row("A1", ["Item", "Apr", "May", "Jun", "Jul", "Trend"], F(bold=True))
    for k, (item, nums) in enumerate([("Life jackets", [70, 64, 52, 45]), ("Paddles", [95, 90, 88, 76]), ("Dry bags", [60, 58, 51, 40])]):
        tr.write(1 + k, 0, item); tr.write_row(1 + k, 1, nums)
        if "sparklines" in done:
            tr.add_sparkline(1 + k, 5, {"range": "Trend!B%d:E%d" % (2 + k, 2 + k), "type": "column"})
    if "old" not in done:
        o = wb.add_worksheet("Old")
        o.add_table(0, 0, 2, 1, {"name": "OldOrder", "columns": [{"header": "Item #"}, {"header": "Quantity"}], "data": [["G-101", 10], ["G-108", 25]]})
    else:
        o = wb.add_worksheet("Old")
        o.write_row(0, 0, ["Item #", "Quantity"]); o.write_row(1, 0, ["G-101", 10]); o.write_row(2, 0, ["G-108", 25])
    if "import" in done:
        d = wb.add_worksheet("E3-SC3-deliveries")
        d.add_table(0, 0, len(DELIVERIES), 2, {"name": "E3_SC3_deliveries", "columns": [{"header": h} for h in ["Delivery", "Item #", "Quantity"]], "data": [list(x) for x in DELIVERIES]})
    ws.activate()
    wb.close()
    if "import" in done:
        add_connection(path, "Query - E3-SC3-deliveries")
    if "filter" in done:
        table_filter(path, 1, 2, ["Safety"])


def write_season_csv(path):
    """A whole season of daily totals for Project 3: date, day, weather, boats out, revenue."""
    rnd = random.Random(2027)
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Day", "Weather", "High (F)", "Solo kayak", "Tandem kayak", "Canoe", "Paddleboard", "Revenue"])
        d = datetime.date(2027, 5, 29)
        while d <= datetime.date(2027, 9, 6):
            weather = rnd.choice(["Sunny", "Sunny", "Partly cloudy", "Cloudy", "Rain"])
            high = rnd.randint(64, 91)
            busy = (1.6 if d.weekday() >= 5 else 1.0) * (0.3 if weather == "Rain" else 1.0) * (high / 80)
            n = [max(0, int(rnd.randint(3, 12) * busy)) for _ in range(4)]
            rev = n[0] * 35 + n[1] * 50 + n[2] * 45 + n[3] * 30
            w.writerow([d.strftime("%m/%d/%Y"), DAYS[d.weekday()], weather, high] + n + [rev])
            d += datetime.timedelta(days=1)


def build_all(base):
    d = os.path.join(base, "week-03")
    os.makedirs(d, exist_ok=True)
    build_l1(os.path.join(d, "E3-L1-import.xlsx"))
    write_csv(os.path.join(d, "rentals-july.csv"), rentals())
    build_l1_web(os.path.join(d, "E3-L1-web.xlsx"))
    build_l2(os.path.join(d, "E3-L2-tables.xlsx"))
    build_l3(os.path.join(d, "E3-L3-sort.xlsx"))
    build_l4(os.path.join(d, "E3-L4-visual.xlsx"))
    build_sc3(os.path.join(d, "E3-SC3-gear.xlsx"))
    write_deliveries_csv(os.path.join(d, "E3-SC3-deliveries.csv"))
    write_season_csv(os.path.join(d, "lakeshore-season.csv"))


if __name__ == "__main__":
    rows = rentals()
    from collections import Counter
    print(Counter(r[3] for r in rows), "hours", sum(r[4] for r in rows), "max hours", max(r[4] for r in rows))
    print("first row", rows[0], "weekend", sum(1 for r in rows if r[2] in ("Sat", "Sun")))
    print("cost>100", sum(1 for r in rows if r[4] * r[5] > 100))
    print("gear max cost", max(g[4] for g in GEAR), "cats", sorted(set(g[2] for g in GEAR)))
