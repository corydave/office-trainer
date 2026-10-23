# -*- coding: utf-8 -*-
"""
Week 2 practice workbooks: Formulas That Work.

build_*(path, done) as in xl_w1.py. Formulas are written with the result Excel
would have saved (XlsxWriter's "value"), so the checker can read the answers.
Everything is fictional.
"""
import math
from xlsxkit import *

ITEMS = [("Solo kayak, half day", 2, 35), ("Tandem kayak, half day", 1, 50), ("Canoe, full day", 1, 70),
         ("Paddleboard, 2 hours", 3, 25), ("Dry bag rental", 2, 5)]


def money(F):
    return F(num_format='"$"#,##0.00')


# ---------------------------------------------------------------------------
# Lesson 1: Formulas and references (D2.01 - D2.04)
# ---------------------------------------------------------------------------
TAX_ITEMS = [("Solo kayak", 35), ("Tandem kayak", 50), ("Canoe", 45), ("Paddleboard", 30), ("Life jacket", 12), ("Dry bag", 5)]
GRID_PRICES = [35, 50, 45, 30, 25]
WEEKLY = [("Mon", 14), ("Tue", 11), ("Wed", 16), ("Thu", 19), ("Fri", 27), ("Sat", 41), ("Sun", 38)]


def order_totals():
    lines = [q * p for _, q, p in ITEMS]
    sub = sum(lines)
    return lines, sub, round((sub - 10) * 1.08, 2)


def build_l1(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    m = money(F)
    # --- Order: arithmetic and order of operations ---
    o = wb.add_worksheet("Order")
    o.write("A1", "Order #4417", F(bold=True, font_size=14))
    o.write_row("A3", ["Item", "Qty", "Price", "Line total"], head)
    lines, sub, total = order_totals()
    for i, (item, q, p) in enumerate(ITEMS):
        o.write(3 + i, 0, item); o.write_number(3 + i, 1, q); o.write_number(3 + i, 2, p, m)
        if "lines" in done:
            o.write_formula(3 + i, 3, "=B%d*C%d" % (4 + i, 4 + i), m, lines[i])
    o.write("C9", "Subtotal", F(bold=True))
    o.write("C10", "Coupon", F(bold=True)); o.write_number("D10", 10, m)
    o.write("C11", "Total with 8% tax", F(bold=True))
    if "subtotal" in done:
        o.write_formula("D9", "=D4+D5+D6+D7+D8", m, sub)
    if "total" in done:
        o.write_formula("D11", "=(D9-D10)*1.08", m, total)
    o.set_column("A:A", 24); o.set_column("B:D", 12)
    # --- Tax: one absolute reference ---
    t = wb.add_worksheet("Tax")
    t.write("A1", "Tax rate", F(bold=True)); t.write_number("B1", 0.08, F(num_format="0%"))
    t.write_row("A3", ["Item", "Price", "Tax"], head)
    for i, (item, p) in enumerate(TAX_ITEMS):
        t.write(3 + i, 0, item); t.write_number(3 + i, 1, p, m)
        if "absolute" in done:
            t.write_formula(3 + i, 2, "=B%d*$B$1" % (4 + i), m, round(p * 0.08, 2))
        elif "relative" in done:   # the classic mistake: filled without the $ signs
            t.write_formula(3 + i, 2, "=B%d*B%d" % (4 + i, 1 + i), m, 0)
    t.set_column("A:A", 14); t.set_column("B:C", 10)
    # --- Grid: mixed references ---
    g = wb.add_worksheet("Grid")
    g.write("A1", "Price by quantity", F(bold=True, font_size=14))
    g.write("A3", "Rate \\ Boats", head)
    for j in range(5):
        g.write_number(2, 1 + j, j + 1, head)
    for i, p in enumerate(GRID_PRICES):
        g.write_number(3 + i, 0, p, m)
        if "mixed" in done:
            for j in range(5):
                g.write_formula(3 + i, 1 + j, "=$A%d*%s$3" % (4 + i, col_letters(2 + j)), m, p * (j + 1))
    g.set_column("A:A", 13); g.set_column("B:F", 10)
    # --- Weekly: a typed total and a #DIV/0! to fix ---
    w = wb.add_worksheet("Weekly")
    w.write_row("A1", ["Day", "Rentals"], head)
    for i, (d, n) in enumerate(WEEKLY):
        w.write(1 + i, 0, d); w.write_number(1 + i, 1, n)
    total_w = sum(n for _, n in WEEKLY)
    w.write("A9", "Total", F(bold=True))
    if "fixtotal" in done:
        w.write_formula("B9", "=B2+B3+B4+B5+B6+B7+B8", None, total_w)
    else:
        w.write_number("B9", 160)          # typed by hand, and wrong
    w.write("A10", "Days open", F(bold=True))
    w.write("A11", "Average per day", F(bold=True))
    if "days" in done:
        w.write_number("B10", 7)
        w.write_formula("B11", "=B9/B10", F(num_format="0.0"), round((total_w if "fixtotal" in done else 160) / 7, 6))
    else:
        w.write_formula("B11", "=B9/B10", F(num_format="0.0"), "#DIV/0!")
    w.set_column("A:A", 16)
    o.activate()
    wb.close()
    if "showformulas" in done:
        show_formulas(path, 3, True)


# ---------------------------------------------------------------------------
# Lesson 2: Everyday functions (D2.05 - D2.07)
# ---------------------------------------------------------------------------
WEEKS = [("Solo kayak", [22, 25, 19, 31]), ("Tandem kayak", [12, 15, 9, 14]), ("Canoe", [18, 16, 21, 24]), ("Paddleboard", [9, 14, 11, 17])]
SIGNUPS = [("Avery Chen", 25), ("Sam Patel", None), ("", None), ("Riley Brooks", 25), ("Jordan Lee", "paid cash"),
           ("Casey Morgan", 20), ("", None), ("Taylor Kim", 25), ("Morgan Diaz", None), ("Drew Allen", 25),
           ("Quinn Harper", 15), ("", None), ("Jamie Fox", 25), ("Robin Hale", None), ("Skyler Moss", 25)]
RETURNS = [("1101", 12, 12), ("1102", 13, 14), ("1103", 15, 15), ("1104", 11, 12), ("1105", 16, 16), ("1106", 14, 13),
           ("1107", 17, 18), ("1108", 12, 12), ("1109", 15, 17), ("1110", 13, 13), ("1111", 16, 15), ("1112", 18, 19)]


def build_l2(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    s = wb.add_worksheet("Weeks")
    s.write("A1", "Rentals by week, June", F(bold=True, font_size=14))
    s.write_row("A3", ["Boat", "Week 1", "Week 2", "Week 3", "Week 4", "Total"], head)
    for i, (b, nums) in enumerate(WEEKS):
        s.write(3 + i, 0, b)
        for j, n in enumerate(nums):
            s.write_number(3 + i, 1 + j, n)
        if "rowsum" in done:
            s.write_formula(3 + i, 5, "=SUM(B%d:E%d)" % (4 + i, 4 + i), None, sum(nums))
    labels = ["Total", "Average", "Highest", "Lowest"]
    fns = ["SUM", "AVERAGE", "MAX", "MIN"]
    for k, lab in enumerate(labels):
        s.write(7 + k, 0, lab, F(bold=True))
        if fns[k].lower() in done:
            for j in range(4):
                col = [w[1][j] for w in WEEKS]
                val = {"SUM": sum(col), "AVERAGE": sum(col) / 4, "MAX": max(col), "MIN": min(col)}[fns[k]]
                s.write_formula(7 + k, 1 + j, "=%s(%s4:%s7)" % (fns[k], col_letters(2 + j), col_letters(2 + j)), F(num_format="0.0") if fns[k] == "AVERAGE" else None, val)
    s.set_column("A:A", 14)
    # --- Signups: COUNT family ---
    g = wb.add_worksheet("Signups")
    g.write("A1", "Paddle clinic sign-ups", F(bold=True, font_size=14))
    g.write_row("A3", ["Name", "Paid"], head)
    for i, (n, p) in enumerate(SIGNUPS):
        if n:
            g.write(3 + i, 0, n)
        if isinstance(p, int):
            g.write_number(3 + i, 1, p)
        elif p:
            g.write(3 + i, 1, p)
    g.write("D3", "Count", head)
    g.write("D4", "Payments recorded (numbers)"); g.write("D5", "Names entered"); g.write("D6", "Blank payment cells")
    paid_numbers = sum(1 for n, p in SIGNUPS if isinstance(p, int))
    names = sum(1 for n, p in SIGNUPS if n)
    blanks = sum(1 for n, p in SIGNUPS if p is None)
    if "count" in done:
        g.write_formula("E4", "=COUNT(B4:B18)", None, paid_numbers)
    if "counta" in done:
        g.write_formula("E5", "=COUNTA(A4:A18)", None, names)
    if "countblank" in done:
        g.write_formula("E6", "=COUNTBLANK(B4:B18)", None, blanks)
    g.set_column("A:A", 16); g.set_column("D:D", 28)
    # --- Returns: IF ---
    r = wb.add_worksheet("Returns")
    r.write("A1", "Return times (24-hour clock)", F(bold=True, font_size=14))
    r.write_row("A3", ["Rental", "Due back", "Returned", "Status"], head)
    for i, (rid, due, back) in enumerate(RETURNS):
        r.write(3 + i, 0, rid); r.write_number(3 + i, 1, due); r.write_number(3 + i, 2, back)
        if "if" in done:
            r.write_formula(3 + i, 3, '=IF(C%d>B%d,"Late","On time")' % (4 + i, 4 + i), None, "Late" if back > due else "On time")
    r.set_column("A:D", 11)
    s.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Lesson 3: Names and links (D2.08, D2.09)
# ---------------------------------------------------------------------------
DEPOSITS = [40, 25, 60, 25, 40, 35, 50]


def build_l3(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    m = money(F)
    inv = wb.add_worksheet("Invoice")
    det = wb.add_worksheet("Details")
    summ = wb.add_worksheet("Summary")
    inv.write("A1", "Invoice", F(bold=True, font_size=14))
    inv.write("E2", "Tax rate", F(bold=True)); inv.write_number("F2", 0.08, F(num_format="0%"))
    inv.write_row("A4", ["Item", "Qty", "Price", "Amount"], head)
    lines, sub, _ = order_totals()
    for i, (item, q, p) in enumerate(ITEMS):
        inv.write(4 + i, 0, item); inv.write_number(4 + i, 1, q); inv.write_number(4 + i, 2, p, m)
        inv.write_formula(4 + i, 3, "=B%d*C%d" % (5 + i, 5 + i), m, lines[i])
    inv.write("C10", "Subtotal", F(bold=True)); inv.write_formula("D10", "=SUM(D5:D9)", m, sub)
    inv.write("C11", "Tax", F(bold=True))
    if "usename" in done:
        inv.write_formula("D11", "=D10*TaxRate", m, round(sub * 0.08, 2))
    else:
        inv.write_formula("D11", "=D10*F2", m, round(sub * 0.08, 2))
    inv.write("C12", "Total", F(bold=True)); inv.write_formula("D12", "=D10+D11", m, round(sub * 1.08, 2))
    inv.set_column("A:A", 24); inv.set_column("B:D", 11); inv.set_column("E:E", 10)
    if "name" in done:
        wb.define_name("TaxRate", "=Invoice!$F$2")
    wb.define_name("Deposits", "=Details!$B$3:$B$9")
    det.write("A1", "Deposits taken", F(bold=True, font_size=14))
    det.write_row("A2", ["Rental", "Deposit"], head)
    for i, d in enumerate(DEPOSITS):
        det.write(2 + i, 0, "R-%d" % (201 + i)); det.write_number(2 + i, 1, d, m)
    summ.write("A1", "Summary", F(bold=True, font_size=14))
    summ.write("A3", "Total deposits"); summ.write_formula("B3", "=SUM(Deposits)", m, sum(DEPOSITS))
    if "linksheet" in done:
        summ.write_url("A5", "internal:Details!A1", string="See every deposit")
    if "linkweb" in done:
        summ.write_url("A6", "https://lakeshorepaddle.example/rates", string="Current rental rates")
    if "removebroken" not in done:
        summ.write_url("A8", "http://oldsite.lakeshorepaddle.example/2019", string="Old price list (2019)")
    summ.set_column("A:A", 26)
    inv.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Lesson 4: Text and generated numbers (D2.10 - D2.12)
# ---------------------------------------------------------------------------
MEMBERS = [("jordan", "RIVERA", "LPC-27-4417", "jrivera@example.com"), ("avery", "chen", "LPC-26-1093", "achen@example.com"),
           ("Sam", "Patel", "LPC-27-5520", "spatel@example.com"), ("RILEY", "brooks", "LPC-25-0381", "rbrooks@example.com"),
           ("casey", "Morgan", "LPC-27-6645", "cmorgan@example.com"), ("Taylor", "KIM", "LPC-26-2290", "tkim@example.com"),
           ("drew", "allen", "LPC-27-7701", "dallen@example.com"), ("Quinn", "Harper", "LPC-26-3158", "qharper@example.com")]


def build_l4(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    ws = wb.add_worksheet("Members")
    ws.write("A1", "Paddle Pass members", F(bold=True, font_size=14))
    ws.write_row("A3", ["First", "Last", "Member code", "Email", "Prefix", "Year", "Last 4", "Last name", "Code length", "Full name"], head)
    for i, (first, last, code, mail) in enumerate(MEMBERS):
        r = 3 + i
        n = r + 1
        ws.write(r, 0, first); ws.write(r, 1, last); ws.write(r, 2, code); ws.write(r, 3, mail.upper() if i % 3 == 0 else mail)
        if "left" in done:
            ws.write_formula(r, 4, "=LEFT(C%d,3)" % n, None, code[:3])
        if "mid" in done:
            ws.write_formula(r, 5, "=MID(C%d,5,2)" % n, None, code[4:6])
        if "right" in done:
            ws.write_formula(r, 6, "=RIGHT(C%d,4)" % n, None, code[-4:])
        if "upper" in done:
            ws.write_formula(r, 7, "=UPPER(B%d)" % n, None, last.upper())
        if "len" in done:
            ws.write_formula(r, 8, "=LEN(C%d)" % n, None, len(code))
        if "concat" in done:
            ws.write_formula(r, 9, '=_xlfn.CONCAT(A%d," ",B%d)' % (n, n), None, first + " " + last)
    last_row = 3 + len(MEMBERS)
    ws.write(last_row + 1, 0, "Email list", F(bold=True))
    if "lower" in done:
        ws.write(2, 10, "Email (lowercase)", head)
        for i, m in enumerate(MEMBERS):
            mail = m[3].upper() if i % 3 == 0 else m[3]
            ws.write_formula(3 + i, 10, "=LOWER(D%d)" % (4 + i), None, mail.lower())
    if "textjoin" in done:
        joined = "; ".join(m[3].upper() if i % 3 == 0 else m[3] for i, m in enumerate(MEMBERS))
        ws.write_formula(last_row + 1, 1, '=_xlfn.TEXTJOIN("; ",TRUE,D4:D%d)' % (last_row), None, joined)
    ws.set_column("A:B", 10); ws.set_column("C:D", 22); ws.set_column("E:K", 11)
    # --- Practice data: SEQUENCE and RANDBETWEEN ---
    p = wb.add_worksheet("Practice")
    p.write_row("A1", ["Line", "Practice hours"], head)
    if "sequence" in done:
        p.write_dynamic_array_formula("A2:A11", "=_xlfn.SEQUENCE(10)", None, 1)
        for k in range(2, 11):
            p.write_number(k, 0, k)
    if "randbetween" in done:
        for k in range(10):
            p.write_formula(1 + k, 1, "=RANDBETWEEN(1,4)", None, 1 + (k * 3) % 4)
    ws.activate()
    wb.close()


# ---------------------------------------------------------------------------
# Skill Check 2
# ---------------------------------------------------------------------------
SC2_ORDERS = [("O-301", "avery chen", 3, 35), ("O-302", "SAM PATEL", 1, 50), ("O-303", "riley brooks", 2, 45), ("O-304", "Jordan Lee", 4, 25),
              ("O-305", "casey MORGAN", 2, 30), ("O-306", "Taylor Kim", 5, 25), ("O-307", "drew allen", 1, 70), ("O-308", "Quinn Harper", 3, 35)]


def build_sc2(path, done=frozenset()):
    wb = new_book(path)
    F = Fmt(wb)
    head = F(bold=True, bottom=1)
    m = money(F)
    o = wb.add_worksheet("Orders")
    s = wb.add_worksheet("Settings")
    o.write("A1", "June orders", F(bold=True, font_size=14))
    o.write_row("A3", ["Order", "Customer", "Boats", "Rate", "Amount", "Discount", "Status", "Prefix", "Customer (caps)", "Label"], head)
    for i, (oid, cust, n, rate) in enumerate(SC2_ORDERS):
        r = 3 + i
        k = r + 1
        if "sequence" in done:
            pass
        o.write(r, 0, oid); o.write(r, 1, cust); o.write_number(r, 2, n); o.write_number(r, 3, rate, m)
        amt = n * rate
        if "amount" in done:
            o.write_formula(r, 4, "=C%d*D%d" % (k, k), m, amt)
        if "discount" in done:
            o.write_formula(r, 5, "=E%d*Settings!$B$2" % k, m, round(amt * 0.1, 2))
        if "if" in done:
            o.write_formula(r, 6, '=IF(C%d>=3,"Group","Single")' % k, None, "Group" if n >= 3 else "Single")
        if "left" in done:
            o.write_formula(r, 7, "=LEFT(A%d,1)" % k, None, oid[0])
        if "upper" in done:
            o.write_formula(r, 8, "=UPPER(B%d)" % k, None, cust.upper())
        if "concat" in done:
            o.write_formula(r, 9, '=_xlfn.CONCAT(A%d," - ",B%d)' % (k, k), None, oid + " - " + cust)
    end = 3 + len(SC2_ORDERS)
    o.write(end, 3, "Total", F(bold=True))
    o.write(end + 1, 3, "Average", F(bold=True))
    o.write(end + 2, 3, "Customers", F(bold=True))
    amounts = [n * r for _, _, n, r in SC2_ORDERS]
    if "sum" in done:
        o.write_formula(end, 4, "=SUM(E4:E11)", m, sum(amounts))
    if "average" in done:
        o.write_formula(end + 1, 4, "=AVERAGE(E4:E11)", m, sum(amounts) / len(amounts))
    if "counta" in done:
        o.write_formula(end + 2, 4, "=COUNTA(B4:B11)", None, len(SC2_ORDERS))
    o.write("L3", "Week", head)
    if "sequence" in done:
        o.write_dynamic_array_formula("L4:L7", "=_xlfn.SEQUENCE(4)", None, 1)
        for k in range(1, 4):
            o.write_number(3 + k, 11, k + 1)
    if "link" in done:
        o.write_url("A15", "internal:Settings!A1", string="Discount settings")
    o.set_column("A:A", 8); o.set_column("B:B", 16); o.set_column("E:F", 10); o.set_column("I:J", 22)
    s.write("A1", "Settings", F(bold=True, font_size=14))
    s.write("A2", "Discount rate"); s.write_number("B2", 0.1, F(num_format="0%"))
    wb.define_name("Rates", "=Orders!$D$4:$D$11")    # for the Name Box task
    if "name" in done:
        wb.define_name("Discount", "=Settings!$B$2")
        s.write("A4", "Discount on $100"); s.write_formula("B4", "=100*Discount", m, 10)
    else:
        s.write("A4", "Discount on $100")
    s.set_column("A:A", 18)
    o.activate()
    wb.close()
    if "showformulas" in done:
        show_formulas(path, 0, True)


def build_all(base):
    import os
    d = os.path.join(base, "week-02")
    os.makedirs(d, exist_ok=True)
    build_l1(os.path.join(d, "E2-L1-orders.xlsx"))
    build_l2(os.path.join(d, "E2-L2-functions.xlsx"))
    build_l3(os.path.join(d, "E2-L3-invoice.xlsx"))
    build_l4(os.path.join(d, "E2-L4-members.xlsx"))
    build_sc2(os.path.join(d, "E2-SC2-orders.xlsx"))


if __name__ == "__main__":
    lines, sub, total = order_totals()
    print("order lines", lines, "sub", sub, "total", total)
    print("tax C9", round(TAX_ITEMS[-1][1] * 0.08, 2), "grid F8", GRID_PRICES[-1] * 5)
    print("weekly total", sum(n for _, n in WEEKLY), "avg", sum(n for _, n in WEEKLY) / 7)
    for j in range(4):
        col = [w[1][j] for w in WEEKS]
        print("week", j + 1, sum(col), sum(col) / 4, max(col), min(col))
    print("row totals", [sum(n) for _, n in WEEKS])
    print("count", sum(1 for n, p in SIGNUPS if isinstance(p, int)), "counta", sum(1 for n, p in SIGNUPS if n), "blank", sum(1 for n, p in SIGNUPS if p is None))
    print("returns", [("Late" if b > d else "On time") for _, d, b in RETURNS])
    print("deposits", sum(DEPOSITS), "invoice tax", round(sub * 0.08, 2))
    amounts = [n * r for _, _, n, r in SC2_ORDERS]
    print("sc2 amounts", amounts, sum(amounts), sum(amounts) / len(amounts))
