# make-starters.py - builds the Week 1 starter documents.
# Each starter has the exact problems its drills ask learners to fix, so the
# checks in the lesson files can be written against a known starting point.
# Run from the word/ folder:  python3 tools/make-starters.py
# Needs python-docx and Pillow (pip install python-docx pillow).
# Week 1 is built here; Weeks 2 to 5 are in starters_w2to5.py.
#
# Business names and people in these files are fictional.
import pathlib
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docxkit import set_compat

OUT = pathlib.Path(__file__).resolve().parent.parent / "starters" / "week-01"
OUT.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------- helpers
def new_doc(font="Aptos", size=11, after=8, lang="en-US"):
    """A blank document with a known Normal style, Letter size, 1 in margins and empty properties."""
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = font
    normal.font.size = Pt(size)
    # make the font apply to all scripts, not only Latin text
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), font)
    set_lang(rpr, lang)
    pf = normal.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(after); pf.line_spacing = 1.0
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.orientation = WD_ORIENT.PORTRAIT
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(1))
    cp = doc.core_properties
    cp.author = ""; cp.title = ""; cp.keywords = ""; cp.subject = ""; cp.comments = ""; cp.last_modified_by = ""
    set_compat(doc, 15)  # current Word; python-docx's template would open in Compatibility Mode
    return doc


def set_lang(rpr, lang):
    el = rpr.find(qn("w:lang"))
    if el is None:
        el = OxmlElement("w:lang"); rpr.append(el)
    el.set(qn("w:val"), lang)


def para(doc, text, **kw):
    """Add a paragraph. kw: bold, font, size, color, highlight, align, after, line, lang."""
    p = doc.add_paragraph()
    if text:
        r = p.add_run(text)
        if kw.get("bold"): r.bold = True
        if kw.get("font"):
            r.font.name = kw["font"]
            r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), kw["font"])
        if kw.get("size"): r.font.size = Pt(kw["size"])
        if kw.get("color"): r.font.color.rgb = RGBColor.from_string(kw["color"])
        if kw.get("highlight"): r.font.highlight_color = kw["highlight"]
        if kw.get("lang"): set_lang(r._element.get_or_add_rPr(), kw["lang"])
    if "align" in kw: p.alignment = kw["align"]
    if "after" in kw: p.paragraph_format.space_after = Pt(kw["after"])
    if "line" in kw: p.paragraph_format.line_spacing = kw["line"]
    return p


def empty(doc, n=1):
    for _ in range(n):
        doc.add_paragraph()


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def save(doc, name):
    path = OUT / name
    doc.save(path)
    print("wrote", path.relative_to(OUT.parent.parent))


# ------------------------------------------------- Lesson 1: a clean letter
def lesson1():
    d = new_doc()
    para(d, "Lakeshore Paddle Company", bold=True, size=16, after=0)
    para(d, "12 Harbor Street, Canandaigua, NY 14424", after=18)
    para(d, "October 20, 2026")
    para(d, "Lakeshore Paddle Company Volunteers", after=0)
    para(d, "Canandaigua, NY 14424")
    para(d, "Dear Volunteers:")
    para(d, "Thank you for helping at last spring's shoreline cleanup. This letter explains how to sign up for this year's volunteer day.")
    para(d, "Volunteer day starts at 9:00 a.m. on Saturday, April 18, at the boat launch. Please reply by April 1 so we can order enough gloves and lunches.")
    para(d, "Sincerely,", after=36)
    para(d, "Jordan Rivera", after=0)
    para(d, "Volunteer Coordinator")
    save(d, "W1-L1-letter.docx")


# --------------------------------------- Lesson 2: a letter that needs cleanup
def lesson2():
    d = new_doc()
    # Planted: 4 x "Acme Outdoor Supply", 3 x "deadline", 1 x "cafe",
    # empty paragraphs, double spaces, and a misplaced pasted paragraph.
    para(d, "Acme Outdoor Supply", bold=True, size=16, after=0)
    para(d, "12 Harbor Street, Canandaigua, NY 14424")
    empty(d, 2)
    para(d, "October 20, 2026")
    empty(d)
    para(d, "Dear Volunteers:")
    empty(d)
    para(d, "Thank you for helping Acme Outdoor Supply at last spring's shoreline cleanup.  This year the cleanup moves to Saturday, April 18.  The sign-up deadline is April 1.")
    empty(d)
    para(d, "We will meet at the cafe next to the boat launch at 9:00 a.m.  Acme Outdoor Supply will provide gloves, bags and lunch.  Please meet the deadline so we can order enough of each.")
    empty(d, 2)
    para(d, "If you can't meet the deadline, call the shop and Acme Outdoor Supply will add you to the waiting list.")
    empty(d)
    para(d, "Sincerely,", after=36)
    para(d, "Jordan Rivera", after=0)
    para(d, "Volunteer Coordinator")
    empty(d)
    # pasted from email: wrong font, red, highlighted, below the signature
    para(d, "We look forward to seeing you on the water!", font="Courier New", size=12, color="C00000", highlight=WD_COLOR_INDEX.YELLOW)
    d.sections[0].footer.paragraphs[0].text = "2026 Lakeshore Paddle Company"
    save(d, "W1-L2-letter.docx")


# ------------------------------- Lesson 3: letter + memo that need formatting
def lesson3():
    d = new_doc(font="Times New Roman", size=12, after=0)
    para(d, "Lakeshore Paddle Company", size=18, after=0)
    para(d, "12 Harbor Street, Canandaigua, NY 14424", after=12)
    para(d, "October 20, 2026", after=12)
    para(d, "Subject: Spring volunteer day", after=12)
    para(d, "Dear Volunteers,", after=12)
    j = WD_ALIGN_PARAGRAPH.JUSTIFY
    para(d, "Thank you for helping at last spring's shoreline cleanup. Your work filled forty bags with trash and made the launch safer for everyone who paddles there.", align=j, line=1.5, after=0)
    para(d, "Volunteer day starts at 9:00 a.m. on Saturday, April 18, at the boat launch. Please reply by April 1 so we can order enough gloves and lunches.", align=j, line=1.5, after=0)
    para(d, "Volunteers should bring water, sunscreen and shoes that can get wet.", font="Comic Sans MS", size=14, color="7030A0", highlight=WD_COLOR_INDEX.BRIGHT_GREEN)
    para(d, "Sincerely,", after=36)
    para(d, "Jordan Rivera", after=0)
    para(d, "Volunteer Coordinator")
    page_break(d)
    para(d, "MEMORANDUM", after=12)
    para(d, "TO:\tShop staff", after=0)
    para(d, "FROM:\tJordan Rivera", after=0)
    para(d, "DATE:\tOctober 20, 2026", after=0)
    para(d, "SUBJECT:\tVolunteer day schedule", after=12)
    para(d, "Please review the volunteer day schedule before Friday. We need two people at the launch by 8:30 a.m. to set up tables, and one person at the shop all day.")
    save(d, "W1-L3-letter.docx")


# -------------------------- Lesson 4: wrong page setup, typos, UK English
def lesson4():
    d = new_doc(lang="en-GB")
    sec = d.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Mm(297), Mm(210)  # A4 landscape
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.75))
    L = "en-GB"
    para(d, "Lakeshore Paddle Company", bold=True, size=16, after=0, lang=L)
    para(d, "12 Harbor Street, Canandaigua, NY 14424", after=18, lang=L)
    para(d, "October 20, 2026", lang=L)
    para(d, "Dear Volunteers:", lang=L)
    para(d, "Thank you for helping at last spring's shoreline cleanup. You will recieve a T-shirt and lunch on volunteer day.", lang=L)
    para(d, "Please bring a seperate bag for recycling. If it rains, we will move the event to tommorow's backup date, Sunday, April 19.", lang=L)
    para(d, "Sincerely,", after=36, lang=L)
    para(d, "Jordan Rivera", after=0, lang=L)
    para(d, "Volunteer Coordinator", lang=L)
    save(d, "W1-L4-letter.docx")


# --------------------------------------------- Skill check 1: a messy memo
def skillcheck1():
    d = new_doc()
    sec = d.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Mm(297), Mm(210)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.5))
    # Planted: 4 x "refund", 3 x "Northside Rentals", empty paragraphs, double spaces
    para(d, "Spring Rental Policy", size=20, after=6)
    para(d, "Northside Rentals - staff memo", after=12)
    empty(d)
    para(d, "Pickup times", bold=True, size=13, color="0046AD", after=4)
    para(d, "Boats are ready at 9:00 a.m.  Customers with a PaddlePass skip the line.", after=8)
    empty(d, 2)
    para(d, "Late returns", after=4)
    para(d, "Rentals must be back by 6:00 p.m. A boat returned late is charged one extra hour, and Northside Rentals does not give a refund for unused time.", after=0)
    empty(d)
    para(d, "Damage fees are charged only after a manager inspects the boat. Customers may ask for a refund within 7 days.", font="Arial Black", size=12, color="C00000", highlight=WD_COLOR_INDEX.YELLOW)
    para(d, "Weather closures: if Northside Rentals closes the lake for weather, every customer gets a full refund.  A PaddlePass holder can choose a refund or a free day.", after=8)
    save(d, "W1-SC1-memo.docx")


if __name__ == "__main__":
    lesson1(); lesson2(); lesson3(); lesson4(); skillcheck1()
    # Weeks 2 to 5 live in their own file to keep this one readable.
    import starters_w2to5
    starters_w2to5.build_all()
