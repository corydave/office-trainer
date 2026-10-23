# make-test-docs.py - builds the answer keys and test documents for the checker.
#
# For each Week 1 starter it writes:
#   *-solved.docx   the starter with every drill done (an answer key)
#   *-partial.docx  some steps done, some not (to prove checks can fail)
# check-tests.html then runs every check in the lesson files against these
# files and the starters, and compares the results with test-docs/expected.json.
#
# Run from the word/ folder:  python3 tools/make-test-docs.py
# Needs python-docx (pip install python-docx).
#
# These are built by code, not saved from Word. Word writes some things
# differently (theme fonts, effect presets), so add real Word-saved files to
# test-docs/ as you collect them, with their expected results.
import copy
import pathlib
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

WORD = pathlib.Path(__file__).resolve().parent.parent
STARTERS = WORD / "starters" / "week-01"
OUT = WORD / "tools" / "test-docs"
OUT.mkdir(parents=True, exist_ok=True)

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"


# ---------------------------------------------------------------- helpers
def open_starter(name):
    return Document(str(STARTERS / name))


def find(doc, starts):
    """The first body paragraph that starts with this text."""
    for p in doc.paragraphs:
        if p.text.startswith(starts):
            return p
    raise ValueError("no paragraph starts with " + starts)


def delete_paragraph(p):
    p._element.getparent().remove(p._element)


def remove_empty_paragraphs(doc):
    for p in list(doc.paragraphs):
        has_break = p._element.xpath(".//w:br[@w:type='page']")
        if p.text.strip() == "" and not has_break:
            delete_paragraph(p)


def replace_in_runs(doc, old, new, paragraphs=None):
    """Replace text run by run (fine for these files: each paragraph is one run)."""
    for p in paragraphs if paragraphs is not None else doc.paragraphs:
        for r in p.runs:
            if old in r.text:
                r.text = r.text.replace(old, new)


def clear_run_formatting(p):
    """What Clear All Formatting / Keep Text Only leaves: runs with no rPr."""
    for r in p.runs:
        rpr = r._element.find(qn("w:rPr"))
        if rpr is not None:
            r._element.remove(rpr)


def add_glow(run):
    """A glow text effect, the way Word 2010+ saves it (w14:glow)."""
    rpr = run._element.get_or_add_rPr()
    rpr.append(parse_xml(
        f'<w14:glow xmlns:w14="{W14}" w14:rad="63500"><w14:schemeClr w14:val="accent1">'
        f'<w14:alpha w14:val="60000"/><w14:satMod w14:val="175000"/></w14:schemeClr></w14:glow>'))


def set_lang_all(doc, lang):
    for p in doc.paragraphs:
        for r in p.runs:
            rpr = r._element.get_or_add_rPr()
            el = rpr.find(qn("w:lang"))
            if el is None:
                el = OxmlElement("w:lang"); rpr.append(el)
            el.set(qn("w:val"), lang)


def letter_page(doc):
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(1))


def save(doc, name):
    doc.save(str(OUT / name))
    print("wrote", (OUT / name).relative_to(WORD))


# ------------------------------------------------------------- Lesson 2
def lesson2_solved():
    d = open_starter("W1-L2-letter.docx")
    remove_empty_paragraphs(d)                                   # D1.03
    for p in d.paragraphs:
        for r in p.runs:
            while "  " in r.text:
                r.text = r.text.replace("  ", " ")
    replace_in_runs(d, "Acme Outdoor Supply", "Lakeshore Paddle Company")   # D1.04
    footer = d.sections[0].footer.paragraphs[0]                  # D1.05
    footer.text = "© 2026 Lakeshore Paddle Company"
    replace_in_runs(d, "cafe", "café")
    moved = find(d, "We look forward")                           # D1.08: cut, paste above Sincerely, Keep Text Only
    clear_run_formatting(moved)
    find(d, "Sincerely,")._element.addprevious(moved._element)
    save(d, "W1-L2-solved.docx")


def lesson2_partial():
    d = open_starter("W1-L2-letter.docx")
    for p in d.paragraphs:                                       # double spaces fixed...
        for r in p.runs:
            r.text = r.text.replace("  ", " ")
    replace_in_runs(d, "Acme Outdoor Supply", "Lakeshore Paddle Company", d.paragraphs[:3])  # ...only the letterhead renamed
    moved = find(d, "We look forward")                           # moved, but kept its red email font
    find(d, "Sincerely,")._element.addprevious(moved._element)
    save(d, "W1-L2-partial.docx")


# ------------------------------------------------------------- Lesson 3
BODY_STARTS = ["Dear Volunteers", "Thank you for", "Volunteer day starts", "Volunteers should bring", "Sincerely,"]


def lesson3_solved():
    d = open_starter("W1-L3-letter.docx")
    for s in BODY_STARTS:                                        # D1.06 step 1: Aptos 11
        for r in find(d, s).runs:
            r.font.name = "Aptos"; r.font.size = Pt(11)
    for r in find(d, "Subject:").runs:                           # step 2: bold
        r.bold = True
    head = find(d, "Lakeshore Paddle Company")                   # step 3: text effect
    add_glow(head.runs[0])
    memo = find(d, "MEMORANDUM")                                 # step 4: Format Painter copies the run formatting
    memo.runs[0]._element.insert(0, copy.deepcopy(head.runs[0]._element.find(qn("w:rPr"))))
    for s in ["Thank you for", "Volunteer day starts"]:          # D1.07: single, 8 pt after, left
        p = find(d, s)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(8)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    find(d, "Please review").paragraph_format.first_line_indent = Inches(0.5)
    clear_run_formatting(find(d, "Volunteers should bring"))     # Clear Formatting
    save(d, "W1-L3-solved.docx")


def lesson3_partial():
    d = open_starter("W1-L3-letter.docx")
    for r in find(d, "Thank you for").runs:                      # only one of the two body paragraphs changed
        r.font.name = "Aptos"; r.font.size = Pt(11)
    head = find(d, "Lakeshore Paddle Company")
    add_glow(head.runs[0])                                       # effect added, Format Painter not used
    p = find(d, "Thank you for")
    p.paragraph_format.line_spacing = 1.15                       # 1.15 instead of single
    p.paragraph_format.space_after = Pt(8)
    find(d, "Please review").paragraph_format.left_indent = Inches(0.5)   # left indent instead of first-line
    save(d, "W1-L3-partial.docx")


# ------------------------------------------------------------- Lesson 4
def lesson4_solved():
    d = open_starter("W1-L4-letter.docx")
    letter_page(d)                                               # D1.09
    replace_in_runs(d, "recieve", "receive")
    replace_in_runs(d, "seperate", "separate")
    replace_in_runs(d, "tommorow", "tomorrow")
    d.core_properties.title = "Volunteer Day Letter"             # D1.10
    d.core_properties.author = "Test Learner"
    d.core_properties.keywords = "letter; volunteer"
    set_lang_all(d, "en-US")
    save(d, "W1-L4-solved.docx")


def lesson4_partial():
    d = open_starter("W1-L4-letter.docx")
    sec = d.sections[0]                                          # orientation fixed, but still A4 with 0.75 in margins
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    replace_in_runs(d, "recieve", "receive")
    d.core_properties.title = "Letter"                           # wrong title, no author
    save(d, "W1-L4-partial.docx")


# ------------------------------------------------------- Skill check 1
def sc1_solved():
    d = open_starter("W1-SC1-memo.docx")
    remove_empty_paragraphs(d)
    for p in d.paragraphs:
        for r in p.runs:
            r.text = r.text.replace("  ", " ")
    replace_in_runs(d, "Northside Rentals", "Lakeshore Paddle Company")
    replace_in_runs(d, "PaddlePass skip", "PaddlePass® skip")
    add_glow(find(d, "Spring Rental Policy").runs[0])
    late = find(d, "Late returns")
    late.runs[0]._element.insert(0, copy.deepcopy(find(d, "Pickup times").runs[0]._element.find(qn("w:rPr"))))
    p = find(d, "Rentals must be")
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.5)
    clear_run_formatting(find(d, "Damage fees"))
    letter_page(d)
    d.core_properties.title = "Spring Rental Policy"
    d.core_properties.keywords = "policy; memo"
    save(d, "W1-SC1-solved.docx")




# =========================================================================== Weeks 2 to 5
# Answer keys for the later weeks, built with the helpers in docxkit.py.
import sys
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))
import docxkit as K
from docx.enum.section import WD_ORIENT as ORIENT


def starter(week, name):
    return Document(str(WORD / "starters" / ("week-%02d" % week) / name))


def restyle(p, style):
    """Apply a paragraph style and take off the hand formatting (what clicking a style + Clear Formatting does)."""
    p.style = style
    clear_run_formatting(p)


def section_break_before(doc, p, landscape_after=True):
    """Insert a Next Page section break just before paragraph p; the part after it becomes landscape."""
    body_sect = doc.element.body.sectPr
    new_p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pPr.append(copy.deepcopy(body_sect))        # the section that ENDS here keeps the old (portrait) setup
    new_p.append(pPr)
    p._p.addprevious(new_p)
    if landscape_after:
        pgSz = body_sect.find(qn("w:pgSz"))
        w, h = pgSz.get(qn("w:w")), pgSz.get(qn("w:h"))
        pgSz.set(qn("w:w"), h); pgSz.set(qn("w:h"), w); pgSz.set(qn("w:orient"), "landscape")


def header_with_page(doc, name="Rivera", where="header"):
    sec = doc.sections[0]
    hp = (sec.header if where == "header" else sec.footer).paragraphs[0]
    hp.text = name + " "
    K.add_simple_field(hp, "PAGE", "2")
    sec.different_first_page_header_footer = True


def move_section(doc, heading, before):
    """Move a heading and everything up to the next Heading 1 to just before another paragraph."""
    paras = doc.paragraphs
    start = next(i for i, p in enumerate(paras) if p.text == heading)
    end = next((i for i in range(start + 1, len(paras)) if paras[i].style.name == "Heading 1"), len(paras))
    target = next(p for p in paras if p.text == before)._p
    for p in paras[start:end]:
        target.addprevious(p._p)


def week2_solved():
    # ---- Lesson 1: styles
    d = starter(2, "W2-L1-report.docx")
    for kind, text in __import__("starters_w2to5").REPORT:
        if kind in ("title", "h1", "h2"):
            restyle(K.find(d, text), {"title": "Title", "h1": "Heading 1", "h2": "Heading 2"}[kind])
    normal = d.styles["Normal"]
    normal.font.name = "Times New Roman"; normal.font.size = Pt(12)
    normal.element.rPr.find(qn("w:rFonts")).set(qn("w:ascii"), "Times New Roman")
    normal.element.rPr.find(qn("w:rFonts")).set(qn("w:hAnsi"), "Times New Roman")
    h1 = d.styles["Heading 1"]
    h1.font.size = Pt(14); h1.font.bold = True; h1.font.color.rgb = __import__("docx").shared.RGBColor(0, 0, 0)
    h1.paragraph_format.space_before = Pt(12); h1.paragraph_format.keep_with_next = True
    save(d, "W2-L1-solved.docx")
    # partial: headings styled, but only Normal modified
    d = starter(2, "W2-L1-report.docx")
    for text in ["Introduction", "Methods", "Recommendations", "Appendix: Sampling Data"]:
        restyle(K.find(d, text), "Heading 1")
    d.styles["Normal"].font.size = Pt(12)
    save(d, "W2-L1-partial.docx")

    # ---- Lesson 2: navigation
    d = starter(2, "W2-L2-report.docx")
    move_section(d, "Methods", "Results")
    K.bookmark(K.find(d, "Methods"), "Methods")
    intro = K.find(d, "This report explains")
    text = intro.text
    before, after = text.split("how we collected our samples")
    for r in list(intro.runs):
        r._r.getparent().remove(r._r)
    intro.add_run(before)
    K.hyperlink(d, intro, "how we collected our samples", anchor="Methods")
    intro.add_run(after)
    save(d, "W2-L2-solved.docx")

    # ---- Lesson 3: lists, headers, breaks
    d = starter(2, "W2-L3-report.docx")
    steps = K.add_numbering(d, "decimal", "Step %1.")
    for t in ["Walk to the marked", "Prepare the sample", "Label it with", "Rinse it twice", "Fill the bottle", "Store it in the cooler"]:
        K.set_list(K.find(d, t), steps, 1 if t.startswith(("Label", "Rinse")) else 0)
    check = K.add_numbering(d, "bullet", "", bullet_font="Wingdings")
    for t in ["Sample bottles", "Secchi disk", "Clipboard and pencil", "Cooler with ice packs"]:
        K.set_list(K.find(d, t), check)
    recs_num = int(K.find(d, "Reduce single-use")._p.pPr.numPr.numId.val)
    fresh = K.restart_num(d, recs_num, 1)
    for t in ["Reduce single-use", "Add two trash", "Run a volunteer"]:
        K.set_list(K.find(d, t), fresh)
    lab_num = int(K.find(d, "Filter each sample")._p.pPr.numPr.numId.val)
    five = K.restart_num(d, lab_num, 5)
    for t in ["Filter each sample", "Count and photograph", "Enter the results"]:
        K.set_list(K.find(d, t), five)
    header_with_page(d)
    section_break_before(d, K.find(d, "Appendix: Sampling Data"))
    save(d, "W2-L3-solved.docx")
    # partial: list made but levels and symbols not changed; numbering not fixed; header without first-page setting
    d = starter(2, "W2-L3-report.docx")
    steps = K.add_numbering(d, "decimal", "%1.")
    for t in ["Walk to the marked", "Prepare the sample", "Label it with", "Rinse it twice", "Fill the bottle", "Store it in the cooler"]:
        K.set_list(K.find(d, t), steps)
    sec = d.sections[0]
    K.add_simple_field(sec.header.paragraphs[0], "PAGE", "1")
    sec.orientation = ORIENT.LANDSCAPE   # whole document turned, not just the appendix
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    save(d, "W2-L3-partial.docx")

    # ---- Lesson 4: notes, contents, citations
    d = starter(2, "W2-L4-report.docx")
    K.footnotes(d, [(K.find(d, "Clarity was measured"), "A Secchi disk is a black and white disk lowered until it disappears from view."),
                    (K.find(d, "Volunteers counted"), "Counts include every item larger than a bottle cap.")], fmt="lowerRoman")
    restyle(K.find(d, "Limitations"), "Heading 1")
    note = K.find(d, "Practice report")
    toc = OxmlElement("w:p"); note._p.addnext(toc)
    from docx.text.paragraph import Paragraph
    tocp = Paragraph(toc, note._parent)
    K.add_field(tocp, 'TOC \\o "1-2" \\h \\z \\u', "")
    last = tocp
    for h in ["Introduction", "Background", "Methods", "Results", "Discussion", "Limitations", "Recommendations", "Appendix: Sampling Data"]:
        e = OxmlElement("w:p"); last._p.addnext(e)
        last = Paragraph(e, note._parent)
        last._p.get_or_add_pPr().append(K.parse_xml('<w:pStyle %s w:val="TOC1"/>' % K.NSW))
        last.add_run(h + "\t1")
    for t in ["Two earlier student surveys", "Clarity was measured", "Run a volunteer cleanup"]:
        K.add_field(K.find(d, t), "CITATION Src1 \\l 1033", "(Lin, 2024)")
    K.sources(d, 3)
    end = d.add_paragraph()
    K.add_field(end, "BIBLIOGRAPHY", "References")
    save(d, "W2-L4-solved.docx")

    # ---- Skill check 2
    d = starter(2, "W2-SC2-handbook.docx")
    for t in ["Welcome", "Your first week", "Rentals", "Safety on the water", "Timeline", "Budget", "Appendix: Contact List"]:
        restyle(K.find(d, t), "Heading 1")
    for t in ["Opening the shop", "Closing the shop"]:
        restyle(K.find(d, t), "Heading 2")
    d.styles["Heading 1"].font.size = Pt(16)
    move_section(d, "Budget", "Timeline")
    K.bookmark(K.find(d, "Budget"), "Budget")
    first = K.find(d, "Read this handbook")
    before, after = first.text.split("see the budget")
    for r in list(first.runs):
        r._r.getparent().remove(r._r)
    first.add_run(before); K.hyperlink(d, first, "see the budget", anchor="Budget"); first.add_run(after)
    num = K.add_numbering(d, "decimal", "%1.")
    for t in ["Unlock the boat racks", "Count the cash drawer.", "Check the life jacket bins", "Turn on the dock lights"]:
        K.set_list(K.find(d, t), num, 1 if t == "Count the cash drawer." else 0)
    close_num = int(K.find(d, "Lock the boat racks")._p.pPr.numPr.numId.val)
    fresh = K.restart_num(d, close_num, 1)
    for t in ["Lock the boat racks", "Count the cash drawer again", "Hang up the wet"]:
        K.set_list(K.find(d, t), fresh)
    header_with_page(d, where="footer")
    section_break_before(d, K.find(d, "Appendix: Contact List"))
    K.footnotes(d, [(K.find(d, "Every renter must wear"), "Children under 12 wear a jacket even on the dock.")], fmt="lowerRoman")
    p = K.find(d, "Read this handbook")
    toc = OxmlElement("w:p"); p._p.addnext(toc)
    from docx.text.paragraph import Paragraph
    K.add_field(Paragraph(toc, p._parent), 'TOC \\o "1-2" \\h \\z \\u', "")
    save(d, "W2-SC2-solved.docx")
    # partial: headings and lists done, nothing else
    d = starter(2, "W2-SC2-handbook.docx")
    for t in ["Welcome", "Your first week", "Rentals", "Safety on the water", "Timeline", "Budget", "Appendix: Contact List"]:
        restyle(K.find(d, t), "Heading 1")
    for t in ["Opening the shop", "Closing the shop"]:
        restyle(K.find(d, t), "Heading 2")
    num = K.add_numbering(d, "decimal", "%1.")
    for t in ["Unlock the boat racks", "Count the cash drawer.", "Check the life jacket bins", "Turn on the dock lights"]:
        K.set_list(K.find(d, t), num)
    save(d, "W2-SC2-partial.docx")


def alt_by_name(doc, alts, decorative=()):
    """Set alt text on drawings by their name (Selection pane name); names in `decorative` get marked decorative."""
    WPD = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
    for dp in doc.element.body.iter(WPD):
        name = dp.get("name")
        if any(name.startswith(n) for n in decorative):
            dp.append(K.parse_xml('<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
                                  '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>'))
        elif alts is not None:
            dp.set("descr", alts.get(name, "Description of " + name))


def picture_index(doc, name):
    WPD = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
    pics = [dp for dp in doc.element.body.iter(WPD) if dp.getparent().find(".//{http://schemas.openxmlformats.org/drawingml/2006/picture}pic") is not None]
    return [dp.get("name") for dp in pics].index(name)


def lines_to_table(doc, first, count, cols):
    """Replace `count` tab-separated paragraphs (starting at `first`) with a table (Convert Text to Table)."""
    start = K.find(doc, first)
    paras = doc.paragraphs
    i = [p._p for p in paras].index(start._p)
    rows = [p.text.split("\t") for p in paras[i:i + count]]
    t = doc.add_table(rows=len(rows), cols=cols)
    t.style = "Table Grid"
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            t.cell(r, c).text = text
    start._p.addprevious(t._tbl)
    for p in paras[i:i + count]:
        p._p.getparent().remove(p._p)
    return t


def table_to_text(doc, table):
    for row in table.rows:
        p = OxmlElement("w:p"); table._tbl.addprevious(p)
        from docx.text.paragraph import Paragraph
        Paragraph(p, table._parent).add_run("\t".join(c.text for c in row.cells))
    table._tbl.getparent().remove(table._tbl)


def table_with(doc, text):
    return next(t for t in doc.tables if any(text in c.text for row in t.rows for c in row.cells))


def sort_rows(table, first_data_row, key):
    rows = list(table.rows)[first_data_row:]
    trs = [r._tr for r in rows]
    parent = trs[0].getparent()
    for tr in trs:
        parent.remove(tr)
    for r in sorted(rows, key=key):
        parent.append(r._tr)


def split_table(doc, table, first_row_text):
    """Split Table: rows from the one starting with first_row_text go into a new table after an empty paragraph."""
    rows = list(table.rows)
    k = next(i for i, r in enumerate(rows) if r.cells[0].text.startswith(first_row_text))
    new = copy.deepcopy(table._tbl)
    for tr in new.findall(qn("w:tr"))[:k]:
        new.remove(tr)
    for r in rows[k:]:
        table._tbl.remove(r._tr)
    gap = OxmlElement("w:p")
    table._tbl.addnext(gap)
    gap.addnext(new)


def columns_section(doc, last_para_of_one_col, cols=2):
    """Continuous section break after a paragraph: the part before stays one column; the rest gets `cols`."""
    body = doc.element.body.sectPr
    first = copy.deepcopy(body)
    for c in first.findall(qn("w:cols")):
        first.remove(c)
    last_para_of_one_col._p.get_or_add_pPr().append(first)
    K.set_columns(doc.sections[-1], cols)
    K.set_section_type(doc.sections[-1], "continuous")


def week3_solved():
    from datetime import datetime
    pics = WORD / "starters" / "week-03"
    # ---- Lesson 1: graphics
    d = starter(3, "W3-L1-flyer.docx")
    d.add_picture(str(pics / "kayak.png"), width=Inches(3))
    list(d.element.body.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))[-1].set("name", "Kayak")
    K.picture_effects(d, picture_index(d, "Kayak"), style=True, bg_removed=True)
    K.picture_effects(d, picture_index(d, "Lake photo"), artistic=True)
    d.add_picture(str(pics / "divider.png"), width=Inches(0.6))
    list(d.element.body.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))[-1].set("name", "Icon boat")
    K.make_icon(d, picture_index(d, "Icon boat"))
    K.add_shape(K.find(d, "Saturday, June 12"), "Rectangle 1", text="Free demos at 10 a.m.", fill="0E7C86", outline="073B40", halign="left")
    K.add_smartart(d, K.find(d, "Reserve a time"), ["Reserve", "Check in", "Safety talk", "Launch"], layout="process1", colors="colorful1", style="simple3")
    K.add_shape(K.find(d, "Last year 214"), "Text Box 2", text="Best morning of my summer.", textbox=True, wrap="square", halign="right", valign="top")
    K.add_model3d(d.add_paragraph())
    alt_by_name(d, {"Lake photo": "Calm blue lake at sunrise with green hills behind it."}, decorative=("Divider line",))
    save(d, "W3-L1-solved.docx")
    # partial: kayak in, no style; shape with default colors; text box not wrapped; no alt text
    d = starter(3, "W3-L1-flyer.docx")
    d.add_picture(str(pics / "kayak.png"), width=Inches(3))
    K.add_shape(K.find(d, "Saturday, June 12"), "Rectangle 1", text="Free demos at 10 a.m.")
    K.add_shape(K.find(d, "Last year 214"), "Text Box 2", text="Best morning of my summer.", textbox=True, wrap="topAndBottom", halign="left", valign="top")
    save(d, "W3-L1-partial.docx")

    # ---- Lesson 2: tables
    def date_key(r):
        return datetime.strptime(r.cells[0].text, "%B %d, %Y")
    d = starter(3, "W3-L2-schedule.docx")
    lines_to_table(d, "Time\tActivity", 7, 3)
    table_to_text(d, table_with(d, "Kayak (half day)"))
    tours = table_with(d, "Summer tours 2027")
    sort_rows(tours, 2, date_key)
    row0 = tours.rows[0].cells
    row0[0].merge(row0[3])
    tours.cell(0, 0).paragraphs[0].runs[0].text = "Summer tours 2027"
    for extra in tours.cell(0, 0).paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    K.column_width(tours, 0, 1.5)
    K.cell_margins(tours, 0.1)
    K.header_row(tours)
    tours.rows[1]._tr.get_or_add_trPr().append(K.parse_xml('<w:tblHeader %s/>' % K.NSW))
    split_table(d, tours, "August")
    signup = d.add_table(rows=6, cols=4); signup.style = "Table Grid"
    save(d, "W3-L2-solved.docx")
    # partial: schedule converted, but the old table kept; title merged; not sorted; no header rows
    d = starter(3, "W3-L2-schedule.docx")
    lines_to_table(d, "Time\tActivity", 7, 3)
    tours = table_with(d, "Summer tours 2027")
    row0 = tours.rows[0].cells
    row0[0].merge(row0[3])
    save(d, "W3-L2-partial.docx")

    # ---- Lesson 3: chart and caption
    d = starter(3, "W3-L3-survey.docx")
    spot = K.find(d, "Insert the chart below")
    chart_p = OxmlElement("w:p"); spot._p.addnext(chart_p)
    from docx.text.paragraph import Paragraph
    cp = Paragraph(chart_p, spot._parent)
    K.add_chart(d, cp, "col", alt="Column chart: kayaks were the top choice with 42 votes, then paddleboards with 31.")
    cap = OxmlElement("w:p"); chart_p.addnext(cap)
    capp = Paragraph(cap, spot._parent)
    capp.style = "Caption"
    capp.add_run("Figure ")
    K.add_field(capp, "SEQ Figure \\* ARABIC", "1")
    capp.add_run(". Boat preferences from the Paddle Day survey")
    save(d, "W3-L3-solved.docx")

    # ---- Lesson 4: desktop publishing
    d = starter(3, "W3-L4-newsletter.docx")
    columns_section(d, K.find(d, "News from Lakeshore"))
    K.column_break(K.find(d, "214 people tried"))
    K.theme_colors(d, "Lakeshore", "0E7C86")
    for s in d.sections:
        K.page_border(s)
    K.watermark(d.sections[-1], "DRAFT")
    K.drop_cap(K.find(d, "Welcome to the summer issue"))
    K.page_color(d, "EAF6F7")
    save(d, "W3-L4-solved.docx")
    # partial: the whole document in two columns; a CONFIDENTIAL watermark; nothing else
    d = starter(3, "W3-L4-newsletter.docx")
    K.set_columns(d.sections[0], 2)
    K.watermark(d.sections[0], "CONFIDENTIAL")
    save(d, "W3-L4-partial.docx")

    # ---- Skill check 3
    d = starter(3, "W3-SC3-flyer.docx")
    K.picture_effects(d, picture_index(d, "Sunset photo"), style=True)
    alt_by_name(d, {"Sunset photo": "Orange sunset over a calm lake."})
    K.add_shape(K.find(d, "Every Thursday"), "Text Box 1", text="Bring a friend for free", textbox=True, wrap="square", halign="right", valign="top")
    K.add_shape(K.find(d, "Every Thursday"), "Rectangle 2", text="Limited spots", halign="left")
    K.add_smartart(d, K.find(d, "Every Thursday"), ["Book", "Check in", "Paddle"])
    lines_to_table(d, "Date\tSunset tour", 5, 3)
    prices = table_with(d, "Sunset series prices")
    sort_rows(prices, 2, lambda r: float(r.cells[1].text.strip("$")))
    prices.rows[0].cells[0].merge(prices.rows[0].cells[1])
    K.header_row(prices)
    prices.rows[1]._tr.get_or_add_trPr().append(K.parse_xml('<w:tblHeader %s/>' % K.NSW))
    columns_section(d, K.find(d, "About the series"))
    K.watermark(d.sections[-1], "SAMPLE")
    save(d, "W3-SC3-solved.docx")


def fill_controls(doc, values):
    """Type into each content control: the placeholder goes away (showingPlcHdr removed)."""
    sdts = list(doc.element.body.iter(qn("w:sdt")))
    for sdt, value in zip(sdts, values):
        pr = sdt.find(qn("w:sdtPr"))
        ph = pr.find(qn("w:showingPlcHdr"))
        if ph is not None:
            pr.remove(ph)
        sdt.find(".//" + qn("w:t")).text = value


def replace_with_field(doc, placeholder, code, result):
    """Swap bracketed placeholder text in a paragraph for a field (what Insert Merge Field does)."""
    p = next(x for x in doc.paragraphs if placeholder in x.text)
    run = next(r for r in p.runs if placeholder in r.text)      # only that run changes
    before, after = run.text.split(placeholder, 1)
    run.text = before
    anchor = run._r
    pieces = K.field_xml(code, result).split("</w:r>")[:-1]
    for piece in pieces:
        el = K.parse_xml(piece + "</w:r>")
        anchor.addnext(el)
        anchor = el
    if after:
        tail = copy.deepcopy(run._r)
        for t in tail.findall(qn("w:t")):
            t.text = after
        anchor.addnext(tail)


def tab_line(doc, start, pos_in=6.5):
    """Spaces before the dates become one tab, with a right tab stop and dot leader (D4.03)."""
    p = K.find(doc, start)
    r = p.runs[0]
    title, dates = [x for x in r.text.split("  ") if x.strip()][0].strip(), r.text.split()[-3:]
    r.text = title + "\t" + " ".join(dates)
    p._p.get_or_add_pPr().append(K.parse_xml('<w:tabs %s><w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>' % (K.NSW, int(pos_in * 1440))))


def week4_solved():
    S = __import__("starters_w2to5")
    # ---- Lesson 1
    d = starter(4, "W4-L1-resume-template.docx")
    fill_controls(d, ["Alex Morgan", "alex.morgan@example.com · 555-0142 · Canandaigua, NY", "Office Assistant, Lakeshore Paddle Company", "2025 to 2026",
                      "Scheduled 20 seasonal staff and balanced the rental desk every day.", "Finger Lakes Community College", "A.A.S. Business Administration, expected May 2027",
                      "Scheduling, customer service, Microsoft Word"])
    save(d, "W4-L1-resume-solved.docx")
    d = starter(4, "W4-L1-cover-letter.docx")
    K.find(d, "[Contact line]").runs[0].text = "Alex Morgan · alex.morgan@example.com · 555-0142"
    save(d, "W4-L1-cover-solved.docx")
    # ---- Lesson 2
    d = starter(4, "W4-L2-resume-ats.docx")
    for start in ["Office Assistant", "Volunteer Coordinator", "Cashier, Main Street Market"]:
        tab_line(d, start)
    t = d.tables[0]
    p = OxmlElement("w:p"); t._tbl.addprevious(p)
    from docx.text.paragraph import Paragraph
    Paragraph(p, t._parent).add_run("Scheduling, Customer service, Microsoft Word and Excel, Cash handling")
    t._tbl.getparent().remove(t._tbl)
    save(d, "W4-L2-solved.docx")
    # ---- Lesson 3
    d = starter(4, "W4-employers.docx")
    t = d.tables[0]
    for row in [["Lakeview Library", "Sam Chen", "Mr.", "1 Library Lane", "Penn Yan", "NY", "14527", "Library Clerk", "No"],
                ["Harbor Health", "Jo Rivera", "Ms.", "40 Harbor Street", "Geneva", "NY", "14456", "Scheduler", "Yes"]]:
        cells = t.add_row().cells
        for c, text in zip(cells, row):
            c.text = text
    save(d, "W4-employers-solved.docx")
    d = starter(4, "W4-L3-letter.docx")
    K.mail_merge(d, "formLetters")
    replace_with_field(d, "[Address Block]", 'ADDRESSBLOCK \\f "<<_TITLE0_ >><<_FIRST0_>><< _LAST0_>>" \\l 1033 \\c 2 \\e "United States" \\d', "«AddressBlock»")
    replace_with_field(d, "[Greeting Line]", 'GREETINGLINE \\f "<<_BEFORE_ Dear >><<_TITLE0_ >><<_LAST0_>><<_AFTER_ :>>" \\l 1033 \\e "Dear Sir or Madam:"', "«GreetingLine»")
    replace_with_field(d, "[Position]", "MERGEFIELD Position", "«Position»")
    replace_with_field(d, "[Remote sentence]", 'IF Yes = "Yes" "I\'m experienced with remote work tools." "I\'m glad to work on site."', "I'm glad to work on site.")
    save(d, "W4-L3-letter-solved.docx")
    # the finished merge: five letters, no fields, a section break between letters
    d = K.new_doc()
    names = [r[0] for r in S.EMPLOYERS[1:]] + ["Lakeview Library", "Harbor Health"]
    for i, company in enumerate(names):
        K.para(d, "Alex Morgan")
        K.para(d, company)
        K.para(d, "I'm writing to apply for the position at " + company + ".")
        if i < len(names) - 1:
            p = K.para(d, "Sincerely, Alex Morgan")
            p._p.get_or_add_pPr().append(copy.deepcopy(d.element.body.sectPr))
    save(d, "W4-L3-merged.docx")
    # ---- Lesson 4: labels and a directory
    d = starter(4, "W4-L4-blank.docx")
    K.mail_merge(d, "mailingLabels")
    t = d.add_table(rows=1, cols=3)
    for i, cell in enumerate(t.rows[0].cells):
        p = cell.paragraphs[0]
        if i:
            K.add_field(p, "NEXT", "«Next Record»")
        K.add_field(p, "ADDRESSBLOCK", "«AddressBlock»")
    save(d, "W4-L4-labels.docx")
    d = starter(4, "W4-L4-blank.docx")
    K.mail_merge(d, "catalog")
    p = d.add_paragraph()
    for f in ["Company", "Contact", "Position"]:
        K.add_field(p, "MERGEFIELD " + f, "«" + f + "»")
        p.add_run("\t")
    save(d, "W4-L4-directory.docx")
    # ---- Lesson 5
    d = starter(4, "W4-L5-resume.docx")
    p = K.find(d, "Portfolio:")
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.add_run("Portfolio: ")
    K.hyperlink(d, p, "my portfolio", url="https://www.flcc.edu")
    save(d, "W4-L5-solved.docx")
    # ---- Skill check 4
    d = starter(4, "W4-SC4-letter.docx")
    K.mail_merge(d, "formLetters")
    replace_with_field(d, "[Address Block]", "ADDRESSBLOCK", "«AddressBlock»")
    replace_with_field(d, "[Greeting Line]", "GREETINGLINE", "«GreetingLine»")
    replace_with_field(d, "[Level]", "MERGEFIELD Level", "«Level»")
    replace_with_field(d, "[RenewBy]", "MERGEFIELD RenewBy", "«RenewBy»")
    replace_with_field(d, "[Level sentence]", 'IF Gold = "Gold" "Gold members get a free sunset tour." "Upgrade to Gold for a free sunset tour."', "Gold members get a free sunset tour.")
    d.core_properties.title = "Renewal Letter"
    save(d, "W4-SC4-solved.docx")


# ---- helpers for comments and tracked changes in existing files
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def comment_parts(doc):
    comments = ext = None
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/comments"):
            comments = rel.target_part
        if rel.reltype.endswith("/commentsExtended"):
            ext = rel.target_part
    return comments, ext


def _ext_root(ext):
    from lxml import etree
    return etree.fromstring(ext.blob)


def _ext_save(ext, root):
    from lxml import etree
    ext._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def comment_op(doc, contains, op, reply_text="Thanks, I'll check it.", author="Test Learner"):
    """op: 'reply', 'resolve' or 'delete' on the comment whose text contains `contains`."""
    comments, ext = comment_parts(doc)
    root = comments.element
    target = next(c for c in root.findall(qn("w:comment")) if contains in "".join(t.text or "" for t in c.iter(qn("w:t"))))
    cid = target.get(qn("w:id"))
    pid = target.findall(qn("w:p"))[-1].get("{%s}paraId" % K.W14)
    ex = _ext_root(ext)
    entry = next(e for e in ex if e.get("{%s}paraId" % W15) == pid)
    if op == "resolve":
        entry.set("{%s}done" % W15, "1")
    elif op == "delete":
        root.remove(target)
        ex.remove(entry)
        body = doc.element.body
        for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
            for e in list(body.iter(qn(tag))):
                if e.get(qn("w:id")) == cid:
                    e.getparent().remove(e)
        for ref in list(body.iter(qn("w:commentReference"))):
            if ref.get(qn("w:id")) == cid:
                run = ref.getparent(); run.getparent().remove(run)
    elif op == "reply":
        new_id = str(max(int(c.get(qn("w:id"))) for c in root.findall(qn("w:comment"))) + 1)
        new_pid = "%08X" % (0x20000000 + int(new_id))
        root.append(K.parse_xml('<w:comment %s w:id="%s" w:author="%s" w:date="%s" w:initials="TL"><w:p w14:paraId="%s" w14:textId="77777777"><w:r><w:t>%s</w:t></w:r></w:p></w:comment>'
                                % (K.NSW, new_id, author, K.DATE, new_pid, reply_text)))
        from lxml import etree
        e = etree.SubElement(ex, "{%s}commentEx" % W15)
        e.set("{%s}paraId" % W15, new_pid); e.set("{%s}paraIdParent" % W15, pid); e.set("{%s}done" % W15, "0")
    _ext_save(ext, ex)


def add_comment_on(doc, paragraph_start, text, author="Test Learner"):
    """A new top-level comment on a paragraph, added to the existing comment parts."""
    comments, ext = comment_parts(doc)
    root = comments.element
    new_id = str(max([int(c.get(qn("w:id"))) for c in root.findall(qn("w:comment"))] + [0]) + 10)
    pid = "%08X" % (0x30000000 + int(new_id))
    root.append(K.parse_xml('<w:comment %s w:id="%s" w:author="%s" w:date="%s" w:initials="TL"><w:p w14:paraId="%s" w14:textId="77777777"><w:r><w:t>%s</w:t></w:r></w:p></w:comment>'
                            % (K.NSW, new_id, author, K.DATE, pid, text)))
    ex = _ext_root(ext)
    from lxml import etree
    e = etree.SubElement(ex, "{%s}commentEx" % W15); e.set("{%s}paraId" % W15, pid); e.set("{%s}done" % W15, "0")
    _ext_save(ext, ex)
    p = K.find(doc, paragraph_start)._p
    pPr = p.find(qn("w:pPr"))
    start = K.parse_xml('<w:commentRangeStart %s w:id="%s"/>' % (K.NSW, new_id))
    (pPr.addnext(start) if pPr is not None else p.insert(0, start))
    p.append(K.parse_xml('<w:commentRangeEnd %s w:id="%s"/>' % (K.NSW, new_id)))
    p.append(K.parse_xml('<w:r %s><w:commentReference w:id="%s"/></w:r>' % (K.NSW, new_id)))


def review(doc, accept=True, kinds=("ins", "del", "fmt"), containing=None):
    """Accept or reject tracked changes (optionally only those whose text contains `containing`)."""
    body = doc.element.body
    def text_of(e):
        return "".join((t.text or "") for t in e.iter(qn("w:t"), qn("w:delText")))
    if "fmt" in kinds:
        for ch in list(body.iter(qn("w:rPrChange"))):
            rpr = ch.getparent()
            if not accept:   # reject: put back the old formatting
                old = ch.find(qn("w:rPr"))
                for c in list(rpr):
                    if c is not ch:
                        rpr.remove(c)
                for c in list(old):
                    rpr.append(c)
            rpr.remove(ch)
    for tag in (["w:ins"] if "ins" in kinds else []) + (["w:del"] if "del" in kinds else []):
        for e in list(body.iter(qn(tag))):
            if e.getparent().tag == qn("w:rPr") or (containing and containing not in text_of(e)):
                continue
            keep = (tag == "w:ins") == accept
            if keep:
                for r in list(e):
                    for dt in r.iter(qn("w:delText")):
                        dt.tag = qn("w:t")
                    e.addprevious(r)
            e.getparent().remove(e)


def week5_solved():
    # ---- Lesson 1: comments and changes on the policy
    d = starter(5, "W5-L1-policy.docx")
    add_comment_on(d, "Paddle Pass holders", "How far ahead can non-members book?")
    comment_op(d, "Is 48 hours right", "reply")
    comment_op(d, "must instead of should", "resolve")
    comment_op(d, "Old note", "delete")
    review(d, accept=True, kinds=("fmt",))
    review(d, accept=False, kinds=("ins", "del"), containing="refund")
    review(d, accept=True)
    save(d, "W5-L1-policy-solved.docx")
    # partial: a new comment and a reply, nothing else; formatting accepted, but the refund deletion accepted too
    d = starter(5, "W5-L1-policy.docx")
    add_comment_on(d, "Paddle Pass holders", "How far ahead can non-members book?")
    review(d, accept=True, kinds=("fmt",))
    review(d, accept=True, kinds=("ins", "del"), containing="refund")
    save(d, "W5-L1-policy-partial.docx")
    # ---- Lesson 1, drill 2: tracked edits, locked
    d = starter(5, "W5-L1-draft.docx")
    K.tracking_on(d, locked=True)
    for t, ins in [("Greet every customer", " with a smile"), ("Check the waiver", " and dated"), ("Fit each life jacket", " twice")]:
        K.tracked_insert(K.find(d, t), ins, "Test Learner")
    K.tracked_delete(K.find(d, "Remind renters"), " where the no-wake zone starts", "Test Learner")
    save(d, "W5-L1-draft-solved.docx")
    # ---- Lesson 2: a comparison and a combination
    d = K.new_doc()
    for line, (old, new) in zip(["Dear Ms. Lee:", "I'm writing to apply.", "For two seasons I ran the desk.", "Thank you."],
                                [("", ""), ("", " which I saw on your website"), ("staff", "twenty staff"), ("I would love to", "I'd welcome the chance to")]):
        p = K.para(d, line)
        if new:
            K.tracked_delete(p, old or "x", "Alex Morgan")
            K.tracked_insert(p, new, "Alex Morgan")
    save(d, "W5-L2-compared.docx")
    d = K.new_doc()
    p = K.para(d, "I would love to talk about how I can help your team grow.")
    K.tracked_insert(p, " I'd welcome the chance.", "Priya Shah")
    p = K.para(d, "For two seasons I ran the rental desk.")
    K.tracked_insert(p, " I scheduled twenty staff.", "Sam Ortiz")
    save(d, "W5-L2-combined.docx")
    # ---- Lesson 3: shipped clean
    S = __import__("starters_w2to5")
    d = starter(5, "W5-L3-ship.docx")
    for c in ["parking info", "Link checked"]:
        comment_op(d, c, "delete")
    d.core_properties.author = ""; d.core_properties.last_modified_by = ""
    hidden = K.find(d, "Internal note")
    hidden._p.getparent().remove(hidden._p)
    K.set_compat(d, 15)
    alt_by_name(d, {"Picture 1": "Calm lake at sunrise where Paddle Day takes place."})
    K.header_row(d.tables[0])
    K.find(d, "What to wear").style = "Heading 2"
    link = next(d.element.body.iter(qn("w:hyperlink")))
    next(link.iter(qn("w:t"))).text = "the Paddle Day sign-up page"
    d.core_properties.title = "Paddle Day Volunteer Guide"
    K.protect(d, "comments", password=True)
    save(d, "W5-L3-solved.docx")
    # partial: comments gone and alt text added, nothing else
    d = starter(5, "W5-L3-ship.docx")
    for c in ["parking info", "Link checked"]:
        comment_op(d, c, "delete")
    alt_by_name(d, {"Picture 1": "Calm lake at sunrise."})
    save(d, "W5-L3-partial.docx")
    # ---- Skill check 5
    d = starter(5, "W5-SC5-report.docx")
    comment_op(d, "rental numbers", "reply")
    comment_op(d, "Love this section", "resolve")
    comment_op(d, "Remove before sharing", "delete")
    add_comment_on(d, "Next season", "Should we add a date for the second Paddle Day?")
    review(d, accept=True, kinds=("fmt",))
    review(d, accept=False, kinds=("ins", "del"), containing="volunteers")
    review(d, accept=False, kinds=("ins",), containing="helped by")
    review(d, accept=True)
    h = K.find(d, "Draft numbers")
    h._p.getparent().remove(h._p)
    alt_by_name(d, {"Picture 2": "The lake at sunrise."})
    K.tracking_on(d, locked=True)
    save(d, "W5-SC5-solved.docx")
    # ---- Capstone practice exam
    d = starter(5, "W5-CAP-exam.docx")
    for p in d.paragraphs:
        for r in p.runs:
            if "Lakeshore Paddle Co." in r.text:   # (setting text on other runs would wipe out pictures)
                r.text = r.text.replace("Lakeshore Paddle Co.", "Lakeshore Paddle Company")
    K.letter_page(d) if hasattr(K, "letter_page") else letter_page(d)
    restyle(K.find(d, "Lakeshore Paddle Company Annual Report"), "Title")
    effect_run = K.find(d, "Lakeshore Paddle Company Annual Report").runs[0]
    effect_run._r.get_or_add_rPr().append(K.parse_xml('<w14:shadow xmlns:w14="%s" w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl"><w14:srgbClr w14:val="000000"/></w14:shadow>' % K.W14))
    for t in ["Our season", "Highlights", "Rentals", "Goals for next year", "Staff", "Next steps"]:
        restyle(next(p for p in d.paragraphs if p.text == t), "Heading 1")
    ps = K.find(d, "Lakeshore Paddle Company had its busiest")
    ps.paragraph_format.line_spacing = 1.15; ps.paragraph_format.space_after = Pt(6)
    clear_run_formatting(K.find(d, "Thank you to every renter"))
    fp = d.sections[0].footer.paragraphs[0]
    fp.text = "© 2026  "
    K.add_simple_field(fp, "PAGE", "1")
    d.core_properties.title = "Annual Report 2026"
    prep = K.find(d, "Prepared for")
    toc = OxmlElement("w:p"); prep._p.addnext(toc)
    from docx.text.paragraph import Paragraph
    K.add_field(Paragraph(toc, prep._parent), 'TOC \\o "1-3" \\h \\z \\u', "")
    num = K.add_numbering(d, "decimal", "%1.")
    for g in ["Add four kayaks", "Hold a second Paddle Day", "Train every guide"]:
        K.set_list(K.find(d, g), num)
    lines_to_table(d, "Name\tRole", 4, 3)
    K.header_row(table_with(d, "Paddleboard"))
    hl = K.find(d, "Paddle Day brought first-time")
    hl.runs[0].text = hl.runs[0].text.replace("Paddle Day brought", "Paddle Day™ brought")
    # Highlights paragraph in its own two-column section
    before = K.find(d, "Highlights")
    first = copy.deepcopy(d.element.body.sectPr)
    for c in first.findall(qn("w:cols")): first.remove(c)
    before._p.get_or_add_pPr().append(first)
    two = copy.deepcopy(d.element.body.sectPr)
    K.set_section_type(d.sections[-1], "continuous")
    hl._p.get_or_add_pPr().append(two)
    two_cols = hl._p.pPr.find(qn("w:sectPr"))
    cols = two_cols.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols"); two_cols.append(cols)
    cols.set(qn("w:num"), "2")
    t = OxmlElement("w:type"); t.set(qn("w:val"), "continuous"); two_cols.insert(0, t)
    K.footnotes(d, [(K.find(d, "Lakeshore Paddle Company had its busiest"), "The sunset series and Paddle Day.")])
    alt_by_name(d, {"Report photo": "Kayaks lined up on the beach at sunrise."})
    i = picture_index(d, "Report photo")
    K.picture_effects(d, i, style=True)
    K.float_picture(d, i, wrap="square", halign="right", valign="top")
    comment_op(d, "meeting date", "reply")
    review(d, accept=True)
    save(d, "W5-CAP-solved.docx")


if __name__ == "__main__":
    lesson2_solved(); lesson2_partial()
    lesson3_solved(); lesson3_partial()
    lesson4_solved(); lesson4_partial()
    sc1_solved()
    week2_solved()
    week3_solved()
    week4_solved()
    week5_solved()
