# -*- coding: utf-8 -*-
"""
xlsxkit.py - helpers for making the Excel course's practice workbooks.

Practice workbooks are written with XlsxWriter, which makes files exactly the
way Excel does (including the saved results of formulas, which the checker
reads). A few things XlsxWriter can't write - named cell styles, threaded
comments, scenarios, data tables, icons, chart style parts, workbook
protection and the "remove personal information" setting - are added
afterwards by small functions below that edit the file's XML directly.

Everything in these files is fictional: businesses, people and data.
Style rule: hyphens, never em dashes.
"""
import io, re, zipfile, uuid
import xlsxwriter
from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


# ---------------------------------------------------------------------------
# Making a workbook
# ---------------------------------------------------------------------------
def new_book(path, title="", author="Lakeshore Paddle Company", keywords=""):
    """A new XlsxWriter workbook with plain, predictable properties."""
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False, "strings_to_urls": False})
    props = {"author": author, "company": "Lakeshore Paddle Company (fictional)"}
    if title:
        props["title"] = title
    if keywords:
        props["keywords"] = keywords
    wb.set_properties(props)
    return wb


class Fmt:
    """Cache XlsxWriter formats so the same look is one format: F(wb)(bold=True)."""
    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def __call__(self, **props):
        key = tuple(sorted(props.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(props)
        return self.cache[key]


def col_num(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s


def rc(ref):
    """'B3' -> (row, col) counted from 0, the way XlsxWriter wants them."""
    m = re.match(r"^\$?([A-Z]+)\$?(\d+)$", ref.upper())
    return int(m.group(2)) - 1, col_num(m.group(1)) - 1


# ---------------------------------------------------------------------------
# Editing a finished file's XML
# ---------------------------------------------------------------------------
def _read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write_zip(path, parts):
    buf = io.BytesIO()
    order = ["[Content_Types].xml", "_rels/.rels"] + sorted(n for n in parts if n not in ("[Content_Types].xml", "_rels/.rels"))
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            if n in parts:
                z.writestr(n, parts[n])
    with open(path, "wb") as f:
        f.write(buf.getvalue())


def _xml(data):
    return etree.fromstring(data)


def _bytes(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _sheet_part(parts, index):
    """The part name of the sheet at tab position `index` (0 = first)."""
    wb = _xml(parts["xl/workbook.xml"])
    sheets = wb.find("{%s}sheets" % NS_MAIN)
    rid = sheets[index].get("{%s}id" % NS_R)
    rels = _xml(parts["xl/_rels/workbook.xml.rels"])
    for r in rels:
        if r.get("Id") == rid:
            return "xl/" + r.get("Target").lstrip("/").replace("xl/", "")
    raise KeyError(index)


def _rels_name(part):
    d, n = part.rsplit("/", 1)
    return d + "/_rels/" + n + ".rels"


def _add_rel(parts, part, rel_type, target, external=False):
    """Add a relationship from `part`; returns the new id."""
    name = _rels_name(part)
    if name in parts:
        root = _xml(parts[name])
    else:
        root = etree.Element("{%s}Relationships" % NS_REL, nsmap={None: NS_REL})
    ids = {r.get("Id") for r in root}
    n = 1
    while "rId%d" % n in ids:
        n += 1
    rel = etree.SubElement(root, "{%s}Relationship" % NS_REL)
    rel.set("Id", "rId%d" % n)
    rel.set("Type", rel_type if rel_type.startswith("http") else REL_BASE + rel_type)
    rel.set("Target", target)
    if external:
        rel.set("TargetMode", "External")
    parts[name] = _bytes(root)
    return "rId%d" % n


def _add_content_type(parts, part_name=None, content_type=None, extension=None):
    ct = _xml(parts["[Content_Types].xml"])
    if extension:
        if not any(e.get("Extension") == extension for e in ct if e.tag.endswith("Default")):
            el = etree.SubElement(ct, "{%s}Default" % NS_CT)
            el.set("Extension", extension)
            el.set("ContentType", content_type)
    else:
        el = etree.SubElement(ct, "{%s}Override" % NS_CT)
        el.set("PartName", "/" + part_name)
        el.set("ContentType", content_type)
    parts["[Content_Types].xml"] = _bytes(ct)


# The order Excel expects the children of <worksheet> in. New elements are
# inserted in the right place so the file stays valid.
WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
            "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
            "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
            "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
            "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def _insert_ws_child(ws, el):
    name = etree.QName(el).localname
    pos = WS_ORDER.index(name)
    for i, child in enumerate(ws):
        cname = etree.QName(child).localname
        if cname in WS_ORDER and WS_ORDER.index(cname) > pos:
            ws.insert(i, el)
            return
    ws.append(el)


def patch(path, fn):
    """Open the file, let fn(parts) change any parts, and save it again."""
    parts = _read_zip(path)
    fn(parts)
    _write_zip(path, parts)


# ---------------------------------------------------------------------------
# Patches
# ---------------------------------------------------------------------------
BUILTIN_STYLE_IDS = {"Title": 15, "Heading 1": 16, "Heading 2": 17, "Heading 3": 18, "Heading 4": 19, "Input": 20,
                     "Output": 21, "Calculation": 22, "Total": 25, "Good": 26, "Bad": 27, "Neutral": 28, "Note": 10}


def cell_style(path, sheet_index, cells, name, builtin=True):
    """Give cells a named cell style (Home > Cell Styles), the way Excel saves it."""
    def fn(parts):
        st = _xml(parts["xl/styles.xml"])
        q = lambda t: "{%s}%s" % (NS_MAIN, t)
        style_xfs = st.find(q("cellStyleXfs"))
        cell_xfs = st.find(q("cellXfs"))
        styles = st.find(q("cellStyles"))
        # Reuse the style if it is already there.
        existing = [s for s in styles if s.get("name") == name]
        if existing:
            xf_id = int(existing[0].get("xfId"))
        else:
            sx = etree.SubElement(style_xfs, q("xf"))
            for k in ("numFmtId", "fontId", "fillId", "borderId"):
                sx.set(k, "0")
            style_xfs.set("count", str(len(style_xfs)))
            xf_id = len(style_xfs) - 1
            cs = etree.SubElement(styles, q("cellStyle"))
            cs.set("name", name)
            cs.set("xfId", str(xf_id))
            if builtin and name in BUILTIN_STYLE_IDS:
                cs.set("builtinId", str(BUILTIN_STYLE_IDS[name]))
            styles.set("count", str(len(styles)))
        # Each styled cell gets its own copy of its format, pointing at the style.
        sp = _sheet_part(parts, sheet_index)
        ws = _xml(parts[sp])
        want = {c.upper() for c in cells}
        for c in ws.iter(q("c")):
            if c.get("r") in want:
                old = int(c.get("s") or 0)
                new = etree.fromstring(etree.tostring(cell_xfs[old]))
                new.set("xfId", str(xf_id))
                cell_xfs.append(new)
                c.set("s", str(len(cell_xfs) - 1))
        cell_xfs.set("count", str(len(cell_xfs)))
        parts["xl/styles.xml"] = _bytes(st)
        parts[sp] = _bytes(ws)
    patch(path, fn)


def set_theme(path, name, colors_name=None, accent1=None):
    """Rename the theme (as Page Layout > Themes does) and optionally its colors."""
    def fn(parts):
        t = parts["xl/theme/theme1.xml"].decode("utf8")
        t = re.sub(r'(<a:theme [^>]*name=")[^"]*"', r'\g<1>%s"' % name, t, count=1)
        if colors_name:
            t = re.sub(r'(<a:clrScheme name=")[^"]*"', r'\g<1>%s"' % colors_name, t, count=1)
        if accent1:
            t = re.sub(r'(<a:accent1>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}', r'\g<1>%s' % accent1.lstrip("#"), t, count=1)
        parts["xl/theme/theme1.xml"] = t.encode("utf8")
    patch(path, fn)


def add_windows(path, n=2):
    """Save the workbook with n windows open (View > New Window)."""
    def fn(parts):
        wb = _xml(parts["xl/workbook.xml"])
        views = wb.find("{%s}bookViews" % NS_MAIN)
        first = views[0]
        for i in range(n - len(views)):
            v = etree.fromstring(etree.tostring(first))
            v.set("xWindow", str(9600 + 200 * i))
            views.append(v)
        parts["xl/workbook.xml"] = _bytes(wb)
    patch(path, fn)


def remove_personal_info(path):
    """What Inspect Document > Remove All (Document Properties and Personal Information) does."""
    def fn(parts):
        wb = _xml(parts["xl/workbook.xml"])
        pr = wb.find("{%s}workbookPr" % NS_MAIN)
        if pr is None:
            pr = etree.Element("{%s}workbookPr" % NS_MAIN)
            wb.insert(1, pr)
        pr.set("filterPrivacy", "1")
        parts["xl/workbook.xml"] = _bytes(wb)
        core = parts["docProps/core.xml"].decode("utf8")
        core = re.sub(r"<dc:creator>[^<]*</dc:creator>", "", core)
        core = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "", core)
        parts["docProps/core.xml"] = core.encode("utf8")
    patch(path, fn)


def set_last_modified_by(path, name):
    def fn(parts):
        core = parts["docProps/core.xml"].decode("utf8")
        if "<cp:lastModifiedBy>" in core:
            core = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "<cp:lastModifiedBy>%s</cp:lastModifiedBy>" % name, core)
        else:
            core = core.replace("</cp:coreProperties>", "<cp:lastModifiedBy>%s</cp:lastModifiedBy></cp:coreProperties>" % name)
        parts["docProps/core.xml"] = core.encode("utf8")
    patch(path, fn)


def show_formulas(path, sheet_index, on=True):
    """Formulas > Show Formulas, saved on the sheet."""
    def fn(parts):
        sp = _sheet_part(parts, sheet_index)
        ws = _xml(parts[sp])
        sv = ws.find("{%s}sheetViews/{%s}sheetView" % (NS_MAIN, NS_MAIN))
        if on:
            sv.set("showFormulas", "1")
        elif "showFormulas" in sv.attrib:
            del sv.attrib["showFormulas"]
        parts[sp] = _bytes(ws)
    patch(path, fn)


def protect_structure(path, password=True):
    """Review > Protect Workbook (structure)."""
    def fn(parts):
        wb = _xml(parts["xl/workbook.xml"])
        p = etree.Element("{%s}workbookProtection" % NS_MAIN)
        p.set("lockStructure", "1")
        if password:
            p.set("workbookPassword", "CC4B")   # Excel's short hash for "paddle" style practice passwords
        # workbookProtection comes right after workbookPr / before bookViews
        bv = wb.find("{%s}bookViews" % NS_MAIN)
        wb.insert(list(wb).index(bv), p)
        parts["xl/workbook.xml"] = _bytes(wb)
    patch(path, fn)


def add_scenarios(path, sheet_index, scenarios, sqref=None):
    """Data > What-If Analysis > Scenario Manager: [{name, inputs: {"B2": 30}}]."""
    def fn(parts):
        sp = _sheet_part(parts, sheet_index)
        ws = _xml(parts[sp])
        sc = etree.Element("{%s}scenarios" % NS_MAIN)
        if sqref:
            sc.set("sqref", sqref)
        for s in scenarios:
            e = etree.SubElement(sc, "{%s}scenario" % NS_MAIN)
            e.set("name", s["name"])
            e.set("locked", "1")
            e.set("count", str(len(s["inputs"])))
            e.set("user", "Learner")
            for ref, val in s["inputs"].items():
                i = etree.SubElement(e, "{%s}inputCells" % NS_MAIN)
                i.set("r", ref)
                i.set("val", str(val))
        _insert_ws_child(ws, sc)
        parts[sp] = _bytes(ws)
    patch(path, fn)


def add_consolidate(path, sheet_index, refs, function="sum", labels=True, links=False):
    """Data > Consolidate: refs = [("Priya", "A3:C8"), ...]. Excel keeps the settings on the result sheet."""
    def fn(parts):
        sp = _sheet_part(parts, sheet_index)
        ws = _xml(parts[sp])
        dc = etree.Element("{%s}dataConsolidate" % NS_MAIN)
        if function != "sum":
            dc.set("function", function)
        if labels:
            dc.set("leftLabels", "1")
            dc.set("topLabels", "1")
        if links:
            dc.set("link", "1")
        drs = etree.SubElement(dc, "{%s}dataRefs" % NS_MAIN)
        drs.set("count", str(len(refs)))
        for sheet, ref in refs:
            d = etree.SubElement(drs, "{%s}dataRef" % NS_MAIN)
            d.set("ref", ref)
            d.set("sheet", sheet)
        _insert_ws_child(ws, dc)
        parts[sp] = _bytes(ws)
    patch(path, fn)


def add_data_table(path, sheet_index, ref, input_cell, by_row=False):
    """Turn already-written result cells into a What-If data table: {=TABLE(,B3)}."""
    def fn(parts):
        sp = _sheet_part(parts, sheet_index)
        ws = _xml(parts[sp])
        first = ref.split(":")[0]
        for c in ws.iter("{%s}c" % NS_MAIN):
            if c.get("r") == first:
                f = etree.Element("{%s}f" % NS_MAIN)
                f.set("t", "dataTable")
                f.set("ref", ref)
                f.set("dt2D", "0")
                f.set("dtr", "1" if by_row else "0")
                f.set("r1", input_cell)
                c.insert(0, f)
        parts[sp] = _bytes(ws)
    patch(path, fn)


def add_connection(path, name="Query - rentals"):
    """The data connection a Get Data (Power Query) import leaves in the workbook."""
    def fn(parts):
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<connections xmlns="%s"><connection id="1" keepAlive="1" name="%s" description="Connection to the query in the workbook." '
               'type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=rentals;" '
               'command="SELECT * FROM [rentals]"/></connection></connections>') % (NS_MAIN, name)
        parts["xl/connections.xml"] = xml.encode("utf8")
        _add_rel(parts, "xl/workbook.xml", "connections", "connections.xml")
        _add_content_type(parts, "xl/connections.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml")
    patch(path, fn)


def chart_style(path, chart_number, style_id):
    """Excel 2013+ saves each chart's style (Chart Design > Chart Styles) in its own part."""
    def fn(parts):
        name = "xl/charts/style%d.xml" % chart_number
        parts[name] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" '
                       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="%d"/>' % style_id).encode("utf8")
        _add_rel(parts, "xl/charts/chart%d.xml" % chart_number, "http://schemas.microsoft.com/office/2011/relationships/chartStyle", "style%d.xml" % chart_number)
        _add_content_type(parts, name, "application/vnd.ms-office.chartstyle+xml")
    patch(path, fn)


SVG_ICON = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"><path d="M8 60h80l-12 14H20z" fill="#0E7C86"/><path d="M48 14v40M48 14l20 30H48" fill="none" stroke="#0E7C86" stroke-width="5"/></svg>'


def make_icon(path, drawing_number=1, picture_index=0):
    """Turn an inserted PNG into an icon (Insert > Icons saves an SVG plus a PNG copy)."""
    def fn(parts):
        dname = "xl/drawings/drawing%d.xml" % drawing_number
        d = parts[dname].decode("utf8")
        parts["xl/media/icon%d.svg" % drawing_number] = SVG_ICON
        _add_content_type(parts, content_type="image/svg+xml", extension="svg")
        rid = _add_rel(parts, dname, "image", "../media/icon%d.svg" % drawing_number)
        blips = [m.start() for m in re.finditer(r"<a:blip ", d)]
        start = blips[picture_index]
        end = d.index("/>", start) if d.find("/>", start) < d.find(">", start) + 1 else d.index(">", start)
        ext = ('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
               '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="%s"/></a:ext></a:extLst>' % rid)
        tag_end = d.index(">", start)
        if d[tag_end - 1] == "/":
            d = d[:tag_end - 1] + ">" + ext + "</a:blip>" + d[tag_end + 1:]
        else:
            d = d[:tag_end + 1] + ext + d[tag_end + 1:]
        parts[dname] = d.encode("utf8")
    patch(path, fn)


def mark_decorative(path, drawing_number=1, object_index=0):
    """Edit Alt Text > Mark as decorative."""
    def fn(parts):
        dname = "xl/drawings/drawing%d.xml" % drawing_number
        d = parts[dname].decode("utf8")
        starts = [m.start() for m in re.finditer(r"<xdr:cNvPr ", d)]
        s = starts[object_index]
        e = d.index(">", s)
        ext = ('<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
               '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>')
        if d[e - 1] == "/":
            d = d[:e - 1] + ">" + ext + "</xdr:cNvPr>" + d[e + 1:]
        else:
            d = d[:e + 1] + ext + d[e + 1:]
        parts[dname] = d.encode("utf8")
    patch(path, fn)


def add_threaded_comments(path, sheet_index, threads):
    """
    Modern (threaded) comments: [{ref, author, text, replies: [(author, text)], done}].
    Write the cell's classic note with write_comment(..., "[Threaded comment]...") first,
    as Excel does for older versions; this adds the threaded part and the people list.
    """
    def fn(parts):
        people = {}
        for t in threads:
            for a in [t["author"]] + [r[0] for r in t.get("replies", [])]:
                people.setdefault(a, "{%s}" % str(uuid.uuid5(uuid.NAMESPACE_DNS, a)).upper())
        if "xl/persons/person.xml" not in parts:
            px = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><personList xmlns="%s" xmlns:x="%s">' % (NS_TC, NS_MAIN)]
            for name, pid in people.items():
                px.append('<person displayName="%s" id="%s" userId="%s" providerId="None"/>' % (name, pid, name))
            px.append("</personList>")
            parts["xl/persons/person.xml"] = "".join(px).encode("utf8")
            _add_rel(parts, "xl/workbook.xml", "http://schemas.microsoft.com/office/2017/10/relationships/person", "persons/person.xml")
            _add_content_type(parts, "xl/persons/person.xml", "application/vnd.ms-excel.person+xml")
        sp = _sheet_part(parts, sheet_index)
        n = 1
        while "xl/threadedComments/threadedComment%d.xml" % n in parts:
            n += 1
        name = "xl/threadedComments/threadedComment%d.xml" % n
        tx = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ThreadedComments xmlns="%s" xmlns:x="%s">' % (NS_TC, NS_MAIN)]
        for i, t in enumerate(threads):
            tid = "{%08X-0000-4000-8000-%012X}" % (sheet_index + 1, i + 1)
            done = ' done="1"' if t.get("done") else ""
            tx.append('<threadedComment ref="%s" dT="2027-05-0%dT10:00:00.00" personId="%s" id="%s"%s><text>%s</text></threadedComment>'
                      % (t["ref"], (i % 8) + 1, people[t["author"]], tid, done, _esc(t["text"])))
            for j, (who, text) in enumerate(t.get("replies", [])):
                rid = "{%08X-0000-4000-9000-%012X}" % (sheet_index + 1, (i + 1) * 100 + j)
                tx.append('<threadedComment ref="%s" dT="2027-05-0%dT11:00:00.00" personId="%s" id="%s" parentId="%s"><text>%s</text></threadedComment>'
                          % (t["ref"], (i % 8) + 1, people[who], rid, tid, _esc(text)))
        tx.append("</ThreadedComments>")
        parts[name] = "".join(tx).encode("utf8")
        _add_rel(parts, sp, "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", "../threadedComments/threadedComment%d.xml" % n)
        _add_content_type(parts, name, "application/vnd.ms-excel.threadedcomments+xml")
    patch(path, fn)


def _esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def table_filter(path, table_number, col_id, values):
    """A filter on one column of an Excel table (the arrow button > check boxes)."""
    def fn(parts):
        name = "xl/tables/table%d.xml" % table_number
        t = _xml(parts[name])
        af = t.find("{%s}autoFilter" % NS_MAIN)
        fc = etree.SubElement(af, "{%s}filterColumn" % NS_MAIN)
        fc.set("colId", str(col_id))
        fs = etree.SubElement(fc, "{%s}filters" % NS_MAIN)
        for v in values:
            f = etree.SubElement(fs, "{%s}filter" % NS_MAIN)
            f.set("val", v)
        parts[name] = _bytes(t)
    patch(path, fn)


def external_link(path, file_name, sheet_name, cells):
    """
    A link to another workbook, the way Excel saves ='[Review-Priya.xlsx]Budget'!B5:
    the formula says [1], and an "external link" part says what [1] is.
    cells: {"B5": 120.0} - the values Excel remembers from the other file.
    """
    def fn(parts):
        n = 1
        while "xl/externalLinks/externalLink%d.xml" % n in parts:
            n += 1
        name = "xl/externalLinks/externalLink%d.xml" % n
        rows = {}
        for ref, v in cells.items():
            r, c = rc(ref)
            rows.setdefault(r + 1, []).append((ref, v))
        data = "".join('<row r="%d">%s</row>' % (r, "".join('<cell r="%s"><v>%s</v></cell>' % (ref, v) for ref, v in sorted(cs))) for r, cs in sorted(rows.items()))
        parts[name] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="%s" xmlns:r="%s"><externalBook r:id="rId1">'
                       '<sheetNames><sheetName val="%s"/></sheetNames><sheetDataSet><sheetData sheetId="0">%s</sheetData></sheetDataSet></externalBook></externalLink>'
                       % (NS_MAIN, NS_R, _esc(sheet_name), data)).encode("utf8")
        parts[_rels_name(name)] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s"><Relationship Id="rId1" '
                                   'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="%s" TargetMode="External"/></Relationships>'
                                   % (NS_REL, _esc(file_name))).encode("utf8")
        rid = _add_rel(parts, "xl/workbook.xml", "externalLink", "externalLinks/externalLink%d.xml" % n)
        _add_content_type(parts, name, "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
        wb = _xml(parts["xl/workbook.xml"])
        ext = wb.find("{%s}externalReferences" % NS_MAIN)
        if ext is None:
            ext = etree.Element("{%s}externalReferences" % NS_MAIN)
            sheets = wb.find("{%s}sheets" % NS_MAIN)
            wb.insert(list(wb).index(sheets) + 1, ext)
        e = etree.SubElement(ext, "{%s}externalReference" % NS_MAIN)
        e.set("{%s}id" % NS_R, rid)
        parts["xl/workbook.xml"] = _bytes(wb)
    patch(path, fn)


