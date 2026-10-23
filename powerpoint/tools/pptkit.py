# -*- coding: utf-8 -*-
"""
pptkit.py - helpers for making the PowerPoint course's practice decks.

Practice decks are written with python-pptx, which makes ordinary slides,
text, pictures, tables, charts and video the way PowerPoint does. Things
python-pptx can't write - transitions, animations, sections, custom shows,
slide show settings, modern comments, handout masters, new layouts, Zoom,
SmartArt, 3D models, ink, icons, embedded fonts, Mark as Final and the
"remove personal information" setting - are added by the functions below,
which edit the slide XML (before saving) or the file's parts (after saving).

Everything in these files is fictional: businesses, people and data.
Style rule: hyphens, never em dashes.
"""
import io, os, re, struct, subprocess, uuid, zipfile, json, math
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from PIL import Image, ImageDraw, ImageFont

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p188": "http://schemas.microsoft.com/office/powerpoint/2018/8/main",
}
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
TEAL = (14, 124, 134)
NAVY = (31, 56, 100)


def qn(tag):
    """'a:t' -> '{namespace}t' for lxml."""
    pre, name = tag.split(":")
    return "{%s}%s" % (NS[pre], name)


def X(xml):
    """Parse a fragment that uses the a:, p:, r:, p14:... prefixes."""
    decl = " ".join('xmlns:%s="%s"' % (k, v) for k, v in NS.items())
    root = etree.fromstring("<wrap %s>%s</wrap>" % (decl, xml))
    return list(root)


def esc(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


# ---------------------------------------------------------------------------
# A new deck
# ---------------------------------------------------------------------------
def new_deck(widescreen=True, title="", author="Lakeshore Paddle Company"):
    """A new presentation from python-pptx's Office Theme template.
    The template is 4:3; widescreen=True makes it 16:9 and stretches the master
    and layouts sideways, the way PowerPoint does when you change the size."""
    prs = Presentation()
    if widescreen:
        factor = 12192000 / 9144000
        prs.slide_width = Emu(12192000)
        prs.slide_height = Emu(6858000)
        pres = prs.part._element
        sz = pres.find(qn("p:sldSz"))
        if "type" in sz.attrib:
            del sz.attrib["type"]
        for part_el in [prs.slide_master._element] + [l._element for l in prs.slide_layouts]:
            for off in part_el.iter(qn("a:off")):
                off.set("x", str(int(int(off.get("x")) * factor)))
            for ext in part_el.iter(qn("a:ext")):
                if ext.get("cx") is not None:
                    ext.set("cx", str(int(int(ext.get("cx")) * factor)))
    cp = prs.core_properties
    cp.author = author
    cp.last_modified_by = author
    cp.title = title
    cp.keywords = ""
    cp.comments = ""
    cp.subject = ""
    return prs


def layout(prs, name):
    for l in prs.slide_layouts:
        if l.name == name:
            return l
    raise KeyError(name)


def add_slide(prs, layout_name, title=None, body=None, notes=None, body_idx=1):
    """A slide with the named layout. body: list of text or (text, level) pairs."""
    s = prs.slides.add_slide(layout(prs, layout_name))
    if title is not None and s.shapes.title is not None:
        s.shapes.title.text = title
    if body is not None:
        ph = [p for p in s.placeholders if p.placeholder_format.idx == body_idx][0]
        fill_text(ph.text_frame, body)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def fill_text(tf, lines):
    """Put lines into a text frame: each item is text or (text, level)."""
    tf.clear()
    first = True
    for item in lines:
        text, lvl = (item, 0) if isinstance(item, str) else item
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        para.text = text
        para.level = lvl
        first = False


def text_box(slide, x, y, w, h, lines, size=None, font=None, bold=False, name=None, color=None):
    """A text box (Insert > Text Box) with plain paragraphs (no bullets)."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    fill_text(tf, lines)
    for p in tf.paragraphs:
        for r in p.runs:
            if size: r.font.size = Pt(size)
            if font: r.font.name = font
            if bold: r.font.bold = True
            if color: r.font.color.rgb = RGBColor(*color)
    if name:
        tb.name = name
    return tb


def find_shape(slide, says=None, name=None):
    for sh in slide.shapes:
        if name and sh.name == name:
            return sh
        if says and sh.has_text_frame and says.lower() in sh.text_frame.text.lower():
            return sh
    raise KeyError(says or name)


def set_alt(shape, text="", decorative=False):
    """View Alt Text: a description, or Mark as decorative."""
    nv = shape._element.find(".//" + qn("p:cNvPr"))
    if text:
        nv.set("descr", text)
    if decorative:
        nv.append(X('<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>')[0])


def hide(slide):
    slide._element.set("show", "0")


def set_runs(shape, size=None, font=None, bold=None, italic=None, color=None, wordart=False):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if size: r.font.size = Pt(size)
            if font: r.font.name = font
            if bold is not None: r.font.bold = bold
            if italic is not None: r.font.italic = italic
            if color: r.font.color.rgb = RGBColor(*color)
            if wordart:
                rPr = r._r.get_or_add_rPr()
                # WordArt "Fill: Teal, Accent color; Outline" style: an outline and a shadow on the letters.
                ln = X('<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln>')[0]
                eff = X('<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>')[0]
                rPr.insert(0, ln)
                # effectLst must come after fills; put it right after ln if there is no fill
                fill = rPr.find(qn("a:solidFill"))
                (fill if fill is not None else ln).addnext(eff)


def columns(shape, n, spacing_in=0.3):
    body = shape.text_frame._txBody.find(qn("a:bodyPr"))
    body.set("numCol", str(n))
    body.set("spcCol", str(int(spacing_in * 914400)))


def bullets(shape, kind="bullet", levels=None):
    """Turn a text box's paragraphs into a list. kind: bullet or number."""
    for i, p in enumerate(shape.text_frame.paragraphs):
        if levels:
            p.level = levels[i]
        pPr = p._p.get_or_add_pPr()
        for old in pPr.findall(qn("a:buNone")) + pPr.findall(qn("a:buChar")) + pPr.findall(qn("a:buAutoNum")):
            pPr.remove(old)
        pPr.set("marL", str(342900 + 457200 * (levels[i] if levels else 0)))
        pPr.set("indent", "-342900")
        pPr.append(X('<a:buFont typeface="Arial"/>')[0] if kind == "bullet" else X('<a:buFont typeface="+mj-lt"/>')[0])
        pPr.append(X('<a:buChar char="&#8226;"/>')[0] if kind == "bullet" else X('<a:buAutoNum type="arabicPeriod"/>')[0])


def footer_placeholders(prs, slide, footer_text=None, number=True, date=False):
    """What Insert > Header & Footer puts on a slide: copies of the layout's footer placeholders."""
    lay = slide.slide_layout
    tree = slide.shapes._spTree
    for ph in lay.placeholders:
        t = ph.placeholder_format.type
        tname = {13: "sldNum", 15: "ftr", 16: "dt"}.get(int(t)) if t is not None else None
        if tname is None or (tname == "sldNum" and not number) or (tname == "ftr" and footer_text is None) or (tname == "dt" and not date):
            continue
        sid = max(int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))) + 1
        text = '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:t>%d</a:t></a:fld>' % (list(prs.slides).index(slide) + 1) if tname == "sldNum" else \
            ('<a:r><a:t>%s</a:t></a:r>' % esc(footer_text) if tname == "ftr" else '<a:fld id="{1A2B3C4D-0000-4000-8000-000000000001}" type="datetime1"><a:t>6/1/2027</a:t></a:fld>')
        tree.append(X('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s Placeholder %d"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="%s" sz="quarter" idx="%d"/></p:nvPr></p:nvSpPr>'
                      '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>'
                      % (sid, {"sldNum": "Slide Number", "ftr": "Footer", "dt": "Date"}[tname], sid, tname, ph.placeholder_format.idx, text))[0])


# ---------------------------------------------------------------------------
# Pictures and media made on the spot (all original, all fictional)
# ---------------------------------------------------------------------------
def _font(size):
    for f in ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"]:
        if os.path.exists(f):
            return ImageFont.truetype(f, size)
    return ImageFont.load_default()


def make_photo(path, w=900, h=600, seed=1, label=None):
    """A simple 'photo' of a lake with a kayak: sky, hills, water and a boat."""
    im = Image.new("RGB", (w, h), (170, 210, 235))
    d = ImageDraw.Draw(im)
    for i in range(h // 2):   # sky gradient
        c = (150 + i * 60 // (h // 2), 200 + i * 30 // (h // 2), 240)
        d.line([(0, i), (w, i)], fill=c)
    d.polygon([(0, h * 0.55), (w * 0.3, h * 0.35), (w * 0.55, h * 0.5), (w * 0.8, h * 0.32), (w, h * 0.5), (w, h * 0.6), (0, h * 0.6)], fill=(70, 120, 80))
    d.rectangle([0, h * 0.58, w, h], fill=(40, 110, 150))
    for k in range(12):
        y = h * 0.62 + k * h * 0.03
        d.line([(w * 0.05 * (k % 5), y), (w * 0.05 * (k % 5) + w * 0.2, y)], fill=(90, 160, 190), width=2)
    bx = w * (0.25 + 0.1 * (seed % 4))
    d.polygon([(bx, h * 0.74), (bx + w * 0.32, h * 0.74), (bx + w * 0.28, h * 0.79), (bx + w * 0.04, h * 0.79)], fill=(230, 110, 40))
    d.ellipse([bx + w * 0.13, h * 0.66, bx + w * 0.17, h * 0.72], fill=(250, 220, 180))
    d.line([(bx + w * 0.05, h * 0.62), (bx + w * 0.27, h * 0.78)], fill=(60, 60, 60), width=4)
    if label:
        d.text((20, 20), label, fill=(255, 255, 255), font=_font(28))
    im.save(path, quality=88)


def make_logo(path, w=400, h=160):
    im = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(im)
    d.ellipse([10, 10, h - 10, h - 10], fill=TEAL)
    d.polygon([(30, h * 0.6), (h - 30, h * 0.6), (h - 45, h * 0.72), (45, h * 0.72)], fill=(255, 255, 255))
    d.text((h + 10, h * 0.25), "Lakeshore", fill=NAVY, font=_font(40))
    d.text((h + 10, h * 0.58), "Paddle Company", fill=TEAL, font=_font(26))
    im.save(path)


def make_icon_png(path, kind="wave", size=128, color=TEAL):
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    if kind == "wave":
        for k in range(3):
            d.arc([10, 20 + k * 30, size - 10, 70 + k * 30], 200, 340, fill=color, width=8)
    elif kind == "sun":
        d.ellipse([34, 34, size - 34, size - 34], fill=color)
        for a in range(0, 360, 45):
            x, y = size / 2 + 50 * math.cos(math.radians(a)), size / 2 + 50 * math.sin(math.radians(a))
            d.line([(size / 2, size / 2), (x, y)], fill=color, width=6)
    else:  # paddle
        d.line([(30, size - 30), (size - 40, 40)], fill=color, width=10)
        d.ellipse([size - 60, 14, size - 14, 60], fill=color)
    im.save(path)


SVG_ICON = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"><path d="M10 60 Q48 30 86 60" stroke="#0E7C86" stroke-width="8" fill="none"/></svg>'
# The SVG copies of the three practice icons (they match make_icon_png's pictures).
SVG_ICONS = {
    "wave": b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128" fill="none" stroke="#0E7C86" stroke-width="8"><path d="M12 40 Q64 10 116 40"/><path d="M12 70 Q64 40 116 70"/><path d="M12 100 Q64 70 116 100"/></svg>',
    "sun": b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128"><circle cx="64" cy="64" r="30" fill="#0E7C86"/><g stroke="#0E7C86" stroke-width="6"><path d="M64 4V24M64 104V124M4 64H24M104 64H124M21 21L35 35M93 93L107 107M21 107L35 93M93 35L107 21"/></g></svg>',
    "paddle": b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128"><path d="M30 98L88 40" stroke="#0E7C86" stroke-width="10"/><ellipse cx="91" cy="37" rx="23" ry="23" fill="#0E7C86"/></svg>',
}


def make_video(path, seconds=8, text="Lakeshore Paddle Company"):
    """A short, silent practice video (a colored card with moving text)."""
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "lavfi", "-i", "color=c=0x0E7C86:s=640x360:d=%d" % seconds,
                    "-vf", "drawtext=text='%s':fontcolor=white:fontsize=34:x=(w-tw)/2:y=h/2-20+20*sin(t)" % text.replace("'", ""),
                    "-c:v", "libx264", "-pix_fmt", "yuv420p", "-movflags", "+faststart", path], check=True)


def make_audio(path, seconds=10):
    """A short, gentle practice tune (two soft tones), as .m4a."""
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "lavfi", "-i", "sine=frequency=392:duration=%d" % seconds,
                    "-f", "lavfi", "-i", "sine=frequency=523:duration=%d" % seconds, "-filter_complex", "amix=inputs=2,volume=0.3",
                    "-c:a", "aac", "-b:a", "64k", path], check=True)


def poster(path, w=640, h=360, text="Video"):
    im = Image.new("RGB", (w, h), TEAL)
    ImageDraw.Draw(im).text((30, h / 2 - 20), text, fill=(255, 255, 255), font=_font(34))
    im.save(path)


def make_glb(path):
    """A tiny 3D model (a kayak-colored box) as a binary glTF (.glb) file."""
    pos = [-1, -0.2, -0.4, 1, -0.2, -0.4, 1, 0.2, -0.4, -1, 0.2, -0.4, -1, -0.2, 0.4, 1, -0.2, 0.4, 1, 0.2, 0.4, -1, 0.2, 0.4]
    idx = [0, 1, 2, 0, 2, 3, 4, 6, 5, 4, 7, 6, 0, 4, 5, 0, 5, 1, 1, 5, 6, 1, 6, 2, 2, 6, 7, 2, 7, 3, 3, 7, 4, 3, 4, 0]
    pbin = struct.pack("<%df" % len(pos), *pos)
    ibin = struct.pack("<%dH" % len(idx), *idx)
    ibin += b"\0" * ((4 - len(ibin) % 4) % 4)
    binary = pbin + ibin
    gltf = {"asset": {"version": "2.0", "generator": "Office Trainer practice file"}, "scene": 0, "scenes": [{"nodes": [0]}],
            "nodes": [{"mesh": 0, "name": "Kayak"}], "materials": [{"pbrMetallicRoughness": {"baseColorFactor": [0.9, 0.43, 0.16, 1]}}],
            "meshes": [{"primitives": [{"attributes": {"POSITION": 0}, "indices": 1, "material": 0}]}],
            "buffers": [{"byteLength": len(binary)}],
            "bufferViews": [{"buffer": 0, "byteOffset": 0, "byteLength": len(pbin)}, {"buffer": 0, "byteOffset": len(pbin), "byteLength": len(idx) * 2}],
            "accessors": [{"bufferView": 0, "componentType": 5126, "count": 8, "type": "VEC3", "min": [-1, -0.2, -0.4], "max": [1, 0.2, 0.4]},
                          {"bufferView": 1, "componentType": 5123, "count": len(idx), "type": "SCALAR"}]}
    js = json.dumps(gltf).encode()
    js += b" " * ((4 - len(js) % 4) % 4)
    out = struct.pack("<III", 0x46546C67, 2, 12 + 8 + len(js) + 8 + len(binary))
    out += struct.pack("<II", len(js), 0x4E4F534A) + js + struct.pack("<II", len(binary), 0x004E4942) + binary
    open(path, "wb").write(out)


def make_outline_docx(path, outline):
    """A Word outline for Home > New Slide > Slides from Outline: Heading 1 = slide titles, Heading 2 = bullets."""
    from docx import Document
    doc = Document()
    for title, points in outline:
        doc.add_heading(title, level=1)
        for pt in points:
            doc.add_heading(pt, level=2)
    doc.core_properties.author = "Lakeshore Paddle Company"
    doc.save(path)


# ---------------------------------------------------------------------------
# Shapes, pictures and tables: formatting
# ---------------------------------------------------------------------------
def crop(pic, l=0, t=0, r=0, b=0):
    pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = l, t, r, b


def picture_style(pic):
    """A built-in picture style ("Simple Frame, White"-like): a white border and a soft shadow."""
    spPr = pic._element.spPr
    spPr.append(X('<a:ln w="88900" cap="sq"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:miter lim="800000"/></a:ln>')[0])
    spPr.append(X('<a:effectLst><a:outerShdw blurRad="55000" dist="18000" dir="5400000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>')[0])


def glow(shape):
    spPr = shape._element.spPr
    spPr.append(X('<a:effectLst><a:glow rad="101600"><a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr></a:glow></a:effectLst>')[0])


def shape_style(shape, fill="3", ln="2", effect="2"):
    """Shape Format > Shape Styles: a theme style (Intense Effect is fill 3, line 2, effect 2)."""
    st = shape._element.find(qn("p:style"))
    st.find(qn("a:fillRef")).set("idx", fill)
    st.find(qn("a:lnRef")).set("idx", ln)
    st.find(qn("a:effectRef")).set("idx", effect)


def table_style(table_shape, style_id):
    tbl = table_shape._element.find(".//" + qn("a:tbl"))
    tblPr = tbl.find(qn("a:tblPr"))
    old = tblPr.find(qn("a:tableStyleId"))
    if old is None:
        old = etree.SubElement(tblPr, qn("a:tableStyleId"))
    old.text = style_id


MEDIUM_2_ACCENT_1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"   # PowerPoint's default table style
MEDIUM_2_ACCENT_2 = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}"
LIGHT_1_ACCENT_1 = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"


def to_front(slide, shape):
    tree = slide.shapes._spTree
    tree.remove(shape._element)
    tree.append(shape._element)


def group(slide, shapes, name="Group"):
    """Group shapes the way Shape Format > Group does (child coordinates match the slide)."""
    tree = slide.shapes._spTree
    xs = [s.left for s in shapes]; ys = [s.top for s in shapes]
    x2 = [s.left + s.width for s in shapes]; y2 = [s.top + s.height for s in shapes]
    x, y, cx, cy = min(xs), min(ys), max(x2) - min(xs), max(y2) - min(ys)
    gid = max(int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))) + 1
    g = X('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="%d" name="%s"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
          '<a:chOff x="%d" y="%d"/><a:chExt cx="%d" cy="%d"/></a:xfrm></p:grpSpPr></p:grpSp>' % (gid, name, x, y, cx, cy, x, y, cx, cy))[0]
    first = shapes[0]._element
    first.addprevious(g)
    for s in shapes:
        tree.remove(s._element)
        g.append(s._element)
    return g


# ---------------------------------------------------------------------------
# Transitions and animations (edit the slide XML before saving)
# ---------------------------------------------------------------------------
def transition(slide, kind="fade", dur=None, after=None, click=True, sound=False, option=None):
    """
    Transitions tab. kind: fade, push, wipe, split, cut (basic, p: namespace) or
    vortex, ripple, prism (Cube), gallery, switch, flip (3D, p14 namespace).
    dur: seconds. after: advance after seconds. option: dict of effect option attributes.
    """
    el = slide._element
    for old in el.findall(qn("p:transition")) + [e for e in el.findall(qn("mc:AlternateContent")) if e.find(".//" + qn("p:transition")) is not None]:
        el.remove(old)
    attrs = ""
    if after is not None:
        attrs += ' advTm="%d"' % int(after * 1000)
    if not click:
        attrs += ' advClick="0"'
    # With no option, write what PowerPoint saves for the default effect option (Push "From Bottom" is dir="u").
    defaults = {"push": {"dir": "u"}, "split": {"orient": "vert"}, "randomBar": {"dir": "vert"}}
    opts = "".join(' %s="%s"' % (k, v) for k, v in (option if option is not None else defaults.get(kind, {})).items())
    snd = '<p:sndAc><p:stSnd><p:snd r:embed="%s" name="chimes.wav"/></p:stSnd></p:sndAc>' % _sound_rid(slide) if sound else ""
    basic = kind in ("fade", "push", "wipe", "split", "cut", "cover", "pull", "randomBar", "zoom", "dissolve")
    if basic and dur is None:
        xml = '<p:transition%s><p:%s%s/>%s</p:transition>' % (attrs, kind, opts, snd)
        node = X(xml)[0]
    else:
        inner = '<p:%s%s/>' % (kind, opts) if basic else '<p14:%s%s/>' % (kind, opts)
        d = ' p14:dur="%d"' % int(dur * 1000) if dur is not None else ""
        xml = ('<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow"%s%s>%s%s</p:transition></mc:Choice>'
               '<mc:Fallback><p:transition spd="slow"%s><p:fade/>%s</p:transition></mc:Fallback></mc:AlternateContent>') % (d, attrs, inner, snd, attrs, snd)
        node = X(xml)[0]
    # A transition comes after <p:clrMapOvr> (or <p:cSld> when there is none) and before <p:timing>.
    anchor = el.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = el.find(qn("p:cSld"))
    anchor.addnext(node)


def make_wav(seconds=0.6, freq=880):
    """A short chime (16-bit mono WAV), the kind of sound a transition can play."""
    import wave
    rate = 22050
    buf = io.BytesIO()
    with wave.open(buf, "wb") as w:
        w.setnchannels(1); w.setsampwidth(2); w.setframerate(rate)
        frames = bytearray()
        for i in range(int(rate * seconds)):
            fade = 1 - i / (rate * seconds)
            frames += struct.pack("<h", int(9000 * fade * math.sin(2 * math.pi * freq * i / rate)))
        w.writeframes(bytes(frames))
    return buf.getvalue()


def _sound_rid(slide):
    """Embed a chime in the slide's part (Transitions > Sound) and return its relationship id."""
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    pkg = slide.part.package
    n = 1
    while any(str(p.partname) == "/ppt/media/sound%d.wav" % n for p in pkg.iter_parts()):
        n += 1
    part = Part(PackURI("/ppt/media/sound%d.wav" % n), "audio/x-wav", pkg, make_wav())
    return slide.part.relate_to(part, REL + "audio")


_ids = {"n": 1}


def animate(slide, effects):
    """
    Animations tab. effects: list of dicts, in order:
      shape (the python-pptx shape), cls (entr/emph/exit/path), preset (PowerPoint's preset id),
      subtype, start (click/with/after), delay (s), dur (s), paragraphs (True = one per paragraph),
      path ("M 0 0 L 0.25 0 E" for motion paths).
    Builds the same timing tree PowerPoint saves.
    """
    el = slide._element
    old = el.find(qn("p:timing"))
    keep_media = []
    if old is not None:
        keep_media = old.findall(".//" + qn("p:video")) + old.findall(".//" + qn("p:audio"))
        el.remove(old)
    cid = [1]

    def nid():
        cid[0] += 1
        return cid[0]

    # Expand "paragraphs" into one effect per paragraph.
    flat = []
    for e in effects:
        if e.get("paragraphs"):
            n = len([p for p in e["shape"].text_frame.paragraphs if p.text.strip()])
            for i in range(n):
                f = dict(e); f["para"] = i; f["start"] = e.get("start", "click") if i == 0 else e.get("next_start", "click")
                flat.append(f)
        else:
            flat.append(e)
    # Group into click groups: a "click" effect starts a new group; with/after join the current one.
    groups = []
    for e in flat:
        if e.get("start", "click") == "click" or not groups:
            groups.append([[e]])
        elif e["start"] == "with":
            groups[-1][-1].append(e)
        else:  # after: a new time offset inside the click group
            groups[-1].append([e])
    seq_children = ""
    builds = {}
    for g in groups:
        click_id = nid()
        first_start = g[0][0].get("start", "click")
        offset = 0
        offs_xml = ""
        for k, offgroup in enumerate(g):
            off_id = nid()
            eff_xml = ""
            longest = 0
            for e in offgroup:
                sp = e["shape"].shape_id
                cls = e.get("cls", "entr")
                node = {"click": "clickEffect", "with": "withEffect", "after": "afterEffect"}[e.get("start", "click")]
                if k == 0 and first_start != "click" and e is offgroup[0]:
                    node = {"with": "withEffect", "after": "afterEffect"}[first_start]
                dur = int(e.get("dur", 0.5) * 1000)
                delay = int(e.get("delay", 0) * 1000)
                longest = max(longest, delay + dur)
                tgt = '<p:spTgt spid="%s">%s</p:spTgt>' % (sp, '<p:txEl><p:pRg st="%d" end="%d"/></p:txEl>' % (e["para"], e["para"]) if "para" in e else "")
                if e.get("para") is not None or e.get("paragraphs"):
                    builds[sp] = "p"
                elif cls in ("entr", "exit", "emph", "path"):
                    builds.setdefault(sp, "whole") if e["shape"].has_text_frame and e["shape"].text_frame.text.strip() else None
                if cls == "path":
                    behav = ('<p:animMotion origin="layout" path="%s" pathEditMode="relative" ptsTypes="AA"><p:cBhvr><p:cTn id="%d" dur="%d" fill="hold"/>'
                             '<p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>') % (e["path"], nid(), dur, tgt)
                elif cls == "emph":
                    behav = ('<p:animRot by="21600000"><p:cBhvr><p:cTn id="%d" dur="%d" fill="hold"/><p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>') % (nid(), dur, tgt)
                else:
                    vis = "visible" if cls == "entr" else "hidden"
                    behav = ('<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst><p:cond delay="%d"/></p:stCondLst></p:cTn><p:tgtEl>%s</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                             '<p:to><p:strVal val="%s"/></p:to></p:set><p:animEffect transition="%s" filter="fade"><p:cBhvr><p:cTn id="%d" dur="%d"/><p:tgtEl>%s</p:tgtEl></p:cBhvr></p:animEffect>') % (
                        nid(), 0 if cls == "entr" else dur - 1, tgt, vis, "in" if cls == "entr" else "out", nid(), dur, tgt)
                eff_xml += ('<p:par><p:cTn id="%d" presetID="%d" presetClass="%s" presetSubtype="%d" fill="hold"%s nodeType="%s"><p:stCondLst><p:cond delay="%d"/></p:stCondLst>'
                            '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par>') % (nid(), e.get("preset", 10), cls, e.get("subtype", 0), ' grpId="0"' if sp in builds else "", node, delay, behav)
            offs_xml += '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>' % (off_id, offset, eff_xml)
            offset += longest
        click_delay = "indefinite" if first_start == "click" else "0"
        cond = '<p:cond delay="%s"/>' % click_delay if first_start == "click" else '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
        seq_children += '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst>%s</p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>' % (click_id, cond, offs_xml)
    media_xml = "".join(etree.tostring(m).decode() for m in keep_media)
    bld = "".join('<p:bldP spid="%s" grpId="0"%s/>' % (sp, ' build="p"' if b == "p" else "") for sp, b in builds.items() if b)
    xml = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
           '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst></p:cTn>'
           '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>%s'
           '</p:childTnLst></p:cTn></p:par></p:tnLst>%s</p:timing>') % (seq_children, media_xml, "<p:bldLst>%s</p:bldLst>" % bld if bld else "")
    # cTn ids must be unique: renumber them in document order.
    node = X(xml)[0]
    for i, c in enumerate(node.iter(qn("p:cTn")), start=1):
        c.set("id", str(i))
    # <p:timing> comes after <p:transition> (and its AlternateContent) and before <p:extLst>.
    ext = el.find(qn("p:extLst"))
    if ext is not None:
        ext.addprevious(node)
    else:
        el.append(node)


def media_options(slide, movie, trim=None, fade_in=None, start="click", loop=False, across=False, audio=False):
    """Playback tab: trim (start ms, end ms), fade in (ms), start automatically, loop, play across slides."""
    nv = movie._element.find(".//" + qn("p:nvPr"))
    media = nv.find(".//" + qn("p14:media"))
    if trim:
        media.append(X('<p14:trim st="%d" end="%d"/>' % trim)[0])
    if fade_in:
        media.append(X('<p14:fade in="%d"/>' % fade_in)[0])
    timing = slide._element.find(qn("p:timing"))
    node = timing.find(".//" + qn("p:video")) if not audio else timing.find(".//" + qn("p:audio"))
    if node is None:
        node = timing.find(".//" + qn("p:video"))
    if audio and node is not None and node.tag == qn("p:video"):
        node.tag = qn("p:audio")
    cmn = node.find(qn("p:cMediaNode"))
    if across:
        cmn.set("numSld", "999")
    ctn = cmn.find(qn("p:cTn"))
    if loop:
        ctn.set("repeatCount", "indefinite")
    if start == "auto":
        # A play command in the main sequence that starts with the slide.
        spid = movie.shape_id
        root_ctn = timing.find(".//" + qn("p:cTn"))
        kids = root_ctn.find(qn("p:childTnLst"))
        seq = kids.find(qn("p:seq"))
        if seq is None:
            seq = X('<p:seq concurrent="1" nextAc="seek"><p:cTn id="900" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>'
                    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>')[0]
            kids.insert(0, seq)
        lst = seq.find(qn("p:cTn")).find(qn("p:childTnLst"))
        lst.insert(0, X('<p:par><p:cTn id="901" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
                        '<p:par><p:cTn id="902" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                        '<p:par><p:cTn id="903" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                        '<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="904" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="%s"/></p:tgtEl></p:cBhvr></p:cmd>'
                        '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>' % spid)[0])
        for i, c in enumerate(timing.iter(qn("p:cTn")), start=1):
            c.set("id", str(i))


# ---------------------------------------------------------------------------
# Part-level patches (after saving)
# ---------------------------------------------------------------------------
def _read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write_zip(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts["[Content_Types].xml"])
        for n, data in parts.items():
            if n != "[Content_Types].xml":
                z.writestr(n, data)


def patch(path, fn):
    parts = _read_zip(path)
    fn(parts)
    _write_zip(path, parts)


def _x(data):
    return etree.fromstring(data)


def _b(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_name(part):
    d, n = part.rsplit("/", 1)
    return d + "/_rels/" + n + ".rels"


def _add_rel(parts, part, rel_type, target, external=False):
    name = _rels_name(part)
    root = _x(parts[name]) if name in parts else etree.Element("{%s}Relationships" % NS_REL, nsmap={None: NS_REL})
    ids = {r.get("Id") for r in root}
    n = 1
    while "rId%d" % n in ids:
        n += 1
    rel = etree.SubElement(root, "{%s}Relationship" % NS_REL)
    rel.set("Id", "rId%d" % n)
    rel.set("Type", rel_type if rel_type.startswith("http") else REL + rel_type)
    rel.set("Target", target)
    if external:
        rel.set("TargetMode", "External")
    parts[name] = _b(root)
    return "rId%d" % n


def _ct(parts, part_name=None, content_type=None, extension=None):
    ct = _x(parts["[Content_Types].xml"])
    if extension:
        if not any(e.get("Extension") == extension for e in ct if e.tag.endswith("Default")):
            el = etree.SubElement(ct, "{%s}Default" % NS_CT); el.set("Extension", extension); el.set("ContentType", content_type)
    else:
        el = etree.SubElement(ct, "{%s}Override" % NS_CT); el.set("PartName", "/" + part_name); el.set("ContentType", content_type)
    parts["[Content_Types].xml"] = _b(ct)


def _slide_parts(parts):
    """Slide part names in show order, and their ids."""
    pres = _x(parts["ppt/presentation.xml"])
    rels = {r.get("Id"): r.get("Target") for r in _x(parts["ppt/_rels/presentation.xml.rels"])}
    out = []
    for s in pres.find(qn("p:sldIdLst")):
        out.append(("ppt/" + rels[s.get(qn("r:id"))].lstrip("/").replace("ppt/", ""), s.get("id"), s.get(qn("r:id"))))
    return out


def _pres_ext(pres, uri, xml):
    """Add (or replace) an extension in presentation.xml's <p:extLst>."""
    ext_lst = pres.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, qn("p:extLst"))
    for e in ext_lst:
        if e.get("uri") == uri:
            ext_lst.remove(e)
    ext = etree.SubElement(ext_lst, qn("p:ext")); ext.set("uri", uri)
    ext.append(X(xml)[0])


def sections(path, secs):
    """Right-click > Add Section. secs: [(name, [slide numbers])]."""
    def fn(parts):
        pres = _x(parts["ppt/presentation.xml"])
        slides = _slide_parts(parts)
        inner = ""
        for i, (name, nums) in enumerate(secs):
            inner += '<p14:section name="%s" id="{%08X-1111-4222-8333-%012X}"><p14:sldIdLst>%s</p14:sldIdLst></p14:section>' % (
                esc(name), i + 1, i + 1, "".join('<p14:sldId id="%s"/>' % slides[n - 1][1] for n in nums))
        _pres_ext(pres, "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}", '<p14:sectionLst>%s</p14:sectionLst>' % inner)
        parts["ppt/presentation.xml"] = _b(pres)
    patch(path, fn)


def custom_show(path, name, nums):
    """Slide Show > Custom Slide Show > Custom Shows > New."""
    def fn(parts):
        pres = _x(parts["ppt/presentation.xml"])
        slides = _slide_parts(parts)
        lst = pres.find(qn("p:custShowLst"))
        if lst is None:
            lst = X('<p:custShowLst/>')[0]
            # custShowLst comes after notesSz/smartTags/embeddedFontLst and before photoAlbum/custDataLst/kinsoku/defaultTextStyle
            pres.find(qn("p:defaultTextStyle")).addprevious(lst)
        lst.append(X('<p:custShow name="%s" id="%d"><p:sldLst>%s</p:sldLst></p:custShow>' % (esc(name), len(lst), "".join('<p:sld r:id="%s"/>' % slides[n - 1][2] for n in nums)))[0])
        parts["ppt/presentation.xml"] = _b(pres)
    patch(path, fn)


def show_settings(path, loop=False, kiosk=False, narration=True, use_timings=True):
    """Slide Show > Set Up Slide Show."""
    def fn(parts):
        pp = _x(parts["ppt/presProps.xml"])
        old = pp.find(qn("p:showPr"))
        if old is not None:
            pp.remove(old)
        attrs = (' loop="1"' if loop else "") + ("" if narration else ' showNarration="0"') + ("" if use_timings else ' useTimings="0"')
        kind = "<p:kiosk/>" if kiosk else "<p:present/>"
        pp.insert(0, X('<p:showPr%s>%s<p:sldAll/><p:penClr><a:prstClr val="red"/></p:penClr></p:showPr>' % (attrs, kind))[0])
        parts["ppt/presProps.xml"] = _b(pp)
    patch(path, fn)


def last_view(path, view="sldThumbnailView", guides=False):
    """The view the presentation was saved in (View tab), and whether guides show."""
    def fn(parts):
        vp = _x(parts["ppt/viewProps.xml"])
        vp.set("lastView", view)
        csv = vp.find(".//" + qn("p:cSldViewPr"))
        if guides:
            csv.set("showGuides", "1")
        parts["ppt/viewProps.xml"] = _b(vp)
    patch(path, fn)


def pres_flags(path, embed_fonts=False, remove_personal=False, modify_password=False):
    def fn(parts):
        pres = _x(parts["ppt/presentation.xml"])
        if embed_fonts:
            pres.set("embedTrueTypeFonts", "1")
            pres.set("saveSubsetFonts", "1")
        if remove_personal:
            pres.set("removePersonalInfoOnSave", "1")
        if modify_password:
            mv = X('<p:modifyVerifier cryptProviderType="rsaAES" cryptAlgorithmClass="hash" cryptAlgorithmType="typeAny" cryptAlgorithmSid="14" '
                   'spinCount="100000" saltData="ZmxjY29mZmljZXRyYWluZQ==" hashData="cHJhY3RpY2Vvbmx5cHJhY3RpY2Vvbmx5cHJhY3RpY2VvbmFhYWFhYWFhYWFhYWFhYWFhYWFhYWE="/>')[0]
            pres.find(qn("p:defaultTextStyle")).addprevious(mv)
        parts["ppt/presentation.xml"] = _b(pres)
        if remove_personal:
            core = parts["docProps/core.xml"].decode()
            core = re.sub(r"<dc:creator>[^<]*</dc:creator>", "<dc:creator></dc:creator>", core)
            core = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>", core)
            # Inspect Document removes the other properties too, the Title among them.
            core = re.sub(r"<dc:title>[^<]*</dc:title>", "<dc:title></dc:title>", core)
            parts["docProps/core.xml"] = core.encode()
    patch(path, fn)


def mark_final(path):
    """File > Info > Protect Presentation > Mark as Final (a custom property)."""
    def fn(parts):
        parts["docProps/custom.xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                                        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal">'
                                        '<vt:bool>true</vt:bool></property></Properties>').encode()
        _ct(parts, "docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
        rels = _x(parts["_rels/.rels"])
        r = etree.SubElement(rels, "{%s}Relationship" % NS_REL)
        r.set("Id", "rIdCustom1"); r.set("Type", REL + "custom-properties"); r.set("Target", "docProps/custom.xml")
        parts["_rels/.rels"] = _b(rels)
    patch(path, fn)


def set_author(path, name):
    def fn(parts):
        core = parts["docProps/core.xml"].decode()
        core = re.sub(r"<dc:creator>[^<]*</dc:creator>", "<dc:creator>%s</dc:creator>" % name, core)
        core = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "<cp:lastModifiedBy>%s</cp:lastModifiedBy>" % name, core)
        parts["docProps/core.xml"] = core.encode()
    patch(path, fn)


def comments(path, threads):
    """
    Modern comments (Review > New Comment). threads: [{slide, author, text, replies: [(author, text)], resolved}].
    Each slide with comments gets a modernComment part; authors go in ppt/authors.xml.
    """
    def fn(parts):
        slides = _slide_parts(parts)
        people = {}
        for t in threads:
            for a in [t["author"]] + [r[0] for r in t.get("replies", [])]:
                people.setdefault(a, "{%s}" % str(uuid.uuid5(uuid.NAMESPACE_DNS, a)).upper())
        authors = '<p188:authorLst xmlns:a="%s" xmlns:r="%s" xmlns:p188="%s">%s</p188:authorLst>' % (NS["a"], NS["r"], NS["p188"],
            "".join('<p188:author id="%s" name="%s" initials="%s" userId="%s" providerId="None"/>' % (pid, esc(n), "".join(w[0] for w in n.split()), esc(n)) for n, pid in people.items()))
        parts["ppt/authors.xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + authors).encode()
        _ct(parts, "ppt/authors.xml", "application/vnd.ms-powerpoint.authors+xml")
        _add_rel(parts, "ppt/presentation.xml", "http://schemas.microsoft.com/office/2018/10/relationships/authors", "authors.xml")
        by_slide = {}
        for t in threads:
            by_slide.setdefault(t["slide"], []).append(t)
        for n, items in by_slide.items():
            sp, sid, _ = slides[n - 1]
            # A comment points at its slide by id and by creationId (an id PowerPoint gives each slide); add one if missing.
            sx0 = _x(parts[sp])
            cre = sx0.find(".//" + qn("p14:creationId"))
            if cre is None:
                ext0 = sx0.find(qn("p:extLst"))
                if ext0 is None:
                    ext0 = etree.SubElement(sx0, qn("p:extLst"))
                e0 = etree.SubElement(ext0, qn("p:ext")); e0.set("uri", "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}")
                cre = etree.SubElement(e0, qn("p14:creationId")); cre.set("val", str(1000000000 + int(sid) * 7919))
                parts[sp] = _b(sx0)
            cid_val = cre.get("val")
            body = ""
            for i, t in enumerate(items):
                cid = "{%08X-AAAA-4BBB-8CCC-%012X}" % (n, i + 1)
                replies = "".join('<p188:reply id="{%08X-DDDD-4EEE-8FFF-%012X}" authorId="%s" created="2027-05-02T10:%02d:00.000"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></p188:txBody></p188:reply>'
                                  % (n, (i + 1) * 10 + j, people[a], j, esc(txt)) for j, (a, txt) in enumerate(t.get("replies", [])))
                status = ' status="resolved"' if t.get("resolved") else ""
                body += ('<p188:cm id="%s" authorId="%s" created="2027-05-01T09:%02d:00.000"%s><pc:sldMkLst xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docMk/><pc:sldMk cId="%s" sldId="%s"/></pc:sldMkLst>'
                         '%s<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></p188:txBody></p188:cm>') % (
                    cid, people[t["author"]], i, status, cid_val, sid, "<p188:replyLst>%s</p188:replyLst>" % replies if replies else "", esc(t["text"]))
            name = "ppt/comments/modernComment_%s_0.xml" % sid
            parts[name] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p188:cmLst xmlns:a="%s" xmlns:r="%s" xmlns:p188="%s">%s</p188:cmLst>' % (NS["a"], NS["r"], NS["p188"], body)).encode()
            _ct(parts, name, "application/vnd.ms-powerpoint.comments+xml")
            rid = _add_rel(parts, sp, "http://schemas.microsoft.com/office/2018/10/relationships/comments", "../comments/" + name.split("/")[-1])
            sx = _x(parts[sp])
            ext = sx.find(qn("p:extLst"))
            if ext is None:
                ext = etree.SubElement(sx, qn("p:extLst"))
            e = etree.SubElement(ext, qn("p:ext")); e.set("uri", "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}")
            e.append(X('<p188:commentRel r:id="%s"/>' % rid)[0])
            parts[sp] = _b(sx)
    patch(path, fn)


def handout_master(path, header=None, date=True, footer=None):
    """View > Handout Master (python-pptx's template has none, so make one)."""
    def fn(parts):
        def ph(i, t, name, x, y, text=""):
            return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="%s" sz="quarter"%s/></p:nvPr></p:nvSpPr>'
                    '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="2971800" cy="458788"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
                    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p>%s</a:p></p:txBody></p:sp>') % (i, name, t, ' idx="%d"' % i if i > 2 else "", x, y,
                                                                                            '<a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r>' % esc(text) if text else '<a:endParaRPr lang="en-US"/>')
        shapes = ph(2, "hdr", "Header Placeholder 1", 0, 0, header or "")
        if date:
            shapes += ph(3, "dt", "Date Placeholder 2", 3884613, 0)
        shapes += ph(4, "ftr", "Footer Placeholder 3", 0, 8685213, footer or "") + ph(5, "sldNum", "Slide Number Placeholder 4", 3884613, 8685213)
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:handoutMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
               '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>%s</p:spTree></p:cSld>'
               '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:handoutMaster>') % (NS["a"], NS["r"], NS["p"], shapes)
        parts["ppt/handoutMasters/handoutMaster1.xml"] = xml.encode()
        _ct(parts, "ppt/handoutMasters/handoutMaster1.xml", "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml")
        # A handout master has its own copy of the theme.
        parts["ppt/theme/theme9.xml"] = parts["ppt/theme/theme1.xml"]
        _ct(parts, "ppt/theme/theme9.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
        _add_rel(parts, "ppt/handoutMasters/handoutMaster1.xml", "theme", "../theme/theme9.xml")
        rid = _add_rel(parts, "ppt/presentation.xml", "handoutMaster", "handoutMasters/handoutMaster1.xml")
        pres = _x(parts["ppt/presentation.xml"])
        hml = X('<p:handoutMasterIdLst><p:handoutMasterId r:id="%s"/></p:handoutMasterIdLst>' % rid)[0]
        nml = pres.find(qn("p:notesMasterIdLst"))
        (nml if nml is not None else pres.find(qn("p:sldMasterIdLst"))).addnext(hml)
        parts["ppt/presentation.xml"] = _b(pres)
    patch(path, fn)


def new_layout(path, name, picture=True, base="slideLayout6.xml"):
    """View > Slide Master > Insert Layout (copy of Title Only), renamed, with a Picture placeholder."""
    def fn(parts):
        src = _x(parts["ppt/slideLayouts/" + base])
        src.set("userDrawn", "1")
        if "type" in src.attrib:
            del src.attrib["type"]
        src.set("preserve", "1")
        src.find(qn("p:cSld")).set("name", name)
        if picture:
            tree = src.find(".//" + qn("p:spTree"))
            tree.append(X('<p:sp><p:nvSpPr><p:cNvPr id="20" name="Picture Placeholder 19"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr>'
                          '<p:spPr><a:xfrm><a:off x="609600" y="1600200"/><a:ext cx="4572000" cy="4114800"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>')[0])
        n = 12
        while "ppt/slideLayouts/slideLayout%d.xml" % n in parts:
            n += 1
        lp = "ppt/slideLayouts/slideLayout%d.xml" % n
        parts[lp] = _b(src)
        _ct(parts, lp, "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")
        _add_rel(parts, lp, "slideMaster", "../slideMasters/slideMaster1.xml")
        rid = _add_rel(parts, "ppt/slideMasters/slideMaster1.xml", "slideLayout", "../slideLayouts/slideLayout%d.xml" % n)
        m = _x(parts["ppt/slideMasters/slideMaster1.xml"])
        lst = m.find(qn("p:sldLayoutIdLst"))
        top = max(int(e.get("id")) for e in lst)
        lst.append(X('<p:sldLayoutId id="%d" r:id="%s"/>' % (top + 1, rid))[0])
        parts["ppt/slideMasters/slideMaster1.xml"] = _b(m)
    patch(path, fn)


def edit_layout(path, layout_file, remove_types=(), rename=None):
    def fn(parts):
        lp = "ppt/slideLayouts/" + layout_file
        l = _x(parts[lp])
        tree = l.find(".//" + qn("p:spTree"))
        for sp in list(tree):
            ph = sp.find(".//" + qn("p:ph"))
            if ph is not None and ph.get("type") in remove_types:
                tree.remove(sp)
        if rename:
            l.find(qn("p:cSld")).set("name", rename)
        parts[lp] = _b(l)
    patch(path, fn)


def master_title_font(prs, font):
    m = prs.slide_master._element
    for latin in m.find(qn("p:txStyles")).find(qn("p:titleStyle")).iter(qn("a:latin")):
        latin.set("typeface", font)


def master_background(prs, color=(232, 244, 245), gradient=False):
    m = prs.slide_master._element
    cSld = m.find(qn("p:cSld"))
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    fill = ('<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="%02X%02X%02X"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' % color
            if gradient else '<a:solidFill><a:srgbClr val="%02X%02X%02X"/></a:solidFill>' % color)
    cSld.insert(0, X('<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>' % fill)[0])


def slide_background(slide, kind="gradient", image_rid=None, color=(14, 124, 134)):
    cSld = slide._element.find(qn("p:cSld"))
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    if kind == "gradient":
        fill = '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="%02X%02X%02X"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' % color
    elif kind == "picture":
        fill = '<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>' % image_rid
    else:
        fill = '<a:solidFill><a:srgbClr val="%02X%02X%02X"/></a:solidFill>' % color
    cSld.insert(0, X('<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>' % fill)[0])


def image_rel(slide, image_path):
    """Add a picture to a slide's part (for a background) and return its relationship id."""
    image_part, rid = slide.part.get_or_add_image_part(image_path)
    return rid


def theme(path, name="Lakeshore", major="Georgia", minor="Verdana", accent1="0E7C86", colors_name="Lakeshore", fonts_name=None):
    """Design > Themes: a different theme (here: renamed, new fonts and accent color).
    fonts_name names the font set (Design > Variants > Fonts); it defaults to the theme name."""
    def fn(parts):
        for tp in [n for n in parts if re.match(r"ppt/theme/theme\d+\.xml$", n)]:
            t = _x(parts[tp])
            t.set("name", name)
            cs = t.find(".//" + qn("a:clrScheme")); cs.set("name", colors_name)
            a1 = cs.find(qn("a:accent1"))
            for c in list(a1):
                a1.remove(c)
            etree.SubElement(a1, qn("a:srgbClr")).set("val", accent1)
            fs = t.find(".//" + qn("a:fontScheme")); fs.set("name", fonts_name or name)
            fs.find(qn("a:majorFont")).find(qn("a:latin")).set("typeface", major)
            fs.find(qn("a:minorFont")).find(qn("a:latin")).set("typeface", minor)
            parts[tp] = _b(t)
    patch(path, fn)


def _graphic_frame(parts, sp, xml_choice, requires, fallback_name, x, y, cx, cy, shape_id, name, alt=""):
    """Add a newer object (Zoom, 3D model) to a slide inside AlternateContent, with a picture fallback."""
    sx = _x(parts[sp])
    tree = sx.find(".//" + qn("p:spTree"))
    frag = X(('<mc:AlternateContent><mc:Choice %s><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s" descr="%s"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
              '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm><a:graphic>%s</a:graphic></p:graphicFrame></mc:Choice>'
              '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></mc:Fallback></mc:AlternateContent>')
             % (requires, shape_id, name, esc(alt), x, y, cx, cy, xml_choice, shape_id, fallback_name, x, y, cx, cy))[0]
    tree.append(frag)
    parts[sp] = _b(sx)


def zoom(path, slide_n, kind, target_slides, return_to=True, shape_id=90):
    """Insert > Zoom. kind: summary, section or slide."""
    uri = {"summary": "http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom",
           "section": "http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom",
           "slide": "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"}[kind]
    pre = {"summary": "psuz", "section": "psez", "slide": "pslz"}[kind]
    def fn(parts):
        slides = _slide_parts(parts)
        sp = slides[slide_n - 1][0]
        objs = ""
        for i, t in enumerate(target_slides):
            tag = {"summary": "summaryZmObj", "section": "sectionZmObj", "slide": "sldZmObj"}[kind]
            idattr = 'sldId="%s"' % slides[t - 1][1] if kind == "slide" else 'sectionId="{%08X-1111-4222-8333-%012X}"' % (i + 1, i + 1)
            objs += '<%s:%s %s><%s:zmPr returnToParent="%d"/></%s:%s>' % (pre, tag, idattr, pre, 1 if return_to else 0, pre, tag)
        inner = '<a:graphicData uri="%s"><%s:%s xmlns:%s="%s">%s</%s:%s></a:graphicData>' % (
            uri, pre, {"summary": "summaryZm", "section": "sectionZm", "slide": "sldZm"}[kind], pre, uri, objs, pre, {"summary": "summaryZm", "section": "sectionZm", "slide": "sldZm"}[kind])
        _graphic_frame(parts, sp, inner, 'xmlns:%s="%s" Requires="%s"' % (pre, uri, pre), "Zoom fallback", 914400, 1828800, 6096000, 2286000, shape_id, "%s Zoom 1" % kind.title())
    patch(path, fn)


def model3d(path, slide_n, glb_path, rotated=False, shape_id=91, alt="A 3D model of a kayak"):
    def fn(parts):
        slides = _slide_parts(parts)
        sp = slides[slide_n - 1][0]
        parts["ppt/media/model3d1.glb"] = open(glb_path, "rb").read()
        _ct(parts, content_type="model/gltf-binary", extension="glb")
        rid = _add_rel(parts, sp, "http://schemas.microsoft.com/office/2017/06/relationships/model3d", "../media/model3d1.glb")
        rot = '<am3d:rot ax="1500000" ay="-2700000" az="0"/>' if rotated else '<am3d:rot ax="0" ay="0" az="0"/>'
        inner = ('<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d"><am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" r:embed="%s">'
                 '<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3048000" cy="2286000"/></a:xfrm></am3d:spPr><am3d:camera><am3d:pos x="0" y="0" z="60000000"/><am3d:up dx="0" dy="36000000" dz="0"/><am3d:lookAt x="0" y="0" z="0"/>'
                 '<am3d:perspective fov="2700000"/></am3d:camera><am3d:trans><am3d:meterPerModelUnit n="1" d="1"/><am3d:preTrans dx="0" dy="0" dz="0"/><am3d:scale><a:sx n="1000000" d="1000000"/><a:sy n="1000000" d="1000000"/><a:sz n="1000000" d="1000000"/></am3d:scale>%s<am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>'
                 '<am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="%s"/></am3d:raster><am3d:objViewport viewportSz="3048000"/><am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr><am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight></am3d:model3d></a:graphicData>') % (rid, rot, rid)
        _graphic_frame(parts, sp, inner, 'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" Requires="am3d"', "3D Model fallback", 6096000, 1828800, 3048000, 2286000, shape_id, "3D Model 1", alt)
    patch(path, fn)


def smartart(path, slide_n, items, layout_id="process1", colors="accent1_2", shape_id=92, alt=""):
    """A SmartArt graphic (data, layout, colors and style parts), as PowerPoint links them."""
    D = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
    def fn(parts):
        slides = _slide_parts(parts)
        sp = slides[slide_n - 1][0]
        pts = '<dgm:pt modelId="{00000000-0000-4000-8000-000000000000}" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
        cxn = ""
        for i, t in enumerate(items):
            pts += '<dgm:pt modelId="{00000000-0000-4000-8000-0000000000%02d}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/><dgm:t><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>%s</a:t></a:r></a:p></dgm:t></dgm:pt>' % (i + 1, esc(t))
            cxn += '<dgm:cxn modelId="{10000000-0000-4000-8000-0000000000%02d}" srcId="{00000000-0000-4000-8000-000000000000}" destId="{00000000-0000-4000-8000-0000000000%02d}" srcOrd="%d" destOrd="0"/>' % (i + 1, i + 1, i)
        data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel xmlns:dgm="%s" xmlns:a="%s"><dgm:ptLst>%s</dgm:ptLst><dgm:cxnLst>%s</dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>' % (D, NS["a"], pts, cxn)
        lo = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:layoutDef xmlns:dgm="%s" xmlns:a="%s" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/%s"><dgm:title val=""/><dgm:desc val=""/><dgm:layoutNode name="diagram"/></dgm:layoutDef>' % (D, NS["a"], layout_id)
        cs = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:colorsDef xmlns:dgm="%s" xmlns:a="%s" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/%s"><dgm:title val=""/><dgm:desc val=""/></dgm:colorsDef>' % (D, NS["a"], colors)
        qs = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:styleDef xmlns:dgm="%s" xmlns:a="%s" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"><dgm:title val=""/><dgm:desc val=""/></dgm:styleDef>' % (D, NS["a"])
        n = 1
        while "ppt/diagrams/data%d.xml" % n in parts:
            n += 1
        ct = "application/vnd.openxmlformats-officedocument.drawingml.diagram"
        rids = {}
        for key, nm, body, kind in [("dm", "data", data, "Data"), ("lo", "layout", lo, "Layout"), ("qs", "quickStyle", qs, "Style"), ("cs", "colors", cs, "Colors")]:
            pn = "ppt/diagrams/%s%d.xml" % (nm, n)
            parts[pn] = body.encode()
            _ct(parts, pn, ct + kind + "+xml")
            rids[key] = _add_rel(parts, sp, "diagram" + {"Style": "QuickStyle"}.get(kind, kind), "../diagrams/%s%d.xml" % (nm, n))
        sx = _x(parts[sp])
        tree = sx.find(".//" + qn("p:spTree"))
        tree.append(X(('<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Diagram %d" descr="%s"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
                       '<p:xfrm><a:off x="1524000" y="1828800"/><a:ext cx="9144000" cy="3200400"/></p:xfrm><a:graphic><a:graphicData uri="%s">'
                       '<dgm:relIds xmlns:dgm="%s" r:dm="%s" r:lo="%s" r:qs="%s" r:cs="%s"/></a:graphicData></a:graphic></p:graphicFrame>')
                      % (shape_id, n, esc(alt), D, D, rids["dm"], rids["lo"], rids["qs"], rids["cs"]))[0])
        parts[sp] = _b(sx)
    patch(path, fn)


def ink(path, slide_n, shape_id=93):
    """Draw > a pen stroke: an ink content part, with a picture fallback."""
    def fn(parts):
        slides = _slide_parts(parts)
        sp = slides[slide_n - 1][0]
        parts["ppt/ink/ink1.xml"] = (b'<?xml version="1.0" encoding="UTF-8"?><inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>10 0, 9 1, 7 2, 5 3, 3 5, 2 7, 1 9, 0 10, 1 12, 3 14, 6 15, 9 16, 12 15, 15 14, 17 12, 18 10, 17 7, 15 5, 13 3, 11 1</inkml:trace></inkml:ink>')
        _ct(parts, "ppt/ink/ink1.xml", "application/inkml+xml")
        rid = _add_rel(parts, sp, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", "../ink/ink1.xml")
        sx = _x(parts[sp])
        tree = sx.find(".//" + qn("p:spTree"))
        tree.append(X(('<mc:AlternateContent><mc:Choice Requires="p14"><p:contentPart p14:bwMode="auto" r:id="%s"><p14:nvContentPartPr><p14:cNvPr id="%d" name="Ink %d"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr>'
                       '<p14:xfrm><a:off x="3000000" y="2000000"/><a:ext cx="1200000" cy="1200000"/></p14:xfrm></p:contentPart></mc:Choice>'
                       '<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="%d" name="Ink fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent>') % (rid, shape_id, shape_id, shape_id))[0])
        parts[sp] = _b(sx)
    patch(path, fn)


def icons(path, slide_n, names):
    """Turn pictures into icons (Insert > Icons saves an SVG with a PNG copy)."""
    def fn(parts):
        slides = _slide_parts(parts)
        sp = slides[slide_n - 1][0]
        _ct(parts, content_type="image/svg+xml", extension="svg")
        sx = _x(parts[sp])
        for pic in sx.iter(qn("p:pic")):
            name = pic.find(".//" + qn("p:cNvPr")).get("name")
            if name in names:
                # Each icon gets its own SVG (wave, sun or paddle, by its name), like Insert > Icons saves.
                kind = next((k for k in SVG_ICONS if k in name.lower()), "wave")
                n = 1
                while "ppt/media/icon%d.svg" % n in parts:
                    n += 1
                parts["ppt/media/icon%d.svg" % n] = SVG_ICONS[kind]
                rid = _add_rel(parts, sp, "image", "../media/icon%d.svg" % n)
                blip = pic.find(".//" + qn("a:blip"))
                blip.append(X('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="%s"/></a:ext></a:extLst>' % rid)[0])
        parts[sp] = _b(sx)
    patch(path, fn)


def size_43_to_169(path):
    """Design > Slide Size > Widescreen (what changes in presentation.xml)."""
    def fn(parts):
        pres = _x(parts["ppt/presentation.xml"])
        sz = pres.find(qn("p:sldSz")); sz.set("cx", "12192000"); sz.set("cy", "6858000")
        if "type" in sz.attrib:
            del sz.attrib["type"]
        parts["ppt/presentation.xml"] = _b(pres)
    patch(path, fn)


def strip_thumbnail(prs_path):
    """python-pptx's template has a thumbnail of an empty slide; drop it so file previews don't mislead."""
    def fn(parts):
        if "docProps/thumbnail.jpeg" in parts:
            del parts["docProps/thumbnail.jpeg"]
            r = _x(parts["_rels/.rels"])
            for rel in list(r):
                if rel.get("Target", "").endswith("thumbnail.jpeg"):
                    r.remove(rel)
            parts["_rels/.rels"] = _b(r)
    patch(prs_path, fn)


def save(prs, path):
    prs.save(path)
    strip_thumbnail(path)


# ---------------------------------------------------------------------------
# A whole deck from a plain list of slides (what the week builders use)
# ---------------------------------------------------------------------------
def build_deck(path, slides, sections_by_title=None, widescreen=True, title="", author="Lakeshore Paddle Company", before_save=None, after_save=None):
    """
    Make a deck from a list of slide dicts, in order:
      layout   - the layout name ("Title and Content")
      title    - the title text
      body     - list of lines for the body placeholder (text or (text, level))
      subtitle - the subtitle on a Title Slide
      notes    - speaker notes
      hidden   - True hides the slide
      fn       - fn(prs, slide) to add anything else (pictures, text boxes ...)
    sections_by_title: [(section name, title of its first slide), ...], in order.
    before_save(prs) runs on the python-pptx deck; after_save(path) can patch the file.
    """
    prs = new_deck(widescreen, title=title, author=author)
    made = []
    for spec in slides:
        s = add_slide(prs, spec.get("layout", "Title and Content"), spec.get("title"), spec.get("body"), spec.get("notes"))
        if spec.get("subtitle") is not None:
            sub = [p for p in s.placeholders if p.placeholder_format.idx == 1][0]
            fill_text(sub.text_frame, [spec["subtitle"]])
        if spec.get("hidden"):
            hide(s)
        if spec.get("fn"):
            spec["fn"](prs, s)
        made.append(s)
    if before_save:
        before_save(prs)
    save(prs, path)
    if sections_by_title:
        titles = [spec.get("title") or "" for spec in slides]
        starts = [titles.index(first) for _, first in sections_by_title]
        secs = []
        for i, (name, _) in enumerate(sections_by_title):
            end = starts[i + 1] if i + 1 < len(starts) else len(slides)
            secs.append((name, list(range(starts[i] + 1, end + 1))))
        sections(path, secs)
    if after_save:
        after_save(path)
    return path


def move(slides, title, after_title=None, to=None):
    """Move the slide dict with this title so it follows after_title (or to index `to`)."""
    i = [s.get("title") for s in slides].index(title)
    item = slides.pop(i)
    if to is not None:
        slides.insert(to if to >= 0 else len(slides) + 1 + to, item)
    else:
        j = [s.get("title") for s in slides].index(after_title)
        slides.insert(j + 1, item)
    return slides


def by_title(slides, title):
    return [s for s in slides if s.get("title") == title][0]


def picture(slide, image_path, x, y, w=None, h=None, name=None, alt="", decorative=False):
    """Insert > Pictures. python-pptx puts the file name in the alt text; PowerPoint doesn't, so clear it."""
    pic = slide.shapes.add_picture(image_path, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    pic._element.find(".//" + qn("p:cNvPr")).set("descr", "")
    if name:
        pic.name = name
    if alt or decorative:
        set_alt(pic, alt, decorative)
    return pic


def notes_master_text(prs, footer=None, header=None):
    """View > Notes Master: type in its footer (and header) placeholders."""
    nm = prs.notes_master
    for ph in nm.placeholders:
        t = int(ph.placeholder_format.type)
        if t == 15 and footer is not None:   # footer
            ph.text_frame.text = footer
        if t == 14 and header is not None:   # header
            ph.text_frame.text = header


def set_text_color(shape, color):
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor(*color)


def media(name):
    """A picture, icon, video, sound or 3D model used inside the practice files, made once in a temp folder.
    lake1.jpg ... lake4.jpg: photos; logo.png; icon-wave.png, icon-sun.png, icon-paddle.png; tour.mp4; music.m4a; kayak.glb; poster.png."""
    import tempfile
    folder = os.path.join(tempfile.gettempdir(), "office-trainer-pp-media")
    os.makedirs(folder, exist_ok=True)
    p = os.path.join(folder, name)
    if not os.path.exists(p):
        if name.startswith("lake"):
            make_photo(p, seed=int(name[4]) if name[4:5].isdigit() else 1)
        elif name == "logo.png":
            make_logo(p)
        elif name.startswith("icon-"):
            make_icon_png(p, name[5:-4])
        elif name.endswith(".mp4"):
            make_video(p, 8)
        elif name.endswith(".m4a"):
            make_audio(p, 10)
        elif name.endswith(".glb"):
            make_glb(p)
        elif name == "poster.png":
            poster(p, text="Lakeshore tour")
        else:
            raise KeyError(name)
    return p


def master_picture(prs, image_path, x, y, w, name="Logo", alt=""):
    """Insert > Pictures while the slide master is selected (python-pptx has no add_picture for masters)."""
    part = prs.slide_master.part
    image_part, rid = part.get_or_add_image_part(image_path)
    iw, ih = Image.open(image_path).size
    cx = int(w * 914400); cy = int(cx * ih / iw)
    tree = prs.slide_master._element.find(".//" + qn("p:spTree"))
    sid = max(int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))) + 1
    tree.append(X('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s" descr="%s"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
                  '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                  '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
                  % (sid, esc(name), esc(alt), rid, int(x * 914400), int(y * 914400), cx, cy))[0])


def autoshape(slide, mso_shape, x, y, w, h, text=None, name=None):
    """Insert > Shapes. python-pptx gives new shapes a strong style (fill 3, line 1, effect 2);
    PowerPoint's default is fill 1, line 2, effect 0, so set that, as a learner's new shape would have."""
    sh = slide.shapes.add_shape(mso_shape, Inches(x), Inches(y), Inches(w), Inches(h))
    st = sh._element.find(qn("p:style"))
    st.find(qn("a:lnRef")).set("idx", "2")
    st.find(qn("a:fillRef")).set("idx", "1")
    st.find(qn("a:effectRef")).set("idx", "0")
    ln_clr = st.find(qn("a:lnRef")).find(qn("a:schemeClr"))
    if ln_clr is not None and ln_clr.find(qn("a:shade")) is None:
        etree.SubElement(ln_clr, qn("a:shade")).set("val", "50000")
    if text is not None:
        sh.text = text
    if name:
        sh.name = name
    return sh


class Ref:
    """Stands in for a shape that is added after saving (a 3D model, a Zoom), so animate() can target its id."""
    def __init__(self, shape_id):
        self.shape_id = shape_id
        self.has_text_frame = False


def add_audio(slide, audio_path, x=11.5, y=0.4, size=0.6, icon=None):
    """Insert > Audio > Audio on My PC: python-pptx only knows video, so add it as a movie and relabel it as audio."""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    icon = icon or media("icon-wave.png")
    mv = slide.shapes.add_movie(audio_path, Inches(x), Inches(y), Inches(size), Inches(size), poster_frame_image=icon, mime_type="audio/mp4")
    vf = mv._element.find(".//" + qn("a:videoFile"))
    vf.tag = qn("a:audioFile")
    rid = vf.get(qn("r:link"))
    slide.part.rels[rid]._reltype = RT.AUDIO
    mv.name = os.path.basename(audio_path)
    return mv
