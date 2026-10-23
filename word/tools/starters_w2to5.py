# starters_w2to5.py - builds the practice documents for Weeks 2 to 5.
# make-starters.py runs this after Week 1:  python3 tools/make-starters.py
#
# Each starter holds exactly the problems its drills ask learners to fix, so the
# checks in the lesson files can be written against a known starting point.
# Everything here is fictional: people, businesses, data and the practice sources.
import pathlib
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from PIL import Image, ImageDraw
import docxkit as K

WORD = pathlib.Path(__file__).resolve().parent.parent
ST = WORD / "starters"


def folder(week):
    return ST / ("week-%02d" % week)


def save(doc, week, name):
    K.save(doc, folder(week), name, WORD)


# =========================================================================== pictures
def make_pictures():
    """Simple drawn 'photos' for Week 3 (and a picture for Week 5), so no one's photo rights are involved."""
    out = folder(3)
    out.mkdir(parents=True, exist_ok=True)
    # A lake at sunrise
    lake = Image.new("RGB", (1200, 600), "#8fc3e8")
    d = ImageDraw.Draw(lake)
    for y in range(0, 300):
        d.line([(0, y), (1200, y)], fill=(143 + y // 6, 195 - y // 10, 232 - y // 8))
    d.ellipse([820, 90, 960, 230], fill="#ffd36e")
    d.polygon([(0, 300), (200, 190), (420, 300)], fill="#4d7c4a")
    d.polygon([(300, 300), (620, 150), (900, 300)], fill="#3f6b3d")
    d.polygon([(760, 300), (1000, 210), (1200, 300)], fill="#557f50")
    d.rectangle([0, 300, 1200, 600], fill="#2f6f99")
    for y in range(320, 600, 28):
        d.line([(80, y), (1120, y)], fill="#4a8bb5", width=3)
    lake.save(out / "lake.jpg", quality=88)
    # A kayak on a plain light background (so Remove Background has something to remove)
    kayak = Image.new("RGB", (900, 500), "#f1efe8")
    d = ImageDraw.Draw(kayak)
    d.polygon([(60, 260), (450, 200), (840, 260), (450, 320)], fill="#e2572b", outline="#8a2905")
    d.ellipse([380, 235, 520, 285], fill="#2b2b2b")
    d.line([(250, 150), (650, 370)], fill="#6b4f2a", width=12)
    d.ellipse([220, 120, 290, 190], fill="#f2b705")
    d.ellipse([610, 330, 680, 400], fill="#f2b705")
    kayak.save(out / "kayak.png")
    # A wavy divider line (decorative)
    div = Image.new("RGBA", (1200, 40), (255, 255, 255, 0))
    d = ImageDraw.Draw(div)
    pts = [(x, 20 + (8 if (x // 40) % 2 else -8)) for x in range(0, 1201, 40)]
    d.line(pts, fill="#0e7c86", width=6)
    div.save(out / "divider.png")
    # A sunset (for Skill Check 3)
    sun = Image.new("RGB", (1200, 600), "#f6a96b")
    d = ImageDraw.Draw(sun)
    for y in range(0, 330):
        d.line([(0, y), (1200, y)], fill=(246 - y // 8, 169 - y // 4, 107 + y // 6))
    d.ellipse([500, 230, 700, 430], fill="#ffdf7a")
    d.rectangle([0, 330, 1200, 600], fill="#3d4f7a")
    sun.save(out / "sunset.jpg", quality=88)
    print("wrote pictures in", out.relative_to(WORD))


def picture(doc, path, width, name, after=6):
    """Insert a picture in its own paragraph and give it a name (shown in the Selection pane)."""
    doc.add_picture(str(path), width=Inches(width))
    pics = list(doc.element.body.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"))
    pics[-1].set("name", name)
    doc.paragraphs[-1].paragraph_format.space_after = Pt(after)


# =========================================================================== Week 2: the field report
# The report every Week 2 lesson works on. (kind, text). Kinds: title, byline, note, h1, h2, p, label,
# steps (the sampling steps), bullets (equipment), findings, recs, lab (numbered lists), table.
REPORT = [
    ("title", "Keeping the Lake Clean: A Student Field Study of the North Shore"),
    ("byline", "ENV 110 Field Team, Finger Lakes Community College, Fall 2026"),
    ("note", "Practice report: the sites, data and studies in it are made up for this course."),
    ("h1", "Introduction"),
    ("p", "Every summer, thousands of people swim, fish and paddle along the north shore of the lake. Our team spent six weeks measuring how clean that shoreline really is, and what visitors leave behind."),
    ("p", "This report explains what we measured, what we found and what we recommend. Read how we collected our samples before you compare our numbers with other studies, because small changes in method change the results."),
    ("h1", "Background"),
    ("h2", "Why the shoreline matters"),
    ("p", "The shallow water near shore is where young fish feed and where most people enter the lake. Litter and runoff collect there first. A clean shoreline is the first sign of a healthy lake, and the first thing a visitor notices."),
    ("h2", "What earlier surveys found"),
    ("p", "Two earlier student surveys counted litter at the public beach. Both found that plastic wrappers and bottle caps made up most of the items. Neither measured water clarity, which is why we added it this year."),
    ("h1", "Methods"),
    ("p", "We collected samples at six sites along a two-mile stretch of shore. Each site was visited every Saturday morning in September and October, for twelve visits in all. Litter was counted in a 10-meter strip at each site."),
    ("label", "Sampling steps"),
    ("steps", ["Walk to the marked sampling site.", "Prepare the sample bottle.", "Label it with the site number.", "Rinse it twice with lake water.", "Fill the bottle 10 cm below the surface.", "Store it in the cooler and record the time."]),
    ("label", "Equipment"),
    ("bullets", ["Sample bottles", "Secchi disk", "Clipboard and pencil", "Cooler with ice packs"]),
    ("h1", "Results"),
    ("h2", "Water clarity"),
    ("p", "Clarity was measured with Secchi disk readings at every visit. The average reading was 4.1 meters, with the clearest water at the two sites farthest from the boat launch."),
    ("h2", "Litter counts"),
    ("p", "Volunteers counted 1,284 pieces of litter over the twelve visits. Counts were highest after warm weekends and lowest after storms, which wash litter away before it can be counted."),
    ("label", "Findings"),
    ("findings", ["Plastic wrappers were the most common item at every site.", "The boat launch had three times more litter than any other site.", "Clarity was lowest where the most litter was found."]),
    ("h1", "Discussion"),
    ("p", "Our results match the earlier surveys: plastic is the main problem, and it gathers where people gather. The link between litter and clarity is new, but twelve visits is not enough to say one causes the other."),
    ("p", "A longer study over a full year, with more sites, would show whether the pattern holds when the beach is quiet."),
    ("h1", "Limitations"),
    ("p", "We sampled only in the fall, only on Saturday mornings, and only along one stretch of shore. Volunteers changed from week to week, so some counts may not be comparable."),
    ("h1", "Recommendations"),
    ("p", "Based on these findings, we make three recommendations to the town and to the shoreline association."),
    ("recs", ["Reduce single-use plastic at the beach snack bar.", "Add two trash and recycling stations near the boat launch.", "Run a volunteer cleanup on the first Saturday of each month."]),
    ("label", "Lab follow-up"),
    ("lab", ["Filter each sample through a 0.5 mm screen.", "Count and photograph anything the screen catches.", "Enter the results in the shared spreadsheet."]),
    ("h1", "Appendix: Sampling Data"),
    ("table", [["Site", "Date", "Time", "Clarity (m)", "Temp (°C)", "Litter items", "Notes"],
               ["1", "Sept 6", "8:10", "4.6", "19", "41", "Calm"],
               ["2", "Sept 6", "8:35", "4.3", "19", "58", "Light wind"],
               ["3 (launch)", "Sept 6", "9:05", "3.2", "20", "187", "Busy weekend"],
               ["4", "Sept 6", "9:30", "4.0", "20", "66", "Calm"],
               ["5", "Sept 6", "9:55", "4.7", "20", "35", "Calm"],
               ["6", "Sept 6", "10:20", "4.8", "21", "29", "Geese on shore"]]),
]

HEADINGS = {"h1": "Heading 1", "h2": "Heading 2", "title": "Title"}


def build_report(styled=True, swap_methods=False, plain_steps=False, limitations_plain=False, lists="good", notes=False):
    """The field report. Options set up each lesson's starting point:
    styled: headings use real heading styles (False = bold Normal text, as if formatted by hand)
    swap_methods: Methods comes after Results (for the Navigation pane drill)
    plain_steps: the sampling steps are plain paragraphs, not a list
    limitations_plain: the Limitations heading is plain text
    lists: "good" (every list right) or "messy" (recommendations continue the findings' numbers, lab list starts at 1)"""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    items = list(REPORT)
    if swap_methods:
        a = next(i for i, x in enumerate(items) if x == ("h1", "Methods"))
        b = next(i for i, x in enumerate(items) if x == ("h1", "Results"))
        c = next(i for i, x in enumerate(items) if x == ("h1", "Discussion"))
        methods, results = items[a:b], items[b:c]
        items = items[:a] + results + methods + items[c:]
    bullet = K.add_numbering(doc, "bullet", "", bullet_font="Symbol", levels=[("bullet", "o"), ("bullet", "")])
    steps_num = K.add_numbering(doc, "decimal", "%1.")
    findings_num = K.add_numbering(doc, "decimal", "%1.")
    lab_num = K.add_numbering(doc, "decimal", "%1.")
    recs_num = findings_num if lists == "messy" else K.restart_num(doc, findings_num, 1)
    if lists == "good":
        lab_num = K.restart_num(doc, steps_num, 5)
    for kind, value in items:
        if kind in ("title", "h1", "h2"):
            plain = (not styled) or (kind == "h1" and value == "Limitations" and limitations_plain)
            if plain:
                size = {"title": 20, "h1": 14, "h2": 12}[kind]
                K.para(doc, value, bold=True, size=size, after=6, before=12 if kind != "title" else 0)
            else:
                doc.add_paragraph(value, style=HEADINGS[kind])
        elif kind == "byline":
            K.para(doc, value, italic=True)
        elif kind == "note":
            K.para(doc, value, size=9, color="4B5563")
        elif kind == "label":
            K.para(doc, value, bold=True, after=4)
        elif kind == "p":
            K.para(doc, value)
        elif kind == "steps":
            for i, text in enumerate(value):
                p = K.para(doc, text, after=2)
                if not plain_steps:
                    K.set_list(p, steps_num, 1 if text.startswith(("Label it", "Rinse it")) else 0)
        elif kind == "bullets":
            for text in value:
                K.set_list(K.para(doc, text, after=2), bullet)
        elif kind in ("findings", "recs", "lab"):
            num = {"findings": findings_num, "recs": recs_num, "lab": lab_num}[kind]
            for text in value:
                K.set_list(K.para(doc, text, after=2), num)
        elif kind == "table":
            t = doc.add_table(rows=len(value), cols=len(value[0]))
            t.style = "Table Grid"
            for r, row in enumerate(value):
                for c, text in enumerate(row):
                    t.cell(r, c).text = text
            K.para(doc, "")
    return doc


def week2():
    # Lesson 1: styles. Headings are formatted by hand; Normal is still Aptos.
    save(build_report(styled=False), 2, "W2-L1-report.docx")
    # Lesson 2: navigation. Real headings, but Methods sits after Results.
    save(build_report(swap_methods=True), 2, "W2-L2-report.docx")
    # Lesson 3: lists, headers and breaks. Plain sampling steps, numbering that runs on, no header.
    save(build_report(plain_steps=True, lists="messy"), 2, "W2-L3-report.docx")
    # Lesson 4: notes, contents and citations. Limitations isn't a heading yet; no table of contents.
    save(build_report(limitations_plain=True), 2, "W2-L4-report.docx")
    save(handbook(), 2, "W2-SC2-handbook.docx")


HANDBOOK = [
    ("title", "Lakeshore Paddle Company Staff Handbook"),
    ("p", "Read this handbook before your first shift. If anything here doesn't match what you see at the shop, ask the manager on duty; you can also see the budget for this season below."),
    ("h1", "Welcome"),
    ("p", "Lakeshore Paddle Company rents kayaks, canoes and paddleboards from May to October. We hire about twenty seasonal staff each year, many of them college students."),
    ("h1", "Your first week"),
    ("h2", "Opening the shop"),
    ("label", "Opening checklist"),
    ("open", ["Unlock the boat racks.", "Count the cash drawer.", "Check the life jacket bins.", "Turn on the dock lights."]),
    ("h2", "Closing the shop"),
    ("label", "Closing checklist"),
    ("close", ["Lock the boat racks.", "Count the cash drawer again.", "Hang up the wet life jackets."]),
    ("h1", "Rentals"),
    ("p", "Every rental starts with a signed waiver and a life jacket fitting. Rentals run by the hour, the half day or the full day."),
    ("h1", "Safety on the water"),
    ("p", "Every renter must wear a life jacket at all times. Staff watch the weather radio and call boats in when wind gusts pass 20 miles per hour."),
    ("h1", "Timeline"),
    ("p", "Training runs the last week of April. The shop opens on Memorial Day weekend and closes on Columbus Day."),
    ("h1", "Budget"),
    ("p", "Our seasonal budget covers wages, new life jackets, two new kayaks and dock repairs. Staff can suggest purchases at the monthly meeting."),
    ("h1", "Appendix: Contact List"),
    ("table", [["Name", "Role", "Phone", "Email", "Days", "Shift", "Notes"],
               ["Jordan Rivera", "Manager", "555-0100", "jordan@example.com", "Mon-Fri", "8-4", "Keys"],
               ["Priya Shah", "Owner", "555-0101", "priya@example.com", "Sat-Sun", "9-5", "Payroll"],
               ["Sam Ortiz", "Lead guide", "555-0102", "sam@example.com", "Wed-Sun", "7-3", "First aid"]]),
]


def handbook():
    """Skill Check 2: headings by hand, Budget after Timeline, plain opening list, closing list runs on."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    num = K.add_numbering(doc, "decimal", "%1.")
    for kind, value in HANDBOOK:
        if kind == "title":
            doc.add_paragraph(value, style="Title")
        elif kind in ("h1", "h2"):
            K.para(doc, value, bold=True, size=14 if kind == "h1" else 12, before=12, after=6)
        elif kind == "label":
            K.para(doc, value, bold=True, after=4)
        elif kind == "p":
            K.para(doc, value)
        elif kind == "open":
            for text in value:
                K.para(doc, text, after=2)
        elif kind == "close":
            # The closing list is numbered, but (below) set to start at 5: the learner restarts it at 1.
            for text in value:
                K.set_list(K.para(doc, text, after=2), num)
        elif kind == "table":
            t = doc.add_table(rows=len(value), cols=len(value[0]))
            t.style = "Table Grid"
            for r, row in enumerate(value):
                for c, text in enumerate(row):
                    t.cell(r, c).text = text
            K.para(doc, "")
    # Make the closing list show 5, 6, 7: a new instance of its list that starts at 5.
    restarted = K.restart_num(doc, num, 5)
    closing = dict(HANDBOOK)["close"]
    for p in doc.paragraphs:
        if p.text in closing:
            K.set_list(p, restarted)
    return doc


# =========================================================================== Week 3: Paddle Day flyer and newsletter
def table_from(doc, rows, style="Table Grid"):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = style
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            t.cell(r, c).text = text
    return t


def flyer():
    """Lesson 1: a flyer with two pictures (no alt text) where shapes, SmartArt and a text box will go."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("Paddle Day at Lakeshore", style="Title")
    picture(doc, folder(3) / "lake.jpg", 6.0, "Lake photo")
    K.para(doc, "Saturday, June 12, 2027 · 9 a.m. to 3 p.m. · North Shore boat launch", bold=True)
    picture(doc, folder(3) / "divider.png", 6.0, "Divider line")
    K.para(doc, "Try a kayak, a canoe or a paddleboard for free. Our staff will help you choose a boat, fit a life jacket and launch safely. No experience needed.")
    doc.add_paragraph("How it works", style="Heading 1")
    K.para(doc, "Reserve a time, check in at the tent, listen to the five-minute safety talk, then launch.")
    doc.add_paragraph("Meet the fleet", style="Heading 1")
    K.para(doc, "We'll have twelve sit-on-top kayaks, four canoes and six paddleboards on the beach, plus two tandem kayaks for families.")
    doc.add_paragraph("What paddlers say", style="Heading 1")
    K.para(doc, "Last year 214 people tried paddling at our first Paddle Day. One of them wrote: Best morning of my summer. I didn't know I could do this.")
    K.para(doc, "Questions? Call the shop at 555-0100 or stop by 12 Harbor Street.")
    return doc


def schedule():
    """Lesson 2: tab-separated text (to table), an old table (to text), and the tour table to fix."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("Paddle Day Schedule", style="Title")
    doc.add_paragraph("Schedule", style="Heading 1")
    for line in ["Time\tActivity\tWhere", "9:00\tCheck-in opens\tTent", "9:30\tSafety talk\tBeach", "10:00\tFree demos\tBoat launch",
                 "12:00\tLunch break\tPicnic area", "1:00\tFamily paddle\tNorth cove", "3:00\tLast boats in\tBoat launch"]:
        K.para(doc, line, after=0)
    doc.add_paragraph("Old price list", style="Heading 1")
    K.para(doc, "We no longer use this table on the website. Turn it back into plain lines of text.")
    table_from(doc, [["Rental", "Price"], ["Kayak (1 hour)", "$20"], ["Kayak (half day)", "$55"], ["Canoe (1 hour)", "$25"], ["Paddleboard (1 hour)", "$22"]])
    doc.add_paragraph("Summer tours", style="Heading 1")
    rows = [["Summer tours 2027", "", "", ""], ["Date", "Tour", "Guide", "Spots"],
            ["July 10, 2027", "Sunrise cove", "Sam", "8"], ["June 12, 2027", "Paddle Day demo", "Priya", "20"],
            ["August 21, 2027", "Lighthouse loop", "Sam", "6"], ["June 26, 2027", "Wildlife watch", "Jordan", "10"],
            ["July 24, 2027", "Sunset paddle", "Priya", "12"], ["August 7, 2027", "Family float", "Jordan", "14"],
            ["June 19, 2027", "Beginner basics", "Sam", "10"], ["July 31, 2027", "North shore cleanup", "Jordan", "16"],
            ["August 28, 2027", "Season finale", "Priya", "20"]]
    table_from(doc, rows)
    K.para(doc, "")
    doc.add_paragraph("Sign-up sheet", style="Heading 1")
    K.para(doc, "Insert the sign-up table below this line.")
    K.para(doc, "")
    return doc


def survey():
    """Lesson 3: survey numbers to chart."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("Paddler Survey Results", style="Title")
    K.para(doc, "After Paddle Day, 100 visitors told us which boat they would rent first.")
    table_from(doc, [["Boat", "Votes"], ["Kayak", "42"], ["Paddleboard", "31"], ["Canoe", "18"], ["Pedal boat", "9"]])
    K.para(doc, "")
    K.para(doc, "Insert the chart below this line, then add its caption under it.")
    K.para(doc, "")
    K.para(doc, "Kayaks were the clear favorite, so we'll add four more before next season.")
    return doc


NEWSLETTER_BODY = [
    ("h2", "Welcome"),
    ("p", "Welcome to the summer issue of The Shoreline. It has been our busiest spring ever: more than 1,900 rentals in May and June, and two new guides on the team. Here's what's new on the water."),
    ("h2", "New boats on the rack"),
    ("p", "Four sit-on-top kayaks and two family canoes arrived in May. They're lighter than our old boats, so they're easier to carry to the water."),
    ("h2", "Paddle Day recap"),
    ("p", "214 people tried paddling for free at our first Paddle Day. Thanks to the twelve volunteers who fitted life jackets, held boats steady and cheered every first launch."),
    ("h2", "Safety corner"),
    ("p", "Afternoon winds on the lake often reach 15 miles per hour. Plan to paddle into the wind first, so the trip back is the easy part."),
    ("h2", "Staff picks"),
    ("p", "Sam's favorite route is the sunrise loop past the lighthouse. Priya recommends the north cove for spotting herons. Jordan's tip: bring a dry bag for your phone."),
    ("p", "Contact us: 555-0100 · 12 Harbor Street, Canandaigua, NY · Open daily 9 to 6."),
]


def newsletter():
    """Lesson 4: one column all the way through, no theme colors of our own, no page design."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("The Shoreline", style="Title")
    K.para(doc, "News from Lakeshore Paddle Company · Summer 2027", italic=True, after=12)
    for kind, text in NEWSLETTER_BODY:
        if kind == "h2":
            doc.add_paragraph(text, style="Heading 2")
        else:
            K.para(doc, text)
    return doc


def sc3_flyer():
    """Skill Check 3: a flyer with a photo, tab-separated text, a prices table and an About section."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("Sunset Paddle Series", style="Title")
    picture(doc, folder(3) / "sunset.jpg", 5.5, "Sunset photo")
    K.para(doc, "Every Thursday in July · 7 p.m. · North Shore boat launch", bold=True)
    doc.add_paragraph("Dates", style="Heading 1")
    for line in ["Date\tSunset tour\tGuide", "July 1\tLighthouse loop\tSam", "July 8\tNorth cove\tPriya", "July 15\tIsland circle\tJordan", "July 22\tLighthouse loop\tSam"]:
        K.para(doc, line, after=0)
    doc.add_paragraph("Prices", style="Heading 1")
    table_from(doc, [["Sunset series prices", ""], ["Option", "Price"], ["Paddleboard", "$30"], ["Single kayak", "$25"], ["Family canoe", "$45"], ["Tandem kayak", "$40"]])
    K.para(doc, "")
    doc.add_paragraph("About the series", style="Heading 1")
    K.para(doc, "Each tour lasts about ninety minutes and ends back at the launch before dark. Guides bring lights, a first aid kit and a radio.")
    K.para(doc, "No experience needed. We'll teach you a forward stroke, a turn and how to get back on a paddleboard if you fall off.")
    K.para(doc, "Book online or at the shop. Tours run in light rain but not in thunder, wind over 15 miles per hour or fog.")
    return doc


def week3():
    make_pictures()
    save(flyer(), 3, "W3-L1-flyer.docx")
    save(schedule(), 3, "W3-L2-schedule.docx")
    save(survey(), 3, "W3-L3-survey.docx")
    save(newsletter(), 3, "W3-L4-newsletter.docx")
    save(sc3_flyer(), 3, "W3-SC3-flyer.docx")


# =========================================================================== Week 4: the job application packet
PERSONA = "Alex Morgan"   # a made-up job seeker learners may use instead of themselves


def resume_template():
    """Lesson 1: a resume template with content controls still showing their placeholder text."""
    doc = K.new_doc(font="Aptos", size=11, after=4)
    for placeholder, style in [("[Your full name]", "Title"), ("[Email · Phone · City, State]", None)]:
        p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        K.content_control(p, placeholder)
    doc.add_paragraph("Experience", style="Heading 1")
    for placeholder in ["[Job title], [Company]", "[Start year] to [End year]", "[One sentence about what you did and what changed because of it]"]:
        K.content_control(doc.add_paragraph(), placeholder)
    doc.add_paragraph("Education", style="Heading 1")
    for placeholder in ["[School name]", "[Degree or program, expected graduation]"]:
        K.content_control(doc.add_paragraph(), placeholder)
    doc.add_paragraph("Skills", style="Heading 1")
    K.content_control(doc.add_paragraph(), "[Three or four skills, separated by commas]")
    return doc


def cover_letter_quickpart():
    doc = K.new_doc(font="Aptos", size=11, after=8)
    K.para(doc, "[Contact line]", after=12)
    K.para(doc, "October 20, 2026")
    K.para(doc, "Hiring Manager", after=0)
    K.para(doc, "Finger Lakes Outfitters", after=0)
    K.para(doc, "88 Lake Road, Geneva, NY 14456")
    K.para(doc, "Dear Hiring Manager:")
    K.para(doc, "I'm writing to apply for the Office Coordinator position. For two seasons I kept the rental desk at Lakeshore Paddle Company running: scheduling staff, answering customers and keeping the books balanced at the end of each day.")
    K.para(doc, "I'd welcome the chance to talk about how I can help your team. Thank you for your time.")
    K.para(doc, "Sincerely,", after=36)
    K.para(doc, PERSONA)
    return doc


def ats_resume():
    """Lesson 2: dates pushed right with spaces, and a layout table that hiring software can misread."""
    doc = K.new_doc(font="Aptos", size=11, after=4)
    K.para(doc, PERSONA, bold=True, size=16)
    K.para(doc, "alex.morgan@example.com · 555-0142 · Canandaigua, NY", after=10)
    doc.add_paragraph("Experience", style="Heading 1")
    for title, dates, desc in [("Office Assistant, Lakeshore Paddle Company", "2025 to 2026", "Scheduled 20 seasonal staff and balanced the rental desk every day."),
                               ("Volunteer Coordinator, North Shore Cleanup", "2023 to 2025", "Recruited and trained 40 volunteers for monthly cleanups."),
                               ("Cashier, Main Street Market", "2022 to 2023", "Handled about 150 transactions a shift with no drawer errors.")]:
        K.para(doc, title + "          " + dates, bold=True, after=0)
        K.para(doc, desc)
    doc.add_paragraph("Skills", style="Heading 1")
    table_from(doc, [["Scheduling", "Customer service"], ["Microsoft Word and Excel", "Cash handling"]], style="Normal Table")
    K.para(doc, "")
    doc.add_paragraph("Education", style="Heading 1")
    K.para(doc, "Finger Lakes Community College, A.A.S. Business Administration, expected May 2027")
    return doc


EMPLOYERS = [["Company", "Contact", "Title", "Address", "City", "State", "ZIP", "Position", "Remote"],
             ["Finger Lakes Outfitters", "Dana Lee", "Ms.", "88 Lake Road", "Geneva", "NY", "14456", "Office Coordinator", "No"],
             ["Canandaigua Community Center", "Chris Patel", "Mr.", "20 Main Street", "Canandaigua", "NY", "14424", "Program Assistant", "No"],
             ["Seneca Shore Inn", "Morgan Brooks", "Mx.", "5 Harbor Way", "Watkins Glen", "NY", "14891", "Guest Services Lead", "Yes"]]


def employers():
    doc = K.new_doc(font="Aptos", size=10, after=0)
    table_from(doc, EMPLOYERS)
    return doc


def merge_letter():
    """Lesson 3: a cover letter with bracketed spots where merge fields will go."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    K.para(doc, PERSONA, bold=True, after=0)
    K.para(doc, "alex.morgan@example.com · 555-0142 · Canandaigua, NY", after=12)
    K.para(doc, "October 20, 2026")
    K.para(doc, "[Address Block]")
    K.para(doc, "[Greeting Line]")
    K.para(doc, "I'm writing to apply for the [Position] position. For two seasons I kept the rental desk at Lakeshore Paddle Company running: scheduling staff, answering customers and balancing the books at the end of each day.")
    K.para(doc, "[Remote sentence]")
    K.para(doc, "I'd welcome the chance to talk about how I can help your team. Thank you for your time.")
    K.para(doc, "Sincerely,", after=36)
    K.para(doc, PERSONA)
    return doc


def blank():
    doc = K.new_doc(font="Aptos", size=11, after=8)
    return doc


def web_resume():
    """Lesson 5: a resume with a raw web address that needs descriptive link text."""
    doc = K.new_doc(font="Aptos", size=11, after=4)
    K.para(doc, PERSONA, bold=True, size=18)
    K.para(doc, "alex.morgan@example.com · 555-0142 · Canandaigua, NY", after=4)
    K.para(doc, "Portfolio: [add a link to your portfolio or LinkedIn profile here]", after=10)
    doc.add_paragraph("Experience", style="Heading 1")
    K.para(doc, "Office Assistant, Lakeshore Paddle Company, 2025 to 2026", bold=True, after=0)
    K.para(doc, "Scheduled 20 seasonal staff and balanced the rental desk every day.")
    doc.add_paragraph("Education", style="Heading 1")
    K.para(doc, "Finger Lakes Community College, A.A.S. Business Administration, expected May 2027")
    return doc


MEMBERS = [["First", "Last", "Address", "City", "State", "ZIP", "Level", "RenewBy"],
           ["Jamie", "Cole", "14 Pine Street", "Canandaigua", "NY", "14424", "Gold", "June 1"],
           ["Taylor", "Nguyen", "302 Elm Avenue", "Geneva", "NY", "14456", "Silver", "June 15"],
           ["Riley", "Park", "9 Shore Drive", "Naples", "NY", "14512", "Gold", "July 1"],
           ["Avery", "Diaz", "77 Mill Road", "Victor", "NY", "14564", "Silver", "July 15"]]


def sc4_letter():
    doc = K.new_doc(font="Aptos", size=11, after=8)
    K.para(doc, "Lakeshore Paddle Club", bold=True, size=14, after=0)
    K.para(doc, "12 Harbor Street, Canandaigua, NY 14424", after=12)
    K.para(doc, "[Address Block]")
    K.para(doc, "[Greeting Line]")
    K.para(doc, "Thank you for paddling with us this year. Your [Level] membership is due for renewal by [RenewBy].")
    K.para(doc, "[Level sentence]")
    K.para(doc, "See you on the water,", after=36)
    K.para(doc, "Jordan Rivera, Club Coordinator")
    return doc


def week4():
    save(resume_template(), 4, "W4-L1-resume-template.docx")
    save(cover_letter_quickpart(), 4, "W4-L1-cover-letter.docx")
    save(ats_resume(), 4, "W4-L2-resume-ats.docx")
    save(merge_letter(), 4, "W4-L3-letter.docx")
    e = K.new_doc(font="Aptos", size=10, after=0); table_from(e, EMPLOYERS); save(e, 4, "W4-employers.docx")
    save(blank(), 4, "W4-L4-blank.docx")
    save(web_resume(), 4, "W4-L5-resume.docx")
    save(sc4_letter(), 4, "W4-SC4-letter.docx")
    m = K.new_doc(font="Aptos", size=10, after=0); table_from(m, MEMBERS); save(m, 4, "W4-SC4-members.docx")


# =========================================================================== Week 5: review and ship
POLICY = [
    ("title", "Rental Safety Policy (Draft 3)"),
    ("p", "This policy applies to every rental from the Lakeshore Paddle Company dock."),
    ("h1", "Life jackets"),
    ("p", "Every renter should wear a life jacket that fits, from launch to return."),
    ("h1", "Refunds"),
    ("p", "Cancellations made at least 48 hours ahead are fully refundable. Later cancellations get a credit for another day."),
    ("h1", "Weather"),
    ("p", "Staff call boats in when gusts pass 20 miles per hour, when thunder is heard, and so on."),
    ("h1", "Paddle Pass"),
    ("p", "Paddle Pass holders may reserve a boat up to two weeks ahead."),
]


def policy():
    """Lesson 1: three comments to manage, and Sam's tracked changes to review."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    paras = {}
    for kind, text in POLICY:
        if kind == "title":
            paras[text] = doc.add_paragraph(text, style="Title")
        elif kind == "h1":
            paras[text] = doc.add_paragraph(text, style="Heading 1")
        else:
            paras[text[:12]] = K.para(doc, text)
    # Sam's tracked changes: two insertions, two deletions, two formatting changes
    first = paras["This policy "]
    K.tracked_insert(first, " Read it before your first shift.", "Sam Ortiz")
    weather = paras["Staff call b"]
    weather.runs[0].text = "Staff call boats in when gusts pass 20 miles per hour, when thunder is heard"
    K.tracked_delete(weather, ", and so on", "Sam Ortiz")
    K.raw_run(weather, '<w:r %s><w:t>.</w:t></w:r>' % K.NSW)
    refunds = paras["Cancellation"]
    refunds.runs[0].text = "Cancellations made at least 48 hours ahead are "
    K.tracked_delete(refunds, "fully refundable", "Sam Ortiz")
    K.tracked_insert(refunds, "refunded minus a $5 fee", "Sam Ortiz")
    K.raw_run(refunds, '<w:r %s><w:t xml:space="preserve">. Later cancellations get a credit for another day.</w:t></w:r>' % K.NSW)
    jackets = paras["Every renter"]
    jackets.runs[0].text = "Every renter "
    r = jackets.add_run("should")
    K.tracked_bold(r, "Sam Ortiz")
    jackets.add_run(" wear a life jacket that fits, from launch to return.")
    K.tracked_bold(paras["Paddle Pass"].runs[0], "Sam Ortiz")
    K.comments(doc, [
        {"anchor": jackets, "author": "Priya Shah", "text": "Should this say must instead of should?"},
        {"anchor": refunds, "author": "Sam Ortiz", "text": "Is 48 hours right? Check with Priya."},
        {"anchor": weather, "author": "Priya Shah", "text": "Old note: delete me before we share this."},
    ])
    return doc


def checklist_draft():
    """Lesson 1, drill 2: a clean draft to edit with Track Changes on."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.add_paragraph("Rental Desk Checklist (Draft)", style="Title")
    for line in ["Greet every customer within one minute.", "Check the waiver is signed before handing over a paddle.",
                 "Fit each life jacket and check the buckle.", "Write the boat number and return time on the rental card.",
                 "Remind renters where the no-wake zone starts."]:
        K.para(doc, line)
    return doc


LETTER = ["Dear Ms. Lee:", "I'm writing to apply for the Office Coordinator position at Finger Lakes Outfitters.",
          "For two seasons I ran the rental desk at Lakeshore Paddle Company. I scheduled staff, answered customers and balanced the books every day.",
          "I would love to talk about how I can help your team grow.", "Sincerely,", "Alex Morgan"]


def letter_version(edits=None, author=None):
    """The cover letter: the original, a revised copy, or a reviewer's copy with tracked changes."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.core_properties.author = author or PERSONA
    for line in LETTER:
        p = K.para(doc, line)
        for (old, new, by) in (edits or []):
            if line.startswith(old[:20]) and by:
                p.runs[0].text = ""
                K.tracked_delete(p, old, by)
                K.tracked_insert(p, new, by)
    return doc


def letter_revised():
    doc = K.new_doc(font="Aptos", size=11, after=8)
    doc.core_properties.author = PERSONA
    revised = ["Dear Ms. Lee:", "I'm writing to apply for the Office Coordinator position at Finger Lakes Outfitters, which I saw on your website.",
               "For two seasons I ran the rental desk at Lakeshore Paddle Company, where I scheduled twenty staff and balanced the books every day.",
               "I'd welcome the chance to talk about how I can help your team.", "Sincerely,", "Alex Morgan"]
    for line in revised:
        K.para(doc, line)
    return doc


def ship():
    """Lesson 3: everything that shouldn't leave the building: comments, hidden text, author info,
    Compatibility Mode, and accessibility problems (no alt text, no table header, no title, a skipped
    heading level, a 'click here' link)."""
    doc = K.new_doc(font="Aptos", size=11, after=8, compat=14)
    doc.core_properties.author = "Jordan Rivera"
    doc.core_properties.last_modified_by = "Jordan Rivera"
    doc.add_paragraph("Paddle Day Volunteer Guide", style="Heading 1")
    p1 = K.para(doc, "Thank you for volunteering at Paddle Day. This guide covers what to wear, where to be and who to ask.")
    hidden = K.para(doc, "Internal note: the staff discount code is PADDLE50.")
    hidden.runs[0].font.hidden = True
    picture(doc, folder(3) / "lake.jpg", 5.0, "Picture 1")
    doc.add_paragraph("What to wear", style="Heading 3")
    K.para(doc, "Wear clothes that can get wet, sunscreen and shoes that stay on in the water.")
    doc.add_paragraph("Shifts", style="Heading 1")
    table_from(doc, [["Shift", "Time", "Lead"], ["Morning", "8:30 to 12:00", "Sam"], ["Afternoon", "12:00 to 3:30", "Priya"]])
    K.para(doc, "")
    p2 = K.para(doc, "To sign up for a shift, ")
    K.hyperlink(doc, p2, "click here", url="https://www.example.org/paddle-day")
    p2.add_run(".")
    K.comments(doc, [
        {"anchor": p1, "author": "Priya Shah", "text": "Can we add the parking info here?"},
        {"anchor": p2, "author": "Sam Ortiz", "text": "Link checked, works."},
    ])
    return doc


SEASON = [
    ("title", "Season Report: Lakeshore Paddle Company"),
    ("p", "This was our strongest season yet, thanks to our volunteers and a dry, warm July."),
    ("h1", "Rentals"),
    ("p", "We logged 6,420 rentals, up 12 percent on last year. Kayaks made up more than half of them."),
    ("h1", "Programs"),
    ("p", "Paddle Day brought 214 first-time paddlers to the shore. The sunset series sold out every week."),
    ("h1", "Next season"),
    ("p", "We plan to add four kayaks and a second Paddle Day in August."),
]


def season_report():
    """Skill Check 5: comments, tracked changes, hidden text and a picture with no alt text."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    paras = {}
    for kind, text in SEASON:
        style = {"title": "Title", "h1": "Heading 1"}.get(kind)
        paras[text[:12]] = doc.add_paragraph(text, style=style) if style else K.para(doc, text)
    intro = paras["This was our"]
    intro.runs[0].text = "This was our strongest season yet, "
    K.tracked_delete(intro, "thanks to our volunteers", "Sam Ortiz")
    K.tracked_insert(intro, "helped by", "Sam Ortiz")
    K.raw_run(intro, '<w:r %s><w:t xml:space="preserve"> a dry, warm July.</w:t></w:r>' % K.NSW)
    K.tracked_bold(paras["Rentals"].runs[0], "Sam Ortiz")
    K.tracked_bold(paras["Programs"].runs[0], "Sam Ortiz")
    rentals = paras["We logged 6,"]
    K.tracked_insert(rentals, " Paddleboards were close behind.", "Priya Shah")
    hidden = K.para(doc, "Draft numbers: check with accounting before sharing.")
    hidden.runs[0].font.hidden = True
    picture(doc, folder(3) / "lake.jpg", 4.5, "Picture 2")
    programs = paras["Paddle Day b"]
    K.comments(doc, [
        {"anchor": rentals, "author": "Sam Ortiz", "text": "Can you double-check the rental numbers?"},
        {"anchor": programs, "author": "Priya Shah", "text": "Love this section."},
        {"anchor": paras["We plan to a"], "author": "Priya Shah", "text": "Remove before sharing: budget not approved yet."},
    ])
    return doc


def exam():
    """The capstone practice exam: one document with something to do from every MO-110 area."""
    doc = K.new_doc(font="Aptos", size=11, after=8)
    sec = doc.sections[0]                       # A4 with narrow margins, to be fixed
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(0.6))
    K.para(doc, "Lakeshore Paddle Company Annual Report", bold=True, size=20, after=12)
    K.para(doc, "Prepared for the Lakeshore Paddle Co. board, October 2026")
    for title, body in [("Our season", "Lakeshore Paddle Co. had its busiest season yet, with more than six thousand rentals and two new programs."),
                        ("Highlights", "Paddle Day brought first-time paddlers to the shore. The sunset series sold out. Lakeshore Paddle Co. added six boats and hired four new guides. We also rebuilt the north dock before Memorial Day.")]:
        K.para(doc, title, bold=True, size=14, before=12, after=6)
        K.para(doc, body)
    picture(doc, folder(3) / "lake.jpg", 3.0, "Report photo")
    K.para(doc, "Rentals", bold=True, size=14, before=12, after=6)
    table_from(doc, [["Boat", "Rentals", "Change"], ["Kayak", "3,410", "+14%"], ["Paddleboard", "1,780", "+9%"], ["Canoe", "1,230", "+6%"]])
    K.para(doc, "")
    K.para(doc, "Goals for next year", bold=True, size=14, before=12, after=6)
    for g in ["Add four kayaks.", "Hold a second Paddle Day.", "Train every guide in first aid."]:
        K.para(doc, g, after=2)
    K.para(doc, "Staff", bold=True, size=14, before=12, after=6)
    for line in ["Name\tRole\tSeasons", "Jordan Rivera\tManager\t6", "Priya Shah\tOwner\t11", "Sam Ortiz\tLead guide\t4"]:
        K.para(doc, line, after=0)
    pasted = K.para(doc, "Thank you to every renter, volunteer and neighbor who made this season possible.", font="Comic Sans MS", size=13, color="7030A0")
    K.para(doc, "Next steps", bold=True, size=14, before=12, after=6)
    nxt = K.para(doc, "The board will review this report at the November meeting.")
    K.tracked_insert(nxt, " Bring questions.", "Priya Shah")
    K.comments(doc, [{"anchor": nxt, "author": "Priya Shah", "text": "Can you confirm the meeting date?"}])
    return doc


def week5():
    save(policy(), 5, "W5-L1-policy.docx")
    save(checklist_draft(), 5, "W5-L1-draft.docx")
    save(letter_version(), 5, "W5-L2-letter-original.docx")
    save(letter_revised(), 5, "W5-L2-letter-revised.docx")
    save(letter_version([("I would love to talk about how I can help your team grow.", "I'd welcome the chance to talk about how I can help your team.", "Priya Shah")], author="Priya Shah"), 5, "W5-L2-review-priya.docx")
    save(letter_version([("For two seasons I ran the rental desk at Lakeshore Paddle Company. I scheduled staff, answered customers and balanced the books every day.",
                          "For two seasons I ran the rental desk at Lakeshore Paddle Company, where I scheduled twenty staff and balanced the books every day.", "Sam Ortiz")], author="Sam Ortiz"), 5, "W5-L2-review-sam.docx")
    save(ship(), 5, "W5-L3-ship.docx")
    save(season_report(), 5, "W5-SC5-report.docx")
    save(exam(), 5, "W5-CAP-exam.docx")


def build_all():
    week2(); week3(); week4(); week5()


if __name__ == "__main__":
    build_all()
