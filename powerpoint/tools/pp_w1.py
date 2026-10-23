# -*- coding: utf-8 -*-
"""
Week 1 practice presentations: The Clear Presentation.

Each build_* function makes one presentation. `done` is a set of step names:
build(path) makes the starter a learner opens; build(path, done={"views", ...})
makes the same presentation with those steps finished, for testing the checks.
Everything is fictional (Lakeshore Paddle Company, its staff and its club).
"""
import os
from pptkit import *

ORG = "Lakeshore Paddle Company"


# ---------------------------------------------------------------------------
# Lesson 1: Meet PowerPoint  (D1.01 save, D1.02 views, D1.03 QAT)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    slides = [
        dict(layout="Title Slide", title="Welcome to Lakeshore Paddle Company", subtitle="Summer staff orientation"),
        dict(title="Your first day at the beach" if "outline" in done else "Your first day",
             body=["Meet at the boathouse at 8:30 a.m.", "Pick up your staff shirt and name tag", "Shadow a guide for the morning"]),
        dict(title="Where things are", body=["Boathouse: boats, paddles and life jackets", "Front desk: rentals and returns", "First aid kit: behind the desk"],
             notes="Show the map of the beach." if "notes" in done else None),
        dict(title="Questions?", body=["Ask your shift lead", "Or email staff@lakeshore.example"]),
    ]
    after = (lambda p: last_view(p, "sldSorterView")) if "sorter" in done else None
    return build_deck(path, slides, title="", after_save=after)


# ---------------------------------------------------------------------------
# Lesson 2: Slides, Layouts and Sections  (D1.04 layouts, D1.05 reuse/outline, D1.06 sections)
# ---------------------------------------------------------------------------
REUSE = [("Our boats", ["Single and tandem kayaks", "Canoes for up to three", "Stand-up paddleboards"]),
         ("Rental prices", ["Kayak: $20 an hour", "Canoe: $25 an hour", "Paddleboard: $18 an hour"])]
OUTLINE = [("Weather rules", ["Check the forecast at 8 a.m.", "Thunder heard: everyone off the water", "Wind over 15 mph: no paddleboards"]),
           ("Lost and found", ["Tag it with the date", "Keep it in the bin behind the desk", "Donate after 30 days"])]


def build_l2_extras(folder):
    """The slides to reuse (P1-L2-reuse.pptx) and the Word outline (P1-L2-outline.docx)."""
    build_deck(os.path.join(folder, "P1-L2-reuse.pptx"), [dict(title=t, body=b) for t, b in REUSE], title="Boats and prices")
    make_outline_docx(os.path.join(folder, "P1-L2-outline.docx"), OUTLINE)


def build_l2(path, done=frozenset()):
    slides = [
        dict(layout="Title Slide", title="Welcome to the team", subtitle=ORG + " staff orientation"),
        dict(title="Agenda", body=["Safety first", "Gear checkout", "Closing time"]),
        dict(title="Safety first", body=["Everyone wears a life jacket", "Guides carry a whistle and a radio"]),
        dict(title="Old schedule (2026)", body=["Open 9 a.m. to 5 p.m.", "Closed Mondays"], hidden="hide" in done),
        dict(title="Gear checkout", body=["Check each boat for damage", "Match paddles to the paddler's height", "Log every rental"]),
        dict(title="Questions", body=["Ask your shift lead"], layout="Section Header" if "layout" in done else "Title and Content"),
        dict(title="Closing time", body=["Rinse the boats", "Lock the boathouse", "Count the life jackets"]),
    ]
    if "layout" in done:
        by_title(slides, "Questions")["body"] = None
    if "newslides" in done:
        i = [s["title"] for s in slides].index("Safety first")
        slides.insert(i + 1, dict(title="Life jacket check", body=["Buckles closed", "Snug at the shoulders"]))
        slides.insert(i + 2, dict(layout="Two Content", title="Kayaks and canoes", body=["Kayaks: one or two people"]))
    if "duplicate" in done:
        i = [s["title"] for s in slides].index("Gear checkout")
        slides.insert(i + 1, dict(slides[i]))
    if "reuse" in done:
        for t, b in REUSE:
            slides.insert(len(slides) - 1, dict(title=t, body=b))
    if "outline" in done:
        for t, b in OUTLINE:
            slides.insert(len(slides) - 1, dict(title=t, body=b))
    if "reorder" in done:
        move(slides, "Questions", to=-1)   # Questions becomes the last slide
    secs = [("Section 1", "Welcome to the team")]
    if "sections" in done:
        secs = [("Getting started", "Welcome to the team"), ("On the water", "Safety first"), ("Gear", "Gear checkout"), ("End of day", "Questions" if "reorder" not in done else "Closing time")]
        if "movesection" in done:
            # Move the Gear section above On the water: its slides now come before Safety first.
            titles = [s["title"] for s in slides]
            g0 = titles.index("Gear checkout")
            g1 = g0 + 1
            while g1 < len(slides) and slides[g1]["title"] in ("Gear checkout",):
                g1 += 1
            gear = slides[g0:g1]
            del slides[g0:g1]
            w = [s["title"] for s in slides].index("Safety first")
            slides[w:w] = gear
            secs = [("Getting started", "Welcome to the team"), ("Gear", "Gear checkout"), ("On the water", "Safety first"), ("End of day", secs[3][1])]
    elif "rename" in done:
        secs = [("Getting started", "Welcome to the team")]
    return build_deck(path, slides, sections_by_title=secs, title="")


# ---------------------------------------------------------------------------
# Lesson 3: Text and Lists  (D1.07 format text, WordArt, columns; D1.08 lists)
# ---------------------------------------------------------------------------
GEAR = ["Life jacket", "Paddle", "Whistle", "Dry bag", "Water bottle", "Sunscreen", "Hat", "Sunglasses with a strap", "Snack", "Phone in a pouch"]


def build_l3(path, done=frozenset()):
    def title_slide(prs, s):
        if "titlefmt" in done:
            set_runs(s.shapes.title, size=54, bold=True)
        tb = text_box(s, 4.2, 5.9, 5, 0.8, ["Paddle smart"], size=32, name="Tagline")
        if "wordart" in done:
            set_runs(tb, wordart=True)

    def what_to_bring(prs, s):
        lines = ["Sun protection", "Hat", "Sunscreen", "Water", "Dry clothes"]
        tb = text_box(s, 0.8, 1.6, 8, 4, lines, size=24, name="Bring list")
        if "bullets" in done:
            bullets(tb, "bullet", [0, 1, 1, 0, 0])

    def launch(prs, s):
        tb = text_box(s, 0.8, 1.6, 8, 4, ["Put on your life jacket", "Carry the boat to the water", "Sit down in the middle", "Push off gently"], size=24, name="Steps")
        if "numbers" in done:
            bullets(tb, "number")

    def gear(prs, s):
        if "columns" in done:
            columns([p for p in s.placeholders if p.placeholder_format.idx == 1][0], 2)

    slides = [
        dict(layout="Title Slide", title="Gear and safety", subtitle=ORG, fn=title_slide),
        dict(layout="Title Only", title="What to bring", fn=what_to_bring),
        dict(layout="Title Only", title="Launch steps", fn=launch),
        dict(title="Gear list", body=GEAR, fn=gear),
    ]
    return build_deck(path, slides, title="")


# ---------------------------------------------------------------------------
# Lesson 4: Size, Footers, Notes and Print  (D1.09 size, D1.10 footers, D1.11 notes, properties, print)
# ---------------------------------------------------------------------------
def build_l4(path, done=frozenset()):
    def footers(prs, s):
        if "footer" in done and s.slide_layout.name != "Title Slide":
            footer_placeholders(prs, s, ORG, number=True)

    slides = [
        dict(layout="Title Slide", title=ORG, subtitle="Summer staff orientation"),
        dict(title="Your first day", body=["Meet at the boathouse at 8:30 a.m.", "Shadow a guide for the morning"], fn=footers),
        dict(title="Safety first", body=["Life jackets on the water", "Guides carry a whistle and a radio"], fn=footers,
             notes="Everyone wears a life jacket on the water, staff too." if "notes" in done else None),
        dict(title="Gear checkout", body=["Check each boat", "Log every rental"], fn=footers),
        dict(title="Questions", body=["Ask your shift lead"], fn=footers),
    ]
    after = (lambda p: size_43_to_169(p)) if "size" in done else None
    return build_deck(path, slides, widescreen=False, title="Summer staff orientation" if "title" in done else "", after_save=after)


# ---------------------------------------------------------------------------
# Skill Check 1: club orientation
# ---------------------------------------------------------------------------
def build_sc1(path, done=frozenset()):
    def subtitle(prs, s):
        if "italic" in done:
            set_runs([p for p in s.placeholders if p.placeholder_format.idx == 1][0], italic=True)

    def bring(prs, s):
        if "columns" in done:
            columns([p for p in s.placeholders if p.placeholder_format.idx == 1][0], 2)

    def signup(prs, s):
        tb = text_box(s, 0.8, 1.6, 9, 4, ["Pick a trip on the calendar", "Write your name on the sheet", "Pay the trip fee at the desk", "Bring a signed waiver"], size=24, name="Sign-up steps")
        if "numbers" in done:
            bullets(tb, "number")

    def footers(prs, s):
        if "numbers_footer" in done and s.slide_layout.name != "Title Slide":
            footer_placeholders(prs, s, None, number=True)

    slides = [
        dict(layout="Title Slide", title="Lakeshore Paddle Club", subtitle="New member orientation", fn=subtitle),
        dict(title="Meeting times", body=["Tuesdays at 6 p.m.", "Saturdays at 9 a.m."], fn=footers),
        dict(title="2026 results", body=["14 trips", "62 members"], hidden="hide" in done, fn=footers),
        dict(title="Questions", body=["Ask any club officer"], fn=footers),
        dict(title="Trip calendar", body=["June: Seneca sunrise paddle", "July: Keuka overnight", "August: Canandaigua race"], fn=footers),
        dict(layout="Title Only", title="Trip sign-up", fn=lambda p, s: (signup(p, s), footers(p, s))),
        dict(title="What to bring", body=GEAR, fn=lambda p, s: (bring(p, s), footers(p, s))),
    ]
    if "newslide" in done:
        slides.insert(3, dict(layout="Two Content", title="Club rules", body=["Paddle with a buddy"], fn=footers))
    if "duplicate" in done:
        i = [s["title"] for s in slides].index("Meeting times")
        slides.insert(i + 1, dict(slides[i]))
    if "reorder" in done:
        move(slides, "Trip calendar", after_title="Trip sign-up")
    first = "Welcome" if "rename" in done else "Untitled Section"
    trips = [("Trips", "Trip calendar" if "reorder" not in done else "Trip sign-up")]
    secs = [(first, "Lakeshore Paddle Club"), ("Wrap-up", "Questions")] + trips
    if "movesection" in done:
        # Wrap-up (Questions) moves to the end, after the Trips slides.
        move(slides, "Questions", to=-1)
        secs = [(first, "Lakeshore Paddle Club")] + trips + [("Wrap-up", "Questions")]
    if "addsection" in done:
        # A "Packing" section that starts with "What to bring".
        i = [n for n, _ in secs].index("Trips")
        secs.insert(i + 1, ("Packing", "What to bring"))
    after = (lambda p: size_43_to_169(p)) if "size" in done else None
    return build_deck(path, slides, sections_by_title=secs, widescreen=False, title="Paddle Club orientation" if "title" in done else "", after_save=after)


def build_all(starters):
    out = os.path.join(starters, "week-01")
    os.makedirs(out, exist_ok=True)
    build_l1(os.path.join(out, "P1-L1-welcome.pptx"))
    build_l2(os.path.join(out, "P1-L2-orientation.pptx"))
    build_l2_extras(out)
    build_l3(os.path.join(out, "P1-L3-text.pptx"))
    build_l4(os.path.join(out, "P1-L4-print.pptx"))
    build_sc1(os.path.join(out, "P1-SC1-club.pptx"))
