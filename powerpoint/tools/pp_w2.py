# -*- coding: utf-8 -*-
"""
Week 2 practice presentations: Design That Works.

build(path) makes the starter; build(path, done={...}) makes it with those
steps finished, for testing the checks. Everything is fictional.
"""
import os
from pptkit import *

ORG = "Lakeshore Paddle Company"


def facet(path, fonts=False):
    """Design > Themes > Facet (a built-in theme), optionally with the Georgia font set."""
    theme(path, name="Facet", major="Georgia" if fonts else "Trebuchet MS", minor="Georgia" if fonts else "Trebuchet MS",
          accent1="90C226", colors_name="Facet", fonts_name="Georgia" if fonts else "Facet")


# ---------------------------------------------------------------------------
# Lesson 1: Themes and Backgrounds  (D2.01 theme, D2.02 backgrounds)
# ---------------------------------------------------------------------------
def build_l1(path, done=frozenset()):
    def title_bg(prs, s):
        if "picturebg" in done:
            slide_background(s, "picture", image_rel(s, media("lake1.jpg")))

    def thanks_bg(prs, s):
        if "solidbg" in done:
            slide_background(s, "solid", color=(31, 56, 100))

    slides = [
        dict(layout="Title Slide", title="Season kickoff", subtitle=ORG, fn=title_bg),
        dict(title="This summer", body=["Open June 1 to Labor Day", "Two new guides", "Sunset tours on Fridays"]),
        dict(title="New boats", body=["Four tandem kayaks", "Two touring canoes"]),
        dict(title="Staff picnic", body=["Saturday, June 12, at noon", "Bring a side dish"]),
        dict(layout="Title Only", title="Thank you", fn=thanks_bg),
    ]
    after = None
    if "theme" in done or "fonts" in done:
        after = lambda p: facet(p, fonts="fonts" in done)

    def before(prs):
        if "applyall" in done:
            # Format Background > Apply to All: the picture goes on the slide master.
            _, rid = prs.slide_master.part.get_or_add_image_part(media("lake1.jpg"))
            cSld = prs.slide_master._element.find(qn("p:cSld"))
            old = cSld.find(qn("p:bg"))
            if old is not None:
                cSld.remove(old)
            cSld.insert(0, X('<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>' % rid)[0])
    return build_deck(path, slides, title="", before_save=before, after_save=after)


# ---------------------------------------------------------------------------
# Lesson 2: The Slide Master and Layouts  (D2.03 master, D2.04 layouts)
# ---------------------------------------------------------------------------
def build_l2(path, done=frozenset(), sc=False):
    def master(prs):
        if "masterbg" in done:
            master_background(prs, (232, 244, 245))
        if "masterlogo" in done:
            master_picture(prs, media("logo.png"), 10.9, 6.3, 2, alt="Lakeshore Paddle Company logo")
        if "titlefont" in done:
            master_title_font(prs, "Georgia")

    layout_name = "Team Photo" if sc else "Quote"
    slides = [
        dict(layout="Title Slide", title="Season kickoff" if not sc else "Company kickoff", subtitle=ORG),
        dict(title="This summer", body=["Open June 1 to Labor Day", "Sunset tours on Fridays"]),
        dict(layout="Title Only" if ("uselayout" not in done or sc) else layout_name, title="A word from our founder"),
        dict(layout="Section Header", title="The season ahead"),
        dict(title="Staff picnic", body=["Saturday, June 12, at noon"]),
    ]
    if sc:
        slides.insert(3, dict(layout="Title Only" if "uselayout" not in done else "Team Photo", title="Our crew"))
        slides.append(dict(layout="Title Only", title="Thank you", fn=lambda p, s: slide_background(s, "solid", color=(31, 56, 100)) if "solidbg" in done else None))
        slides[0]["fn"] = lambda p, s: slide_background(s, "picture", image_rel(s, media("lake2.jpg"))) if "picturebg" in done else None

    def after(p):
        if "newlayout" in done or "uselayout" in done:
            new_layout(p, layout_name, picture=True)
        if "editlayout" in done:
            edit_layout(p, "slideLayout2.xml" if sc else "slideLayout3.xml", remove_types=("dt",))
        if "handout" in done:
            handout_master(p, header="Season kickoff" if sc else ORG)

    def before(prs):
        master(prs)
        if "notesfooter" in done:
            notes_master_text(prs, footer=ORG if sc else "Staff use only")
        if "uselayout" in done:
            pass

    if "uselayout" in done:
        # The layout must exist before a slide can use it: build in two passes.
        tmp = path + ".tmp.pptx"
        build_deck(tmp, [dict(s, layout=("Title Only" if s.get("layout") == layout_name else s.get("layout", "Title and Content"))) for s in slides], title="")
        new_layout(tmp, layout_name, picture=True)
        prs = Presentation(tmp)
        target = [l for l in prs.slide_layouts if l.name == layout_name][0]
        for sl in prs.slides:
            if sl.shapes.title is not None and sl.shapes.title.text == ("Our crew" if sc else "A word from our founder"):
                sl.part.rels[[k for k, r in sl.part.rels.items() if r.reltype.endswith("/slideLayout")][0]]._target = target.part
        master(prs)
        if sc and "picturebg" in done:
            slide_background(prs.slides[0], "picture", image_rel(prs.slides[0], media("lake2.jpg")))
        if sc and "solidbg" in done:
            slide_background(prs.slides[len(prs.slides) - 1], "solid", color=(31, 56, 100))
        if "notesfooter" in done:
            notes_master_text(prs, footer=ORG if sc else "Staff use only")
        prs.save(path)
        os.remove(tmp)
        if "editlayout" in done:
            edit_layout(path, "slideLayout2.xml" if sc else "slideLayout3.xml", remove_types=("dt",))
        if "handout" in done:
            handout_master(path, header="Season kickoff" if sc else ORG)
        return path
    return build_deck(path, slides, title="", before_save=before, after_save=after)


# ---------------------------------------------------------------------------
# Lesson 3: Handouts, Notes and Templates  (D2.05 handout and notes masters, D2.06 template)
# ---------------------------------------------------------------------------
def build_l3(path, done=frozenset()):
    def before(prs):
        if "notesfooter" in done:
            notes_master_text(prs, footer="Staff use only")

    slides = [
        dict(layout="Title Slide", title="Season kickoff", subtitle=ORG),
        dict(title="This summer", body=["Open June 1 to Labor Day", "Sunset tours on Fridays"], notes="Thank the returning guides."),
        dict(title="Staff picnic", body=["Saturday, June 12, at noon"], notes="Sign-up sheet is by the time clock."),
    ]
    after = (lambda p: handout_master(p, header=ORG)) if "handout" in done else None
    return build_deck(path, slides, title="", before_save=before, after_save=after)


# ---------------------------------------------------------------------------
# Lesson 4: Design That Reads  (D2.07 cleanup)
# ---------------------------------------------------------------------------
RULES = ["Wear a life jacket at all times", "Stay inside the buoys", "No glass on the beach", "Return boats by 6 p.m.",
         "Paddleboards: stay within 100 yards of shore", "Children under 12 paddle with an adult", "Report damage at the desk",
         "Don't stand up in a canoe", "Keep the dock clear", "Watch for the red flag: everyone in"]


def build_l4(path, done=frozenset()):
    novelty = None if "fonts" in done else "Comic Sans MS"

    def fontify(prs, s):
        for sh in s.shapes:
            if sh.has_text_frame and novelty:
                set_runs(sh, font=novelty)

    def hours(prs, s):
        fontify(prs, s)
        ph = [p for p in s.placeholders if p.placeholder_format.idx == 1][0]
        if "contrast" not in done:
            set_text_color(ph, (200, 200, 200))

    slides = [
        dict(layout="Title Slide", title="Renter guide", subtitle=ORG, fn=fontify),
        dict(title="Hours", body=["Open daily, 9 a.m. to 7 p.m.", "Last rental at 5 p.m."], fn=hours),
    ]
    if "split" in done:
        slides.append(dict(title="Rules for renters (1 of 2)", body=RULES[:5], fn=fontify))
        slides.append(dict(title="Rules for renters (2 of 2)", body=RULES[5:], fn=fontify))
    else:
        slides.append(dict(title="Rules for renters", body=RULES, fn=fontify))
    slides.append(dict(title="Thanks for paddling with us", body=["Tag us in your photos"], fn=fontify))
    return build_deck(path, slides, title="")


# ---------------------------------------------------------------------------
# Skill Check 2: company kickoff (masters, layouts, backgrounds)
# ---------------------------------------------------------------------------
def build_sc2(path, done=frozenset()):
    return build_l2(path, done, sc=True)


def build_all(starters):
    out = os.path.join(starters, "week-02")
    os.makedirs(out, exist_ok=True)
    make_photo(os.path.join(out, "lake.jpg"), seed=1)
    make_logo(os.path.join(out, "logo.png"))
    build_l1(os.path.join(out, "P2-L1-design.pptx"))
    build_l2(os.path.join(out, "P2-L2-master.pptx"))
    build_l3(os.path.join(out, "P2-L3-handouts.pptx"))
    build_l4(os.path.join(out, "P2-L4-cluttered.pptx"))
    build_sc2(os.path.join(out, "P2-SC2-kickoff.pptx"))
