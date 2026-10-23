# -*- coding: utf-8 -*-
"""
build.py - checks coverage, then writes:
  out/word-course-map.md        (the course map - the main source document)
  out/word-course-map.html      (standalone interactive page with full HTML skeleton)
  out/artifact.html             (same page without the skeleton, for publishing)
"""
import json, re, sys, pathlib
from data import *

OUT = pathlib.Path("out"); OUT.mkdir(exist_ok=True)
problems = []

# ---------- index the MO-110 sub-skills ----------
SKILLS = {}   # code -> dict
for d in MO110:
    for o in d["objectives"]:
        for code, name in o["skills"]:
            SKILLS[code] = {"code": code, "name": name, "obj": o["id"], "objName": o["name"],
                            "domain": d["id"], "domainName": d["name"], "drills": [], "weeks": [], "checks": []}

WEEK_IDS = [w["id"] for w in WEEKS]
CLO_IDS = [c["id"] for c in CLOS]

# ---------- walk the weeks ----------
drilled_so_far = set()
for w in WEEKS:
    for did, task, mos, extra, where in w["drills"]:
        for code in mos:
            if code not in SKILLS:
                problems.append(f"{did}: unknown MO-110 code {code}")
                continue
            SKILLS[code]["drills"].append(did)
            if w["id"] not in SKILLS[code]["weeks"]:
                SKILLS[code]["weeks"].append(w["id"])
            drilled_so_far.add(code)
        if where not in ("web", "desktop", "both-check"):
            problems.append(f"{did}: bad platform tag {where}")
    sc = w["skillcheck"]
    for code in sc["covers"]:
        if code not in SKILLS:
            problems.append(f"{sc['id']}: unknown code {code}")
        elif code not in drilled_so_far:
            problems.append(f"{sc['id']} tests {code} before it is drilled")
        else:
            SKILLS[code]["checks"].append(sc["id"])
    pts = sum(c[3] for c in w["project"]["criteria"])
    if pts != 100:
        problems.append(f"{w['project']['id']} rubric totals {pts}, not 100")
    for oid, text, clos in w["objectives"]:
        for c in clos:
            if c not in CLO_IDS:
                problems.append(f"{oid}: unknown CLO {c}")
            if c not in w["clos"]:
                problems.append(f"{oid} maps {c} but week {w['id']} matrix does not list it")

for code, s in SKILLS.items():
    if not s["drills"]:
        problems.append(f"MO-110 {code} {s['name']} is never drilled")

# ---------- CLO evidence ----------
CLO_ARTIFACT = {
    "CLO1": "Project 1 letter and memo (.docx and .pdf); Project 2 paper layout; Project 4 resumes; clean-and-ship checklist in Project 5.",
    "CLO2": "Skill checks SC1 to SC5 (feature use under time); Project 2 styles, citations and template; capstone practice exam.",
    "CLO3": "Project 3 graphics, chart and alt text; SC3; capstone domain 5.",
    "CLO4": "Project 3 schedule table; Project 4 ATS resume structure; SC3; capstone domain 3.",
    "CLO5": "Project 3 newsletter layout (columns, sections, theme, page elements); Project 4 designed resume.",
    "CLO6": "Project 4 web page version of the resume (.htm) and the format-comparison reflection.",
    "CLO7": "Project 4 merged cover letters, envelopes or labels, and directory; SC4 merge check.",
    "CLO8": "Project 5 tracked review, comment responses, compare and combine, and clean-and-ship; SC5.",
}
EVIDENCE = []
for c in CLOS:
    cid = c["id"]
    row = {"id": cid, "text": c["text"], "bloom": c["bloom"], "I": [], "R": [], "A": [],
           "objectives": [], "criteria": [], "checks": [], "artifact": CLO_ARTIFACT[cid]}
    for w in WEEKS:
        lv = w["clos"].get(cid, "")
        for letter in "IRA":
            if letter in lv:
                row[letter].append(w["id"])
        for oid, text, clos in w["objectives"]:
            if cid in clos:
                row["objectives"].append(oid)
        for crit in w["project"]["criteria"]:
            if cid in crit[2]:
                row["criteria"].append(f"{crit[0]}")
        if "A" in lv:
            row["checks"].append(w["skillcheck"]["id"])
    row["checks"].append("CAP")
    if not row["A"]:
        problems.append(f"{cid} is never assessed")
    if not row["criteria"]:
        problems.append(f"{cid} has no rubric criterion")
    EVIDENCE.append(row)

for t in TOPICS:
    if not t["weeks"]:
        problems.append(f"{t['id']} not covered")

if sum(g[1] for g in GRADING) != 100:
    problems.append("Grading does not total 100")

# ---------- no em dashes anywhere ----------
blob = json.dumps([COURSE, CLOS, TOPICS, TOPIC_HOW, MO110, MO111, WEEKS, NORTHSTAR, NACE, DIGCOMP,
                   WCAG, PLATFORM, CONVENTIONS, DECISIONS], ensure_ascii=False)
if "—" in blob:
    problems.append("Em dash found in data")

# ---------- summary numbers ----------
total_sub = len(SKILLS)
drilled = sum(1 for s in SKILLS.values() if s["drills"])
checked = sum(1 for s in SKILLS.values() if s["checks"])
# How many exam sub-skills the capstone practice exam lists (its "mos:" line).
_cap = next(pathlib.Path(__file__).resolve().parent.parent.glob("content/week-05/*capstone*.md"), None)
_capm = re.search(r"^mos: (.*)$", _cap.read_text(encoding="utf8"), re.M) if _cap else None
CAP_SKILLS = len([x for x in _capm.group(1).split(",") if x.strip()]) if _capm else 0
STATS = {"capSkills": CAP_SKILLS, "clos": len(CLOS), "closAssessed": sum(1 for e in EVIDENCE if e["A"]),
         "topics": len(TOPICS), "topicsCovered": sum(1 for t in TOPICS if t["weeks"]),
         "sub": total_sub, "subDrilled": drilled, "subChecked": checked,
         "objectives": sum(len(d["objectives"]) for d in MO110),
         "drills": sum(len(w["drills"]) for w in WEEKS),
         "mo111Full": sum(1 for m in MO111 if m["status"] == "full"),
         "mo111Partial": sum(1 for m in MO111 if m["status"] == "partial"),
         "mo111None": sum(1 for m in MO111 if m["status"] == "none"),
         "webDrills": sum(1 for w in WEEKS for d in w["drills"] if d[4] == "web"),
         "deskDrills": sum(1 for w in WEEKS for d in w["drills"] if d[4] == "desktop"),
         "bothDrills": sum(1 for w in WEEKS for d in w["drills"] if d[4] == "both-check")}

print("STATS", json.dumps(STATS))
if problems:
    print("PROBLEMS:"); [print(" -", p) for p in problems]
    sys.exit(1)
print("Coverage checks passed.")

# ======================================================================
# Markdown
# ======================================================================
WHERE = {"web": "Web + desktop", "desktop": "Desktop only", "both-check": "Any (self-check)"}
LEVEL = {"I": "Introduced", "R": "Reinforced", "A": "Assessed"}
def wk(wid): return f"Week {wid[1:]}"
def cell(lv): return ", ".join(lv) if lv else ""
md = []
A = md.append
A(f"# {COURSE['title']}\n")
A(f"**{COURSE['college']}** · {COURSE['version']} · {COURSE['date']} · {COURSE['author']}\n")
A(f"- **Format:** {COURSE['format']}")
A(f"- **Certification target:** {COURSE['cert']}")
A(f"- **Coverage:** {STATS['closAssessed']}/{STATS['clos']} course learning outcomes assessed · {STATS['topicsCovered']}/{STATS['topics']} outline topics · {STATS['subDrilled']}/{STATS['sub']} MO-110 sub-skills practiced ({STATS['subChecked']} also in weekly skill checks; {STATS['capSkills']} sampled across every domain in the capstone) · {STATS['mo111Full']} MO-111 objectives fully and {STATS['mo111Partial']} partly covered")
A(f"- **Companion:** `word-course-map.html` shows this map interactively. This Markdown file is the master copy. Both files are generated from the same data (`_build/data.py`), so they always agree; after changing the map, regenerate both with `python3 build.py`.\n")

A("## How to read this map\n")
A("- **CLO** = course learning outcome from the official course description (verbatim). **T** = topic outline item (verbatim).")
A("- **I / R / A** = Introduced, Reinforced, Assessed (standard curriculum mapping).")
A("- **Drills** are short, outcome-checked practice tasks for the Office Trainer (it checks the result, not the clicks). Each is tagged *Web + desktop* or *Desktop only*.")
A("- **MO-110 codes** (for example 2.2.3) follow Microsoft's skills outline. **E** codes (for example E4.3) are MO-111 Expert objectives. **N** codes are Northstar Digital Literacy Word standards.\n")

A("## Course at a glance\n")
A("| Week | Theme | Project | CLOs (I/R/A) | Outline topics | MO-110 domains drilled |")
A("|---|---|---|---|---|---|")
for w in WEEKS:
    doms = sorted({SKILLS[c]["domain"] for d in w["drills"] for c in d[2]})
    clos = "; ".join(f"{k} {v}" for k, v in w["clos"].items())
    A(f"| {w['n']} | **{w['title']}** - {w['tagline']} | {w['project']['id']} {w['project']['title']} | {clos} | {', '.join(w['topics'])} | {', '.join(doms) if doms else 'Mostly MO-111'} |")
A("")

A("## Decisions and flags\n")
for title, kind, text in DECISIONS:
    tag = "FLAG" if kind == "flag" else "NOTE"
    A(f"- **[{tag}] {title}.** {text}")
A("")

A("## Grading (suggested)\n")
A("| Component | Weight |"); A("|---|---|")
for g, p in GRADING: A(f"| {g} | {p}% |")
A("| **Total** | **100%** |\n")

A("## Weekly plan\n")
for w in WEEKS:
    A(f"### Week {w['n']}: {w['title']}\n")
    A(f"*{w['tagline']}*  ")
    A(f"**Outline topics:** {', '.join(t['id'] + ' ' + t['text'] for t in TOPICS if t['id'] in w['topics'])}  ")
    clo_txt = "; ".join(k + " (" + ", ".join(LEVEL[x] for x in v) + ")" for k, v in w["clos"].items())
    A(f"**CLOs:** {clo_txt}  ")
    A(f"**Workload:** {w['hours']}\n")
    A("**Why it matters**\n")
    A(f"- *College:* {w['why']['college']}")
    A(f"- *Work:* {w['why']['work']}")
    A(f"- *Life:* {w['why']['life']}\n")
    A("**Module objectives** - by the end of the week, learners can:\n")
    for oid, text, clos in w["objectives"]:
        A(f"- **{oid}** {text} *({', '.join(clos)})*")
    A("")
    A("**Practice drills (Office Trainer)**\n")
    A("| ID | Task (checked outcome) | MO-110 | Other | Platform |"); A("|---|---|---|---|---|")
    for did, task, mos, extra, where in w["drills"]:
        A(f"| {did} | {task} | {', '.join(mos) or '-'} | {', '.join(extra) or '-'} | {WHERE[where]} |")
    A("")
    sc = w["skillcheck"]
    A(f"**Skill check {sc['id']}:** {sc['desc']} Covers MO-110 {', '.join(sc['covers'])}.\n")
    p = w["project"]
    A(f"**Project {p['id']}: {p['title']}**\n")
    A(p["desc"] + "\n")
    A("| Criterion | Description | CLOs | Points |"); A("|---|---|---|---|")
    for cid, text, clos, pts in p["criteria"]:
        A(f"| {cid} | {text} | {', '.join(clos)} | {pts} |")
    A("| | **Total** | | **100** |\n")
    if "capstone" in w:
        c = w["capstone"]
        A(f"**{c['title']}**\n")
        A(c["desc"] + f" Evidence for {', '.join(c['clos'])}.\n")
    A(f"**AI moment:** {w['ai']}\n")
    if "privacy" in w: A(f"**Privacy:** {w['privacy']}\n")
    A(f"**Platform:** {w['platform']}\n")
    A(f"**Standards:** NACE - {', '.join(w['nace'])}. DigComp 3.0 - {', '.join(w['digcomp'])}. WCAG 2.1 - {', '.join(w['wcag'])}.\n")
    A(f"**Learner takeaway:** {w['life']}\n")

A("## Alignment matrix\n")
A("### Course learning outcomes by week\n")
A("| CLO | " + " | ".join(f"Wk {w['n']}" for w in WEEKS) + " |")
A("|---|" + "---|" * len(WEEKS))
for c in CLOS:
    A(f"| **{c['id']}** {c['text']} | " + " | ".join(cell(w['clos'].get(c['id'], '')) for w in WEEKS) + " |")
A("")
A("### Topic outline by week\n")
A("| Topic | " + " | ".join(f"Wk {w['n']}" for w in WEEKS) + " |")
A("|---|" + "---|" * len(WEEKS))
for t in TOPICS:
    A(f"| **{t['id']}** {t['text']} | " + " | ".join("X" if w['id'] in t['weeks'] else "" for w in WEEKS) + " |")
A("")
A("### Module objectives to course outcomes\n")
A("| Objective | " + " | ".join(CLO_IDS) + " |")
A("|---|" + "---|" * len(CLO_IDS))
for w in WEEKS:
    for oid, text, clos in w["objectives"]:
        A(f"| {oid} | " + " | ".join("X" if c in clos else "" for c in CLO_IDS) + " |")
A("")

A("## Evidence for each course learning outcome\n")
for e in EVIDENCE:
    A(f"### {e['id']}: {e['text']}\n")
    A(f"- **Bloom level:** {e['bloom']}")
    A(f"- **Introduced:** {', '.join(wk(x) for x in e['I']) or '-'} · **Reinforced:** {', '.join(wk(x) for x in e['R']) or '-'} · **Assessed:** {', '.join(wk(x) for x in e['A'])}")
    A(f"- **Taught through objectives:** {', '.join(e['objectives'])}")
    A(f"- **Rubric criteria:** {', '.join(e['criteria'])}")
    A(f"- **Skill checks and capstone:** {', '.join(e['checks'])}")
    A(f"- **Evidence artifact:** {e['artifact']}\n")
A("### Topic outline evidence\n")
for t in TOPICS:
    A(f"- **{t['id']} {t['text']}:** {TOPIC_HOW[t['id']]}")
A("")

A("## MOS crosswalk: MO-110 Word Associate (Microsoft 365 Apps)\n")
A(f"All {STATS['sub']} sub-skills are practiced in at least one drill; the capstone practice exam samples {STATS['capSkills']} of them from every domain. Wording follows Microsoft's skills outline; check the current study guide before printing.\n")
for d in MO110:
    A(f"### Domain {d['id']}: {d['name']} ({d['weight']})\n")
    A("| Code | Skill | Week | Drills | Skill check |"); A("|---|---|---|---|---|")
    for o in d["objectives"]:
        A(f"| **{o['id']}** | **{o['name']}** | | | |")
        for code, name in o["skills"]:
            s = SKILLS[code]
            A(f"| {code} | {name} | {', '.join(x[1:] for x in s['weeks'])} | {', '.join(s['drills'])} | {', '.join(s['checks']) or 'Capstone'} |")
    A("")

A("## MOS crosswalk: MO-111 Word Expert (stretch content)\n")
A("This course does not prepare learners for MO-111. It covers the Expert objectives that the official course outcomes require.\n")
A("| Code | Objective | Status | Weeks | How |"); A("|---|---|---|---|---|")
ST = {"full": "Covered", "partial": "Partly covered", "none": "Not covered"}
for m in MO111:
    A(f"| E{m['id']} | {m['name']} | {ST[m['status']]} | {', '.join(x[1:] for x in m['weeks']) or '-'} | {m['how']} |")
A("")

A("## Other standards and conventions\n")
A("### Northstar Digital Literacy: Microsoft Word\n")
A("| Code | Standard | Week | Drill |"); A("|---|---|---|---|")
for n in NORTHSTAR: A(f"| {n[0]} | {n[1]} | {n[2][1:]} | {n[3]} |")
A("")
A("### NACE Career Readiness Competencies\n")
A("| Competency | Weeks | Where it shows up |"); A("|---|---|---|")
for n in NACE: A(f"| {n[0]} | {', '.join(x[1:] for x in n[1])} | {n[2]} |")
A("")
A("### DigComp 3.0 (European Digital Competence Framework, December 2025)\n")
A("| Competence area | Weeks | Where it shows up |"); A("|---|---|---|")
for n in DIGCOMP: A(f"| {n[0]} | {', '.join(x[1:] for x in n[1])} | {n[2]} |")
A("")
A("### Accessibility: WCAG 2.1 AA success criteria practiced (ADA Title II)\n")
A("| Criterion | Name | Level | In Word | Weeks |"); A("|---|---|---|---|---|")
for n in WCAG: A(f"| {n[0]} | {n[1]} | {n[2]} | {n[3]} | {', '.join(x[1:] for x in n[4])} |")
A("")
A("### Word for the web vs desktop Word\n")
A("From Microsoft's feature comparison table. Microsoft updates Word for the web often, so recheck each term.\n")
A("| Feature | Word for the web | Desktop Word | Used in |"); A("|---|---|---|---|")
for p in PLATFORM: A(f"| {p[0]} | {p[1]} | {p[2]} | {p[3]} |")
A(f"\nDrill platform split: {STATS['webDrills']} work on the web, {STATS['deskDrills']} need desktop Word, {STATS['bothDrills']} work anywhere with a self-check.\n")
A("### Conventions and assumptions\n")
A("| Item | Type | Note | Weeks |"); A("|---|---|---|---|")
for c in CONVENTIONS: A(f"| {c[0]} | {c[1]} | {c[2]} | {', '.join(c[3])} |")
A("")
A("## Sources\n")
for t, u in SOURCES: A(f"- [{t}]({u})")
A("")
(OUT / "word-course-map.md").write_text("\n".join(md), encoding="utf-8")

# ======================================================================
# HTML - inject data into the template
# ======================================================================
DATA = {"course": COURSE, "clos": CLOS, "topics": TOPICS, "topicHow": TOPIC_HOW, "mo110": MO110,
        "skills": SKILLS, "mo111": MO111, "weeks": WEEKS, "grading": GRADING, "northstar": NORTHSTAR,
        "nace": NACE, "digcomp": DIGCOMP, "wcag": WCAG, "platform": PLATFORM, "conventions": CONVENTIONS,
        "decisions": DECISIONS, "sources": SOURCES, "evidence": EVIDENCE, "stats": STATS}
payload = json.dumps(DATA, ensure_ascii=False).replace("</", "<\\/")
tpl = pathlib.Path("template.html").read_text(encoding="utf-8")
page = tpl.replace("__DATA_JSON__", payload)
(OUT / "artifact.html").write_text(page, encoding="utf-8")
head, body = page.split("<!--BODY-->", 1)   # template marks where <head> content ends
standalone = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
              '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
              + head + '</head>\n<body>\n' + body + '\n</body>\n</html>\n')
(OUT / "word-course-map.html").write_text(standalone, encoding="utf-8")
print("Wrote", [p.name for p in OUT.iterdir()])
