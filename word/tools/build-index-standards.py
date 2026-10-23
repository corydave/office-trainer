# build-index-standards.py - writes the "What each course covers" tabs on index.html.
#
# The home page lists the standards each course meets: the course learning
# outcomes, the Microsoft Office Specialist exam skills, Northstar Digital
# Literacy, NACE career readiness, DigComp and the WCAG accessibility
# criteria. Rather than typing them into the page (where they'd drift out of
# date), this script reads them from the three course maps
# (word/word-course-map.md, excel/excel-course-map.md,
# powerpoint/powerpoint-course-map.md) and each course's outline, and writes
# the HTML between these two lines in index.html:
#     <!-- STANDARDS:START -->   and   <!-- STANDARDS:END -->
# It then updates index.html's security fingerprints (build-index-csp.py).
#
# Run from anywhere:  python3 word/tools/build-index-standards.py
import html, pathlib, re, subprocess, sys

here = pathlib.Path(__file__).resolve().parent
site = here.parent.parent
COURSES = [("word", "Word", "MO-110"), ("excel", "Excel", "MO-210"), ("powerpoint", "PowerPoint", "MO-310")]
e = html.escape


def section(md, heading_start, level="###"):
    """The text under the first heading that starts with heading_start, up to the next heading of the same or higher level."""
    m = re.search(r"^%s %s.*$" % (re.escape(level), re.escape(heading_start)), md, re.M)
    if not m:
        sys.exit("Couldn't find the heading '%s %s'" % (level, heading_start))
    rest = md[m.end():]
    stop = re.search(r"^#{2,%d} " % len(level), rest, re.M)
    return m.group(0), rest[: stop.start()] if stop else rest


def rows(text):
    """The body rows of the first Markdown table in text, as lists of cells (the header and the --- line skipped)."""
    out = []
    for line in text.splitlines():
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-+:?", c) for c in cells):
            continue
        out.append(cells)
    return out[1:]  # drop the header row


def plain(cell):
    """A table cell without Markdown bold marks."""
    return cell.replace("**", "").strip()


def outcomes(folder):
    """The course learning outcomes from content/outline.md (::: outcomes, lines "- CLO1 | text")."""
    md = (site / folder / "content" / "outline.md").read_text(encoding="utf-8")
    block = re.search(r"^::: outcomes\n(.*?)^:::", md, re.S | re.M).group(1)
    return [tuple(p.strip() for p in line[2:].split("|", 1)) for line in block.splitlines() if line.startswith("- ")]


def course_html(folder, name, exam):
    md = (site / folder / f"{folder}-course-map.md").read_text(encoding="utf-8")
    clos = outcomes(folder)

    # Microsoft Office Specialist: the Associate exam's domains, objectives and sub-skills
    mos_title, mos = section(md, "MOS crosswalk: " + exam, "##")
    domains = []
    for dm in re.finditer(r"^### (Domain \d+: .*?)(?: \(.*?\))?$", mos, re.M):
        body = mos[dm.end():]
        nxt = re.search(r"^### ", body, re.M)
        body = body[: nxt.start()] if nxt else body
        groups = []
        for r in rows(body):
            if r[0].startswith("**"):
                groups.append([plain(r[0]) + " " + plain(r[1]), []])
            elif re.fullmatch(r"\d+\.\d+\.\d+", r[0]) and groups:
                groups[-1][1].append(r[0] + " " + r[1])
        domains.append((dm.group(1), groups))
    skills = sum(len(g[1]) for _, gs in domains for g in gs)
    exam_name = re.sub(r"^## MOS crosswalk: ", "", mos_title)

    north = [(r[0], r[1]) for r in rows(section(md, "Northstar")[1])]
    nace = [r[0] for r in rows(section(md, "NACE")[1])]
    dig = [r[0] for r in rows(section(md, "DigComp")[1])]
    wcag = [(r[0], r[1], r[2]) for r in rows(section(md, "Accessibility: WCAG")[1])]

    summary = [
        ("Course learning outcomes", f"{len(clos)} of {len(clos)} assessed", "FLCC's official outcomes for the course, each assessed in a project, skill check or the capstone"),
        (f"Microsoft Office Specialist: {exam_name}", f"{skills} of {skills} exam skills practiced", "Every skill on Microsoft's outline has a drill; weekly skill checks and a timed practice exam cover every domain"),
        (f"Northstar Digital Literacy: Microsoft {name}", f"{len(north)} of {len(north)} standards", "Each standard is practiced in a named drill"),
        ("NACE Career Readiness Competencies", f"{len(nace)} competencies", "Communication, technology, teamwork and more, built into the projects"),
        ("DigComp 3.0 (European Digital Competence Framework)", f"{len(dig)} competence areas", "Content creation, collaboration, safety and problem solving"),
        ("WCAG 2.1 AA (ADA Title II)", f"{len(wcag)} success criteria practiced", "Learners make accessible files as they go: alt text, headings, contrast, reading order"),
    ]
    out = [f'<section class="std-panel" id="std-{folder}" role="tabpanel" aria-labelledby="std-tab-{folder}">',
           f'<h3 class="std-head">Microsoft {name}</h3>',
           f'<div class="table-wrap"><table class="std-table"><caption>What the {name} course covers</caption>',
           '<thead><tr><th scope="col">Standard</th><th scope="col">Coverage</th><th scope="col">How</th></tr></thead><tbody>']
    for std, cov, how in summary:
        out.append(f'<tr><th scope="row">{e(std)}</th><td class="cov">{e(cov)}</td><td>{e(how)}</td></tr>')
    out.append("</tbody></table></div>")

    def details(title, items_html):
        return f'<details class="std-list"><summary>{e(title)}</summary>{items_html}</details>'

    out.append('<div class="std-details">')
    out.append(details(f"The {len(clos)} course learning outcomes", "<ul>" + "".join(f"<li><b>{e(i)}</b> {e(t)}</li>" for i, t in clos) + "</ul>"))
    mos_html = ""
    for dname, groups in domains:
        mos_html += f"<h4>{e(dname)}</h4><ul>" + "".join(
            f"<li><b>{e(g)}</b><ul>" + "".join(f"<li>{e(s)}</li>" for s in subs) + "</ul></li>" for g, subs in groups) + "</ul>"
    out.append(details(f"All {skills} {exam} exam skills", mos_html))
    out.append(details(f"The {len(north)} Northstar standards", "<ul>" + "".join(f"<li><b>{e(c)}</b> {e(t)}</li>" for c, t in north) + "</ul>"))
    out.append(details("NACE competencies and DigComp areas", "<h4>NACE</h4><ul>" + "".join(f"<li>{e(x)}</li>" for x in nace) +
                       "</ul><h4>DigComp 3.0</h4><ul>" + "".join(f"<li>{e(x)}</li>" for x in dig) + "</ul>"))
    out.append(details(f"The {len(wcag)} WCAG criteria", "<ul>" + "".join(f"<li><b>{e(c)}</b> {e(n)} (Level {e(l)})</li>" for c, n, l in wcag) + "</ul>"))
    out.append("</div>")
    out.append(f'<p class="std-more"><a href="{folder}/{folder}-course-map.html">See the full {name} course map</a>: every week, drill, skill and rubric.</p>')
    out.append("</section>")
    return "\n".join(out), skills


parts, totals = [], []
for folder, name, exam in COURSES:
    block, skills = course_html(folder, name, exam)
    parts.append(block)
    totals.append(skills)
tabs = ('<div class="tabs std-tabs" role="tablist" aria-label="Standards by course" hidden>' +
        "".join(f'<button type="button" role="tab" id="std-tab-{f}" aria-controls="std-{f}" aria-selected="{"true" if i == 0 else "false"}"'
                f' tabindex="{0 if i == 0 else -1}">{n}</button>' for i, (f, n, _) in enumerate(COURSES)) + "</div>")
generated = "<!-- STANDARDS:START (made by word/tools/build-index-standards.py from the course maps; don't edit by hand) -->\n" + tabs + "\n" + "\n".join(parts) + "\n<!-- STANDARDS:END -->"

page = site / "index.html"
text = page.read_text(encoding="utf-8")
new, n = re.subn(r"<!-- STANDARDS:START.*?<!-- STANDARDS:END -->", lambda m: generated, text, flags=re.S)
assert n == 1, "index.html needs one <!-- STANDARDS:START --> ... <!-- STANDARDS:END --> pair"
page.write_text(new, encoding="utf-8")
print("Wrote the standards tabs:", ", ".join(f"{n} {s} exam skills" for (_, n, _), s in zip(COURSES, totals)))
subprocess.run([sys.executable, str(here / "build-index-csp.py")], check=True)
