# build-site-pages.py - builds the site's three tool pages, each as ONE file:
#
#   badges.html                 My badges: every badge and certificate, all three courses
#   instructor/sections.html    the section builder: pick lessons, add videos, get a section key
#   instructor/grader.html      the grader: check a folder of student files, make grade and feedback files
#
# Each page starts as a template in tools/src/ (badges.src.html, sections.src.html,
# grader.src.html). This script copies in:
#   /*__SITE_CSS__*/      tools/src/site.css   (shared look)
#   /*__SITE_JS__*/       tools/src/site.js    (shared helpers)
#   /*__TRAINER_MD__*/    tools/trainer-md.js  (reads lessons, section keys, badges)
#   /*__WORD_CHECKER__*/ /*__EXCEL_CHECKER__*/ /*__PPT_CHECKER__*/   the three checkers (grader only)
# and then writes each page's Content Security Policy (__CSP__), which lets the
# browser run ONLY the script and style blocks in that page, identified by their
# sha256 fingerprints. Edit a template, then run this again, or the browser will
# refuse to run the old fingerprints' code.
#
# Run from anywhere:  python3 word/tools/build-site-pages.py
import base64, hashlib, pathlib, re

here = pathlib.Path(__file__).resolve().parent          # word/tools
site = here.parent.parent                                # the top of the site (office-trainer/)
src = here / "src"

JSZIP = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"


def read(path):
    return pathlib.Path(path).read_text(encoding="utf-8")


def video_hosts():
    """Every approved video host in the three courses (for the section builder's preview frames)."""
    hosts = {"www.youtube-nocookie.com"}
    for course in ("word", "excel", "powerpoint"):
        m = re.search(r"^videoHosts:\s*\[(.*?)\]", read(site / course / "course.config.md"), re.M)
        if m:
            hosts.update(h.strip() for h in m.group(1).split(",") if h.strip())
    # Only real host names (letters, digits, dots, hyphens) go into the policy.
    return sorted(h for h in hosts if re.fullmatch(r"[a-z0-9.-]+", h))


def fingerprints(html, tag):
    """sha256 of the exact text inside each <tag>...</tag> (plain tags only: <script src=...> isn't counted)."""
    blocks = re.findall(r"<%s>(.*?)</%s>" % (tag, tag), html, re.S)
    assert blocks, "no <%s> block found" % tag
    return ["'sha256-%s'" % base64.b64encode(hashlib.sha256(b.encode("utf-8")).digest()).decode() for b in blocks]


# (template, output, way back to the top of the site, extra script sources, frame sources)
PAGES = [
    ("badges.src.html", "badges.html", "", [], []),
    ("sections.src.html", "instructor/sections.html", "../", [], ["https://" + h for h in video_hosts()]),
    ("grader.src.html", "instructor/grader.html", "../", [JSZIP], []),
]

pieces = {
    "/*__SITE_CSS__*/": src / "site.css",
    "/*__SITE_JS__*/": src / "site.js",
    "/*__TRAINER_MD__*/": here / "trainer-md.js",
    "/*__WORD_CHECKER__*/": here / "word-checker.js",
    "/*__EXCEL_CHECKER__*/": site / "excel" / "tools" / "excel-checker.js",
    "/*__PPT_CHECKER__*/": site / "powerpoint" / "tools" / "ppt-checker.js",
}

for template, out, root, scripts, frames in PAGES:
    html = read(src / template)
    for marker, path in pieces.items():
        if marker not in html:
            continue
        code = read(path)
        # A closing tag inside the code would end the <script> or <style> early.
        assert "</script" not in code.lower() and "</style" not in code.lower(), f"{path.name} must not contain a closing script or style tag"
        assert html.count(marker) == 1, f"{marker} must appear once in {template}"
        html = html.replace(marker, code.replace("__ROOT__", root) if marker == "/*__SITE_CSS__*/" else code)
    policy = "; ".join([
        "default-src 'none'",
        "script-src " + " ".join(fingerprints(html, "script") + scripts),
        "style-src " + " ".join(fingerprints(html, "style")),
        "img-src 'self' data:",
        "font-src 'self'",
        "connect-src 'self'",                       # only this site's course files can be fetched
        "frame-src " + (" ".join(frames) if frames else "'none'"),
        "base-uri 'none'", "form-action 'none'", "object-src 'none'",
    ])
    assert html.count("__CSP__") == 1, f"{template} needs one __CSP__"
    html = html.replace("__CSP__", policy)
    target = site / out
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(html, encoding="utf-8")
    print("Wrote", out, f"({len(html) // 1024} KB)")
