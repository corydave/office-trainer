# build-index-csp.py - updates the security fingerprints in index.html.
#
# index.html (the product page) has a Content Security Policy that lets the
# browser run only the ONE <script> and the ONE <style> block in the page,
# identified by their sha256 fingerprints. Change either block and the old
# fingerprint no longer matches, so the browser would refuse to run it.
# This script works out the new fingerprints and writes them into the policy.
#
# Run from the office-trainer folder (or anywhere):  python3 word/tools/build-index-csp.py
import base64, hashlib, pathlib, re

page = pathlib.Path(__file__).resolve().parent.parent.parent / "index.html"
html = page.read_text(encoding="utf-8")


def fingerprint(tag):
    """sha256 of the exact text between <tag> and </tag> (there must be exactly one)."""
    # Only a real tag at the start of a line counts (not the word in a comment).
    blocks = re.findall(r"^[ \t]*<%s>(.*?)</%s>" % (tag, tag), html, re.S | re.M)
    assert len(blocks) == 1, "index.html must have exactly one <%s> block (found %d)" % (tag, len(blocks))
    return base64.b64encode(hashlib.sha256(blocks[0].encode("utf-8")).digest()).decode()


script, style = fingerprint("script"), fingerprint("style")
html, n1 = re.subn(r"script-src 'sha256-[^']*'", "script-src 'sha256-%s'" % script, html)
html, n2 = re.subn(r"style-src 'sha256-[^']*'", "style-src 'sha256-%s'" % style, html)
assert n1 == 1 and n2 == 1, "the Content Security Policy in index.html should have one script-src and one style-src"
page.write_text(html, encoding="utf-8")
print("Updated index.html: script sha256-%s, style sha256-%s" % (script, style))
