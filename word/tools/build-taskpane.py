# build-taskpane.py - makes taskpane.html, the Word, Excel and PowerPoint add-in, as ONE file.
#
# It copies five pieces of JavaScript into tools/src/taskpane.src.html:
#   tools/trainer-md.js              reads and draws the lesson files (shared with the editor)
#   tools/word-checker.js            checks a Word document (shared with check-tests.html)
#   ../excel/tools/excel-checker.js  checks an Excel workbook (shared with excel's check-tests.html)
#   ../powerpoint/tools/ppt-checker.js checks a PowerPoint presentation (shared with powerpoint's check-tests.html)
#   tools/src/taskpane.app.js        the add-in's own screens and buttons
# and writes the result to taskpane.html at the top of the office-trainer folder
# (next to manifest.xml), which is where Word, Excel and PowerPoint load it from.
#
# Run from the word/ folder:  python3 tools/build-taskpane.py
# Then bump APP_VERSION in taskpane.app.js if you're uploading a new version.
import pathlib

here = pathlib.Path(__file__).resolve().parent
src = (here / "src" / "taskpane.src.html").read_text(encoding="utf-8")
pieces = {
    "/*__TRAINER_MD__*/": here / "trainer-md.js",
    "/*__WORD_CHECKER__*/": here / "word-checker.js",
    "/*__EXCEL_CHECKER__*/": here.parent.parent / "excel" / "tools" / "excel-checker.js",
    "/*__PPT_CHECKER__*/": here.parent.parent / "powerpoint" / "tools" / "ppt-checker.js",
    "/*__APP__*/": here / "src" / "taskpane.app.js",
}
for marker, path in pieces.items():
    code = path.read_text(encoding="utf-8")
    # A closing script tag inside the code would end the <script> early.
    assert "</script" not in code.lower(), f"{path.name} must not contain a closing script tag"
    assert src.count(marker) == 1, f"{marker} must appear exactly once in taskpane.src.html"
    src = src.replace(marker, code)

out = here.parent.parent / "taskpane.html"
out.write_text(src, encoding="utf-8")
print("Wrote", out, f"({len(src) // 1024} KB)")
