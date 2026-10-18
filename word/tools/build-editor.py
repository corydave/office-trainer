# build-editor.py - makes tools/editor.html, a single file you can open in a browser.
# It copies trainer-md.js into the editor so both always use the same parser.
# Run from the word/ folder:  python3 tools/build-editor.py
import pathlib
here = pathlib.Path(__file__).parent
src = (here / "src" / "editor.src.html").read_text(encoding="utf-8")
lib = (here / "trainer-md.js").read_text(encoding="utf-8")
assert "</script" not in lib.lower(), "trainer-md.js must not contain a closing script tag"
out = src.replace("/*__TRAINER_MD__*/", lib)
(here / "editor.html").write_text(out, encoding="utf-8")
print("Wrote", here / "editor.html", f"({len(out)//1024} KB)")
