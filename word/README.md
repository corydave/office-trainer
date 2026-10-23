# Word course - Office Trainer content

Everything learners see in the Word add-in comes from the Markdown files in this folder. To change a lesson, a video, a step or a check, edit the Markdown and rebuild. Nothing about a lesson lives anywhere else.

## What's here

| Path | What it is |
|---|---|
| `course.config.md` | Course settings: default fonts, approved video hosts, license |
| `content/00-start-here.md` | "How This Course Works": the two modes, activity labels, badges and certificates, privacy. Learners see it first |
| `content/outline.md` | The course outline: every week and stop, including lessons not written yet, plus the course outcomes and the MO-110 skill list. The roadmap is built from it |
| `content/week-01/` ... `week-05/` | All five weeks: 20 lessons, five skill checks, five projects and the capstone (practice exam and portfolio), one file each |
| `sections/` | One file per class section (`WRD-7K3Q.md`): the instructor's replacement or extra videos. `WRD-SAMPLE.md` is a template |
| `starters/week-01/` ... `week-05/` | The starter files the drills open (.docx, plus pictures for Week 3). Checks are written against these. Every person, business and source in them is made up |
| `spec/trainer-markdown.md` | The Trainer Markdown language: directives, steps, checks, validation rules |
| `spec/checker-library.md` | Every check, its parameters, how it works on web and desktop, and what Week 1 can't auto-check |
| `spec/opportunities-log.md` | Ideas and open questions found while building and testing: UX, UI, content, and what still needs real Word |
| `spec/badges-and-certificates.md` | Badges, certificates, the record format and numbering, and the options for verifying them |
| `roadmap.html` | The learner-facing course roadmap, drawn as a subway map (built; don't edit by hand) |
| `certificate.html` | Certificate preview: front and back, printable (built; don't edit by hand) |
| `tools/word-checker.js` | The checker library: opens a .docx and runs every check (one file, no dependencies besides JSZip). The add-in and `check-tests.html` share it |
| `tools/src/taskpane.src.html`, `tools/src/taskpane.app.js` | The add-in's page and code (shared by the Word and Excel courses). `build-taskpane.py` combines them with `trainer-md.js`, `word-checker.js` and `../excel/tools/excel-checker.js` into `../taskpane.html` |
| `tools/build-taskpane.py` | Rebuilds the add-in: `python3 tools/build-taskpane.py` |
| `tools/build-site-pages.py`, `tools/src/badges.src.html`, `sections.src.html`, `grader.src.html`, `site.css`, `site.js` | The badges page (`../badges.html`), the section builder and the grader (`../instructor/`). Rebuild with `python3 tools/build-site-pages.py` |
| `tools/build-index-standards.py` | Writes the home page's standards tabs from the three course maps (and updates its fingerprints) |
| `tools/test-docs/`, `tools/make-test-docs.py` | Answer keys and part-done documents for testing the checks, and the script that makes them |
| `tools/check-tests.html` | Runs every check in all five weeks against the test documents and compares the results with `test-docs/expected.json`. Open it from the website or a local web server |
| `tools/editor.html` | The editor. Open it in Chrome or Edge, then **Open folder** and pick this folder |
| `tools/trainer-md.js` | The parser, validator and renderer (one file, no dependencies). The editor and the add-in share it |
| `tools/validate.mjs` | Command-line validator: `node tools/validate.mjs content/week-01` |
| `tools/make-starters.py` | Rebuilds the starter files: `python3 tools/make-starters.py` (needs python-docx and Pillow). Weeks 2-5 are in `tools/starters_w2to5.py`; both use the helpers in `tools/docxkit.py` (fields, lists, comments, tracked changes, drawings, SmartArt, charts and more, written as raw Word XML) |
| `tools/build-editor.py` | Rebuilds `editor.html` after a change to `trainer-md.js` or `tools/src/editor.src.html` |
| `tools/build-learner-pages.mjs` | Rebuilds `roadmap.html` and `certificate.html` from the outline and lesson files: `node tools/build-learner-pages.mjs` (for Excel: `node tools/build-learner-pages.mjs ../excel`) |
| `word-course-map.md` / `.html` | The 5-week course map (`_build/` regenerates both) |
| `starters/poc/`, `fixtures/` | The proof-of-concept's practice files and test files. The add-in no longer uses them; keep them until you're sure you don't need them |

## Editing a lesson

1. Open `tools/editor.html` in Chrome or Edge and choose **Open folder** (this folder). Safari and Firefox can use **Open file** instead, but can't check that starter files exist.
2. Pick a lesson in the file list. Type on the left; the right side previews the add-in panel at its real width.
3. Use **Insert** for videos, callouts, quizzes, drills, steps and checks. **Check** lists every check in the library with an example.
4. Fix anything under **Problems**. Select a problem to jump to its line. Errors block a file from being marked `status: ready`.
5. Save with **Save** or [Ctrl/Cmd+S].

## Swapping in an instructor's videos

1. In the editor, choose **New > Section video file** (or copy `sections/WRD-SAMPLE.md`) and name it after the section's code.
2. The preview lists every course video and its id. Use **Replace video** or **Add video** in the Insert bar.
3. To see a lesson the way that section will, open the lesson and pick the section under **Preview as**.
4. Upload the section file with the lessons. Learners who enter that code (panel: **Settings > Your section**) get the instructor's videos, and the stops listed in `skip:` are marked "Not in your section"; everyone else keeps the standard course.

Easier for most instructors: the **section builder** (`../instructor/sections.html`) does the same without uploading anything. It makes a **section key** students paste instead of a code. See `spec/sections-and-grading.md`, which also covers the **grader** (`../instructor/grader.html`).

## How the add-in uses these files

- When the panel opens, `../taskpane.html` reads `course.config.md`, `content/outline.md` and every lesson the outline names. **Change a lesson's Markdown and upload it: no rebuild needed.** (Rebuild only after changing `trainer-md.js`, `word-checker.js` or the files in `tools/src/`.)
- Learners see the course home (the weeks as subway lines), then each lesson drawn exactly as the editor previews it, with working drills, quizzes, skill checks and project self-checks.
- **Open it here** puts a drill's starter into the open document. **Check my work** reads the document and marks each step **Done** or **Not yet**, with a note and, inside Word, a **Show me where** button.
- Badges, certificates and progress are kept in the learner's browser only. The **Badges** screen can save and reopen the record as a .json file.
- Links straight to a lesson or a drill (for Brightspace): `taskpane.html?lesson=W1-L2` or `taskpane.html?lesson=W1-L2&activity=D1.04`.

## Status of the course

- **All five weeks are written** (33 files) and validate with **0 errors**. The 58 warnings are expected for a draft: 21 videos still to be recorded (with checked captions) and 16 example screenshots still to be captured.
- 66 badges, 20 lesson certificates and the course certificate. Starter files exist for every drill, skill check and the capstone exam, and each opens in LibreOffice (a stand-in test that the files are valid).
- **Every check passes its tests:** 558 of 558 results match across 74 documents (`tools/check-tests.html`). The same cases also pass when run through the add-in's own buttons, and a full walk-through of the course earns every badge and the course certificate.
- **Not tested in real Word yet.** Everything was tested in a browser with a stand-in for Word, and the test documents were made by code. See `spec/opportunities-log.md` for what to try first.

## Next steps

1. **Try it in Word for the web and desktop Word,** starting with the items marked "real Word" in `spec/opportunities-log.md` (starters loading, Show me where, Compare and mail merge results in a new window, SmartArt and charts saved by Word).
2. Save finished drills from real Word into `tools/test-docs/` and add them to `expected.json`, so the checks are tested against what Word really writes.
3. Record the videos and replace each `url: TODO`; set `captions: verified` once the captions have been checked. Capture the example screenshots.
4. Mark each file `status: ready` once it validates with 0 errors.
