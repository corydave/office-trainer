# PowerPoint course - Office Trainer content

Everything learners see in the PowerPoint course comes from the Markdown files in this folder, just like the Word and Excel courses next door. The add-in is shared: the same `../taskpane.html` opens in Word, Excel and PowerPoint, works out which app opened it, and shows that course. To change a lesson, a video, a step or a check, edit the Markdown here.

## What's here

| Path | What it is |
|---|---|
| `course.config.md` | Course settings. `app: powerpoint` tells the validator, the editor and the add-in to use the PowerPoint checks |
| `content/00-start-here.md` | "How This Course Works". Learners see it first |
| `content/outline.md` | Every week and stop, the seven CSC 136 outcomes and the 73 MO-310 skills. The roadmap is built from it |
| `content/week-01/` ... `week-05/` | All five weeks: 20 lessons, five skill checks, five projects and the capstone (a 20-task MO-310 practice exam and a portfolio) |
| `sections/PP-SAMPLE.md` | Template for an instructor's own videos |
| `starters/week-01/` ... `week-05/` | The practice presentations, plus the files drills need: a Word outline, slides to reuse, a reviewer's copy, photos, a logo, a short video, music and a 3D model. Built by code; every person and business in them is made up |
| `spec/checker-library.md` | Every PowerPoint check, its parameters, how the checker reads a presentation, and what it can't see |
| `spec/opportunities-log.md` | Fixes made, open questions, what to try in real PowerPoint first, and ideas |
| `roadmap.html`, `certificate.html` | The learner roadmap and certificate viewer (built; don't edit by hand) |
| `powerpoint-course-map.md` / `.html` | The 5-week course map: CLOs, topics, MO-310 coverage, Northstar, NACE, DigComp, WCAG. `_build/` regenerates both |
| `tools/ppt-checker.js` | The checker library: opens a .pptx and runs every check (one file; needs only JSZip). The add-in and `check-tests.html` share it |
| `tools/pp_w1.py` ... `pp_w5.py`, `tools/pptkit.py` | The practice-file builders (python-pptx, plus small XML patches for what python-pptx can't write: transitions, animations, sections, comments, Zoom and more) |
| `tools/make-starters.py` | Rebuilds every starter: `python3 tools/make-starters.py` (needs `python-pptx`, `lxml`, `pillow`, `python-docx` and `ffmpeg`) |
| `tools/make-test-docs.py`, `tools/test-docs/` | Answer keys and part-done presentations made by the same builders, and `expected.json` |
| `tools/check-tests.html` | Runs every check against the test presentations and compares with `expected.json`. Open it from the website or a local web server |

The editor (`../word/tools/editor.html`), the validator (`../word/tools/validate.mjs`) and the lesson language (`../word/spec/trainer-markdown.md`) are shared with the other courses.

## Common jobs

| To | Run (from this folder) |
|---|---|
| Check the lessons | `node ../word/tools/validate.mjs content/*.md content/week-0*/*.md` |
| Rebuild the starters | `python3 tools/make-starters.py` |
| Rebuild the test presentations | `python3 tools/make-test-docs.py`, then open `tools/check-tests.html` |
| Rebuild the roadmap and certificate | `node ../word/tools/build-learner-pages.mjs .` |
| Rebuild the course map | `cd _build && python3 build.py` |
| Rebuild the add-in (after a change to `ppt-checker.js`) | `cd ../word && python3 tools/build-taskpane.py` |

## How the add-in uses these files

- In PowerPoint, the panel reads `powerpoint/course.config.md`, `powerpoint/content/outline.md` and every lesson the outline names. In a plain browser, `taskpane.html?app=powerpoint` opens this course (and it's remembered), with a **Course: Word / Excel / PowerPoint** switch on the home page.
- PowerPoint progress is saved separately (`officeTrainer.powerpoint.v1`), so the same drill numbers never mix with Word's or Excel's. The first time someone opens this course, their name and display settings are copied from another course.
- **Open it in PowerPoint** opens a drill's practice file as a **new presentation** (`PowerPoint.createPresentation`, PowerPointApi 1.1), so masters, layouts, sections and show settings all come along. The learner opens Office Trainer in that window to check. **Download** is always there too.
- **Check my work** reads the open presentation (saved or not) and marks each step **Done** or **Not yet**. **Show me where** goes to the slide and selects the shape the check is about (PowerPoint 1.5 APIs; older versions just go to the slide).
- Lesson links for Brightspace: `taskpane.html?lesson=PP1-L2` or `taskpane.html?lesson=PP1-L2&activity=D1.06`.

## Status of the course

- **All five weeks are written** (33 files) and validate with **0 errors**. The 13 warnings are expected for a draft: 3 videos to record (with checked captions) and 7 example screenshots to capture. More video slots can be added per lesson.
- 54 badges, 20 lesson certificates and the course certificate. Every MO-310 skill (73 of 73) is practiced in a drill; 54 are also in a timed skill check, and the capstone samples 22 from every domain.
- **Every check passes its tests:** 680 of 680 results match across 92 test cases (`tools/check-tests.html`). The same cases pass through the add-in's own buttons with a stand-in for PowerPoint, a full walk-through earns all 54 badges and the course certificate, and the accessibility scan (axe-core, WCAG 2.1 AA) finds 0 violations on every page in light and dark mode. The starters open in LibreOffice.
- **Not tested in real PowerPoint yet.** See `spec/opportunities-log.md`, section 2, for what to try first. Three independent content reviews checked every menu path and web-versus-desktop claim; their fixes are in the lessons.

## Decisions to review

- **CLO 7** says "Collaborate on presentations: Compare and Combine presentations". Microsoft removed **Review > Compare** from Microsoft 365 PowerPoint for Windows in 2025 (version 2502); it was never in PowerPoint for Mac or the web. Lesson 5.2 teaches Compare for PowerPoint 2024 and older, and a side-by-side method for everyone else, with the same checked result. Consider rewording CLO 7, for example "Collaborate on presentations: co-author, review, and combine changes from reviewers". Check which PowerPoint the campus labs have.
- **Lakeshore Paddle Company** carries on from Word and Excel, so the three courses feel like one program.
- Several drills are desktop-only because PowerPoint for the web can't do the skill (Outline view, WordArt, text columns, the slide master, charts, SmartArt, motion paths, media playback, Zoom, show settings). Each is labeled, and the start page lists them.
