# Excel course - Office Trainer content

Everything learners see in the Excel course comes from the Markdown files in this folder, exactly like the Word course next door. The add-in is shared: the same `../taskpane.html` opens in Word and in Excel, works out which app opened it, and shows that course. To change a lesson, a video, a step or a check, edit the Markdown here.

## What's here

| Path | What it is |
|---|---|
| `course.config.md` | Course settings. `app: excel` tells the validator, the editor and the add-in to use the Excel checks |
| `content/00-start-here.md` | "How This Course Works". Learners see it first |
| `content/outline.md` | Every week and stop, the seven CSC 135 outcomes and the 64 MO-210 skills. The roadmap is built from it |
| `content/week-01/` ... `week-05/` | All five weeks: 20 lessons, five skill checks, five projects and the capstone (a 20-task MO-210 practice exam and a portfolio) |
| `sections/XL-SAMPLE.md` | Template for an instructor's own videos |
| `starters/week-01/` ... `week-05/` | The practice workbooks (plus a few CSV files, a logo and Review-Priya.xlsx for linking). Built by code; every person and business in them is made up |
| `spec/checker-library.md` | Every Excel check, its parameters, how the checker reads a workbook, and what it can't see |
| `spec/opportunities-log.md` | Fixes made, open questions, what to try in real Excel first, and ideas for UX, UI and content |
| `roadmap.html`, `certificate.html` | The learner roadmap and certificate viewer (built; don't edit by hand) |
| `excel-course-map.md` / `.html` | The 5-week course map: CLOs, topics, MO-210 and MO-211 coverage, Northstar, NACE, DigComp, WCAG. `_build/` regenerates both |
| `tools/excel-checker.js` | The checker library: opens an .xlsx and runs every check (one file; needs only JSZip). The add-in and `check-tests.html` share it |
| `tools/xl_w1.py` ... `xl_w5.py`, `tools/xlsxkit.py` | The practice-workbook builders (XlsxWriter, plus small XML patches for what XlsxWriter can't write) |
| `tools/make-starters.py` | Rebuilds every starter: `python3 tools/make-starters.py` (needs `xlsxwriter` and `lxml`) |
| `tools/make-test-docs.py`, `tools/test-docs/` | Answer keys and part-done workbooks made by the same builders, and `expected.json` |
| `tools/check-tests.html` | Runs every check against the test workbooks and compares with `expected.json`. Open it from the website or a local web server |

The editor (`../word/tools/editor.html`), the validator (`../word/tools/validate.mjs`) and the lesson language (`../word/spec/trainer-markdown.md`) are shared with the Word course.

## Common jobs

| To | Run (from this folder) |
|---|---|
| Check the lessons | `node ../word/tools/validate.mjs content/*.md content/week-0*/*.md` |
| Rebuild the starters | `python3 tools/make-starters.py` |
| Rebuild the test workbooks | `python3 tools/make-test-docs.py`, then open `tools/check-tests.html` |
| Rebuild the roadmap and certificate | `node ../word/tools/build-learner-pages.mjs .` |
| Rebuild the course map | `cd _build && python3 build.py` |
| Rebuild the add-in (after a change to `excel-checker.js`) | `cd ../word && python3 tools/build-taskpane.py` |

## How the add-in uses these files

- In Excel, the panel reads `excel/course.config.md`, `excel/content/outline.md` and every lesson the outline names. In Word it reads `word/` instead. In a plain browser, `taskpane.html?app=excel` opens this course (and it's remembered), with a **Course: Word / Excel** switch on the home page.
- Excel progress is saved separately from Word progress (`officeTrainer.excel.v1`), so the same drill numbers never mix. The first time someone opens the Excel course, their name and display settings are copied from the Word course.
- **Open it in Excel** opens a drill's practice file as a **new workbook** (a new window in desktop Excel, a new tab on the web), so properties, names, links and protection all come along. The learner opens Office Trainer in that window to check. **Download** is always there too.
- **Check my work** reads the open workbook (saved or not) and marks each step **Done** or **Not yet**. **Show me where** goes to the sheet and selects the cells.
- Lesson links for Brightspace: `taskpane.html?lesson=XL1-L2` or `taskpane.html?lesson=XL1-L2&activity=D1.06`. Inside Excel the course is chosen by the app, so the same link works in both.

## Status of the course

- **All five weeks are written** (34 files) and validate with **0 errors**. The 59 warnings are expected for a draft: 24 videos to record (with checked captions) and 11 example screenshots to capture.
- 64 badges, 21 lesson certificates and the course certificate. Every MO-210 skill (64 of 64) is practiced in a drill; 53 are also in a timed skill check.
- **Every check passes its tests:** 621 of 621 results match across 90 workbooks (`tools/check-tests.html`). The same cases pass through the add-in's own buttons with a stand-in for Excel, a full walk-through earns all 64 badges and the course certificate, and the accessibility scan (axe-core, WCAG 2.1 AA) finds 0 violations on every page in light and dark mode. Every starter opens in LibreOffice.
- **Not tested in real Excel yet.** See `spec/opportunities-log.md`, section 2, for what to try first.

## Decisions to review

- **CLO 7** says "Compare and Combine workbooks". Excel's Compare and Merge needs the retired Shared Workbook feature, so Week 5 teaches the modern way (co-authoring, comments, Show Changes, Version History, external references and Consolidate) and shows the legacy command once. Consider rewording CLO 7, for example "Collaborate on workbooks: co-author, review changes, and combine workbooks".
- **CLO 6** ("preview the worksheet as a Web page"): taught with Save as Web Page (desktop) and mentions Embed. It may also be worth rewording.
- Project 4 and Project 5 were reshaped while writing (a business plan; a team event workbook). The course map now matches the lessons.
