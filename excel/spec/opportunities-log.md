# Opportunities log (Excel)

Ideas, open questions and fixes found while building and testing the Excel course (October 2026, add-in 0.6, Excel checker 0.2.0). Each item says what it is, why it matters, and a suggested next step. Priority: **High** (affects learners now), **Medium**, **Low**.

## 1. Fixed while building

Found by testing:

| What was wrong | Fix |
|---|---|
| D1.10 (Format Painter) passed on the starter, because both header rows started plain | The check now also looks for the fill and wrap the learner copies |
| D1.06's delete-cell step passed on the starter | It also checks the title row is back in place |
| Number formats like `[$-409]m/d/yyyy` read as "text", and blank cells failed format checks | The checker ignores locale tags and skips blank cells |
| A text function's answer ("27" from `MID`) failed a numeric `result=` | A number-looking answer also matches the same text |
| D3.05's expected total was a guess (70); the real value is 105 | Fixed, and every expected value was recalculated from the starters by an independent review |
| Formulas with `*` turned into italics in the lesson text | Every formula is now in backticks |
| D4.08's "result is $193.33" step would fail again after Goal Seek changed the price | `formula.is` gained `min=` and `max=`: the step now checks for a positive `PMT` payment |
| A project preflight's `datatable.is` with no sheet only looked at the first sheet | It looks at every sheet when no sheet is named |
| Two checks that failed for the same reason showed the same message twice | Messages are shown once |
| A note's "still there" message named the cell twice ("A note in D9 is still there, in D9") | Fixed |
| A long web address in bold widened the panel past a 320 px screen | Lesson text wraps long words |
| Preflights with no practice file (projects) didn't show their download links (lakeshore-season.csv) | They now show a **Download:** bar |

From two independent content reviews (every fix is in the lesson files):

- Excel for the web shortcuts: **Ctrl+L** (not Ctrl+T) for a table, **Ctrl+G** (not F5) for Go To, **Ctrl+F6** (not F6) to move between the panel and the workbook. F2 works on the web too.
- Power Query's Remove Duplicates only compares the selected columns: the hint now says to select every column first.
- The Quick Access Toolbar menu has no Save As: the hint now goes through More Commands > All Commands.
- "Move but don't size with cells" doesn't keep a logo in place: it's **Don't move or size with cells**.
- Chart Elements (+) and the Styles brush are desktop-only: Lesson 4.2 has a web note (the Format pane).
- Opening a template with File > Open opens the template itself; File > New > Personal (or a double-click) makes a copy. The lesson and quiz say so.
- Skill Check 4 asks for the quick layout before the title (some layouts replace the title). Skill Check 3 asks for the filter last (formatting while filtered can miss hidden rows).
- D5.02: the @mention comment comes after restoring a version (restoring would remove it).
- Check Accessibility doesn't flag default table names, a missing Title, or color alone. Lesson 5.3 now separates "the checker flags" from "also fix yourself" (the trainer's checks test both).
- Inspect Document's categories use Excel's real names (Comments and Annotations; Hidden Rows and Columns; Hidden Worksheets), warn that removing hidden rows deletes data, and say that removing properties clears the Title. Project 5 says to set the Title after inspecting.
- Mac: **Excel > Settings** (Preferences on older versions).
- Co-authoring: **File > Save a Copy** (not Save As) when AutoSave is on. Embed links on college accounts may ask viewers to sign in.
- CSV and PDF on the web: **File > Export > Download this sheet as CSV** and **File > Export > Download as PDF** (Week 1 said Save As).
- Setting the Title property is marked desktop-only (as in Week 1) in D5.05, D5.07 and the capstone.
- D1.08: the currency step includes the total (C5:C11), and the Format Cells dialog step is desktop-only (the review couldn't confirm the web's dialog has Border and Fill tabs; see 2.6).
- Project 2's invoice number no longer uses `TODAY()` (an invoice shouldn't change every day); it explains Ctrl+; for a fixed date.
- Hints explain two answers a beginner would doubt: `COUNTBLANK` = 6 (three unpaid plus three spacer rows) and `CONCAT` showing "jordan RIVERA".
- Week 3's data runs into the first week of August; the scenario now says so.

## 2. Test in real Excel first (High)

Everything ran in a browser with a stand-in for Excel, and the workbooks were made by code. Try these first, in Excel for the web and desktop Excel (Windows and Mac):

1. **The add-in in Excel at all.** Upload the new `manifest.xml` (version 1.3.0.0, now with a Workbook host). Microsoft's online manifest validator couldn't be reached from the build machine; run `npx office-addin-manifest validate manifest.xml` once. Check the button appears on Excel's Home tab, and that Word still works.
2. **Open it in Excel** uses `Excel.createWorkbook` (ExcelApi 1.8). Check that it opens the practice file in a new window (desktop) or tab (web), and that Office Trainer opened there remembers the learner's progress (both panels share this site's storage).
3. **getFileAsync in Excel.** Confirm it includes unsaved changes on the web and on desktop, and that `Office.context.document.url` gives the file name `file.name` needs (D4.07's .xltx, Project file names).
4. **Show me where**: going to a sheet and selecting a range, including a chart sheet (the checker sends just the sheet).
5. **Files Excel really writes.** Save real results for: threaded comments (replies, resolved), notes, scenarios, a data table, Consolidate (`dataConsolidate`), an external link (`externalLink`, and its file path on OneDrive), Inspect Document's "remove personal information" (`filterPrivacy`), sheet and structure protection, a theme change, a custom cell style, icons (SVG), decorative pictures, chart styles (the "not the default style" test assumes style 2 / 201 is the default), chart alt text, a Power Query connection, multiple windows (`window.count`). Add them to `tools/test-docs/` and `expected.json`.
6. **Facts to confirm** (the reviews couldn't verify these): Find All's count in Excel for the web (D1.05's hint), whether the web's Format Cells dialog has Border and Fill tabs, and what a saved .htm page shows as its browser-tab title (D5.07 asks the learner to describe it).

## 3. UX opportunities

| Priority | Idea | Why |
|---|---|---|
| High | **Open the practice file and the trainer together.** Excel opens the practice workbook in a new window, and the learner has to open Office Trainer there. A starter could carry an "open this add-in with the file" setting (a web extension part pointing at the add-in), so the panel opens by itself | The extra step is where beginners get lost |
| High | **A "new window" helper card** with a screenshot of the Home tab button, shown after "Open it in Excel" | Same reason |
| Medium | **Skill check answers in the page** (as in Word): `answer:` values are in the page source. Fine for practice; check by hash in a graded section | Integrity |
| Medium | **One record for both courses.** Badges are per course (they share drill numbers). The Badges screen could show the other course's totals, and "Save my record" could save both | Learners taking both courses |
| Medium | **Preflight results could be saved** like drills | Projects take several sessions |
| Low | The browser-mode **Course: Word / Excel** switch could also appear in Settings | Discoverability |

## 4. UI opportunities

| Priority | Idea |
|---|---|
| Medium | Week 1 Lesson 1 and Week 3 Lesson 2 have four or five drills; on a phone they're long. Collapse finished drills |
| Medium | The selected course chip in the home switch only shows a check mark; a filled style would make it clearer |
| Low | The roadmap's station labels wrap to three lines in places; shorter outline titles would help |

## 5. Content opportunities

| Priority | Idea |
|---|---|
| High | **Videos and screenshots.** 24 videos and 11 example screenshots are TODO |
| Medium | **CLO 7 and CLO 6 wording** (see the README): the course teaches the modern way to compare and combine, and shows the legacy Compare and Merge once |
| Medium | **PivotTables** aren't on MO-210 and aren't taught. They're the most-asked-for Excel skill at work; an optional Week 3 lesson (or MO-211 extension) would be popular |
| Medium | **SmartArt** is detected by the checker but D4.06 uses a screenshot step; a real-Excel SmartArt file would let it be auto-checked |
| Medium | **AI callouts** are only in the five projects. Lessons could add short Copilot in Excel notes where it fits (explain a formula in Week 2, suggest a chart in Week 4), each with a "check it yourself" line |
| Low | **MO-211 extension.** Seven MO-211 objectives are partly covered (what-if, protection, data validation). A short "Expert extras" week could cover the rest (XLOOKUP, advanced conditional formatting, PivotTables, macros) |
| Low | A second persona besides Lakeshore Paddle Company (a school club, a small café) would vary the data |

## 6. Housekeeping

- The Playwright tests (end-to-end with a stand-in for Excel, the course walk, keyboard, axe) live outside the repo. Moving them into `tools/tests/` would let anyone rerun them.
- `make-test-docs.py` writes 90 workbooks; the test-docs folder doesn't need to be on the public site.
- Badge ids include the drill id, so renumbering a drill changes its badge number. Never renumber after launch.
