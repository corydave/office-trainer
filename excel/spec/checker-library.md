# Excel checker library (v0.2.0, all five weeks)

The checker is the only code that decides whether a step passed. Lesson files name a check and pass it values; they can't add logic. The code is `tools/excel-checker.js` (one file; it needs only JSZip). The add-in, `tools/check-tests.html` and the command-line tests all share it.

**Principle: check the outcome, not the clicks.** A check looks at the saved workbook, never at how the learner got there. Typing `=SUM(B2:B9)`, using AutoSum or clicking **Total Row** all pass a "the total is there" check.

## 1. How the checker reads a workbook

- **Inside Excel**, the add-in asks for a copy of the open workbook (`getFileAsync`, the whole .xlsx). This works the same in Excel for the web and desktop Excel, and includes changes that aren't saved yet.
- **In a browser**, the learner uploads the saved .xlsx (or .xltx) and gets the same result.
- The checker opens the file with JSZip and reads the XML parts with the browser's DOMParser. Nothing runs from the file: formulas are read as text, never evaluated.
- **Results come from Excel.** Excel saves the last calculated value of every formula (`<v>` in the sheet XML). The checker compares those, so it never has to recalculate. A `result=` check reads what the learner sees in the cell.
- **Filled formulas** are compared in R1C1 form, so `=B2*C2` in row 2 and `=B9*C9` in row 9 are "the same formula". Shared formulas (Excel saves a filled range once and marks the rest "shared") are rebuilt cell by cell first.
- **Function names** are compared without Excel's hidden prefixes (`_xlfn.`, `_xlws.`, `_xlpm.`), so `UNIQUE`, `SORT`, `CONCAT` and `TEXTJOIN` read as written.
- **Number formats** are sorted into kinds (general, number, currency, accounting, percent, date, time, fraction, scientific, text) by reading the format code, ignoring locale tags like `[$-409]`. Blank cells are skipped by format checks.
- **Comments and notes** are told apart: threaded comments live in `xl/threadedComments`, notes in the classic comments part. Excel also writes a "[Threaded comment]" note copy of every comment for older versions; the checker skips those.
- **Office.js is used for three things only:** getting the file, opening a practice file (`Excel.createWorkbook`, a new window or tab), and "Show me where" (going to the sheet and selecting the cells).

Each check returns `{ pass, message, where }`. Messages are written for learners ("F2 uses the General number format. It should use Currency."). `where` is `{ sheet, range }`, which "Show me where" selects.

## 2. Pointing at cells

Most checks take `sheet=June` plus `cell=B2` or `range=B2:B9`. A range can include its sheet (`range=June!B2:B9`) or be a defined name (`range=TaxRate`). With no sheet, the active sheet is used (or the first worksheet, if a chart sheet is active). Values with spaces go in quotes: `sheet="Rentals Chart"`.

## 3. The checks

`*` = required. "Selector" = takes `cell=` or `range=`.

### Files, properties and sheets

| Check | What it tests | Parameters |
|---|---|---|
| `file.name` | The file's name matches a pattern (`*` wildcards). "Lastname_" fails with a reminder to use your own name | pattern* |
| `props.is` | The Title, Author or Tags (keywords) property | title, author, keywords |
| `self.attest` | Nothing: the learner ticks a box (not allowed in skill checks) | label* |
| `sheet.is` | A sheet exists (or doesn't), its position, tab color, hidden state, worksheet or chart sheet | name*, exists, index, after, tabColor, hidden, kind |
| `sheet.count` | How many sheets (optionally visible only) | min, max, visible |
| `window.count` | How many windows the workbook has open (View > New Window) | min* |
| `doc.is` | What Inspect Document changes: personal information removed (or the "remove on save" flag), hidden sheets | personalInfo (kept/removed), hiddenSheets |
| `access.is` | The accessibility rules: alt text, default sheet names, default table names, Title, table header rows, merged cells in tables, link text | rules (any of alt, sheetNames, tableNames, title, tableHeaders, merged, links; default all) |
| `theme.is` | The workbook theme, or just "not Office" | name, changed, colors |
| `cellstyle.is` | A named cell style exists, and is the learner's own (not built in) | name*, exists, custom |
| `protect.is` | Sheet protection (and password), workbook structure protection, and that input cells are unlocked | sheet, sheetProtected, password, structure, structurePassword, unlocked |
| `connection.count` | Data connections (Get Data / Power Query) | min, max |

### Cells, formulas and text

| Check | What it tests | Parameters |
|---|---|---|
| `cell.is` | Every cell in the range: value, number, date, type, number format and decimals, font, fill, alignment, wrap, indent, rotation, borders, cell style, locked, merged, "same format as" another cell, "no formatting" | Selector (required), plus value, contains, number, tolerance, date, type, blank, numFmt, decimals, strict, bold, italic, underline, fontSize, font, fontColor, fill, align, valign, wrap, indent, rotation, border, style, locked, clean, sameAs, merged |
| `formula.is` | The cells hold formulas: which functions, what text they contain or lack, the result, filled consistently, a spilled dynamic array, cell references instead of typed numbers, structured references, a link to another workbook, a minimum or maximum result | Selector (required), plus uses, contains, lacks, result, filled, spill, hasFormula, refs, structured, external, min, max |
| `text.present` / `text.absent` | Text appears (at least `min` times) or is gone, in a range, a sheet or the whole workbook | value*, sheet, min, max, case, selector |
| `name.is` | A defined name, what it refers to, its scope | name*, refersTo, scope, exists |
| `link.present` | Hyperlinks: to a web address or a place in the workbook, with descriptive text | sheet, url, location, label, min, max, descriptive, selector |
| `col.is` / `row.is` | Column width or row height, hidden | col* or row*, width, minWidth, maxWidth, height, minHeight, hidden |
| `validation.is` | Data validation: the type, the list items, or that none is left | sheet, type, list, exists, selector |
| `comment.is` | A comment or note exists, says something, has a reply, is resolved, or is gone | kind* (comment/note), says, author, state (present/absent/replied/resolved/open), selector |
| `comment.count` | How many comments and/or notes | kind, min, max, selector |

### Views and printing

| Check | What it tests | Parameters |
|---|---|---|
| `view.is` | Freeze panes, Normal / Page Layout / Page Break Preview, Show Formulas, split, gridlines | sheet, frozen, view, showFormulas, split, gridlines |
| `page.is` | Orientation, paper size, fit to pages, scale, margins, centering, print gridlines, print area, print titles | sheet, orientation, size, fitWide, fitTall, scale, margins, top, bottom, left, right, centerH, gridlines, printArea, printTitles |
| `header.is` | Header or footer text and fields (page number, pages, sheet name, file name, date, time, path), and where they sit | sheet, part* (header/footer), says, field, position |

### Tables and data

| Check | What it tests | Parameters |
|---|---|---|
| `table.is` | An Excel table: name, style, style options, total row and its function, columns, row count, no blank rows, or that it's gone (converted to a range) | sheet, name, contains, exists, style, bandedRows, bandedColumns, firstColumn, lastColumn, headerRow, totalRow, total, totalFunction, hasColumn, lacksColumn, rows, minRows, columns, noBlankRows |
| `table.count` | How many tables | sheet, contains, min, max |
| `sort.is` | The rows really are in order (it reads the data, not the sort settings) by one or two columns | by*, order, then, thenOrder, sheet, table, selector |
| `filter.is` | A filter is on, shows a value, and hides the rest | sheet, table, column, active, shows, only, selector |
| `data.is` | A block of data (or the table with given headers, `table=*` for any): its headers, row count, no duplicate rows | sheet, table, headers, rows, minRows, maxRows, noDuplicates, selector |
| `cf.is` | A conditional formatting rule of a type (and operator and value) covers the range | sheet, type, operator, value, exists, covers, selector |
| `cf.count` | How many rules touch the range (0 = cleared) | sheet, min, max, selector |
| `sparkline.is` | Sparklines: where, which type, their source, high point, markers | sheet, location, type, source, high, markers, min |

### Charts and graphics

| Check | What it tests | Parameters |
|---|---|---|
| `chart.is` | A chart: on a chart sheet, type and grouping, number of series, a series name, rows or columns as series, which cells it covers, title, axis titles, data labels, legend position, data table, a style other than the default, alt text | sheet, index, chartSheet, type, grouping, series, minSeries, seriesName, plotBy, hasTitle, title, axisTitles, dataLabels, legend, dataTable, styleChanged, alt, altSays, values, covers |
| `chart.count` | How many charts (optionally chart sheets) | sheet, min, max, chartSheet |
| `object.is` | A picture, icon, shape, text box or SmartArt: its name, its text, alt text or decorative | kind*, sheet, index, name, says, alt, decorative, items |
| `object.count` | How many graphics of a kind | kind*, sheet, min, max |

### What-if and combining

| Check | What it tests | Parameters |
|---|---|---|
| `scenario.is` | Scenario Manager scenarios: their names, how many, and the changing cell | sheet, names, min, changing |
| `datatable.is` | A What-If data table (`{=TABLE(,B5)}`) in the range, with the right input cell and no errors. With no sheet or range it looks on every sheet | sheet, input, selector |
| `consolidate.is` | Data > Consolidate was used on the sheet: the function, how many source ranges, top-row and left-column labels, links to the sources | sheet, function, sources, labels, links |

Goal Seek leaves no trace in the file, so a Goal Seek step checks its **result**: the changing cell holds the answer (`cell.is number=14931.39 tolerance=1`) and the formula shows the goal (`formula.is result=250`).

## 4. What the checker can't see (screenshot or self-check steps)

- The Quick Access Toolbar, and whether a command (like Compare and Merge Workbooks) was added to it
- Co-authoring presence, Show Changes and Version History (they live in OneDrive, not in the file)
- Files saved in other formats from the workbook: the .htm web page, the PDF, the CSV
- The Compatibility Checker's report, and Check Accessibility's own pane (the checker runs its own rules instead, see `access.is`)
- SmartArt is detected (`object.is kind=smartart`), but no starter tests it yet; D4.06 uses a screenshot step

## 5. Testing the checks

- `tools/make-starters.py` builds every practice workbook from `tools/xl_w1.py` ... `xl_w5.py` (XlsxWriter, plus small XML patches from `tools/xlsxkit.py` for things XlsxWriter can't write: cell styles, themes, threaded comments, scenarios, data tables, Consolidate, external links, protection, "remove personal information").
- `tools/make-test-docs.py` uses the same builders to make answer keys and part-done workbooks (`tools/test-docs/`), and `expected.json`: for 90 workbooks, whether each checked step should pass or fail. Every starter must fail every checked step.
- `tools/check-tests.html` runs all of them in a browser: **621 of 621 results match** (October 2026). The same cases also pass end to end through the add-in with a stand-in for Excel.
