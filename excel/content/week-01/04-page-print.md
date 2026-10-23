---
type: lesson
id: XL1-L4
week: 1
order: 4
title: Rows, Columns, Printing and Properties
objectives: [W1.4, W1.5]
clos: [CLO1, CLO6]
minutes: 45
status: draft
---

::: callout why
A worksheet that looks fine on screen can print as six pages with one column on each. A few minutes on column widths and page setup turn it into one clean page or a PDF you can send.
:::

## Column widths and row heights

| To | Do this |
|---|---|
| Fit a column to its contents | Double-click the line to the right of its column letter (**AutoFit**) |
| Set an exact width | Select the column, then [[path: Home > Format > Column Width]] |
| Set an exact row height | Select the row, then [[path: Home > Format > Row Height]] |
| Change several at once | Select all the columns (or rows) first |

Column widths are measured in characters (8.43 is the default); row heights in points.

## Working on several sheets at once

To make the same change on several sheets, **group** them: select the first tab, then hold [[key: Ctrl / Cmd]] and select the others. The title bar says **Group**. Whatever you type or format now happens on every grouped sheet. Right-click a tab and choose **Ungroup Sheets** when you're done.

::: callout warning
Ungroup as soon as you finish. While sheets are grouped, typing in one cell overwrites the same cell on every sheet in the group.
:::

::: drill D1.11
title: Set widths and heights on two sheets at once
badge: Layout Grouper
starter: starters/week-01/E1-L4-report.xlsx
platform: desktop
mos: 1.3.2, 2.2.9
other: N10
minutes: 5

- step: Group the **June** and **July** sheets, then set column **A** to a width of **20** on both at once.
  check: col.is sheet=June col=A width=20
  check: col.is sheet=July col=A width=20
  hint: Select June's tab, [[key: Ctrl / Cmd]]+select July's, select column A, then [[path: Home > Format > Column Width]].
- step: While they're still grouped, set row **3** (the headers) to a height of **32** on both.
  check: row.is sheet=June row=3 height=32
  check: row.is sheet=July row=3 height=32
- step: Ungroup the sheets.
:::

## Page setup

| Setting | Where |
|---|---|
| Orientation, margins, paper size | [[path: Page Layout > Page Setup]] |
| Fit to one page wide | [[path: Page Layout > Scale to Fit]]: set **Width** to **1 page** and **Height** to **Automatic** |
| Print only part of a sheet | Select the cells, then [[path: Page Layout > Print Area > Set Print Area]] |
| Headers and footers | [[path: Insert > Header & Footer]]. Use the **Header & Footer** tab's buttons for the page number, sheet name, file name or date, so they update themselves |
| Check the result | [[path: File > Print]] shows a preview ([[key: Ctrl+P / Cmd+P]]) |

A header or footer has three parts: left, center and right. Click in the part you want before you use a button.

::: platform web
Excel for the web can change orientation, margins, scaling and the print area when you print, but it can't add headers and footers. Use desktop Excel for this drill.
:::

::: video
id: e1-l4-print
url: TODO
title: Page setup, print areas and footers
captions: none
duration: 4:00
:::

::: drill D1.12
title: Set up the June report to print on one page
badge: Page Setter
starter: starters/week-01/E1-L4-report.xlsx
platform: desktop
mos: 1.3.1, 1.3.3, 1.5.1, 1.5.3
other: N17
minutes: 8

- step: Set the **June** sheet to **landscape**.
  check: page.is sheet=June orientation=landscape
- step: Scale it to fit **1 page wide** (height Automatic).
  check: page.is sheet=June fitWide=1 fitTall=0
- step: Set the print area to the report only, **A1:F11**.
  check: page.is sheet=June printArea=A1:F11
- step: Add a footer with the **sheet name** on the left and the **page number** on the right.
  check: header.is sheet=June part=footer field=sheet position=left
  check: header.is sheet=June part=footer field=page position=right
  hint: [[path: Insert > Header & Footer]], scroll to the footer, click its left part and select **Sheet Name**, then click its right part and select **Page Number**.
- step: Add a header with **Lakeshore Paddle Company** in the center.
  check: header.is sheet=June part=header says="Lakeshore Paddle Company" position=center
:::

## Properties and PDF

**Properties** travel with the file: a **Title** becomes the PDF's title and what screen readers announce, and **Tags** help search. Fill them in on [[path: File > Info]] (desktop Excel), on the right side.

To make a PDF, use [[path: File > Export > Create PDF/XPS]] (desktop) or [[path: File > Export > Download as PDF]] (web). Open the PDF to check it before you send it.

::: drill D1.13
title: Add properties and export a PDF
badge: PDF and Properties
starter: starters/week-01/E1-L4-report.xlsx
platform: desktop
mos: 1.4.5, 1.5.2
other: N13
minutes: 5

- step: Set the **Title** property to **June Rental Report** and add the tag **rentals**.
  check: props.is title="June Rental Report"
  check: props.is keywords=rentals
  hint: [[path: File > Info]], then type in the **Title** and **Tags** boxes on the right.
- step: Export the June sheet as a PDF named **Lastname_E1_Report.pdf** and open it.
  screenshot: Your PDF open in a browser or PDF reader, showing the report on one landscape page with the footer.
  example: TODO
  describe: Say how many pages the PDF has and what the footer shows.
:::

::: quiz Q14.1
question: You want the page number in the footer. Why use the Page Number button instead of typing 1?
type: single
- [ ] Typed numbers can't go in a footer
- [x] The button inserts a code that shows the right number on every page
- [ ] The button makes the number bold
explain: The Page Number button adds a code (&[Page]) that Excel replaces with each page's own number when it prints.
:::

## Check your understanding

- What does "Width: 1 page, Height: Automatic" do to a wide sheet?
- Why is it risky to leave sheets grouped?
- Where do a workbook's Title and Tags show up after you share it?
