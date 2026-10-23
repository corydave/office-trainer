---
type: lesson
id: XL1-L3
week: 1
order: 3
title: Format Cells
objectives: [W1.4]
clos: [CLO1, CLO2]
minutes: 55
status: draft
---

::: callout why
The same numbers can look like a mess or like a report. Formatting tells readers what each number is (money, a percentage, a date) and where to look first, without changing the numbers themselves.
:::

## Number formats

A number format changes how a number **looks**, never what it **is**. 0.31 formatted as a percentage shows 31%, but Excel still calculates with 0.31.

| Format | 1234.5 shows as | Use it for |
|---|---|---|
| **General** | 1234.5 | Plain numbers, as typed |
| **Currency** | $1,234.50 | Money (the $ sits next to the number) |
| **Accounting** | $ 1,234.50 | Money in columns (the $ lines up at the left) |
| **Percentage** | 123450% (so use it on 0.31, which shows 31%) | Shares and rates |
| **Short Date** | Dates such as 6/1/2027 | Dates |

Pick one from the list in [[path: Home > Number]], or open **Format Cells** for every option ([[key: Ctrl+1 / Cmd+1]]). The **Increase Decimal** and **Decrease Decimal** buttons sit next to the list.

::: callout tip
A date that shows as a five-digit number, like 46539, is a real date in the wrong format. Excel stores dates as counts of days since 1900. Apply a date format and it shows up properly.
:::

::: video
id: e1-l3-formats
url: TODO
title: Number formats and the Format Cells dialog
captions: none
duration: 3:30
:::

::: drill D1.08
title: Format money, percentages and dates
badge: Number Formatter
starter: starters/week-01/E1-L3-report.xlsx
platform: web
mos: 2.2.5, 2.2.6
other: N7
minutes: 7

- step: Format the **Revenue** numbers, including the total (C5:C11), as **Currency** with 2 decimal places.
  check: cell.is sheet=June range=C5:C11 numFmt=currency decimals=2
  hint: Select the cells, then pick **Currency** in the [[path: Home > Number]] list.
- step: Format the **Share of revenue** numbers (D5:D11) as **Percentage** with no decimal places.
  check: cell.is sheet=June range=D5:D11 numFmt=percent decimals=0
- step: The **First rental** column shows five-digit numbers. Format E5:E10 as a **Short Date**.
  check: cell.is sheet=June range=E5:E10 numFmt=date
- step: Open **Format Cells** for the header row (A3:F3), add a **bottom border** on the Border tab and a light **fill color** on the Fill tab.
  platform: desktop
  check: cell.is sheet=June range=A3:F3 border=bottom fill=any
  hint: Select A3:F3 and press [[key: Ctrl+1 / Cmd+1]]. On the Border tab, choose a line style, then select the bottom edge of the preview.
:::

## Alignment

| Tool | Where | What it does |
|---|---|---|
| **Merge & Center** | [[path: Home > Merge & Center]] | Joins the selected cells into one and centers the text. Use it for titles above a table, not inside data |
| **Wrap Text** | [[path: Home > Wrap Text]] | Long text wraps onto more lines in the same cell |
| **Orientation** | [[path: Home > Orientation]] | Angles or turns text, handy for narrow column labels |
| **Increase Indent** | [[path: Home > Increase Indent]] | Moves text in from the left, to show items belong to a group |

::: callout warning
Merged cells inside a block of data break sorting, filtering and screen readers. Merge titles, never data.
:::

::: drill D1.09
title: Merge the title, wrap and angle labels, indent rows
badge: Alignment Ace
starter: starters/week-01/E1-L3-report.xlsx
platform: web
mos: 2.2.1, 2.2.2, 2.2.4
other: N7
minutes: 6

- step: Merge and center the title across **A1:F1**.
  check: cell.is sheet=June range=A1:F1 merged=true align=center
- step: Turn on **Wrap Text** for the header row (A3:F3), so the long headers fit.
  check: cell.is sheet=June range=A3:F3 wrap=true
- step: Angle the week labels in **B13:E13** counterclockwise.
  check: cell.is sheet=June range=B13:E13 rotation=45
  hint: [[path: Home > Orientation > Angle Counterclockwise]].
- step: Indent the four boat names under their categories (A5, A6, A8 and A10) one level.
  check: cell.is sheet=June range=A5:A6 indent=any
  check: cell.is sheet=June cell=A8 indent=any
  check: cell.is sheet=June cell=A10 indent=any
  hint: Hold [[key: Ctrl / Cmd]] to select all four, then [[path: Home > Increase Indent]].
:::

## Cell styles, Format Painter and Clear

- **Cell styles** ([[path: Home > Cell Styles]]) apply a ready-made look in one click: **Title**, **Heading 1** to **4**, **Total**, **Good**, **Bad**, **Input** and more. They follow the workbook's theme.
- **Format Painter** ([[path: Home > Format Painter]]) copies the look of one cell onto others: select the formatted cell, select Format Painter, then drag across the cells to change. Double-click it to paint several places; press [[key: Esc]] to stop.
- **Clear Formats** ([[path: Home > Clear > Clear Formats]]) takes the formatting off and leaves the contents.

::: drill D1.10
title: Cell styles, Format Painter and Clear Formats
badge: Style Setter
starter: starters/week-01/E1-L3-report.xlsx
platform: web
mos: 2.2.3, 2.2.7, 2.2.8
minutes: 6

Do D1.08 and D1.09 first: this drill copies the header format you made there.

- step: Apply the **Title** cell style to the report title (A1).
  check: cell.is sheet=June cell=A1 style=Title
- step: Apply the **Total** cell style to the Total row (A11:F11).
  check: cell.is sheet=June range=A11:F11 style=Total
- step: Use **Format Painter** to copy the June header row's format (A3:F3) onto the July sheet's header row.
  check: cell.is sheet=July range=A3:F3 sameAs=June!A3
  check: cell.is sheet=July range=A3:F3 fill=any wrap=true
  hint: Select June's A3:F3, select **Format Painter**, switch to the July sheet, then drag across A3:F3.
- step: The **Notes** area (H3:H6) has leftover bold, italic, red and yellow formatting. Clear the formats but keep the notes.
  check: cell.is sheet=June range=H3:H6 clean=true
  check: cell.is sheet=June cell=H3 value=Notes
:::

::: quiz Q13.1
question: Cell C5 shows $1,820.00. What does Excel use when C5 is in a formula?
type: single
- [x] 1820
- [ ] "$1,820.00" as text
- [ ] 1820.00 rounded to whole dollars
- [ ] Whatever the formula bar shows with the $ sign
explain: Number formats change only how a value looks. Formulas always use the value itself.
:::

## Check your understanding

- Why shouldn't you merge cells inside a table of data?
- What's the difference between Clear Formats and pressing Delete?
- When is Accounting format better than Currency?
