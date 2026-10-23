---
type: lesson
id: W3-L2
week: 3
order: 2
title: Tables
objectives: [W3.3]
clos: [CLO4]
minutes: 45
status: draft
---

::: callout why
Schedules, price lists and sign-up sheets are tables. A real table keeps columns lined up no matter what you type, sorts itself, and tells a screen reader which heading goes with each cell. Columns lined up with spaces or tabs do none of that.
:::

## Make a table

| To | Do this |
|---|---|
| Insert a new table | [[path: Insert > Table]], then drag across the grid, or choose **Insert Table** and type the rows and columns |
| Turn tab-separated lines into a table | Select the lines, then [[path: Insert > Table > Convert Text to Table]] (desktop Word) |
| Turn a table back into lines of text | Click in the table, then [[path: Table Layout > Convert to Text]] (desktop Word) |
| Add a row at the end | Click in the last cell and press [[key: Tab]] |

::: video
id: w3-l2-tables
url: TODO
title: Tables from scratch and from text
captions: none
duration: 3:30
:::

::: drill D3.08
title: Convert text to a table, and a table to text
badge: Table Converter
starter: starters/week-03/W3-L2-schedule.docx
platform: desktop
mos: 3.1.1, 3.1.2
minutes: 5

- step: Turn the seven tab-separated schedule lines (from **Time Activity Where** to **Last boats in**) into a table with 3 columns.
  check: table.is contains="Safety talk" cols=3 minRows=7
  hint: Select the seven lines, then [[path: Insert > Table > Convert Text to Table]]. Separate text at **Tabs**.
- step: The **Old price list** table isn't needed as a table any more. Convert it to text, separated by tabs.
  check: table.count contains="Kayak (half day)" max=0
  check: text.present value="Kayak (half day)"
  hint: Click in the table, then [[path: Table Layout > Convert to Text]], and choose **Tabs**.
:::

## Change a table

Click in a table to see the **Table Design** and **Table Layout** tabs.

| Task | Where |
|---|---|
| **Merge** cells into one (a title across the top) | Select them, then [[path: Table Layout > Merge Cells]] |
| **Split** one cell into several | [[path: Table Layout > Split Cells]] |
| **Resize** a column to an exact width | Click in it, then set **Width** in [[path: Table Layout > Cell Size]] |
| **Sort** rows by a column | [[path: Table Layout > Sort]] (desktop Word); choose **Header row** if the table has one |
| **Cell margins** (space inside each cell) | [[path: Table Layout > Cell Margins]] |
| **Repeat header rows** on every page | Select the top row(s), then [[path: Table Layout > Repeat Header Rows]] |
| **Split** the table into two | Click in the row that should start the new table, then [[path: Table Layout > Split Table]] |

::: callout note
**Repeat Header Rows** does two jobs: the headings show at the top of every page the table runs onto, and screen readers know those cells are column headings. Accessibility checkers look for it.
:::

::: drill D3.09
title: Build and fix the tour table
badge: Table Tamer
starter: starters/week-03/W3-L2-schedule.docx
platform: desktop
mos: 3.1.3, 3.2.1, 3.2.2, 3.2.3, 3.2.4, 3.2.5, 3.2.6
other: N14
minutes: 12

- step: Under **Sign-up sheet**, insert a new table with **4 columns** and **6 rows**.
  check: table.is any=true cols=4 rows=6
  hint: [[path: Insert > Table > Insert Table]], then type 4 and 6.
- step: In the **Summer tours** table, merge the four cells of the top row into one title cell.
  check: table.is contains="Summer tours 2027" merged=true
  hint: Select the whole top row, then [[path: Table Layout > Merge Cells]].
- step: Make the **Date** column exactly **1.5 in** wide.
  check: table.is contains="Summer tours 2027" col=1 width=1.5in
  hint: Click in the Date column, then set **Width** to 1.5" in [[path: Table Layout > Cell Size]].
- step: Sort the tours by **Date**, earliest first.
  check: table.is contains="Summer tours 2027" sortedBy=1
  hint: Select the rows from the first date to the last (not the title or the column labels), then [[path: Table Layout > Sort]], **Sort by Column 1**, **Type: Date**, **Ascending**.
- step: Set the table's left and right cell margins to **0.1 in**.
  check: table.is contains="Summer tours 2027" cellMargin=0.1in
  hint: [[path: Table Layout > Cell Margins]], then Left 0.1" and Right 0.1".
- step: Make the top two rows (the title and the column labels) repeat as header rows.
  check: table.is contains="Summer tours 2027" headerRepeat=true
  hint: Select both rows, then [[path: Table Layout > Repeat Header Rows]].
- step: Split the table so the August tours start their own table.
  check: table.is contains="August" lacks="June" minRows=3
  hint: Click in the first August row, then [[path: Table Layout > Split Table]].
:::

::: callout tip
This drill merges a title row because Word's exam asks for merging. In your own tables, put the title in a heading or caption **above** the table instead. Screen readers treat every repeating header row as column headings, so a merged title row gets read out as if it were one.
:::

::: quiz Q32.1
question: Your table runs onto a second page, and readers can't tell which column is which. What's the fix?
type: single
- [ ] Copy the heading row onto the second page by hand
- [x] Select the heading row and turn on Repeat Header Rows
- [ ] Make the table smaller so it fits on one page
- [ ] Merge the heading cells
explain: Repeat Header Rows shows the headings on every page automatically, and it marks them as headings for screen readers. A copied row would end up in the middle of a page as soon as the table changes.
:::

## Check your understanding

- Why is a table better than tabs for a sign-up sheet?
- What should you select before sorting a table with a title row?
- When would you split a table instead of starting a new one?
