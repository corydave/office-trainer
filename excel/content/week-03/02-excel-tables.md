---
type: lesson
id: XL3-L2
week: 3
order: 2
title: Excel Tables
objectives: [W3.2, W3.3]
clos: [CLO1, CLO2]
minutes: 50
status: draft
---

::: callout why
A plain range is just cells. An **Excel table** knows it's a table: it grows when you add a row, keeps its formatting banded, adds filter buttons, totals itself, and lets formulas use column names like `[@Hours]` instead of addresses like E2.
:::

## Making a table

Click anywhere in the data and press [[key: Ctrl+T / Cmd+T]] (in Excel for the web, [[key: Ctrl+L]], because Ctrl+T opens a new browser tab), or use [[path: Insert > Table]]. Check the range Excel picked and that **My table has headers** is ticked.

::: callout warning
Excel picks the range by looking for the first empty row or column. If your data has an empty row in the middle, select the whole range yourself first.
:::

Then, on the **Table Design** tab:

| To | Do this |
|---|---|
| Give it a name formulas can use | Type in **Table Name** (no spaces, like **Rentals**) |
| Change its look | Pick a style in **Table Styles** (Light, Medium or Dark) |
| Show or hide bands, the first column and more | Tick the boxes in **Table Style Options** |
| Add totals | Tick **Total Row**, then choose Sum, Average, Count and so on from each total cell's list |
| Make it a normal range again | **Convert to Range** (Table Tools in some versions) |

To add a column, type a header in the empty cell just right of the table. To add a row, type just below it, or press [[key: Tab]] in the last cell. To delete rows, select them, right-click and choose **Delete > Table Rows**.

::: video
id: e3-l2-tables
url: TODO
title: Create, style and total a table
captions: none
duration: 4:00
:::

::: drill D3.03
title: Turn the rental list into a table
badge: Table Maker
starter: starters/week-03/E3-L2-tables.xlsx
platform: web
mos: 3.1.1, 3.1.2
minutes: 5

- step: The list on **Rentals** has one empty row in the middle. Select the whole list (**A1:G34**) and make it a table with headers.
  check: table.is contains="Rental #" minRows=32
  hint: Select A1, then [[key: Ctrl+Shift+End / Cmd+Shift+End]] to select to the last cell, then [[key: Ctrl+T / Cmd+T]] ([[key: Ctrl+L]] in Excel for the web).
- step: Name the table **Rentals**.
  check: table.is name=Rentals
- step: Apply any **Medium** table style.
  check: table.is name=Rentals style=medium
:::

::: drill D3.04
title: Style options, a total row and table rows
badge: Table Tuner
starter: starters/week-03/E3-L2-tables.xlsx
platform: web
mos: 3.2.1, 3.2.2, 3.2.3
minutes: 7

Do D3.03 first.

- step: Turn on **Banded Columns**.
  check: table.is name=Rentals bandedColumns=true
- step: Turn on the **Total Row**. Make it total the **Hours** (Sum) and average the **Rate** (Average).
  check: table.is name=Rentals totalRow=true total=Hours totalFunction=sum
  check: table.is name=Rentals total=Rate totalFunction=average
- step: Add a column named **Deposit** to the right of **Customer**.
  check: table.is name=Rentals hasColumn=Deposit
  hint: Type Deposit in the empty cell right of the Customer header. The table grows to include it.
- step: Delete the table's empty row.
  check: table.is name=Rentals noBlankRows=true
  hint: Right-click a cell in the empty row, then [[path: Delete > Table Rows]].
:::

## Structured references

In a table, formulas can use column names. `=[@Hours]*[@Rate]` means "this row's Hours times this row's Rate". Type it once in a new column and the table fills the whole column for you. Outside the table, `=SUM(Rentals[Hours])` totals a whole column, however many rows it has.

::: drill D3.05
title: Calculate with column names
badge: Structured Thinker
starter: starters/week-03/E3-L2-tables.xlsx
platform: web
mos: 4.1.2
minutes: 4

Do D3.03 and D3.04 first.

- step: Add a column named **Cost** after **Deposit**, and calculate each rental's cost (Hours times Rate) with a structured reference.
  check: formula.is sheet=Rentals range=I2:I33 structured=true contains="Hours,Rate" filled=true
  check: formula.is sheet=Rentals cell=I2 result=105
  hint: Type Cost as the header, then in the first cell type =, click the Hours cell, type `*` and click the Rate cell. Excel writes `=[@Hours]*[@Rate]` and fills the column.
:::

::: drill D3.06
title: Convert a table back to a range
badge: Range Restorer
starter: starters/week-03/E3-L2-tables.xlsx
platform: web
mos: 3.1.3
minutes: 2

- step: On **Prices**, the old **OldPrices** table isn't needed as a table any more. Convert it to a normal range, keeping the data.
  check: table.is name=OldPrices exists=false
  check: text.present sheet=Prices value=Paddleboard
  hint: Click in it, then [[path: Table Design > Convert to Range]] and **Yes**.
:::

::: quiz Q32.1
question: A table named Rentals has a Hours column. Which formula keeps working when you add ten more rentals?
type: single
- [ ] `=SUM(E2:E33)`
- [x] `=SUM(Rentals[Hours])`
- [ ] `=SUM(E2:E43)` typed after you add them
- [ ] `=E2+E3+E4`
explain: A structured reference means "the whole Hours column", so it includes new rows automatically.
:::

## Check your understanding

- Name three things a table does that a plain range doesn't.
- Why give a table a real name instead of Table1?
- What does the @ in `[@Hours]` mean?
