---
type: project
id: XL2-P2
week: 2
order: 6
title: "Project 2: Invoice and Grade Calculator"
objectives: [W2.1, W2.2, W2.3, W2.4, W2.5]
clos: [CLO1, CLO2]
minutes: 160
status: draft
rubricBadge: Calculator Builder
---

::: callout why
An invoice and a grade calculator are two of the most common spreadsheets in work and school. Built well, you fill in the inputs and every other number takes care of itself.
:::

## What you'll make

One workbook, `Lastname_P2_Calculators.xlsx`, with two sheets:

1. **Invoice.** For a real or made-up small business (or a club event):
   - At least five line items with quantity and price, and a line-total formula filled down
   - A subtotal with `SUM`, tax from a **named** tax-rate cell, and a total
   - An invoice number built with text functions (for example `=CONCAT("INV-",UPPER(LEFT(B2,3)),"-",B3)` from a customer name in B2 and an order number in B3)
   - The invoice date. `=TODAY()` changes every day the file is opened, so for a real invoice type the date, or press [[key: Ctrl+;]] to enter today's date as a fixed value
2. **Grade calculator.** For a real course of yours (or a made-up one):
   - Categories (homework, quizzes, projects, exams) with their weights in one input area
   - Scores for each category, a weighted total that uses **absolute** references or names for the weights
   - `AVERAGE`, `MAX` and `MIN` of the scores, and a `COUNT` of graded items
   - An `IF` that shows **On track** or **Needs attention** based on the total

::: callout privacy
Use made-up names and scores. Never put a classmate's real grades into a practice file or an AI tool.
:::

## Required features

- Every calculated number is a formula; inputs are clearly marked (the **Input** cell style is a good way)
- At least one absolute or mixed reference and at least one named range, each used in a formula
- `SUM`, `AVERAGE`, `MAX` or `MIN`, a `COUNT` function, `IF`, and at least two text functions
- Formats and page setup from Week 1, so both sheets print well
- No #DIV/0! or other errors when the inputs are filled in

::: callout ai
Optional: ask an AI tool to explain a formula you wrote, or to suggest one. Then test it by changing an input and checking the answer yourself. AI tools often get cell references wrong.
:::

## Check before you submit

::: preflight P2-check
title: Preflight for your calculators
starter: none
platform: web

- step: Every sheet has a real name, and the workbook has a Title property.
  check: access.is rules="sheetNames,title"
:::

Then check by hand: change one quantity and one score. Do the totals change? If a number stays the same, it's typed, not calculated.

## How it's graded

::: rubric P2
- P2-a | 25 | CLO2 | Every calculated cell is a formula; nothing is typed that Excel should calculate.
- P2-b | 25 | CLO2 | References are right: absolute or named for fixed values, relative for filled formulas; the sheets still work when inputs change.
- P2-c | 25 | CLO2 | Functions fit the job: SUM, AVERAGE, COUNT family, IF and text functions used correctly.
- P2-d | 25 | CLO1 | The workbook is readable and printable: labels, number formats, a marked input area and page setup.
:::
