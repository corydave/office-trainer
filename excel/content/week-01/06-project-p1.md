---
type: project
id: XL1-P1
week: 1
order: 6
title: "Project 1: Monthly Budget Workbook"
objectives: [W1.2, W1.3, W1.4, W1.5]
clos: [CLO1, CLO2]
minutes: 180
status: draft
rubricBadge: Budget Builder
---

::: callout why
A budget is the spreadsheet almost everyone needs at some point: for a household, a semester, a club or a small business. Building one that's clear and print-ready uses every skill from this week.
:::

## What you'll make

A one-month budget workbook, submitted as **.xlsx** and **.pdf**:

1. **A budget sheet** with an income section and an expense section. Each line has a category, a planned amount and an actual amount. (You'll add the formulas that total them in Week 2; this week, type the totals or leave them for later.)
2. **A second sheet** with notes, or a second month in the same layout.

Pick a budget that matters to you: your own month, a household's, a club's event, or a small business's. Name the files `Lastname_P1_Budget.xlsx` and `Lastname_P1_Budget.pdf`.

::: callout privacy
Made-up or rounded numbers are completely fine; the skills are what's graded. Never include account numbers, passwords or other people's private details.
:::

## Required features

- Sheets with real names (not Sheet1) and the budget sheet's tab colored
- A merged and centered title (above the table, never inside it), with the **Title** cell style
- A header row with wrap text, a border and a fill
- Money in **Currency** or **Accounting** format; any percentages as **Percentage**
- Categories indented under their section labels
- Columns wide enough to read (no ####), and the header row frozen
- Page setup: orientation that fits, **1 page wide**, a print area, a footer with the sheet name and page number
- Workbook properties: **Title** and **Tags**
- A PDF that opens and matches the workbook

::: callout ai
Optional: ask an AI tool to suggest budget categories for your situation, then choose the ones that fit. Never paste real account details or balances into an AI tool.
:::

## Check before you submit

Run these automatic checks on your workbook. They don't grade anything; they catch the problems that cost the most points.

::: preflight P1-check
title: Preflight for your budget
starter: none
platform: web

- step: Every sheet has a real name, and the workbook has a Title property.
  check: access.is rules="sheetNames,title"
- step: The budget sheet has a print area.
  platform: desktop
  check: page.is printArea=any
- step: The budget sheet has a page number in its footer.
  platform: desktop
  check: header.is part=footer field=page
:::

## How it's graded

::: rubric P1
- P1-a | 25 | CLO1 | The budget is organized: clear sections, labeled rows and columns, sensible sheet names, no blank spacer rows inside the data.
- P1-b | 25 | CLO1, CLO2 | Formatting uses Excel tools on purpose: number formats, cell styles, alignment, column widths (no spaces or merged cells as workarounds).
- P1-c | 20 | CLO2 | Navigation and views help the reader: frozen headings and a second sheet for notes or a second month.
- P1-d | 30 | CLO1 | Print-ready: page setup, print area, footer and properties are set; the PDF matches the workbook.
:::

This project is also the evidence for MO-210 objective 1.5.2 (saving and exporting in alternative formats).
