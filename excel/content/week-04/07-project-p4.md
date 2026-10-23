---
type: project
id: XL4-P4
week: 4
order: 7
title: "Project 4: Business Plan Workbook"
objectives: [W4.1, W4.2, W4.3, W4.4, W4.5]
clos: [CLO3, CLO4, CLO5]
minutes: 180
status: draft
rubricBadge: Business Planner
---

::: callout why
A real plan answers "what if?" before anyone spends money. This project puts everything from this week together: a chart that makes the case, a graphic that makes it look finished, and what-if tools that test the numbers.
:::

## What you'll make

One workbook, `Lastname_P4_BusinessPlan.xlsx`, for a small business or club idea of your own (a food truck, a tutoring service, a club fundraiser), with:

1. **An Inputs sheet** with the numbers you can change (price, units, costs, a loan), formatted with a cell style you created.
2. **A Plan sheet** that works out the profit (or the amount raised) with formulas, and a loan payment with `PMT`.
3. **What-if analysis:**
   - **Goal Seek** to find the break-even (or target) price, with a sentence saying what you found
   - **three scenarios** (for example Low, Expected and High) and a **scenario summary** sheet
   - a **one-variable data table** of the profit or payment for at least five values of one input
4. **A chart** that shows the plan's main point, with a clear title, axis titles where they help, and alt text that states the point.
5. **A graphic**: your logo, a picture or icon, or a SmartArt process, with alt text or marked decorative.
6. A **theme** other than Office, real sheet names, and a **Title** property.

::: callout privacy
Use made-up numbers, or public ones (prices from a store's website). Don't put anyone's real financial details in the workbook.
:::

::: callout ai
Optional: ask an AI tool what costs a business like yours might have, then decide which ones to include and estimate them yourself. Write one sentence on the Plan sheet about what you kept and what you dropped.
:::

## Check before you submit

::: preflight P4-check
title: Preflight for your business plan
starter: none
platform: desktop

- step: The workbook has at least one chart.
  check: chart.count min=1
- step: There are at least three scenarios.
  check: scenario.is min=3
- step: The workbook has a one-variable data table.
  check: datatable.is
- step: Charts and graphics have alt text, sheets have real names, and the workbook has a Title.
  check: access.is rules="alt,sheetNames,title"
- step: The workbook uses a theme other than Office.
  check: theme.is changed=true
:::

## How it's graded

::: rubric P4
- P4-a | 25 | CLO3 | Goal Seek, the scenarios with a summary, and the data table are set up correctly and answer real questions about the plan.
- P4-b | 25 | CLO5 | The chart fits the data, shows the main point, and has a title, useful elements and alt text that states the point.
- P4-c | 15 | CLO4 | The graphic is placed and sized well and has alt text or is marked decorative.
- P4-d | 20 | CLO3 | The formulas, including `PMT`, are correct and use cell references, so the what-if tools work.
- P4-e | 15 | CLO1, CLO2 | The workbook uses a theme and a custom cell style, has real sheet names and a Title, and the written findings are accurate.
:::
