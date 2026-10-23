---
type: lesson
id: XL4-L1
week: 4
order: 1
title: Create Charts
objectives: [W4.1]
clos: [CLO5]
minutes: 45
status: draft
---

::: callout why
A table of numbers asks readers to find the pattern themselves. A chart shows it: which boat is most popular, which month is busiest, where the season peaks. Picking the right chart and the right data is most of the work.
:::

## Your scenario

The owner of Lakeshore Paddle Company wants to show the season at the town's business meeting, and to know whether a new boat loan is affordable. This week you'll chart the season, add a graphic to a flyer, and answer what-if questions with Excel.

## Picking a chart type

| You want to show | Use |
|---|---|
| Amounts side by side (boats, months) | **Column** or **Bar** chart |
| A change over time | **Line** chart |
| Parts of one whole (only a few parts) | **Pie** chart |
| How two numbers relate | **Scatter** chart |

Select the data, **including** the headers, then [[path: Insert > Charts]] and pick a type, or [[path: Insert > Recommended Charts]] to see Excel's suggestions. [[key: Alt+F1]] inserts the default chart instantly.

::: video
id: e4-l1-charts
url: TODO
title: Insert a chart and move it to its own sheet
captions: none
duration: 3:30
:::

::: drill D4.01
title: Insert a clustered column chart
badge: Chart Maker
starter: starters/week-04/E4-L1-charts.xlsx
platform: web
mos: 5.1.1
other: N16
minutes: 4

- step: On **Season**, select the boats and the **May to July** numbers (A3:D7) and insert a **Clustered Column** chart.
  check: chart.is type=column grouping=clustered minSeries=3
  hint: Select A3:D7, then [[path: Insert > Insert Column or Bar Chart > Clustered Column]].
:::

## Chart sheets

A chart can sit on a worksheet next to its data, or have a whole sheet to itself, a **chart sheet**, which is handy for printing or presenting. Select the chart, then [[path: Chart Design > Move Chart]] and choose **New sheet**.

::: drill D4.02
title: Move the chart to its own sheet
badge: Chart Sheet Mover
starter: starters/week-04/E4-L1-charts.xlsx
platform: desktop
mos: 5.1.2
minutes: 2

Do D4.01 first.

- step: Move the chart to a new chart sheet named **Rentals Chart**.
  check: chart.is sheet="Rentals Chart" chartSheet=true
  hint: Select the chart, [[path: Chart Design > Move Chart]], choose **New sheet**, type Rentals Chart, **OK**.
:::

## Series, and switching rows and columns

Each set of bars (or each line) is a **data series**. Excel guesses whether your rows or your columns are the series:

- [[path: Chart Design > Select Data]] adds, removes or edits series. You can also drag the colored outline around the data on the sheet to include more.
- [[path: Chart Design > Switch Row/Column]] swaps them: boats along the bottom with a bar for each month, or months along the bottom with a bar for each boat.

::: drill D4.03
title: Add August and switch rows and columns
badge: Series Switcher
starter: starters/week-04/E4-L1-charts.xlsx
platform: web
mos: 5.2.1, 5.2.2
minutes: 5

Do D4.01 first.

- step: Add the **August** numbers (E3:E7) to the chart.
  check: chart.is covers=B4:E7
  hint: [[path: Chart Design > Select Data]], then **Add** (series name E3, values E4:E7). Or, if the chart is still on Season, drag the data outline's corner to include column E.
- step: Switch rows and columns, so the **months** run along the bottom and each **boat** is a series.
  check: chart.is plotBy=rows series=4
  hint: [[path: Chart Design > Switch Row/Column]].
:::

::: quiz Q41.1
question: You want to show how canoe rentals rose and fell from May to August. Which chart fits best?
type: single
- [ ] Pie chart
- [x] Line chart
- [ ] Scatter chart
- [ ] Doughnut chart
explain: A line chart shows change over time. A pie chart shows parts of one whole, which isn't what this question asks.
:::

## Check your understanding

- Why select the headers along with the numbers before you insert a chart?
- When is a chart sheet better than a chart on the worksheet?
- What does Switch Row/Column change, and what stays the same?
