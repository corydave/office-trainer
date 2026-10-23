---
type: lesson
id: PP3-L3
week: 3
order: 3
title: Tables and Charts
objectives: [W3.3, W3.4]
clos: [CLO5, CLO2, CLO3]
minutes: 45
status: draft
---

::: callout why
A table lets people look up a number; a chart shows them what the numbers mean. Use a table for prices and schedules, and a chart when the point is a comparison or a trend.
:::

## Tables

- **Insert a table:** [[path: Insert > Table]], then drag across the grid to pick the number of columns and rows (or **Insert Table** to type them). The table icon in a content placeholder works too.
- **Add rows and columns:** click in a cell, then [[path: Layout > Insert Above]], **Insert Below**, **Insert Left** or **Insert Right**. Pressing [[key: Tab]] in the last cell adds a row.
- **Delete:** click in the row or column, then [[path: Layout > Delete]] > **Delete Rows** or **Delete Columns**.
- **Style:** [[path: Table Design > Table Styles]]. Keep **Header Row** ticked, so the first row is formatted (and read by screen readers) as headings.

::: drill D3.05
title: Build the price table
badge: Table Setter
starter: starters/week-03/P3-L3-data.pptx
platform: web
mos: 4.1.1, 4.1.2, 4.1.3
minutes: 8

The prices to type:

| Boat | 1 hour | Half day | Full day |
|---|---|---|---|
| Kayak | $20 | $45 | $70 |
| Canoe | $25 | $55 | $85 |

- step: On **Prices**, insert a table with **4 columns and 3 rows** and type the prices above.
  check: table.is title="Prices" has="Kayak"
  hint: [[path: Insert > Table]], drag across 4 columns and 3 rows, then type. [[key: Tab]] moves to the next cell.
- step: Add a row at the bottom: **Paddleboard**, **$18**, **$40**, **$60**.
  check: table.is title="Prices" rows=4 has="Paddleboard"
  hint: Click in the last cell and press [[key: Tab]].
- step: The shop stopped full-day rentals. Delete the **Full day** column.
  check: table.is title="Prices" cols=3 lacks="Full day"
  hint: Click in that column, then [[path: Layout > Delete > Delete Columns]].
- step: Apply a different table style, and keep the header row.
  check: table.is title="Prices" styleChanged=true headerRow=true
:::

## Charts

[[path: Insert > Chart]] picks a chart type, then opens a small worksheet with sample data. Type your own data over it (delete the sample rows you don't need) and close the worksheet. To change it later:

- **Change the type:** [[path: Chart Design > Change Chart Type]].
- **Edit the data:** [[path: Chart Design > Edit Data]].
- **Add a title, data labels or a legend:** the **+** (Chart Elements) button beside the chart (Windows), or [[path: Chart Design > Add Chart Element]].
- **Style:** [[path: Chart Design > Chart Styles]] and **Change Colors**.

| Chart | Best for |
|---|---|
| **Column** or **Bar** | Comparing amounts (bar is easier when the labels are long) |
| **Line** | A trend over time |
| **Pie** | Parts of one whole, with only a few parts |

::: drill D3.06
title: Chart the season
badge: Chart Maker
starter: starters/week-03/P3-L3-data.pptx
platform: desktop
mos: 4.2.1, 4.2.2
other: N7
minutes: 10

Rentals by month (the numbers are also in the slide's notes):

| Boat | June | July | August |
|---|---|---|---|
| Kayak | 120 | 180 | 150 |
| Canoe | 60 | 95 | 80 |
| Paddleboard | 90 | 140 | 110 |

- step: On **Season rentals**, insert a **Clustered Column** chart of these rentals.
  check: chart.is title="Season rentals" minSeries=3
  hint: [[path: Insert > Chart > Column > Clustered Column]]. In the worksheet, type the boats in column A and the months across row 1. Drag the blue outline so it covers only your data.
- step: Change it to a **Clustered Bar** chart.
  check: chart.is title="Season rentals" type=bar grouping=clustered
  hint: [[path: Chart Design > Change Chart Type > Bar]].
- step: Add the chart title **Rentals by month** and turn on **data labels**.
  check: chart.is title="Season rentals" chartTitle="Rentals by month" dataLabels=true
- step: Give the chart alt text that states its point (for example, which boat was rented most).
  check: chart.is title="Season rentals" alt=true
:::

::: quiz Q33.1
question: You want to show how kayak rentals rose and fell from May to September. Which chart fits best?
type: single
- [ ] Pie
- [x] Line
- [ ] A table
- [ ] Clustered bar
explain: A line chart shows change over time. A pie shows parts of one whole.
:::

## Check your understanding

- When is a table better than a chart?
- What does the Header Row option do for accessibility?
- Why say the chart's point in its alt text, not just "a bar chart"?
