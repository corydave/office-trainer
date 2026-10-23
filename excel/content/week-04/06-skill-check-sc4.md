---
type: skillcheck
id: XL4-SC4
week: 4
order: 6
title: "Skill Check 4: Charts and What-If"
objectives: [W4.1, W4.2, W4.4, W4.5]
clos: [CLO3, CLO5]
minutes: 30
status: draft
---

One starter workbook, a set of tasks and a time limit, with no hints. Each task is worth one point. Several tasks use the desktop app (chart sheets, quick layouts, Goal Seek and Scenario Manager).

::: skillcheck SC4
title: Season sales
badge: Skill Check 4
starter: starters/week-04/E4-SC4-sales.xlsx
platform: desktop
mos: 2.2.7, 5.1.1, 5.1.2, 5.2.1, 5.2.2, 5.2.3, 5.3.1, 5.3.2, 5.3.3
timeLimit: 30
attempts: 2
passing: 70

- task: On **Sales**, chart **Rentals** and **Lessons** for all four seasons (A3:E5) as a **Clustered Column** chart.
  check: chart.is type=column grouping=clustered minSeries=2
- task: Move the chart to its own chart sheet named **Sales Chart**.
  check: chart.is sheet="Sales Chart" chartSheet=true
- task: Add the **Gear** row (A6:E6) to the chart.
  check: chart.is covers=B4:E6
- task: Switch rows and columns, so each **season** is a series and the sources run along the bottom.
  check: chart.is plotBy=columns series=4
- task: Apply a quick layout that adds a **data table** under the chart.
  check: chart.is dataTable=true
- task: Give the chart the title **Sales by season**.
  check: chart.is title="Sales by season"
- task: Apply a different chart style.
  check: chart.is styleChanged=true
- task: Add alternative text to the chart that states its main point.
  check: chart.is alt=true
- task: On **Sales**, apply the **Title** cell style to A1.
  check: cell.is sheet=Sales cell=A1 style=Title
- task: On **Price**, use **Goal Seek** to find the price per lesson (B3) that makes the **Profit** (B7) exactly **$2,400**.
  check: cell.is sheet=Price cell=B3 number=32.5 tolerance=0.01
  check: formula.is sheet=Price cell=B7 result=2400
- task: On **Price**, add two scenarios that change **Lessons sold** (B4): **Busy** = 160 and **Slow** = 80.
  check: scenario.is sheet=Price names="Busy,Slow" changing=B4
:::
