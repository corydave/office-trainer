---
type: skillcheck
id: XL3-SC3
week: 3
order: 5
title: "Skill Check 3: Tables and Data"
objectives: [W3.1, W3.2, W3.3, W3.4]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

One starter workbook, a set of tasks and a time limit, with no hints. Each task is worth one point. Download the deliveries file (next to the practice file) before you start.

::: skillcheck SC3
title: Gear deliveries
badge: Skill Check 3
starter: starters/week-03/E3-SC3-gear.xlsx
files: starters/week-03/E3-SC3-deliveries.csv
platform: web
mos: 1.1.1, 2.4.1, 2.4.2, 2.4.3, 3.1.1, 3.1.2, 3.1.3, 3.2.1, 3.2.2, 3.2.3, 3.3.1, 3.3.2, 4.1.2, 4.2.4, 4.2.5
timeLimit: 30
attempts: 2
passing: 70

- task: Import **E3-SC3-deliveries.csv** into the workbook with **Get Data**, so it loads as a table.
  platform: desktop
  check: connection.count min=1
  check: data.is table=* headers="Delivery,Item #,Quantity" rows=5
- task: On **Gear**, turn the list (A1:E13) into a table named **Gear**.
  check: table.is name=Gear minRows=12
- task: Apply the table style **Medium 4** (Table Style Medium 4).
  check: table.is name=Gear style=TableStyleMedium4
- task: Turn on the **First Column** style option.
  check: table.is name=Gear firstColumn=true
- task: Add a total row that sums **On hand**.
  check: table.is name=Gear totalRow=true total="On hand" totalFunction=sum
- task: Add a **Value** column that multiplies **On hand** by **Cost** with a structured reference.
  check: table.is name=Gear hasColumn=Value
  check: formula.is sheet=Gear range=F2:F13 structured=true contains="On hand,Cost"
- task: Sort the table by **Category** (A to Z), then by **On hand** (largest first).
  check: sort.is table=Gear by=Category then="On hand" thenOrder=desc
- task: On **Lists**, in A2, list each category once with `UNIQUE`.
  check: formula.is sheet=Lists range=A2:A4 uses=UNIQUE spill=true
  check: cell.is sheet=Lists cell=A5 blank=true
- task: On **Lists**, in C2, list every cost from highest to lowest with `SORT`.
  check: formula.is sheet=Lists range=C2:C13 uses=SORT spill=true
  check: formula.is sheet=Lists cell=C2 result=89
- task: Highlight **On hand** values (D2:D13) **less than 15**.
  check: cf.is sheet=Gear range=D2:D13 type=cellIs operator=lessThan value=15
- task: Remove the data bars from the **Cost** column (E2:E13).
  check: cf.count sheet=Gear range=E2:E13 max=0
- task: On **Trend**, insert **column** sparklines in F2:F4 for each item's April to July numbers.
  platform: desktop
  check: sparkline.is sheet=Trend location=F2:F4 type=column min=3
- task: On **Old**, convert the **OldOrder** table to a normal range.
  check: table.is name=OldOrder exists=false
- task: Filter the table to show only the **Safety** category. (Do this last.)
  check: filter.is table=Gear column=Category shows=Safety
:::
