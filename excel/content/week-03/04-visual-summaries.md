---
type: lesson
id: XL3-L4
week: 3
order: 4
title: Conditional Formatting and Sparklines
objectives: [W3.4]
clos: [CLO1, CLO2]
minutes: 40
status: draft
---

::: callout why
Thirty-two rows of numbers hide their pattern. Conditional formatting colors the cells that matter, and sparklines draw a tiny trend next to each row, so the story jumps out before anyone reads a number.
:::

## Conditional formatting

[[path: Home > Conditional Formatting]] formats cells **based on their values**, and updates as the values change.

| Rule type | Example |
|---|---|
| **Highlight Cells Rules** | Costs greater than $100 in light red |
| **Top/Bottom Rules** | The top 10% of hours |
| **Data Bars** | A bar in each cell, as long as its value |
| **Color Scales** | Low values red, high values green |
| **Icon Sets** | Arrows or flags by value |

To remove rules, select the cells, then [[path: Home > Conditional Formatting > Clear Rules > Clear Rules from Selected Cells]]. **Manage Rules** shows every rule on the sheet.

::: callout note
Color alone doesn't reach everyone: some readers can't tell red from green, and screen readers don't announce fill colors. Pair color with something else, like data bars, icons or a label, and say what the highlight means.
:::

::: video
id: e3-l4-visual
url: TODO
title: Conditional formatting and sparklines
captions: none
duration: 3:30
:::

::: drill D3.09
title: Data bars, a highlight rule, and an old rule to remove
badge: Pattern Spotter
starter: starters/week-03/E3-L4-visual.xlsx
platform: web
mos: 2.4.2, 2.4.3
minutes: 6

- step: Add **Data Bars** to the **Hours** column (E2:E33).
  check: cf.is sheet=Rentals range=E2:E33 type=dataBar
- step: Highlight **Cost** values (H2:H33) **greater than 100**.
  check: cf.is sheet=Rentals range=H2:H33 type=cellIs operator=greaterThan value=100
  hint: [[path: Home > Conditional Formatting > Highlight Cells Rules > Greater Than]], type 100, pick a format, **OK**.
- step: The **Rate** column (F2:F33) still has a color scale from an old report. Remove it.
  check: cf.count sheet=Rentals range=F2:F33 max=0
:::

## Sparklines

A **sparkline** is a chart that fits in one cell. Select the data, then [[path: Insert > Sparklines > Line]] (or **Column** or **Win/Loss**), and choose the cells to put them in. On the **Sparkline** tab, tick **High Point**, **Low Point** or **Markers** to mark them.

::: drill D3.10
title: Trend lines for each boat
badge: Trend Spotter
starter: starters/week-03/E3-L4-visual.xlsx
platform: desktop
mos: 2.4.1
minutes: 4

- step: On **Trend**, insert **line** sparklines in F3:F6 showing each boat's May to August numbers (B3:E6).
  check: sparkline.is sheet=Trend location=F3:F6 type=line min=4
  hint: Select B3:E6, then [[path: Insert > Sparklines > Line]]; Location Range is F3:F6.
- step: Mark each sparkline's **high point**.
  check: sparkline.is sheet=Trend location=F3:F6 high=true
:::

::: quiz Q34.1
question: You highlight costs over $100 in red. Later a cost changes from $90 to $120. What happens?
type: single
- [ ] Nothing: you have to apply the rule again
- [x] The cell turns red, because the rule checks the value whenever it changes
- [ ] The rule is deleted
explain: Conditional formatting rules keep checking the values, so the formatting always matches the data.
:::

## Check your understanding

- Why shouldn't color be the only way a highlight carries meaning?
- How do you find a rule you can't see on screen?
- When is a sparkline better than a full chart?
