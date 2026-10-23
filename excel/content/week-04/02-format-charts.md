---
type: lesson
id: XL4-L2
week: 4
order: 2
title: Format Charts
objectives: [W4.2]
clos: [CLO5]
minutes: 35
status: draft
---

::: callout why
A chart without a title or labels makes readers guess. A few chart elements, a clean style and good alternative text make it clear to everyone, including people who can't see it.
:::

## Chart elements

Select the chart, then the **+** button beside it (**Chart Elements**), or [[path: Chart Design > Add Chart Element]] (desktop):

::: platform web
In Excel for the web there's no **+** button. Double-click the chart (or select it, then [[path: Chart > Format]]) to open the **Format** pane, where you turn titles, axis titles, data labels and the legend on and off, and set the legend's position.
:::

| Element | Use it to |
|---|---|
| **Chart Title** | Say what the chart shows, or better, what it means ("Solo kayaks lead every month") |
| **Axis Titles** | Name the units ("Rentals", "Month") |
| **Data Labels** | Show each bar's number on the bar |
| **Legend** | Say which color is which series. Move it with **Legend > Bottom** |
| **Data Table** | Show the numbers in a small table under the chart |

## Layouts and styles

- **Quick Layout** ([[path: Chart Design > Quick Layout]], desktop) applies a set of elements in one click. **Layout 5**, for example, adds a data table and an axis title.
- **Chart Styles** ([[path: Chart Design > Chart Styles]]) change colors, backgrounds and effects together. In the desktop app, the brush button beside the chart does the same.

## Alternative text

Right-click the chart and choose **Edit Alt Text** (or **View Alt Text**). Write what the chart **shows**, in a sentence or two: the point a sighted reader would take away, not "a column chart".

::: callout tip
Good alt text for a chart: "Solo kayaks were rented most in every month, peaking at 66 in July; paddleboards were rented least." Long data belongs in the worksheet, which screen readers can read.
:::

::: video
id: e4-l2-format
url: TODO
title: Titles, labels, layouts and alt text for charts
captions: none
duration: 4:00
:::

::: drill D4.04
title: Add titles, labels and move the legend
badge: Chart Labeler
starter: starters/week-04/E4-L1-charts.xlsx
platform: web
mos: 5.2.3
minutes: 6

Do the Lesson 1 drills (D4.01 to D4.03) first: this drill works on the same chart.

- step: Give the chart the title **Rentals by boat, May to August**.
  check: chart.is title="Rentals by boat"
- step: Add a title to **each axis** (for example **Month** and **Rentals**).
  check: chart.is axisTitles=true
- step: Add **data labels** to the bars.
  check: chart.is dataLabels=true
- step: Move the legend to the **bottom**.
  check: chart.is legend=bottom
:::

::: drill D4.05
title: Apply a layout and a style, and add alt text
badge: Chart Stylist
starter: starters/week-04/E4-L1-charts.xlsx
platform: desktop
mos: 5.3.1, 5.3.2, 5.3.3
minutes: 6

- step: Apply **Quick Layout 5**, which adds a data table under the chart.
  check: chart.is dataTable=true
  hint: [[path: Chart Design > Quick Layout > Layout 5]]. If it changes your title, retype it.
- step: Apply a different **chart style**.
  check: chart.is styleChanged=true
- step: Add alternative text that states the chart's main point.
  check: chart.is alt=true
:::

::: quiz Q42.1
question: Which is the best alt text for this chart?
type: single
- [ ] Chart 1
- [ ] A clustered column chart with four series and four categories
- [x] Solo kayaks were rented most every month, peaking at 66 in July; paddleboards were rented least.
- [ ] Rentals
explain: Alt text should give the takeaway a sighted reader gets at a glance. The chart type and series count don't tell anyone what the data says.
:::

## Check your understanding

- What's the difference between a quick layout and a chart style?
- When would data labels clutter a chart?
- Where do the full numbers live for a screen reader user?
