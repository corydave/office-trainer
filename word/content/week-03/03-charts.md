---
type: lesson
id: W3-L3
week: 3
order: 3
title: Charts and Captions
objectives: [W3.4]
clos: [CLO3]
minutes: 25
status: draft
---

::: callout why
"42 people chose kayaks and 9 chose pedal boats" is easy to read. Twenty numbers like that are not. A chart shows the pattern at a glance, and a caption tells readers what they're looking at, even when the chart sits pages away from the sentence that mentions it.
:::

## Insert a chart

[[path: Insert > Chart]] (desktop Word) opens a small spreadsheet next to a sample chart. Type or paste your numbers over the sample data, then close the spreadsheet. The chart updates as you type.

| Chart | Best for |
|---|---|
| **Column** or **Bar** | Comparing amounts across a few categories |
| **Line** | Change over time |
| **Pie** | Parts of one whole, with five slices or fewer |

**Chart Elements** (the + button beside a selected chart) adds or removes the title, labels and legend.

::: video
id: w3-l3-charts
url: TODO
title: A column chart from survey numbers
captions: none
duration: 3:00
:::

::: platform web
Word for the web shows and resizes charts made in desktop Word. Depending on your version it may also insert a basic chart ([[path: Insert > Chart]]), but editing a chart's data needs desktop Word, so use desktop Word for this lesson.
:::

## Captions

A **caption** is a numbered label: "Figure 1. Boat preferences from the Paddle Day survey." Use [[path: References > Insert Caption]]. Word numbers figures and tables for you and renumbers them when you add one in the middle.

::: callout note
A chart also needs alt text that gives the main point, such as "Column chart: kayaks were the top choice with 42 votes, then paddleboards with 31." The caption names the figure; the alt text says what it shows.
:::

::: drill D3.10
title: Chart the survey and caption it
badge: Chart Maker
starter: starters/week-03/W3-L3-survey.docx
platform: desktop
other: E3.4
minutes: 8

- step: On the empty line under **Insert the chart below this line**, insert a **Clustered Column** chart of the four boats and their votes.
  check: object.is kind=chart chartType=column
  hint: [[path: Insert > Chart > Column > Clustered Column]]. In the spreadsheet, type the four boats in column A and the votes in column B, delete the extra sample columns, then close it.
- step: Add a caption below the chart that reads **Figure 1. Boat preferences from the Paddle Day survey**.
  check: field.present code="SEQ Figure"
  check: style.is para="Figure 1" name="Caption"
  hint: Select the chart, then [[path: References > Insert Caption]]. Label **Figure**, Position **Below selected item**, and type the rest after "Figure 1".
- step: Give the chart alt text that states its main point.
  check: object.is kind=chart alt=true
  hint: Right-click the chart and choose **View Alt Text** (or **Edit Alt Text**).
:::

::: quiz Q33.1
question: You want to show how rentals changed month by month from May to October. Which chart fits best?
type: single
- [ ] Pie
- [x] Line
- [ ] Clustered column of each boat type
- [ ] A table is always better
explain: A line chart shows change over time. A pie chart shows parts of one whole and gets hard to read with more than about five slices.
:::

## Check your understanding

- What does a caption add that the chart's title doesn't?
- What should the alt text for a chart say?
