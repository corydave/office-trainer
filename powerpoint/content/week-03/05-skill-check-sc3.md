---
type: skillcheck
id: PP3-SC3
week: 3
order: 5
title: "Skill Check 3: Visual Slides"
objectives: [W3.1, W3.2, W3.3, W3.4]
clos: [CLO4, CLO5]
minutes: 35
status: draft
---

One starter presentation, a set of tasks and a time limit, with no hints. Each task is worth one point. The chart and SmartArt tasks need desktop PowerPoint.

::: skillcheck SC3
title: Season report
badge: Skill Check 3
starter: starters/week-03/P3-SC3-report.pptx
platform: web
mos: 3.3.1, 3.3.2, 3.4.1, 3.4.3, 3.4.4, 3.4.5, 3.4.6, 3.4.7, 3.5.1, 3.5.2, 3.5.3, 4.1.1, 4.1.2, 4.1.3, 4.2.1, 4.2.2, 4.3.1, 4.3.3
timeLimit: 35
attempts: 2
passing: 70

- task: On **Our beach**, crop the photo to a square.
  check: object.is title="Our beach" kind=picture cropped=true square=true
- task: Apply a picture style to the same photo.
  check: object.is title="Our beach" kind=picture styled=true
- task: Give the photo alt text.
  check: object.is title="Our beach" kind=picture alt=true
- task: On **Highlights**, insert a **Right Arrow** shape with the text **Busiest day: July 4**.
  check: object.is title="Highlights" kind=shape says="Busiest day"
- task: Make the arrow exactly **4"** wide.
  check: object.is title="Highlights" says="Busiest day" width=4in
- task: Apply a shape style to the arrow.
  check: object.is title="Highlights" says="Busiest day" styleChanged=true
- task: The logo at the top right is behind a box. Bring it to the front.
  check: object.is title="Highlights" name=Logo front=true
- task: Align the three icons along their tops.
  check: align.is title="Highlights" kind=icon edge=top
- task: Group the three icons.
  check: object.is title="Highlights" kind=icon grouped=true
- task: On **Prices**, insert a table with 3 columns and 3 rows: **Boat**, **Hour**, **Day**; **Kayak**, **$20**, **$70**; **Canoe**, **$25**, **$85**.
  check: table.is title="Prices" cols=3 rows=3 has="Hour"
- task: Apply a table style other than the default.
  check: table.is title="Prices" styleChanged=true
- task: On **Rentals**, insert a **Clustered Column** chart of the rentals in the slide's notes.
  platform: desktop
  check: chart.is title="Rentals" type=column grouping=clustered minSeries=3
- task: Give the chart the title **Rentals by month**.
  platform: desktop
  check: chart.is title="Rentals" chartTitle="Rentals by month"
- task: On **Steps**, convert the list to any SmartArt graphic and add a fourth shape: **Paddle!**
  platform: desktop
  check: object.is title="Steps" kind=smartart items=4
:::
