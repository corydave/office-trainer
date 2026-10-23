---
type: project
id: PP3-P3
week: 3
order: 6
title: "Project 3: Data Story"
objectives: [W3.1, W3.2, W3.3, W3.4]
clos: [CLO2, CLO3, CLO4, CLO5]
minutes: 150
status: draft
rubricBadge: Data Storyteller
---

::: callout why
Numbers persuade when people can see them. This project asks you to find one finding in real or realistic data and make it obvious, with a table, a chart, a diagram and pictures that all point the same way.
:::

## What you'll make

One presentation, `Lastname_P3_DataStory.pptx`, of **6 to 8 slides** that tells **one finding** from data: a club's attendance, a part-time job's busy hours (with permission), a sports season, your city's recycling numbers, or a public data set. It needs:

1. A title that **states the finding** ("Saturday mornings bring half our rentals"), not just the topic.
2. A **table** of the key numbers, with a header row and a table style.
3. A **chart** that makes the finding easy to see, with a chart title and labels where they help.
4. A **SmartArt** graphic (a process, a cycle or a list) that explains how the data was gathered or what to do next.
5. At least one **picture** or **icon**, resized without stretching and cropped where it helps.
6. Shapes, text boxes and icons **aligned** and **grouped** neatly.
7. **Alt text** on every picture, chart, SmartArt graphic and icon (or **decorative** where it only decorates).
8. The **source** of the data on the slide or in the notes.

::: callout privacy
Use public or made-up data. Don't show anyone's grades, health, pay or contact details, and use photos you took or have the right to use.
:::

::: callout ai
Optional: ask an AI tool which chart type fits your data. In the notes, say whether you agreed and why.
:::

## Check before you submit

::: preflight P3-check
title: Preflight for your data story
starter: none
platform: web

- step: The presentation has 6 to 8 slides.
  check: slide.count min=6 max=8
- step: There's a table with a header row.
  check: table.is slide=all headerRow=true
- step: There's a chart.
  check: object.count kind=chart min=1
- step: There's a SmartArt graphic.
  platform: desktop
  check: object.count kind=smartart min=1
- step: There's at least one picture or icon.
  check: object.count kind=image min=1
- step: Every graphic has alt text (or is decorative), and every slide has its own title.
  check: access.is rules="alt,titles,uniqueTitles,tableHeaders"
:::

## How it's graded

::: rubric P3
- P3-a | 30 | CLO5, CLO3 | The table and chart fit the data and support one clear finding.
- P3-b | 25 | CLO4 | Pictures, shapes and SmartArt explain, rather than decorate, and are aligned and grouped neatly.
- P3-c | 20 | CLO4 | Every graphic has useful alt text or is marked decorative.
- P3-d | 25 | CLO3, CLO2 | The data and sources are honest and cited on the slide or in the notes.
:::
