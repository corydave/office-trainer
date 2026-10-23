---
type: skillcheck
id: PP1-SC1
week: 1
order: 5
title: "Skill Check 1: The Clear Presentation"
objectives: [W1.1, W1.2, W1.3, W1.4, W1.5]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

One starter presentation, a set of tasks and a time limit, with no hints. Each task is worth one point. Two tasks need desktop PowerPoint (the Title property and columns).

::: skillcheck SC1
title: Paddle club orientation
badge: Skill Check 1
starter: starters/week-01/P1-SC1-club.pptx
platform: web
mos: 1.2.1, 1.2.3, 2.1.3, 2.1.5, 2.2.1, 2.2.3, 2.3.1, 2.3.2, 2.3.3, 3.1.1, 3.1.2, 3.1.3
timeLimit: 30
attempts: 2
passing: 70

- task: Change the slide size to **Widescreen (16:9)**.
  check: size.is ratio=16:9
- task: Set the **Title** property to **Paddle Club orientation**.
  platform: desktop
  check: props.is title="Paddle Club orientation"
- task: Right after **2026 results**, add a **Two Content** slide titled **Club rules**.
  check: slide.is title="Club rules" layout="Two Content" after="2026 results"
- task: Duplicate the **Meeting times** slide.
  check: slide.count title="Meeting times" min=2
- task: Hide the **2026 results** slide.
  check: slide.is title="2026 results" hidden=true
- task: Add slide numbers to every slide except the title slide.
  check: footer.is slideNumbers=true skipTitle=true
- task: Rename the **Untitled Section** section to **Welcome**.
  check: section.is name="Welcome" index=1
- task: Add a section named **Packing** that starts with the **What to bring** slide.
  check: section.is name="Packing" starts="What to bring"
- task: Move the **Wrap-up** section to the end of the presentation.
  check: slide.is slide=last section="Wrap-up"
- task: Move the **Trip calendar** slide so it comes right after **Trip sign-up**.
  check: slide.is title="Trip calendar" after="Trip sign-up"
- task: Make the subtitle on the title slide italic.
  check: text.is slide=1 phType=subTitle italic=true
- task: Set the list on **What to bring** in two columns.
  platform: desktop
  check: text.is title="What to bring" says="Life jacket" columns=2
- task: Turn the steps on **Trip sign-up** into a numbered list.
  check: list.is title="Trip sign-up" kind=number
:::
