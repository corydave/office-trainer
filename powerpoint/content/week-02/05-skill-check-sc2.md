---
type: skillcheck
id: PP2-SC2
week: 2
order: 5
title: "Skill Check 2: Masters, Layouts and Backgrounds"
objectives: [W2.1, W2.2, W2.3]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

One starter presentation, a set of tasks and a time limit, with no hints. Each task is worth one point. Most tasks need desktop PowerPoint (Slide Master, Handout Master and Notes Master views).

::: skillcheck SC2
title: Company kickoff
badge: Skill Check 2
starter: starters/week-02/P2-SC2-kickoff.pptx
files: starters/week-02/logo.png, starters/week-02/lake.jpg
platform: desktop
mos: 1.1.1, 1.1.2, 1.1.3, 1.1.4, 1.1.5, 1.1.6, 2.2.2
timeLimit: 30
attempts: 2
passing: 70

- task: Change the slide master's background to any background style or fill.
  check: master.is background=any
- task: Set the slide master's title font to **Georgia**.
  check: master.is titleFont=Georgia
- task: Put **logo.png** on the slide master, so it appears on every slide.
  check: master.is picture=true
- task: Create a layout named **Team Photo** with a **Picture** placeholder.
  check: layout.is name="Team Photo" custom=true placeholder=pic
- task: Apply the **Team Photo** layout to the **Our crew** slide.
  check: slide.is title="Our crew" layout="Team Photo"
- task: On the **Title and Content** layout, delete the date placeholder.
  check: layout.is name="Title and Content" lacks=dt
- task: In the handout master, type **Season kickoff** in the header.
  check: master.is kind=handout phSays="hdr:Season kickoff"
- task: In the notes master, type **Lakeshore Paddle Company** in the footer.
  check: master.is kind=notes phSays="ftr:Lakeshore Paddle Company"
- task: Give the title slide a picture background using **lake.jpg**. Don't change the other slides.
  check: slide.is slide=1 background=picture
  check: slide.is title="This summer" background=none
- task: Give the **Thank you** slide a solid fill background in a dark color.
  check: slide.is title="Thank you" background=solid
:::
