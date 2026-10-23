---
type: skillcheck
id: PP5-SC5
week: 5
order: 5
title: "Skill Check 5: Present and Collaborate"
objectives: [W5.1, W5.3, W5.4]
clos: [CLO1, CLO3, CLO7]
minutes: 30
status: draft
---

One starter presentation, a set of tasks and a time limit, with no hints. Each task is worth one point. Most tasks need desktop PowerPoint.

::: skillcheck SC5
title: Team retreat
badge: Skill Check 5
starter: starters/week-05/P5-SC5-team.pptx
platform: desktop
mos: 1.4.1, 1.4.2, 1.4.3, 1.5.3, 1.5.4, 1.5.5
timeLimit: 30
attempts: 2
passing: 70

- task: Reply to the comment on **Agenda** with **Yes, June 12.**
  check: comment.is says="date right" state=replied
- task: Resolve that comment's thread.
  check: comment.is says="date right" state=resolved
- task: Delete the comment on **Lunch**.
  check: comment.is says="Pizza or sandwiches" state=absent
- task: On **Thank you**, add the comment **Add our email?**
  check: comment.is title="Thank you" says="email"
- task: Create a custom show named **Highlights** with three slides: **Team retreat**, **Our morning paddle** and **Thank you**.
  check: show.is customShow="Highlights" slides=3 includes="Team retreat,Our morning paddle,Thank you"
- task: Set the show to loop continuously until Esc.
  check: show.is loop=true
- task: Rehearse timings for every slide and keep them.
  check: show.is timings=true
- task: Remove the document properties and personal information. Don't remove the comments.
  check: doc.is personalInfo=removed
- task: Remove the content that's off the edge of a slide.
  check: doc.is offSlide=none
- task: Fix the accessibility error for the photo.
  check: access.is rules=alt
- task: Embed the fonts in the file.
  check: doc.is fontsEmbedded=true
:::
