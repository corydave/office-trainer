---
type: project
id: PP5-P5
week: 5
order: 6
title: "Project 5: Team Presentation"
objectives: [W5.1, W5.2, W5.3, W5.4]
clos: [CLO1, CLO3, CLO7]
minutes: 150
status: draft
rubricBadge: Team Presenter
---

::: callout why
Most presentations at work and in college are team efforts that someone has to merge, rehearse and deliver. This project is the whole cycle: write together, review, combine, rehearse, present, and ship a clean final copy.
:::

## What you'll make

With one or two classmates, a **10 to 12 slide** presentation on a topic you choose together, saved in **OneDrive** and shared with each other with **Can edit**, as `Teamname_P5_Presentation.pptx`. It needs:

1. **Real co-authoring:** each member builds some of the slides in the shared file.
2. **Comment threads:** at least three comments with replies, resolved as you deal with them.
3. **A reviewer's copy, compared and combined:** one member downloads a copy, makes at least three changes in it, and another member compares it with the shared file (with **Compare** if your PowerPoint has it, or side by side) and takes some changes and leaves others.
4. **Speaker notes** on every slide, and **rehearsed timings**.
5. A **custom show**: a 3-minute version of the talk.
6. A **final copy**, inspected (comments and personal information removed), accessible (no errors in the Accessibility Checker), with the fonts embedded, exported as a **PDF** and a **video**.
7. **Delivery:** present it live with Presenter View, or submit the video with narration.

::: callout privacy
Shared decks carry hidden data: names in comments, notes and file properties. Inspect the final copy before you share it outside the team, and share view-only unless people need to edit.
:::

::: callout ai
Optional: Copilot in PowerPoint (or another AI tool) can draft speaker notes or a summary. Treat it like a reviewer's comment: check it, change it, and say in the notes where you used it.
:::

## Check before you submit

Run this on the **final copy**.

::: preflight P5-check
title: Preflight for your final copy
starter: none
platform: desktop

- step: The presentation has 10 to 12 slides, and every slide has notes.
  check: slide.count min=10 max=12
  check: notes.count min=10
- step: Every slide has a rehearsed timing.
  check: show.is timings=true
- step: The comments and personal information are gone.
  check: comment.count max=0
  check: doc.is personalInfo=removed
- step: There are no accessibility errors.
  check: access.is rules="alt,titles,uniqueTitles,tableHeaders,links,readingOrder"
- step: The fonts are embedded.
  check: doc.is fontsEmbedded=true
:::

## How it's graded

::: rubric P5
- P5-a | 25 | CLO7 | The team collaborated for real: comment threads with replies, edits from each member, and a compared and combined reviewer's copy.
- P5-b | 25 | CLO3 | The presentation is professional: a clear story, consistent design, notes for every slide, a custom short version.
- P5-c | 25 | CLO3, CLO1 | The delivery is rehearsed: timings, Presenter View or a recording, within the time limit.
- P5-d | 25 | CLO1, CLO7 | The final copy ships clean: inspected, accessible, fonts embedded, exported as a PDF and a video.
:::
