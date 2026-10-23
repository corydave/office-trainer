---
type: project
id: PP1-P1
week: 1
order: 6
title: "Project 1: Explainer Presentation"
objectives: [W1.1, W1.2, W1.3, W1.4, W1.5]
clos: [CLO1, CLO2, CLO3]
minutes: 180
status: draft
rubricBadge: Clear Explainer
---

::: callout why
Explaining something you know well to people who don't is one of the most useful things a presentation does. This project uses everything from this week: layouts, sections, lists, notes and a handout.
:::

## What you'll make

One presentation, `Lastname_P1_Explainer.pptx`, of **8 to 10 slides** that teaches something you know well: a skill (changing a bike tire, a recipe, a game strategy), a place (your hometown, a park, a campus building) or a process (how your job's busiest day works). It needs:

1. A **title slide** with your title and your name.
2. At least **two named sections** in a logical order, with a **Section Header** slide where a new part starts if it helps.
3. **One idea per slide**, each with a clear, different **title** typed in the title placeholder.
4. At least one **bulleted list with two levels** and one **numbered list** for steps.
5. **Slide numbers** on every slide except the title slide, and a footer with your name or course.
6. **Speaker notes** on at least six slides, holding the detail you'd say out loud.
7. A **Title** property, and a **PDF handout** (3 slides per page) saved next to the presentation.

You may start from your own Word outline (Heading 1 for titles, Heading 2 for points) and use **Slides from Outline**.

::: callout ai
Optional: ask an AI tool to suggest an outline for your topic, then cut it down to what you can explain yourself. Check every fact it adds, and write in your notes which suggestions you kept.
:::

::: callout privacy
Use your own topic or a made-up one. Don't put classmates' names, photos or personal details on your slides.
:::

## Check before you submit

::: preflight P1-check
title: Preflight for your explainer
starter: none
platform: web

- step: The presentation has 8 to 10 slides.
  check: slide.count min=8 max=10
- step: The slides are in at least two named sections.
  check: section.count min=2
- step: Every slide has its own title, and the presentation has a Title property.
  check: access.is rules="titles,uniqueTitles,title"
- step: At least six slides have speaker notes.
  check: notes.count min=6
- step: There's a numbered list somewhere.
  check: list.is slide=all kind=number
- step: Slide numbers are on every slide except the title slide.
  check: footer.is slideNumbers=true skipTitle=true
:::

## How it's graded

::: rubric P1
- P1-a | 25 | CLO3 | The presentation is organized: sections in a logical order, one idea per slide, a clear title on every slide.
- P1-b | 25 | CLO1, CLO2 | Text is formatted with PowerPoint's tools on purpose: layouts and placeholders, lists and levels, readable sizes (no text boxes faking layouts).
- P1-c | 20 | CLO3 | Speaker notes carry the detail, so the slides stay short.
- P1-d | 30 | CLO1 | Ready to share: slide numbers and footer, a Title property, and a PDF handout that matches the slides.
:::
