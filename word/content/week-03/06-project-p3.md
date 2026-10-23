---
type: project
id: W3-P3
week: 3
order: 6
title: "Project 3: Newsletter or Event Flyer"
objectives: [W3.1, W3.2, W3.3, W3.4, W3.5]
clos: [CLO3, CLO4, CLO5]
minutes: 150
status: draft
---

::: callout why
Every club, team, small business and community group needs a flyer or newsletter sooner or later. This project gives you one you can actually use, built so that everyone can read it, including people who use screen readers or can't see color well.
:::

## What you'll make

A **two-page newsletter** or **event flyer** for a real campus club, small business or community group (or a made-up one). Submit the .docx and a PDF.

It must include:

- A **masthead** (title and date line) in one column, over a **two-column** body
- At least **three graphics** (pictures, icons or shapes), placed with wrapping on purpose
- A **SmartArt** graphic (a process, a timeline or an org chart)
- A **schedule or price table**, sorted in a sensible order, with a header row that repeats (put the table's title above it, not in a merged top row)
- A **chart** with a caption
- A consistent **theme** (custom theme colors are a plus)
- **Accessibility:** alt text on every graphic (or marked decorative), real headings, enough color contrast, and a clean result from [[path: Review > Check Accessibility]]

Name the files `Lastname_P3_Newsletter.docx` and `Lastname_P3_Newsletter.pdf`.

::: callout privacy
Use your own photos, stock images from [[path: Insert > Pictures > Stock Images]], or images whose license lets you reuse them. Don't put people's faces in a flyer without their permission.
:::

::: callout ai
Optional: AI image tools and Word's automatic alt text are quick but not always right. If you use an AI-made image, say so in a caption or the footer, check the tool's terms, and rewrite any automatic alt text so it says what matters on your page.
:::

## Check before you submit

::: preflight P3-check
title: Preflight for your newsletter or flyer
starter: none
platform: web

- step: At least three graphics.
  check: object.count kind=any min=3
- step: Every graphic has alt text or is marked decorative, and every table has a header row.
  check: access.is rules="alt,tableHeaders"
- step: A SmartArt graphic.
  check: object.count kind=smartart min=1
  platform: desktop
- step: A chart with a caption.
  check: object.count kind=chart min=1
  check: field.present code="SEQ"
  platform: desktop
- step: Headings use heading styles, with no skipped levels.
  check: access.is rules="headings"
:::

## How it's graded

::: rubric P3
- P3-a | 20 | CLO3 | Graphics are inserted, formatted and positioned on purpose, not by accident.
- P3-b | 20 | CLO4 | The table has a real header row that repeats, is sorted and formatted, and doesn't use merged cells in ways that confuse reading order.
- P3-c | 25 | CLO5 | Columns, sections, the theme and page elements create a clear visual hierarchy.
- P3-d | 20 | CLO3, CLO5 | Accessibility: meaningful alt text, enough contrast, a logical reading order, and a clean Accessibility Checker.
- P3-e | 15 | CLO3 | The chart supports the message and has a caption.
:::
