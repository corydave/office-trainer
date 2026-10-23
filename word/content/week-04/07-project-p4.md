---
type: project
id: W4-P4
week: 4
order: 7
title: "Project 4: Job Application Packet"
objectives: [W4.1, W4.2, W4.3, W4.4, W4.5]
clos: [CLO1, CLO2, CLO4, CLO5, CLO6, CLO7]
minutes: 180
status: draft
---

::: callout why
Internships, scholarships, transfer applications and jobs all ask for the same packet. Build it properly once, and the next application takes an hour instead of a weekend.
:::

## What you'll make

1. **A designed resume** (.docx and .pdf), from your own saved template.
2. **An ATS-friendly resume** (.docx): real headings, tab stops, no text boxes or layout tables, contact details in the body.
3. **Cover letters merged to at least three employers** (real postings, or the Alex Morgan persona), with an Address Block, a Greeting Line, at least one merge field in the body and an If...Then...Else rule.
4. **Envelopes or labels**, and a **directory** of your contacts, from the same recipient list.
5. **A web page version** of your resume (.htm). Word saves its pictures in a folder next to it, so zip the .htm file and that folder together before you submit. Add a short reflection (one paragraph) comparing a web page, a PDF and a share link for sending your resume.

::: callout privacy
Use a made-up persona if you'd rather not share your own details. Your recipient list holds other people's information: keep it in your course folder only, and delete it once the project is graded.
:::

::: callout ai
Optional: AI can tailor a cover letter to a job posting quickly. Your judgment decides what's true about you: never claim a skill or experience you don't have, and keep other people's personal details out of AI prompts.
:::

## Check before you submit

Run this on your **ATS-friendly resume**.

::: preflight P4-check
title: Preflight for your ATS resume
starter: none
platform: web

- step: No text boxes or layout tables.
  check: object.count kind=textbox max=0
  check: table.count max=0
- step: Headings use heading styles, and links say where they go.
  check: access.is rules="headings,links"
- step: The Title property is filled in.
  check: props.is title="*"
  platform: desktop
- step: No double spaces used to push text across the page.
  check: text.noDoubleSpaces
:::

## How it's graded

::: rubric P4
- P4-a | 20 | CLO1, CLO2, CLO5 | Resume content and layout are professional; a personal template is saved and reused.
- P4-b | 15 | CLO1, CLO4 | The ATS version uses simple structure: no text boxes or layout tables, and key details aren't hidden in headers.
- P4-c | 25 | CLO7 | The mail merge is correct: data source, fields, rule, preview, and no merge errors.
- P4-d | 20 | CLO7 | Envelopes or labels and a directory are produced from the same data source.
- P4-e | 20 | CLO6 | The web page version is created and previewed, and the reflection compares output formats.
:::
