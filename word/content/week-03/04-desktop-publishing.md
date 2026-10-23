---
type: lesson
id: W3-L4
week: 3
order: 4
title: Desktop Publishing
objectives: [W3.5]
clos: [CLO5]
minutes: 40
status: draft
---

::: callout why
Newsletters, flyers and programs live or die on layout. Columns, a consistent color theme and a few page elements make a page look designed, and they keep it readable if you use them with a light touch.
:::

## Columns and sections

[[path: Layout > Columns]] splits text into two or three columns. If you select part of the document first, Word adds **Continuous section breaks** around your selection, so a one-column heading (a **masthead**) can sit above two columns of articles.

- A **column break** ([[path: Layout > Breaks > Column]]) pushes the next text to the top of the next column, so columns end where you choose.
- Turn on **Show/Hide** (¶) to see the breaks. Deleting a section break by accident merges two sections' layouts.

::: video
id: w3-l4-columns
url: TODO
title: A newsletter with a masthead and two columns
captions: none
duration: 3:30
:::

::: drill D3.11
title: Two columns under a one-column masthead
badge: Column Crafter
starter: starters/week-03/W3-L4-newsletter.docx
platform: desktop
mos: 2.3.1, 2.3.2, 2.3.3
minutes: 7

- step: Select from **Welcome** down to the end of the **Staff picks** paragraph, and put it in **two columns**. The masthead stays in one column.
  check: section.is para="The Shoreline" columns=1
  check: section.is para="Welcome to the summer issue" columns=2
  hint: [[path: Layout > Columns > Two]] with the text selected. Word adds the section breaks for you.
- step: Check that the two-column section starts with a **Continuous** section break, so it stays on the same page as the masthead.
  check: section.is para="Welcome to the summer issue" breakType=continuous
  hint: Turn on Show/Hide (¶) to see "Section Break (Continuous)" above Welcome.
- step: Insert a **column break** before **Safety corner**, so it starts at the top of the second column.
  check: break.count kind=column min=1
  hint: Click just before Safety corner, then [[path: Layout > Breaks > Column]].
:::

## Themes and page design

A **theme** sets the document's colors, fonts and effects together ([[path: Design > Themes]]). Change just the colors with [[path: Design > Colors]], or make your own set with **Customize Colors**, for example to match a club's brand.

| Element | Where | Use it sparingly for |
|---|---|---|
| **Page color** | [[path: Design > Page Color]] | Screen-only documents (Word doesn't print it unless **Print background colors and images** is on in its options) |
| **Page border** | [[path: Design > Page Borders]] | Certificates, posters, invitations |
| **Watermark** | [[path: Design > Watermark]] (desktop Word) | DRAFT, CONFIDENTIAL or SAMPLE behind the text |
| **Drop cap** | [[path: Insert > Drop Cap]] (desktop Word) | The first letter of the first article |

::: callout warning
Color must have enough contrast with the text on it: at least 4.5 to 1 for normal text (WCAG 1.4.3). Pale yellow text on white looks fine on your screen and disappears on a projector. And never use color as the only way to show something (WCAG 1.4.1).
:::

::: drill D3.12
title: Brand colors and page elements
badge: Page Designer
starter: starters/week-03/W3-L4-newsletter.docx
platform: desktop
mos: 1.2.4
other: E3.2
minutes: 8

- step: Create custom theme colors named **Lakeshore**, with **Accent 1** set to the company teal, **#0E7C86**.
  check: theme.is colors="Lakeshore" accent1=#0E7C86
  hint: [[path: Design > Colors > Customize Colors]]. Open the Accent 1 menu, choose **More Colors > Custom**, and type 0E7C86 in **Hex**. Name the set Lakeshore and select **Save**.
- step: Add a page border to the whole document.
  check: section.is border=true
  hint: [[path: Design > Page Borders]]. Pick a style and color; **Apply to: Whole document**.
- step: Add a **DRAFT** watermark.
  check: page.background watermark="DRAFT"
  hint: [[path: Design > Watermark > DRAFT 1]].
- step: Give the first paragraph (**Welcome to the summer issue...**) a drop cap.
  check: para.is para="Welcome to the summer issue" dropCap=true
  hint: Click in the paragraph, then [[path: Insert > Drop Cap > Dropped]].
- step: Set a light page color.
  check: page.background color=any
  hint: [[path: Design > Page Color]]. Keep it pale, so the text stays easy to read.
:::

## Check your understanding

- Why does putting part of a page in columns need section breaks?
- When is a page color a bad idea?
- What's the minimum contrast for normal-size text?
