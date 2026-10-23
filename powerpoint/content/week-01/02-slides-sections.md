---
type: lesson
id: PP1-L2
week: 1
order: 2
title: Slides, Layouts and Sections
objectives: [W1.2, W1.3]
clos: [CLO1, CLO2, CLO3]
minutes: 50
status: draft
---

::: callout why
A presentation people can follow has one idea per slide, in an order that makes sense, grouped into parts you can name. Layouts give every slide the same structure, and sections let you move a whole part of the talk at once.
:::

## New slides and layouts

A **layout** decides which placeholders a slide has and where they sit: a title, a box for text or a picture, two boxes side by side. Pick the layout when you add the slide, or change it later.

| To | Do this |
|---|---|
| Add a slide | [[path: Home > New Slide]] (the top half of the button adds a slide like the current one; the arrow lets you pick a layout). [[key: Ctrl+M]] also works |
| Change a slide's layout | [[path: Home > Layout]], then pick one. Your text moves into the new placeholders |
| Duplicate a slide | Right-click its thumbnail > **Duplicate Slide**, or [[key: Ctrl+D]] with the thumbnail selected |
| Delete a slide | Select its thumbnail and press [[key: Delete]] |

The common layouts are **Title Slide** (first slide), **Title and Content** (a title and a list, table, chart or picture), **Two Content** (two side-by-side boxes), **Section Header** (a divider between parts), **Title Only** and **Blank**.

::: drill D1.04
title: Add slides with the right layouts
badge: Layout Picker
starter: starters/week-01/P1-L2-orientation.pptx
platform: web
mos: 2.1.3, 2.1.5
other: N3
minutes: 8

- step: Right after the **Safety first** slide, add a **Title and Content** slide titled **Life jacket check**, with two bullets: **Buckles closed** and **Snug at the shoulders**.
  check: slide.is title="Life jacket check" layout="Title and Content" after="Safety first"
  hint: Select the Safety first thumbnail, then [[path: Home > New Slide]] (the arrow) and **Title and Content**. The new slide goes right after the one you selected.
- step: Add a **Two Content** slide titled **Kayaks and canoes**.
  check: slide.is title="Kayaks and canoes" layout="Two Content"
- step: Change the layout of the **Questions** slide to **Section Header**.
  check: slide.is title="Questions" layout="Section Header"
  hint: Select the Questions slide, then [[path: Home > Layout > Section Header]].
- step: Duplicate the **Gear checkout** slide (the shop will add a second page of steps later).
  check: slide.count title="Gear checkout" min=2
  hint: Right-click its thumbnail > **Duplicate Slide**.
:::

## Reuse slides and build slides from an outline

You don't have to retype slides that already exist.

- **Reuse Slides** brings in slides from another presentation saved on your computer: [[path: Home > New Slide > Reuse Slides]] (desktop), **Open a PowerPoint File**, then select the slides you want. Tick **Keep source formatting** only if you want them to keep their old look; otherwise they take on this presentation's theme.
- **Slides from Outline** turns a Word outline into slides: [[path: Home > New Slide > Slides from Outline]] (desktop). Each **Heading 1** paragraph becomes a slide title and each **Heading 2** becomes a bullet on that slide.

::: drill D1.05
title: Reuse slides and import an outline
badge: Slide Recycler
starter: starters/week-01/P1-L2-orientation.pptx
files: starters/week-01/P1-L2-reuse.pptx, starters/week-01/P1-L2-outline.docx
platform: desktop
mos: 2.1.1, 2.1.2
minutes: 8

Download **P1-L2-reuse.pptx** and **P1-L2-outline.docx** first (they're next to the practice file above).

- step: Reuse the **Our boats** and **Rental prices** slides from **P1-L2-reuse.pptx**, so they take on this presentation's look.
  check: slide.is title="Our boats"
  check: slide.is title="Rental prices"
  hint: [[path: Home > New Slide > Reuse Slides]], select **Open a PowerPoint File** (or **Browse** in older versions), pick P1-L2-reuse.pptx from your computer (not OneDrive), then select each slide. Leave **Keep source formatting** cleared.
- step: Build slides from the Word outline **P1-L2-outline.docx**.
  check: slide.is title="Weather rules"
  check: text.present value="Thunder heard"
  hint: [[path: Home > New Slide > Slides from Outline]], then pick the .docx file. Close the outline in Word first, or PowerPoint can't open it.
:::

## Sections

**Sections** are named groups of slides, like chapters. They show as headings in the slide pane and in Slide Sorter, and you can collapse them, move them, or print one.

| To | Do this |
|---|---|
| Add a section | Right-click the first slide of the new section (or the gap above it) > **Add Section**, type a name, then **Rename** |
| Rename a section | Right-click its name > **Rename Section** |
| Move a section | Drag its name, or right-click it > **Move Section Up** or **Move Section Down**. All its slides move with it |
| Remove a section | Right-click its name > **Remove Section** (keeps the slides) |

To reorder slides, drag thumbnails in the slide pane or in Slide Sorter. To keep a slide but skip it in the show, right-click it > **Hide Slide** (its number is crossed out). To show it again, right-click it > **Unhide Slide**.

::: video
id: p1-l2-sections
url: TODO
title: Sections, reordering and hidden slides
captions: none
duration: 3:30
:::

::: drill D1.06
title: Organize the talk into sections
badge: Section Planner
starter: starters/week-01/P1-L2-orientation.pptx
platform: web
mos: 2.2.1, 2.3.1, 2.3.2, 2.3.3
other: N10
minutes: 10

- step: The presentation has one section, **Section 1**. Rename it **Getting started**.
  check: section.is name="Getting started" index=1
  hint: Right-click **Section 1** above the first thumbnail > **Rename Section**.
- step: Add three more sections: **On the water**, starting with **Safety first**; **Gear**, starting with **Gear checkout**; and **End of day**, for **Questions** and **Closing time**.
  check: section.is name="On the water" starts="Safety first"
  check: section.is name="Gear" starts="Gear checkout"
  check: section.is name="End of day" contains="Closing time"
  hint: Right-click the **Safety first** thumbnail > **Add Section**, type the name, then **Rename**. Repeat for the others.
- step: Staff collect their gear before the safety talk. Move the **Gear** section up, so it comes before **On the water**.
  check: section.is name="Gear" index=2
  hint: Right-click the **Gear** section name > **Move Section Up**.
- step: Drag the **Questions** slide so it's the last slide in the presentation.
  check: slide.is slide=last titleIs="Questions"
- step: Hide the **Old schedule (2026)** slide.
  check: slide.is title="Old schedule" hidden=true
  hint: Right-click its thumbnail > **Hide Slide**.
:::

::: quiz Q12.1
question: You move a section to the end of the presentation. What happens to its slides?
type: single
- [ ] They stay where they were; only the section name moves
- [x] They all move with the section, in the same order
- [ ] They're deleted and have to be reused
- [ ] They're hidden
explain: A section is a group of slides. Moving it moves every slide in it, which is why sections make it easy to rearrange a long talk.
:::

## Check your understanding

- When would you pick a Section Header slide instead of a Title and Content slide?
- What's the difference between hiding a slide and deleting it?
- In a Word outline, which paragraphs become slide titles?
