---
type: lesson
id: PP1-L4
week: 1
order: 4
title: Size, Footers, Notes and Printing
objectives: [W1.5]
clos: [CLO1]
minutes: 45
status: draft
---

::: callout why
The last step before you present or hand out a file is the one people notice: slides that fit the screen, numbers to refer to ("go back to slide 4"), notes that keep you on track, and a handout people can write on.
:::

## Slide size

Today's screens are wide, so new presentations are **Widescreen (16:9)**. Older presentations and some projectors use **Standard (4:3)**. Change it with [[path: Design > Slide Size]]. When you go from 4:3 to 16:9, PowerPoint stretches the layouts sideways; check pictures and charts afterward.

::: drill D1.09
title: Make the slides widescreen
badge: Widescreen
starter: starters/week-01/P1-L4-print.pptx
platform: web
mos: 1.2.1
minutes: 3

- step: The orientation deck was made years ago at 4:3. Change it to **Widescreen (16:9)**.
  check: size.is ratio=16:9
  hint: [[path: Design > Slide Size > Widescreen]]. If PowerPoint asks, choose **Ensure Fit**.
:::

## Slide numbers and footers

[[path: Insert > Header & Footer]] (desktop) or [[path: Insert > Footer > Footer]] (PowerPoint for the web) adds a **slide number**, a **footer** (such as your organization's name) and the **date** to slides. **Apply to All** puts them on every slide; **Don't show on title slide** leaves the first slide clean. In desktop PowerPoint, the **Notes and Handouts** tab of the same box sets the header and footer for printed pages.

::: drill D1.10
title: Number the slides
badge: Slide Numberer
starter: starters/week-01/P1-L4-print.pptx
platform: web
mos: 2.2.3
minutes: 4

- step: Add slide numbers and the footer **Lakeshore Paddle Company** to every slide except the title slide.
  check: footer.is slideNumbers=true footer="Lakeshore Paddle Company" skipTitle=true
  hint: [[path: Insert > Header & Footer]]. Tick **Slide number**, tick **Footer** and type the name, tick **Don't show on title slide**, then **Apply to All**. In PowerPoint for the web, [[path: Insert > Footer > Footer]] opens a Footer pane with the same choices.
:::

## Speaker notes and properties

Keep slides short and put the details in the **notes**: the sentence you want to say, a number you might forget, a reminder to pause. You'll see them in Presenter View in Week 5, and they print on notes pages.

**Properties** describe the file. The **Title** property is what a screen reader announces and what shows in search results, so give every presentation one: [[path: File > Info]], then type in the **Title** box (desktop).

## Print and export

[[path: File > Print]] shows a preview (desktop). Under **Settings**, the second box picks what prints. PowerPoint for the web offers **Print Slides**, **Print Handouts** (3 per page) and **Print Notes Pages** in the same menu.

| Choice | Prints |
|---|---|
| **Full Page Slides** | One slide per page |
| **Notes Pages** | One slide per page with its notes under it, your script |
| **Handouts** (1, 2, 3, 4, 6 or 9 slides) | Small slides on each page. **3 Slides** adds lines next to each slide for the audience's own notes |

To share a copy that looks the same on any device, save it as a PDF: [[path: File > Export > Create PDF/XPS]] (desktop) or [[path: File > Export > Download as PDF]] (PowerPoint for the web). In the desktop Publish box, **Options** lets you publish handouts or notes pages instead of full slides.

::: drill D1.11
title: Notes, properties and handouts
badge: Ready to Print
starter: starters/week-01/P1-L4-print.pptx
platform: web
mos: 1.2.3, 1.3.1, 1.3.2, 1.3.3, 1.5.6
other: N15, N16, N17
minutes: 10

- step: On the **Safety first** slide, add the speaker note **Everyone wears a life jacket on the water, staff too.**
  check: slide.is title="Safety first" notes="life jacket"
- step: Set the presentation's **Title** property to **Summer staff orientation**.
  platform: desktop
  check: props.is title="Summer staff orientation"
  hint: [[path: File > Info]], then click **Add a title** on the right and type.
- step: Open [[path: File > Print]] and choose **3 Slides** handouts. How many pages will print for this 5-slide presentation? Type the number.
  answer: 2
  hint: Desktop: the page count is under the preview (for example "1 of 2"). Web: choose **Print Handouts** (it prints 3 slides per page), then read the page count in your browser's print box.
- step: Switch the print setting to **Notes Pages** and take a screenshot of the preview showing the Safety first slide with its note.
  screenshot: The print preview of the Safety first notes page, with the note under the slide.
  example: TODO
  describe: Say which print setting you chose and what appears under the slide.
- step: Save a PDF copy of the presentation in your course folder.
  hint: [[path: File > Export > Create PDF/XPS]] (desktop) or [[path: File > Export > Download as PDF]] (web). Open the PDF to check it.
:::

::: quiz Q14.1
question: You want people to take notes beside each slide during your talk. What do you print?
type: single
- [ ] Notes Pages
- [ ] Full Page Slides
- [x] Handouts, 3 slides per page
- [ ] An outline
explain: The 3 Slides handout is the only one that prints lines next to each slide for the audience's notes. Notes Pages print your own speaker notes.
:::

## Check your understanding

- What might go wrong with a picture when you change a 4:3 deck to 16:9?
- Why leave the slide number off the title slide?
- What's the difference between Notes Pages and Handouts?
