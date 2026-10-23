---
type: lesson
id: PP4-L4
week: 4
order: 4
title: Links, Action Buttons and Zoom
objectives: [W4.4]
clos: [CLO2, CLO3]
minutes: 40
status: draft
---

::: callout why
A presentation doesn't have to go in a straight line. Links, buttons and Zoom let you jump to the part your audience asks about, or let someone at a kiosk explore on their own and always find their way back.
:::

## Hyperlinks

Select the words (or a shape or picture), then [[path: Insert > Link]] ([[key: Ctrl+K]], or [[key: Cmd+K]] on a Mac):

- **Existing File or Web Page** (desktop) or a web address: type the address, such as `https://www.lakeshorepaddle.example/rentals`.
- **Place in This Document** (desktop), or the slides in this presentation (web): pick a slide.

Link words that say where the link goes ("our rental page"), not "click here" or the raw address. Screen readers often read links on their own, out of context.

## Action buttons

**Action buttons** are shapes made for navigation: [[path: Insert > Shapes > Action Buttons]] (the last group), then draw one. The **Action Settings** box opens: choose **Hyperlink to** and a slide (**First Slide**, **Last Slide Viewed**, or **Slide...** for a particular one). You can set the same actions on any shape with [[path: Insert > Action]].

::: drill D4.09
title: Link the tour together
badge: Link Builder
starter: starters/week-04/P4-L4-links.pptx
platform: web
mos: 3.2.1
minutes: 7

- step: On **Book online**, link the words **rental page** to `https://www.lakeshorepaddle.example/rentals`.
  check: link.present title="Book online" url="lakeshorepaddle.example/rentals" says="rental page"
  hint: Select the two words, then [[path: Insert > Link]] and paste the address.
- step: On **Menu**, link the word **Prices** to the **Prices** slide.
  check: link.present title="Menu" toSlide="Prices"
  hint: Select the word, [[path: Insert > Link]], then **Place in This Document** and the Prices slide. (In PowerPoint for the web, pick the Prices slide from the list of slides in this presentation.)
- step: On **Prices**, add an action button that goes back to the **Menu** slide.
  platform: desktop
  check: link.present title="Prices" button=true toSlide="Menu"
  hint: [[path: Insert > Shapes > Action Buttons]], draw one, then **Hyperlink to > Slide...** and pick **Menu**.
:::

## Zoom

**Zoom** ([[path: Insert > Zoom]], desktop) adds live thumbnails you can click during the show:

| Zoom | What it does |
|---|---|
| **Summary Zoom** | Makes a new slide with a thumbnail for each section you pick: a menu for the whole talk. It adds sections if you don't have them |
| **Section Zoom** | A thumbnail on the current slide that jumps to one section |
| **Slide Zoom** | A thumbnail that jumps to one slide |

Summary and Section zooms come back to the zoom slide on their own. For a Slide Zoom, select it and tick **Return to Zoom** on the **Zoom** tab, or the show moves on to the next slide.

::: drill D4.10
title: Make a zoom menu
badge: Zoom Navigator
starter: starters/week-04/P4-L4-links.pptx
platform: desktop
mos: 2.1.4, 3.2.2
minutes: 8

- step: Insert a **Summary Zoom** of the **Boats**, **Prices** and **Book** sections.
  check: zoom.is kind=summary
  hint: [[path: Insert > Zoom > Summary Zoom]], tick the first slide of each section, then **Insert**.
- step: On **Menu**, insert a **Section Zoom** to the **Boats** section.
  check: zoom.is title="Menu" kind=section
  hint: Select the Menu slide, then [[path: Insert > Zoom > Section Zoom]].
- step: On **Kayaks**, insert a **Slide Zoom** to the **Paddleboards** slide, and tick **Return to Zoom**.
  check: zoom.is title="Kayaks" kind=slide returnToZoom=true
  hint: [[path: Insert > Zoom > Slide Zoom]], pick Paddleboards, then [[path: Zoom > Return to Zoom]].
:::

::: quiz Q44.1
question: Which link text is best for a screen reader user?
type: single
- [ ] Click here
- [ ] https://www.lakeshorepaddle.example/rentals
- [x] Book on our rental page
- [ ] Link
explain: Link text should make sense on its own, because screen reader users often jump from link to link.
:::

## Check your understanding

- When is a Summary Zoom better than a menu of hyperlinks?
- What does Return to Zoom do?
- Why avoid "click here" as link text?
