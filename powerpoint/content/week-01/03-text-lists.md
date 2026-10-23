---
type: lesson
id: PP1-L3
week: 1
order: 3
title: Text, Columns and Lists
objectives: [W1.4]
clos: [CLO1]
minutes: 35
status: draft
---

::: callout why
Slides are read from across a room. Big, consistent text and short lists are easier to follow than paragraphs, and real lists (not lines with dashes typed in front) keep their shape when you change the theme or add a point.
:::

## Format text

Select text (or the whole placeholder's border, to change all of it) and use the **Font** group on the **Home** tab: font, size, bold, italic, underline, color, and **Clear All Formatting**. [[key: Ctrl+Shift+>]] and [[key: Ctrl+Shift+<]] grow and shrink the size a step at a time.

- **WordArt styles** add an outline, a shadow, a glow or a gradient to the letters themselves: select the text, then [[path: Shape Format > WordArt Styles]]. Use them for one short line, like a tagline, never for body text. WordArt styles are desktop only.
- **Columns** split a long list into two or three side by side: select the placeholder, then [[path: Home > Add or Remove Columns]] (desktop; the button with two columns of lines in the Paragraph group).

::: callout tip
Keep body text at **24 pt or larger** and titles bigger still. If a list doesn't fit at 24 pt, split it across two slides or into columns rather than shrinking it.
:::

::: drill D1.07
title: Format the title, a tagline and a long list
badge: Text Stylist
starter: starters/week-01/P1-L3-text.pptx
platform: desktop
mos: 3.1.1, 3.1.2
other: N4
minutes: 8

- step: On the title slide, make the title **Gear and safety** bold and **54 pt**.
  check: text.is slide=1 phType=title bold=true size=54
  hint: Select the title placeholder's border (so all of its text is selected), then use **Bold** and the font size box on the **Home** tab.
- step: Apply any **WordArt style** to the tagline **Paddle smart**.
  platform: desktop
  check: text.is slide=1 says="Paddle smart" wordart=true
  hint: Select the text, then pick a style in [[path: Shape Format > WordArt Styles]].
- step: On **Gear list**, set the list in **two columns** so all ten items fit at a readable size.
  platform: desktop
  check: text.is title="Gear list" says="Life jacket" columns=2
  hint: Click in the list, then [[path: Home > Add or Remove Columns > Two Columns]].
:::

## Bulleted and numbered lists

- **Bullets** for points in no particular order: [[path: Home > Bullets]].
- **Numbering** for steps that happen in order: [[path: Home > Numbering]].
- **List levels**: press [[key: Tab]] at the start of a line to indent it under the line above (a sub-point), and [[key: Shift+Tab]] to move it back out. **Increase List Level** and **Decrease List Level** on the Home tab do the same.

The body placeholder on a **Title and Content** slide is already a bulleted list. A text box you draw yourself starts as plain lines, so you turn on bullets or numbers yourself.

::: drill D1.08
title: Make real lists
badge: List Maker
starter: starters/week-01/P1-L3-text.pptx
platform: web
mos: 3.1.3
minutes: 6

- step: On **What to bring**, turn the lines into a **bulleted** list.
  check: list.is title="What to bring" kind=bullet
  hint: Select all the lines, then [[path: Home > Bullets]].
- step: Indent **Hat** and **Sunscreen** one level, under **Sun protection**.
  check: list.is title="What to bring" items="Hat,Sunscreen" level=2
  check: list.is title="What to bring" levels=2
  hint: Click at the start of each line and press [[key: Tab]].
- step: On **Launch steps**, turn the steps into a **numbered** list.
  check: list.is title="Launch steps" kind=number
  hint: Select the steps, then [[path: Home > Numbering]].
:::

::: quiz Q13.1
question: Which text should be a numbered list?
type: single
- [ ] Reasons to rent a kayak
- [x] How to launch a kayak, step by step
- [ ] The colors the boats come in
- [ ] The names of the guides
explain: Use numbers when order matters, like steps. Use bullets when it doesn't.
:::

## Check your understanding

- Why is a real bulleted list better than typing a dash in front of each line?
- What does pressing Tab at the start of a bullet do?
- When would you use columns instead of a second slide?
