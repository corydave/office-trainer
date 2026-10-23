---
type: lesson
id: PP3-L2
week: 3
order: 2
title: Shapes, Ink, Order and Alignment
objectives: [W3.2]
clos: [CLO4, CLO2]
minutes: 45
status: draft
---

::: callout why
Shapes point, label and connect: an arrow to the dock, a callout for this week's deal. Lined up and evenly spaced, they look designed; a few pixels off, they look like an accident.
:::

## Shapes and text boxes

- **Insert a shape:** [[path: Insert > Shapes]], pick one, then click on the slide (or drag to draw it). Hold [[key: Shift]] while you drag for a perfect circle or square.
- **Add text to a shape:** select it and start typing.
- **Text box:** [[path: Insert > Text Box]], then click and type. Use text boxes for labels; use placeholders for the slide's main text.
- **Exact size:** type a **Height** and **Width** in [[path: Shape Format > Size]].
- **Format:** [[path: Shape Format > Shape Styles]] applies a theme style in one click, or set **Shape Fill**, **Shape Outline** and **Shape Effects** yourself.

::: drill D3.03
title: Add a callout and a label
badge: Shape Shaper
starter: starters/week-03/P3-L2-shapes.pptx
platform: web
mos: 3.4.1, 3.4.3, 3.4.4, 3.4.5, 3.4.6
other: N8
minutes: 7

- step: On **Find us**, insert any **callout** shape and type **Free demos Saturday** in it.
  check: object.is title="Find us" kind=shape says="Free demos" geom=callout
  hint: [[path: Insert > Shapes]], then a shape from the **Callouts** group. Select the shape and type.
- step: Make the callout exactly **3"** wide and **1"** tall.
  check: object.is title="Find us" says="Free demos" width=3in height=1in
  hint: [[path: Shape Format > Size]]: type 1 in Height and 3 in Width. (In PowerPoint for the web, the tab is called **Shape**.)
- step: Apply a **shape style** to the callout.
  check: object.is title="Find us" says="Free demos" styleChanged=true
  hint: [[path: Shape Format > Shape Styles]], then any style in the gallery.
- step: Add a text box under the title that says **Open 9 to 7 every day**.
  check: object.is title="Find us" kind=textbox says="Open 9 to 7"
:::

## Order, align and group

| To | Do this |
|---|---|
| Put an object on top | Right-click > **Bring to Front** (or **Bring Forward** one layer). **Send to Back** does the opposite |
| Line objects up | Select them all ([[key: Shift]]+click), then [[path: Shape Format > Align]] > **Align Top**, **Align Middle**, **Align Left** ... |
| Space them evenly | Select three or more, then **Align > Distribute Horizontally** (or **Vertically**) |
| Move them as one | Select them, then [[path: Shape Format > Group > Group]]. **Ungroup** splits them again |
| See alignment tools | [[path: View > Guides]] and **Gridlines** (desktop). Smart guides (red dashed lines) appear on their own while you drag |
| Pick hidden objects | [[path: Home > Arrange > Selection Pane]] lists every object on the slide, top to bottom |

## Draw with ink

The **Draw** tab has pens and highlighters for writing on the slide with a mouse, pen or finger. Ink is saved as an object you can move, resize or delete. Use it to circle something or add a quick sketch.

::: drill D3.04
title: Line up the icons
badge: Aligner
starter: starters/week-03/P3-L2-shapes.pptx
platform: web
mos: 3.4.2, 3.5.1, 3.5.2, 3.5.3, 3.5.4
minutes: 8

- step: On **Rentals**, the logo is hidden behind the banner. Bring it to the front.
  check: object.is title="Rentals" name=Logo front=true
  hint: Find it in [[path: Home > Arrange > Selection Pane]], select it, then **Bring to Front**.
- step: Align the three icons along their **tops**, then **distribute** them so the gaps between them are equal.
  check: align.is title="Rentals" kind=icon edge=top
  check: align.is title="Rentals" kind=icon distributed=horizontal
  hint: Select all three icons, then [[path: Graphics Format > Align > Align Top]], then **Align > Distribute Horizontally**. (For icons the tab is **Graphics Format**; for shapes it's **Shape Format**.) (In PowerPoint for the web, use [[path: Home > Arrange > Align]].)
- step: Group the three icons.
  check: object.is title="Rentals" kind=icon grouped=true
  hint: With all three selected, [[path: Graphics Format > Group > Group]], or right-click > **Group**.
- step: Turn on the **Guides**, then save.
  platform: desktop
  check: view.is guides=true
  hint: [[path: View > Guides]].
- step: On **Rentals**, use a pen on the **Draw** tab to circle the banner.
  check: object.count title="Rentals" kind=ink min=1
  hint: [[path: Draw > Pen]], draw, then press [[key: Esc]] to stop drawing.
:::

::: quiz Q32.1
question: You align three icons, then drag one by accident. What would have prevented it?
type: single
- [ ] Bringing them to the front
- [x] Grouping them, so they move as one object
- [ ] Turning on the gridlines
- [ ] Adding alt text
explain: A group moves and resizes as one, so the spacing you set stays put.
:::

## Check your understanding

- When would you use a text box instead of a placeholder?
- What's the difference between Align Top and Distribute Horizontally?
- How do you select an object that's hidden behind another one?
