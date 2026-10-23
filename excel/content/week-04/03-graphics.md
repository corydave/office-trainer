---
type: lesson
id: XL4-L3
week: 4
order: 3
title: Pictures, Shapes, Icons and SmartArt
objectives: [W4.3]
clos: [CLO4]
minutes: 35
status: draft
---

::: callout why
A logo makes a price list look official, a callout shape points at the one thing readers shouldn't miss, and a SmartArt diagram shows a process faster than a paragraph. Graphics float above the cells, so they don't change your data.
:::

## Graphics float over the grid

Everything on [[path: Insert > Illustrations]] (pictures, shapes, icons, SmartArt) sits **on top of** the cells, not in them. Drag a graphic to move it, drag a corner handle to resize it, and use the **Picture Format** or **Shape Format** tab that appears when it's selected.

| Graphic | Insert it with | Good for |
|---|---|---|
| **Picture** | [[path: Insert > Pictures > This Device]] | A logo or a photo |
| **Shape** | [[path: Insert > Shapes]] | A callout, an arrow, a box with text in it |
| **Text box** | [[path: Insert > Text Box]] | A note that isn't in a cell |
| **Icon** | [[path: Insert > Icons]] | A small, simple symbol that recolors cleanly |
| **SmartArt** | [[path: Insert > SmartArt]] (desktop) | A process, a cycle or an org chart |

To type in a shape, select it and start typing. Hold [[key: Shift]] while you drag a shape to keep it in proportion.

::: callout tip
Right-click a graphic > **Size and Properties** > **Properties** to choose whether it moves and resizes with the cells. "Don't move or size with cells" keeps a logo in place when you sort or insert rows.
:::

## Alternative text for graphics

Every graphic needs alt text **or** a "decorative" mark:

- A logo or a picture that carries meaning: right-click > **Edit Alt Text** and say what it is ("Lakeshore Paddle Company logo").
- A border, a divider or an icon that only repeats nearby text: tick **Mark as decorative**, so screen readers skip it.

::: video
id: e4-l3-graphics
url: TODO
title: Pictures, shapes, icons and SmartArt in a worksheet
captions: none
duration: 4:00
:::

::: drill D4.06
title: Dress up the flyer
badge: Flyer Designer
starter: starters/week-04/E4-L3-graphics.xlsx
files: starters/week-04/logo.png
platform: desktop
minutes: 10

Download **logo.png** first (it's next to the practice file above).

- step: Insert **logo.png** on the **Flyer** sheet, to the right of the price list.
  check: object.count sheet=Flyer kind=picture min=1
  hint: [[path: Insert > Pictures > This Device]], then pick logo.png.
- step: Add a shape (any rectangle or callout) that says **Free demos Saturday, 10 a.m. to 2 p.m.**
  check: object.is sheet=Flyer kind=shape says="Free demos"
  hint: [[path: Insert > Shapes]], drag to draw it, then type.
- step: Insert any boat or water **icon**.
  check: object.count sheet=Flyer kind=icon min=1
  hint: [[path: Insert > Icons]], search for **kayak** or **boat**, then **Insert**.
- step: Insert a SmartArt **Basic Process** with three steps: **Reserve**, **Paddle**, **Return**.
  screenshot: The Flyer sheet with the three-step SmartArt process on it.
  example: TODO
  describe: Name the SmartArt layout you used and the three steps in it.
  hint: [[path: Insert > SmartArt > Process > Basic Process]], then type in the Text Pane.
- step: Give the logo alt text, and give the icon alt text or mark it **decorative**.
  check: object.count sheet=Flyer kind=any min=2
  check: access.is rules=alt
:::

::: quiz Q43.1
question: An icon of a kayak sits beside the words "Kayak rentals". What should its alt text be?
type: single
- [ ] "Icon"
- [ ] "Kayak icon from the Icons gallery"
- [x] Nothing: mark it decorative, because the words beside it already say it
- [ ] "Picture 3"
explain: When a graphic only repeats text that is right next to it, mark it decorative so a screen reader doesn't say the same thing twice.
:::

## Check your understanding

- Why does a logo need alt text when a divider line doesn't?
- What happens to a picture when you sort the rows under it, and how do you control that?
- When is SmartArt clearer than a list of steps in cells?
