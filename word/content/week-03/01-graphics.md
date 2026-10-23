---
type: lesson
id: W3-L1
week: 3
order: 1
title: Pictures, Shapes and SmartArt
objectives: [W3.1, W3.2]
clos: [CLO3]
minutes: 60
status: draft
---

::: callout why
A good picture or diagram says in a second what a paragraph says in a minute. A badly placed one pushes text around, hides behind other things and tells a screen reader user nothing. This lesson is about adding graphics on purpose.
:::

## Your scenario

Lakeshore Paddle Company is holding its first **Paddle Day**: free kayak, canoe and paddleboard demos for anyone who wants to try. You're making the flyer this week, then the event schedule, a chart of the visitor survey and the summer newsletter.

## Pictures

Insert a picture with [[path: Insert > Pictures]]: **This Device** for a file you have, **Stock Images** for Microsoft's free library, or **Online Pictures** (check the license before you use one).

Once a picture is selected, the **Picture Format** tab appears:

| Group | What it does |
|---|---|
| **Picture Styles** | Frames, shadows and shapes in one click |
| **Adjust** | **Remove Background**, **Corrections**, **Color**, **Artistic Effects** (desktop Word) |
| **Arrange** | Wrapping and position (later in this lesson) |
| **Size** | Height, width and **Crop** |

::: video
id: w3-l1-pictures
url: TODO
title: Picture styles, effects and removing backgrounds
captions: none
duration: 4:00
:::

::: drill D3.01
title: Insert a photo, style it, remove its background
badge: Picture Pro
starter: starters/week-03/W3-L1-flyer.docx
files: starters/week-03/kayak.png
platform: desktop
mos: 5.1.2, 5.2.2, 5.2.3
other: N14
minutes: 7

Download the kayak picture first (it's next to the practice file above).

- step: Insert **kayak.png** on the empty line after the **Meet the fleet** paragraph.
  check: object.count kind=picture min=3
  hint: Click at the end of the Meet the fleet paragraph and press [[key: Enter]], then [[path: Insert > Pictures > This Device]].
- step: Apply any picture style to the kayak picture.
  check: object.is kind=picture any=true pictureStyle=true
  hint: Select the picture, then pick a style in [[path: Picture Format > Picture Styles]].
- step: Remove the kayak picture's background, so only the kayak and paddle are left.
  check: object.is kind=picture any=true bgRemoved=true
  hint: [[path: Picture Format > Remove Background]]. Word marks the background in purple; select **Keep Changes**.
:::

::: drill D3.02
title: Apply an artistic effect
badge: Artistic Touch
starter: starters/week-03/W3-L1-flyer.docx
platform: desktop
mos: 5.2.1
minutes: 3

- step: Apply any **artistic effect** (for example Pencil Sketch or Paint Strokes) to the lake photo at the top.
  check: object.is kind=picture name="Lake photo" artistic=true
  hint: Select the photo, then [[path: Picture Format > Artistic Effects]].
:::

## Shapes and icons

**Shapes** ([[path: Insert > Shapes]]) are lines, arrows, boxes, stars and callouts. **Icons** ([[path: Insert > Icons]]) are simple, single-color drawings that stay sharp at any size.

- To put text in a shape, select it and start typing.
- **Shape Fill** and **Shape Outline** ([[path: Shape Format > Shape Styles]]) change its colors.
- Hold [[key: Shift]] while drawing to keep a circle round or a line straight.

::: drill D3.03
title: Add an icon and a labeled shape
badge: Shape Shifter
starter: starters/week-03/W3-L1-flyer.docx
platform: web
mos: 5.1.1, 5.1.7, 5.3.2, 5.2.4
other: N14
minutes: 6

- step: Insert an icon of a boat, a sun or anything that fits the event.
  check: object.count kind=icon min=1
  hint: [[path: Insert > Icons]], search for "kayak" or "boat", select one and choose **Insert**.
  web: In Word for the web, icons are under [[path: Insert > Pictures > Icons]] in some versions. If you don't see them, use desktop Word.
- step: Insert a rectangle or rounded rectangle and type **Free demos at 10 a.m.** in it.
  check: object.is kind=shape any=true says="Free demos at 10 a.m."
  hint: [[path: Insert > Shapes]], draw it, then type while the shape is selected.
- step: Change the shape's fill and outline colors.
  check: object.is kind=shape contains="Free demos" fill=changed outline=changed
  hint: [[path: Shape Format > Shape Fill]] and **Shape Outline**. Pick colors with enough contrast for the text.
:::

## SmartArt

**SmartArt** ([[path: Insert > SmartArt]]) turns a list into a diagram: steps in a process, a cycle, a hierarchy. Type each item in the **Text pane** (the arrow on the left edge of the graphic). **Change Colors** and the **SmartArt Styles** gallery are on the **SmartArt Design** tab.

::: platform web
SmartArt needs desktop Word. Word for the web shows SmartArt but can't insert it.
:::

::: drill D3.04
title: Show the sign-up steps as SmartArt
badge: SmartArt Starter
starter: starters/week-03/W3-L1-flyer.docx
platform: desktop
mos: 5.1.4, 5.3.3, 5.2.5
minutes: 7

- step: Under **How it works**, insert a **Basic Process** SmartArt graphic with four steps: **Reserve**, **Check in**, **Safety talk**, **Launch**.
  check: object.is kind=smartart layout=process items=4
  hint: [[path: Insert > SmartArt > Process > Basic Process]]. Open the Text pane and press [[key: Enter]] to add a fourth box.
- step: Change the graphic's colors.
  check: object.is kind=smartart colors=true
  hint: [[path: SmartArt Design > Change Colors]].
- step: Apply a different SmartArt style.
  check: object.is kind=smartart smartStyle=true
:::

## Text boxes, wrapping and position

A **text box** ([[path: Insert > Text Box]]) holds text you can place anywhere, like a pull quote in a magazine.

How a graphic sits with the text around it is its **wrapping** ([[path: Layout Options]] button, or [[path: Picture Format > Wrap Text]]):

| Wrapping | Effect |
|---|---|
| **In Line with Text** | Moves with the text like one big letter (the default for pictures) |
| **Square** | Text flows around a rectangle around the object |
| **Tight** | Text follows the object's outline |
| **Top and Bottom** | Text above and below, nothing beside it |
| **Behind Text / In Front of Text** | The object floats over or under the text |

**Position** ([[path: Picture Format > Position]]) puts a floating object in one of nine spots on the page, such as top right.

::: drill D3.05
title: Add a pull quote
badge: Quote Placer
starter: starters/week-03/W3-L1-flyer.docx
platform: web
mos: 5.1.6, 5.3.1, 5.4.1, 5.4.2
minutes: 6

- step: Insert a text box and type the paddler's quote: **Best morning of my summer.**
  check: object.is kind=textbox any=true says="Best morning of my summer"
  hint: [[path: Insert > Text Box > Draw Text Box]] (desktop), or [[path: Insert > Text Box]] (web).
  web: If your Word for the web has no Text Box button, do this drill in desktop Word.
- step: Set the text box to **Square** wrapping and position it at the **top right** of the page.
  check: object.is kind=textbox contains="Best morning" wrap=square halign=right valign=top
  hint: [[path: Shape Format > Position > Top Right with Square Text Wrapping]] does both in one click.
:::

## Screenshots and 3D models

- **Screenshot** ([[path: Insert > Screenshot]], desktop Word) inserts a picture of another open window, or a **Screen Clipping** of any part of your screen.
- **3D Models** ([[path: Insert > 3D Models]], desktop Word) can be turned with the rotate handle in the middle, or set to a view in [[path: 3D Model > 3D Model Views]].

::: drill D3.06
title: A screen clipping and a 3D model
badge: Dimension Explorer
starter: starters/week-03/W3-L1-flyer.docx
platform: desktop
mos: 5.1.5, 5.1.3, 5.2.6
minutes: 6

- step: Insert a **Screen Clipping** of any web page about paddling or lakes, at the end of the flyer.
  screenshot: The flyer with your screen clipping in it.
  example: TODO
  describe: Say what web page you clipped and where it is in the flyer.
- step: Insert a 3D model (try the stock **Boats** or **Animals** models).
  check: object.count kind=model3d min=1
  hint: [[path: Insert > 3D Models > Stock 3D Models]].
- step: Rotate the model so it faces a different way, or pick a view from **3D Model Views**.
  screenshot: The 3D model turned to a new angle.
  example: TODO
  describe: Say which way the model faces now.
:::

## Alt text

**Alternative text** (alt text) is what a screen reader says in place of a picture. Good alt text says what matters **in context**, in a sentence or less: not "image of a lake" but "Calm lake at sunrise with green hills behind it."

- Right-click a graphic and choose **View Alt Text** or **Edit Alt Text** (the name depends on your version), or use [[path: Picture Format > Alt Text]].
- Pictures that are only decoration, like a divider line, should be marked **decorative**, so screen readers skip them.
- Word can suggest alt text. Read it and rewrite it; automatic descriptions are often vague or wrong.

::: drill D3.07
title: Alt text for every graphic
badge: Alt Text Ally
starter: starters/week-03/W3-L1-flyer.docx
platform: web
mos: 5.4.3
minutes: 5

- step: Mark the wavy **divider line** under the date as **decorative**.
  check: object.is kind=picture name="Divider line" decorative=true
  hint: Right-click it, choose **View Alt Text** (or **Edit Alt Text**), then check **Mark as decorative**.
- step: Write alt text for every other graphic in the flyer: the lake photo, and anything you added in the drills above.
  check: access.is rules="alt"
  hint: Say what the reader needs to know. For the lake photo: "Calm blue lake at sunrise with green hills behind it."
:::

::: quiz Q31.1
question: Which is the best alt text for the kayak picture on this flyer?
type: single
- [ ] kayak.png
- [ ] Image
- [x] Orange sit-on-top kayak with a yellow-bladed paddle
- [ ] A picture of a kayak that we will have at Paddle Day this year for people to try
explain: Good alt text is short and says what matters. File names and "image" say nothing, and screen readers already announce that it's a picture.
:::

## Check your understanding

- When would you choose Square wrapping, and when In Line with Text?
- What's the difference between a picture style and an artistic effect?
- Which graphics should be marked decorative?
