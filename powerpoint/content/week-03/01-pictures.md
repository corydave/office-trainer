---
type: lesson
id: PP3-L1
week: 3
order: 1
title: Pictures and Screenshots
objectives: [W3.1]
clos: [CLO4]
minutes: 40
status: draft
---

::: callout why
One good photo does more than a paragraph: it shows the boat, the beach, the result. But a stretched, blurry or uncropped photo looks careless, and a photo with no alt text is invisible to anyone using a screen reader.
:::

## Insert, resize and crop

- **Insert a picture:** [[path: Insert > Pictures > This Device]] (or **Stock Images** for free photos you can use). A picture placeholder's icon also works.
- **Resize without stretching:** drag a **corner** handle, not a side handle. For an exact size, type the width in [[path: Picture Format > Size]]; with **Lock aspect ratio** on (the default), the height follows.
- **Crop:** [[path: Picture Format > Crop]], drag the black crop handles, then press [[key: Esc]] or select Crop again. **Crop > Aspect Ratio > 1:1** makes a perfect square, and **Crop to Shape** puts the photo in a circle or any other shape.

::: callout tip
Cropped parts are still in the file until you compress the picture ([[path: Picture Format > Compress Pictures]], desktop; tick **Delete cropped areas of pictures**). Do that before you share a deck full of big photos.
:::

::: drill D3.01
title: Insert, resize and crop a photo
badge: Photo Fixer
starter: starters/week-03/P3-L1-pictures.pptx
files: starters/week-03/kayak.jpg
platform: web
mos: 3.3.1
other: N7
minutes: 6

Download **kayak.jpg** first (it's next to the practice file above).

- step: On **Our fleet**, insert **kayak.jpg**.
  check: object.count title="Our fleet" kind=picture min=1
  hint: Select the slide, then [[path: Insert > Pictures > This Device]].
- step: Resize it from a corner so it's about half the slide wide, without stretching it.
  check: object.is title="Our fleet" kind=picture stretched=false minWidth=4in
- step: Crop it to a square.
  check: object.is title="Our fleet" kind=picture cropped=true square=true
  hint: [[path: Picture Format > Crop > Aspect Ratio > 1:1]], then [[key: Esc]]. (In PowerPoint for the web: [[path: Picture > Crop > Aspect Ratio]].)
:::

## Styles, effects and alt text

- **Picture styles** ([[path: Picture Format > Picture Styles]]) add a frame, a shadow or a soft edge in one click. **Picture Effects** adds just one effect, such as a shadow, glow or reflection.
- **Alt text** describes a picture for people who can't see it: right-click > **View Alt Text** (or [[path: Picture Format > Alt Text]]). Say what matters in a sentence: "A red kayak on a calm lake below green hills", not "kayak.jpg" or "image".
- **Decorative:** if a picture only decorates (a border, a background flourish), tick **Mark as decorative** in the same pane, so screen readers skip it.

## Screenshots and screen clippings

[[path: Insert > Screenshot]] (desktop) pastes a picture of another open window, or choose **Screen Clipping** to drag around just part of the screen. Close or hide anything private first: everything in the clipping becomes part of the slide.

::: drill D3.02
title: Style, describe and clip
badge: Picture Pro
starter: starters/week-03/P3-L1-pictures.pptx
platform: desktop
mos: 3.3.2, 3.3.3, 3.4.7
minutes: 8

- step: Apply a **picture style** to the photo on **Our fleet**, and add a picture **effect** such as a shadow or glow.
  check: object.is title="Our fleet" kind=picture styled=true effects=true
  hint: [[path: Picture Format > Picture Styles]], then [[path: Picture Format > Picture Effects]].
- step: Give both photos (on **Our fleet** and **The beach**) alt text that says what they show.
  check: object.is title="Our fleet" kind=picture alt=true
  check: object.is title="The beach" kind=picture alt=true
  hint: Right-click the photo > **View Alt Text**, then type one sentence.
- step: Open a weather website in your browser. Then, on **Today's weather**, insert a **screen clipping** of the forecast.
  check: object.count title="Today's weather" kind=picture min=1
  hint: Click the browser window, then go straight back to PowerPoint and choose [[path: Insert > Screenshot > Screen Clipping]]. The screen turns pale; drag around the forecast. (Mac: if nothing shows, allow PowerPoint in **System Settings > Privacy & Security > Screen Recording**.)
:::

::: quiz Q31.1
question: Which alt text is best for a photo of the shop's beach on a busy Saturday?
type: single
- [ ] beach.jpg
- [ ] Picture
- [x] Families launching kayaks from a crowded beach on a sunny day
- [ ] A photo I took last summer of the beach where our shop is, which shows a lot of people
explain: Good alt text says what the picture shows, and why it matters on this slide, in one short sentence.
:::

## Check your understanding

- Why drag a corner handle instead of a side handle?
- What happens to the parts of a photo you crop away?
- When should a picture be marked decorative instead of getting alt text?
