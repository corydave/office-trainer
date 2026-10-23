---
type: lesson
id: PP2-L1
week: 2
order: 1
title: Themes and Backgrounds
objectives: [W2.1]
clos: [CLO1, CLO2]
minutes: 35
status: draft
---

::: callout why
A theme sets the colors, fonts and backgrounds of every slide in one click, and keeps them consistent. Change the theme and the whole presentation changes with it, as long as you used placeholders and theme colors instead of formatting each slide by hand.
:::

## Themes and variants

A **theme** is a package of three things: a set of **colors**, a pair of **fonts** (one for headings, one for body text) and **effects**, plus background styles and layouts that use them.

- Apply a theme: [[path: Design > Themes]], then pick one. Point at a theme to preview it on the slide.
- Pick a **variant**: the thumbnails in [[path: Design > Variants]] are versions of the theme with different colors or backgrounds.
- Change only part of it (desktop): [[path: Design > Variants > More > Colors]] or **Fonts**. A font set such as **Georgia** changes every title and every bullet that uses the theme fonts.

::: callout tip
Theme fonts show as **(Headings)** and **(Body)** at the top of the font list. Text in those fonts changes with the theme. Text you set to a named font (say, Arial) stays Arial, which is why it's better to change the theme's fonts than to change each slide.
:::

::: drill D2.01
title: Apply a theme, a variant and new fonts
badge: Theme Setter
starter: starters/week-02/P2-L1-design.pptx
platform: web
other: N5
minutes: 6

- step: Apply any theme other than **Office Theme**, such as **Facet** if your PowerPoint has it.
  check: theme.is changed=true
  hint: [[path: Design > Themes]]. Select the **More** arrow at the bottom right of the gallery to see them all. (Microsoft updates the gallery now and then, so your list may differ from a classmate's.)
- step: Pick one of the theme's other **variants**.
  hint: [[path: Design > Variants]]: select a thumbnail and watch the colors change.
- step: Change the theme fonts to the **Georgia** font set.
  platform: desktop
  check: theme.is fonts=Georgia
  hint: [[path: Design > Variants > More > Fonts]], then **Georgia**.
:::

## Backgrounds for one slide

**Format Background** changes the background of the selected slide only: a **solid fill**, a **gradient fill**, a **picture or texture fill**, or a pattern. **Apply to All** changes every slide, which is better done once in the slide master (next lesson).

- Open it: [[path: Design > Format Background]], or right-click an empty part of the slide > **Format Background**.
- For a picture, choose **Picture or texture fill**, then **Insert** > **From a File** (desktop), or the picture option in the web app's Format Background pane.
- Keep text readable: a busy photo behind text needs **Transparency** turned up, or a plain area where the text sits.

::: drill D2.02
title: Give two slides their own backgrounds
badge: Backdrop
starter: starters/week-02/P2-L1-design.pptx
files: starters/week-02/lake.jpg
platform: web
mos: 2.2.2
minutes: 6

Download **lake.jpg** first (it's next to the practice file above).

- step: Give the title slide a **picture** background using **lake.jpg**.
  check: slide.is slide=1 background=picture
  hint: Select slide 1, then [[path: Design > Format Background > Picture or texture fill]], **Insert**, and pick lake.jpg. Don't select Apply to All.
- step: Give the **Thank you** slide a **solid fill** background in a dark color.
  check: slide.is title="Thank you" background=solid
- step: Make sure the other slides still use the theme's background.
  check: slide.is title="This summer" background=none
  check: slide.is title="Staff picnic" background=none
  hint: If every slide changed, you selected Apply to All. Press [[key: Ctrl+Z]] to undo it, then apply the picture to slide 1 again without Apply to All.
:::

::: quiz Q21.1
question: You set every title to Arial by hand, then apply a new theme. What happens to the titles?
type: single
- [ ] They change to the new theme's heading font
- [x] They stay Arial, because a named font overrides the theme font
- [ ] They turn into WordArt
- [ ] PowerPoint asks which font to keep
explain: Only text in the theme fonts, (Headings) and (Body), follows the theme. Use Home > Clear All Formatting, or change the theme fonts, to bring hand-formatted text back in line.
:::

## Check your understanding

- What three things does a theme set?
- When is Apply to All the wrong choice?
- How do you keep text readable on a picture background?
