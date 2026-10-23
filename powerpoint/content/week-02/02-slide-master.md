---
type: lesson
id: PP2-L2
week: 2
order: 2
title: The Slide Master and Layouts
objectives: [W2.2]
clos: [CLO2]
minutes: 45
status: draft
---

::: callout why
Putting a logo on 20 slides one at a time takes 20 tries and one of them ends up crooked. Putting it on the **slide master** takes one, and every slide, including the ones you add next month, gets it in the same spot.
:::

## How the master works

[[path: View > Slide Master]] (desktop) opens Slide Master view. The thumbnails on the left are:

- the **slide master** at the top (the larger thumbnail): what you change here changes **every** slide,
- the **layouts** under it (Title Slide, Title and Content, Two Content ...): what you change on a layout changes only the slides that use it.

Change the master's background with [[path: Slide Master > Background Styles]] (or **Format Background**), its fonts with **Fonts**, and its title style by selecting the title placeholder and changing the font. Insert a picture on the master to put it on every slide. Then select **Close Master View**.

::: callout note
Slides keep anything you formatted by hand. If a slide doesn't pick up a master change, select it in Normal view and use [[path: Home > Reset]].
:::

::: video
id: p2-l2-master
url: TODO
title: The slide master and its layouts
captions: none
duration: 4:00
:::

::: drill D2.03
title: Brand the slide master
badge: Master Builder
starter: starters/week-02/P2-L2-master.pptx
files: starters/week-02/logo.png
platform: desktop
mos: 1.1.1, 1.1.2
minutes: 8

Download **logo.png** first (it's in the Week 2 practice files).

- step: In Slide Master view, select the slide master (the top thumbnail) and change its background with **Background Styles** or **Format Background**.
  check: master.is background=any
  hint: [[path: View > Slide Master]], select the top thumbnail, then [[path: Slide Master > Background Styles]] and pick any style except the first.
- step: Insert **logo.png** on the slide master, at the bottom right, so it appears on every slide.
  check: master.is picture=true
  hint: With the master selected, [[path: Insert > Pictures > This Device]]. Shrink it and drag it to the corner. Give it alt text: **Lakeshore Paddle Company logo**.
- step: Change the master title style's font to **Georgia**.
  check: master.is titleFont=Georgia
  hint: Select the border of the title placeholder on the master, then pick Georgia in [[path: Home > Font]].
- step: Select **Close Master View** and check that the logo and the font show on every slide.
:::

## Create and change layouts

Layouts are reusable slide designs. In Slide Master view:

| To | Do this |
|---|---|
| Create a layout | [[path: Slide Master > Insert Layout]], then **Rename** it |
| Add a placeholder | [[path: Slide Master > Insert Placeholder]], pick **Content**, **Text**, **Picture**, **Chart**, **Table** or **Media**, then drag to draw it |
| Remove a placeholder | Select its border on the layout and press [[key: Delete]] |
| Hide the title or footers | Clear **Title** or **Footers** in the Master Layout group |

Back in Normal view, the new layout is in [[path: Home > New Slide]] and [[path: Home > Layout]].

::: drill D2.04
title: Make a Quote layout
badge: Layout Designer
starter: starters/week-02/P2-L2-master.pptx
platform: desktop
mos: 1.1.3, 1.1.4
minutes: 10

- step: Insert a new layout and rename it **Quote**.
  check: layout.is name=Quote custom=true
  hint: [[path: View > Slide Master > Insert Layout]], then [[path: Slide Master > Rename]].
- step: Add a **Picture** placeholder to the Quote layout, on the left half.
  check: layout.is name=Quote placeholder=pic
  hint: [[path: Slide Master > Insert Placeholder > Picture]], then drag to draw it.
- step: On the **Section Header** layout, delete the **date** placeholder (bottom left).
  check: layout.is name="Section Header" lacks=dt
- step: Close Master View and apply the **Quote** layout to the slide **A word from our founder**.
  check: slide.is title="A word from our founder" layout=Quote
  hint: Select the slide, then [[path: Home > Layout > Quote]].
:::

::: quiz Q22.1
question: You want a "Draft" label on every slide while the deck is being reviewed. Where should you put it?
type: single
- [ ] On each slide, in a text box
- [x] On the slide master
- [ ] In the speaker notes
- [ ] In a comment
explain: Anything on the slide master appears on every slide, and you can remove it in one place when the review is over.
:::

## Check your understanding

- What's the difference between changing the slide master and changing one layout?
- Why might a slide not change when you edit the master?
- When is a custom layout worth making?
