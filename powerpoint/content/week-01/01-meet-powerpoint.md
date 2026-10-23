---
type: lesson
id: PP1-L1
week: 1
order: 1
title: Meet PowerPoint
objectives: [W1.1]
clos: [CLO1, CLO2]
minutes: 40
status: draft
---

::: callout why
A presentation is a set of slides you talk over, not a document people read. Once you know your way around the window and the views, you can build slides quickly, keep the details in your notes, and hand the file to someone else without confusing them.
:::

## Your scenario

All five weeks follow **Lakeshore Paddle Company**, a small kayak, canoe and paddleboard rental shop on a Finger Lakes beach. This week you'll build the orientation presentation for its new summer staff, then make your own presentation on something you know well.

## The PowerPoint window

| Part | What it's for |
|---|---|
| **Ribbon** | Commands, grouped on tabs such as Home, Insert, Design, Transitions, Animations, Slide Show, Review and View |
| **Slide pane** (left) | Thumbnails of every slide. Select one to edit it, drag one to move it |
| **Slide** (center) | The slide you're editing. Text goes in **placeholders**, the dotted boxes that say "Click to add title" |
| **Notes pane** (below the slide) | Your **speaker notes**: what you'll say. The audience doesn't see them. If it's hidden, select **Notes** on the status bar |
| **Status bar** (bottom) | The slide number, the Notes button, the view buttons and the zoom slider |
| **Quick Access Toolbar** | Your own one-click buttons (desktop PowerPoint). If you don't see it, right-click the Ribbon and choose **Show Quick Access Toolbar** |

A **placeholder** comes from the slide's **layout**. Type in placeholders instead of drawing your own text boxes: the text then follows the theme, appears in Outline view, and is read in the right order by screen readers.

::: video
id: p1-l1-tour
url: TODO
title: A tour of the PowerPoint window
captions: none
duration: 3:00
:::

::: drill D1.01
title: Save the presentation with a clear name
badge: File Namer
starter: starters/week-01/P1-L1-welcome.pptx
platform: web
other: N1
minutes: 3

- step: Open the practice file, then save it in your course folder as **Lastname_P1_Welcome.pptx** (with your own last name).
  check: file.name pattern="*_P1_Welcome.pptx"
  hint: [[path: File > Save As]], choose your course folder, and type the name. In PowerPoint for the web, select the file name at the top to rename it.
:::

## Views

Each view shows the same slides in a different way. Switch on the **View** tab, or with the buttons at the bottom right.

| View | Use it to |
|---|---|
| **Normal** | Build slides: thumbnails, the slide and your notes |
| **Outline** (desktop) | Write and reorganize the words. Every slide title and bullet is listed down the left, like a Word outline |
| **Slide Sorter** | See the whole presentation at once and drag slides into a new order |
| **Notes Page** (desktop) | See one slide with its notes under it, the way notes print. PowerPoint for the web has a **Notes** button that shows the Notes pane instead |
| **Reading View** | Watch the show in a window, without going full screen |
| **Slide Show** | Present: full screen, one slide at a time. In desktop PowerPoint, [[key: F5]] starts from the beginning and [[key: Shift+F5]] from the current slide (in PowerPoint for the web, use the **Slide Show** tab). [[key: Esc]] ends it |

Zoom with the slider at the bottom right, or [[key: Ctrl+mouse wheel]]. **Fit slide to current window** (the button next to the slider) brings the whole slide back into view.

::: drill D1.02
title: Work in every view
badge: View Finder
starter: starters/week-01/P1-L1-welcome.pptx
platform: web
mos: 1.2.2
other: N2, N6, N9
minutes: 8

- step: Change the title of slide 2 to **Your first day at the beach**. In desktop PowerPoint, do it in **Outline** view.
  check: slide.is slide=2 titleIs="Your first day at the beach"
  hint: [[path: View > Outline View]], click at the end of "Your first day" in the outline and type. Outline view is desktop only; in PowerPoint for the web, type the new title on the slide.
- step: On slide 3, add the note **Show the map of the beach.** In desktop PowerPoint, do it in **Notes Page** view.
  check: slide.is slide=3 notes="map of the beach"
  hint: [[path: View > Notes Page]], then click in the box under the slide. In PowerPoint for the web, select **Notes** on the status bar and type in the Notes pane.
- step: Open **Reading View**, move through the slides with the arrow keys, then press [[key: Esc]]. Back in Normal view, zoom in to 150% and then fit the slide to the window.
  hint: In desktop PowerPoint, select the zoom percentage at the bottom right to open the Zoom box and type 150; the button to the right of the slider fits the slide. In PowerPoint for the web, use the + and - buttons next to the zoom percentage.
- step: Switch to **Slide Sorter** view and save the presentation while it's in that view.
  platform: desktop
  check: view.is view=sorter
  hint: [[path: View > Slide Sorter]], then [[key: Ctrl+S]]. PowerPoint remembers the view you saved in. (PowerPoint for the web always opens in Normal view, so this step needs desktop PowerPoint.)
:::

::: quiz Q11.1
question: You want to drag slide 9 so it comes right after slide 2, and see the whole deck while you do it. Which view is best?
type: single
- [ ] Notes Page
- [x] Slide Sorter
- [ ] Reading View
- [ ] Slide Show
explain: Slide Sorter shows every slide as a thumbnail, so you can drag one to its new spot and see the order of the whole presentation.
:::

## The Quick Access Toolbar

The **Quick Access Toolbar** holds the commands you use all the time, one click away whatever tab is open. In desktop PowerPoint, select the arrow at its right end to add common commands, or right-click any button on the Ribbon and choose **Add to Quick Access Toolbar**. You can show it above or below the Ribbon.

::: drill D1.03
title: Customize the Quick Access Toolbar
badge: Toolbar Tuner
starter: starters/week-01/P1-L1-welcome.pptx
platform: desktop
other: N11
minutes: 4

- step: Add **New Slide** and **Start From Beginning** to the Quick Access Toolbar.
  screenshot: The PowerPoint window with the New Slide and Start From Beginning buttons on the Quick Access Toolbar.
  example: TODO
  describe: List the buttons now on your Quick Access Toolbar, in order.
  hint: Right-click **New Slide** on the **Home** tab and choose **Add to Quick Access Toolbar**. For Start From Beginning, select the arrow at the end of the toolbar and pick it from the list.
:::

::: quiz Q11.2
question: Where do you type what you plan to say about a slide, so the audience doesn't see it?
type: single
- [ ] In a small text box at the bottom of the slide
- [x] In the Notes pane under the slide
- [ ] In the slide's title
- [ ] In a comment
explain: Speaker notes live in the Notes pane. You see them in Presenter View and on printed notes pages; the audience sees only the slide.
:::

## Check your understanding

- Why type in a placeholder instead of drawing a text box?
- Which view would you use to rewrite all your slide titles quickly?
- What's the difference between Reading View and Slide Show?
