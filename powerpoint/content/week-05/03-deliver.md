---
type: lesson
id: PP5-L3
week: 5
order: 3
title: Deliver the Show
objectives: [W5.3]
clos: [CLO3, CLO1]
minutes: 40
status: draft
---

::: callout why
The slides are only half of a presentation. Rehearsing with a timer, presenting with your notes in front of you, and having a short version ready when time runs out are what make a talk go well.
:::

## Custom shows

A **custom show** is a named list of slides from the same file, such as a 5-minute version of a 15-minute talk. [[path: Slide Show > Custom Slide Show > Custom Shows]] (desktop) > **New**, name it, add slides, and put them in order. Run it from the same menu, or pick it under **Show slides** in Set Up Slide Show.

## Set Up Slide Show

[[path: Slide Show > Set Up Slide Show]] (desktop) has the options for how the show runs:

| Setting | Choices |
|---|---|
| **Show type** | **Presented by a speaker** (full screen), **Browsed by an individual** (in a window), **Browsed at a kiosk** (full screen, loops, no clicking ahead; it needs rehearsed timings or action buttons, or it stays on the first slide) |
| **Show options** | **Loop continuously until 'Esc'**, **Show without narration**, **Show without animation** |
| **Show slides** | All, a range, or a custom show |
| **Advance slides** | **Manually**, or **Using timings, if present** |

::: drill D5.04
title: A short tour for the lobby screen
badge: Show Planner
starter: starters/week-05/P5-L3-show.pptx
platform: desktop
mos: 1.4.1, 1.4.2
minutes: 7

- step: Create a custom show named **Short tour** with four slides, in this order: **Welcome**, **Our boats**, **Book online**, **Thank you**.
  check: show.is customShow="Short tour" slides=4 includes="Welcome,Our boats,Book online,Thank you" excludes="Prices"
  hint: [[path: Slide Show > Custom Slide Show > Custom Shows > New]]. Tick the slides on the left, then **Add**.
- step: Set the show to run **at a kiosk** and **loop until Esc**.
  check: show.is kiosk=true loop=true
  hint: [[path: Slide Show > Set Up Slide Show]], then **Browsed at a kiosk (full screen)**. Loop is ticked for you.
:::

## Rehearse, record and present

- **Rehearse Timings** ([[path: Slide Show > Rehearse Timings]]) runs the show with a timer. Talk through each slide as you would for real and click to go on. At the end, choose **Yes** to keep the timings; each slide is then set to advance after its time.
- **Recording options**: the **Slide Show** tab's **Play Narrations** and **Use Timings** boxes decide whether a recorded voice and the rehearsed timings are used when the show runs; **Show Media Controls** shows play controls on audio and video. [[path: Slide Show > Record]] records your voice (and camera) with each slide.
- **Presenter View** shows you the current slide, the next slide, your notes and a timer, while the audience sees only the slide. It turns on by itself with a second screen; with one screen, start the show, then right-click > **Show Presenter View** (or press [[key: Alt+F5]] on Windows; on a Mac, [[path: Slide Show > Presenter View]]).

::: drill D5.05
title: Rehearse and present
badge: Rehearsed
starter: starters/week-05/P5-L3-show.pptx
platform: desktop
mos: 1.4.3, 1.4.4, 1.4.5
other: N14
minutes: 15

- step: Rehearse the whole tour, saying a sentence or two per slide, and keep the timings.
  check: show.is timings=true
  hint: [[path: Slide Show > Rehearse Timings]]. At the end, select **Yes**.
- step: This presentation was set up to ignore timings and narration. Turn on **Use Timings** and **Play Narrations**.
  check: show.is useTimings=true narration=true
  hint: Tick both boxes in the **Set Up** group of the **Slide Show** tab (or, in Set Up Slide Show, clear **Show without narration** and choose **Using timings, if present**).
- step: Start the show in **Presenter View** and take a screenshot of it with your notes and the timer showing.
  screenshot: Presenter View with the current slide, the next slide, the notes and the timer.
  example: TODO
  describe: Which slide is next, and how long have you been presenting?
  hint: [[key: F5]], then right-click > **Show Presenter View**, or [[key: Alt+F5]].
:::

::: quiz Q53.1
question: Your talk got cut from 15 minutes to 5 the morning of the event. What's the fastest fix?
type: single
- [ ] Delete two thirds of the slides
- [ ] Save a second copy and edit it
- [x] Make a custom show with just the slides you need
- [ ] Turn on Loop continuously
explain: A custom show uses the same slides without changing the full presentation, so both versions stay ready.
:::

## Check your understanding

- What's the difference between a kiosk show and one browsed by an individual?
- What does Rehearse Timings save in the file?
- What can you see in Presenter View that the audience can't?
