---
type: lesson
id: PP4-L1
week: 4
order: 1
title: Transitions
objectives: [W4.1]
clos: [CLO6]
minutes: 30
status: draft
---

::: callout why
A transition is how one slide gives way to the next. A quick, quiet one keeps attention on what you're saying; a dramatic one marks a big moment. Used everywhere, the dramatic ones wear people out.
:::

## Apply a transition

Select a slide, then pick an effect on the **Transitions** tab. The star next to a thumbnail means it has a transition. To use the same one everywhere, select **Apply To All**.

- **Subtle** transitions (Fade, Push, Wipe, Split, Cut) suit most slides.
- **Exciting** and **Dynamic Content** transitions include the 3D ones (Vortex, Ripple, Cube, Gallery, Flip, Switch). Save them for a title slide or one big reveal.
- **Effect Options** changes the direction or style (Push **From Left**, Vortex **From Right**).
- **Duration** sets how long the effect takes, in seconds.

::: callout note
Fast flashing can trigger seizures for some people, and lots of motion can make others dizzy. Keep transitions short and calm, and never use one that flashes more than three times a second.
:::

::: drill D4.01
title: Fade everywhere, one 3D moment
badge: Smooth Mover
starter: starters/week-04/P4-L1-transitions.pptx
platform: web
mos: 5.1.1, 5.1.2
other: N12
minutes: 6

- step: Apply the **Fade** transition to every slide.
  check: transition.is title="Morning on the lake" type=fade
  check: transition.is title="See you on the water" type=fade
  hint: Select any slide, [[path: Transitions > Fade]], then **Apply To All**.
- step: Give the title slide a **3D** transition instead, such as **Vortex**, **Ripple** or **Cube**.
  check: transition.is slide=1 kind=3d
  hint: Select slide 1, then open the transitions gallery (the **More** arrow) and look in **Exciting**.
- step: Change its **Effect Options** to a different direction.
  check: transition.is slide=1 optionChanged=true
- step: Set its **Duration** to **2** seconds.
  check: transition.is slide=1 duration=2
  hint: [[path: Transitions > Timing > Duration]]: type 02.00.
:::

## Timing and sound

In the **Timing** group:

- **On Mouse Click** waits for you to click or press a key.
- **After** advances on its own after the time you set. Tick both, and the slide moves on after that time unless you click first.
- **Sound** plays a sound with the transition (desktop). Use sound rarely; most rooms find it distracting.

A slideshow that runs by itself (a photo loop in a lobby, a kiosk) uses **After** on every slide.

::: drill D4.02
title: Time the photo slides
badge: Timekeeper
starter: starters/week-04/P4-L1-transitions.pptx
platform: desktop
mos: 5.1.2
other: N12
minutes: 5

- step: Set the three photo slides (**Morning on the lake**, **The boathouse**, **Sunset paddle**) to advance **after 5 seconds**.
  check: transition.is title="Morning on the lake" advanceAfter=5
  check: transition.is title="The boathouse" advanceAfter=5
  check: transition.is title="Sunset paddle" advanceAfter=5
  hint: Select the three thumbnails ([[key: Shift]]+click), tick **After** in [[path: Transitions > Timing]] and type 00:05.00.
- step: Add a transition **sound** to the last slide, **See you on the water**.
  check: transition.is title="See you on the water" sound=true
  hint: [[path: Transitions > Timing > Sound]], then pick a short sound such as **Chime**.
:::

::: quiz Q41.1
question: Which choice fits a 20-slide class presentation best?
type: single
- [ ] A different 3D transition on every slide
- [x] One subtle transition on every slide, with a stronger one only where a new part begins
- [ ] Sound on every transition, so people stay awake
- [ ] No transitions, ever
explain: Consistent, quiet transitions keep attention on you. A stronger one at a section break tells the audience something new is starting.
:::

## Check your understanding

- What's the difference between On Mouse Click and After?
- When would you use Apply To All?
- Why should transitions avoid flashing?
