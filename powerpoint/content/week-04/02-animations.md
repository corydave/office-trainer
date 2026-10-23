---
type: lesson
id: PP4-L2
week: 4
order: 2
title: Animations
objectives: [W4.2]
clos: [CLO6]
minutes: 45
status: draft
---

::: callout why
If your whole list appears at once, people read ahead and stop listening. Revealing one point at a time keeps them with you. Animation is for pacing and pointing, not decoration.
:::

## Four kinds of animation

| Kind | Color in the gallery | What it does |
|---|---|---|
| **Entrance** | Green | Brings something onto the slide (Appear, Fade, Fly In, Wipe) |
| **Emphasis** | Yellow | Draws attention to something already there (Pulse, Spin, Grow/Shrink) |
| **Exit** | Red | Takes something off the slide (Disappear, Fade, Fly Out) |
| **Motion path** | Gray lines | Moves something along a line, an arc or a shape you draw |

Select the object, then pick an effect on the **Animations** tab. To add a **second** effect to the same object, use [[path: Animations > Add Animation]]; picking from the main gallery replaces the first one.

For a list, **Effect Options > Sequence > By Paragraph** makes each bullet appear on its own click.

::: drill D4.03
title: Reveal, point and remove
badge: Animator
starter: starters/week-04/P4-L2-animations.pptx
platform: web
mos: 5.2.1, 5.2.3
minutes: 8

- step: On **Why rent with us**, give the bulleted list an **entrance** effect that brings in one bullet at a time.
  check: anim.is title="Why rent with us" says="Safe gear" class=entrance byParagraph=true
  hint: Select the list's border, pick **Fade** (or any green effect), then [[path: Animations > Effect Options > By Paragraph]].
- step: Give the logo an **emphasis** effect, such as **Spin** or **Grow/Shrink**.
  check: anim.is title="Why rent with us" name=Logo class=emphasis
- step: Give the sticky note **Book by Friday!** an **exit** effect.
  check: anim.is title="Why rent with us" says="Book by Friday" class=exit
:::

## Timing and order

The **Timing** group (and the **Animation Pane**, [[path: Animations > Animation Pane]]) controls when each effect runs:

- **Start: On Click** waits for a click. **With Previous** starts at the same time as the effect before it. **After Previous** starts when the one before it finishes.
- **Duration** is how long the effect takes; **Delay** waits before it starts.
- **Reorder Animation > Move Earlier / Move Later**, or drag in the Animation Pane, changes the order. The numbers on the slide show the click order.

::: drill D4.04
title: Pace the animations
badge: Pace Setter
starter: starters/week-04/P4-L2-animations.pptx
platform: web
mos: 5.2.3, 5.2.5
minutes: 6

Do this after D4.03, in the same file.

- step: Set the logo's effect to start **After Previous**, with a **1-second** delay.
  check: anim.is title="Why rent with us" name=Logo start=after minDelay=1
  hint: Select the logo, then set **Start** and **Delay** in [[path: Animations > Timing]].
- step: Make the list's entrance last **1 second**.
  check: anim.is title="Why rent with us" says="Safe gear" duration=1
- step: Move the logo's animation so it runs **first**.
  check: anim.is title="Why rent with us" name=Logo first=true
  hint: Select the logo, then [[path: Animations > Reorder Animation > Move Earlier]] until it's number 0 or 1, or drag it to the top of the Animation Pane.
:::

## Motion paths

[[path: Animations > Motion Paths]] (more in **Add Animation > More Motion Paths**) moves an object along a line. A green arrow marks where it starts and a red one where it stops. **Effect Options** sets the direction, and **Reverse Path Direction** swaps the start and end.

::: drill D4.05
title: Paddle to the dock
badge: Path Finder
starter: starters/week-04/P4-L2-animations.pptx
platform: desktop
mos: 5.2.4
minutes: 5

- step: On **Our route**, give the kayak a **Lines** motion path.
  check: anim.is title="Our route" name=Kayak class=path
  hint: Select the kayak, then [[path: Animations > Motion Paths > Lines]].
- step: The dock is on the left. Set the path so the kayak moves **left**, toward the dock.
  check: anim.is title="Our route" name=Kayak pathDir=left
  hint: [[path: Animations > Effect Options > Left]], or **Reverse Path Direction** if it's going right.
:::

::: quiz Q42.1
question: You want the logo to spin right after the title appears, without another click. What should its Start be?
type: single
- [ ] On Click
- [ ] With Previous
- [x] After Previous
- [ ] It can't be done without a click
explain: After Previous starts an effect as soon as the one before it finishes. With Previous would run both at the same time.
:::

## Check your understanding

- Why does By Paragraph help an audience listen?
- What's the difference between With Previous and After Previous?
- How do you add a second effect to the same object without replacing the first?
