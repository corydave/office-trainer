---
type: lesson
id: PP5-L2
week: 5
order: 2
title: Compare and Combine Presentations
objectives: [W5.2]
clos: [CLO7]
minutes: 30
status: draft
---

::: callout why
Not every reviewer works in your shared file. Someone downloads a copy, edits it, and sends it back. Your job is to find what they changed, take the changes you want into your own copy, and leave the rest.
:::

::: callout note
**Which PowerPoint has Compare?** Microsoft removed **Review > Compare** from Microsoft 365 PowerPoint for Windows in 2025 (version 2502), and it was never in PowerPoint for Mac or the web. PowerPoint 2024, 2021 and older still have it. If yours does, use it (below). If not, use **side by side**: the result, and the drill's checks, are the same.
:::

## With Compare (PowerPoint 2024 and older)

1. Open **your** presentation.
2. [[path: Review > Compare]] (desktop), pick the reviewer's copy, then **Merge**. The **Revisions** pane opens.
3. The pane lists **Slide Changes** (changes to objects on the current slide) and **Presentation Changes** (slides added, deleted or moved, theme changes). Markers on the slide show each change: select a marker to see the change, and tick its box to apply it.
4. For each change, use [[path: Review > Accept]] or **Reject** (each has a menu: this change, all changes to this slide, or all changes to the presentation). **Previous** and **Next** step through them.
5. Select **End Review**. Changes you didn't accept are discarded, and the markup is removed.

Your file now has the changes you accepted: the two copies are combined into one.

## Side by side (current Microsoft 365)

1. Open **your** presentation and the reviewer's copy, and put the windows side by side ([[path: View > Arrange All]] on Windows, or drag them on a Mac or in two browser tabs).
2. Go through the reviewer's copy slide by slide, in Normal view, and note each change. Their comments, or a short list from them of what they changed, make this much faster: ask for one.
3. Make each change you accept in **your** copy. Copy and paste works for text and objects; [[path: Home > New Slide > Reuse Slides]] brings in a whole slide they added.
4. Close the reviewer's copy without saving, and save yours.

::: callout tip
Save a copy of your presentation before you compare, and accept changes one at a time when you're not sure. Undo works too, until you end the review.
:::

::: drill D5.03
title: Take the reviewer's good changes
badge: Change Merger
starter: starters/week-05/P5-L2-compare.pptx
files: starters/week-05/P5-L2-reviewer.pptx
platform: desktop
minutes: 10

Download **P5-L2-reviewer.pptx** first. Sam reviewed a copy and made three changes: a new title, a new kayak price, and red text on Hours.

- step: Open your copy (the practice file) and compare it with **P5-L2-reviewer.pptx**, with Compare or side by side.
  hint: [[path: Review > Compare]], pick the reviewer's file, then **Merge**. No Compare button? Open both files side by side instead.
- step: Accept the new title (**Summer kickoff 2027**) and the new kayak price (**$22 an hour**).
  check: text.present slide=1 value="Summer kickoff 2027"
  check: text.present title="Prices" value="$22 an hour"
  hint: With Compare: go to each slide, select the change in the Revisions pane (or the marker on the slide), then [[path: Review > Accept]]. Side by side: retype or paste the new title and price into your copy.
- step: Reject the red text color on **Hours**, then finish the review and save.
  check: text.is title="Hours" says="Open daily" colorChanged=false
  hint: With Compare: [[path: Review > Reject]] on that change, then [[path: Review > End Review]] and **Yes**. Side by side: leave the Hours text as it is in your copy, and close the reviewer's copy without saving.
:::

::: quiz Q52.1
question: In PowerPoint's Compare, you accepted one change. What happens to the changes you didn't accept when you select End Review?
type: single
- [ ] They're accepted automatically
- [x] They're discarded
- [ ] They're saved as comments
- [ ] They stay as markup until you delete them
explain: End Review keeps what you accepted and throws away the rest, so make your choices before you end it.
:::

## Check your understanding

- When would you use Compare instead of co-authoring?
- What's the difference between Slide Changes and Presentation Changes?
- Why save a copy before you start?
