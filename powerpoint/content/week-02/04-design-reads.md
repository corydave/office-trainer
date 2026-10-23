---
type: lesson
id: PP2-L4
week: 2
order: 4
title: Design That Reads
objectives: [W2.5]
clos: [CLO3]
minutes: 35
status: draft
---

::: callout why
Most bad slides aren't ugly, they're hard to read: too many words, text that fades into the background, five fonts fighting. A few rules fix most of them, and they're the same rules that make slides accessible.
:::

## A design checklist

| Rule | Why |
|---|---|
| **One idea per slide**, with a title that says it | People read the title first. If the title is the point, they get it even if they miss the rest |
| **Six lines or fewer** in a list, short lines | A wall of text makes people read instead of listen. Split it, or move detail to the notes |
| **Strong contrast** between text and background | Light gray on white, or red on green, disappears on a projector. Aim for a contrast of at least **4.5 to 1** |
| **Two fonts at most**: the theme's heading and body fonts | Novelty fonts are hard to read and look unplanned |
| **24 pt or larger** for body text | The back row has to read it |
| **Color never carries the meaning alone** | Add a word, a label or a pattern, for people who can't tell the colors apart |

::: callout tip
**Replace Fonts** changes one font to another everywhere in the presentation: [[path: Home > Replace > Replace Fonts]] (desktop). In PowerPoint for the web, select the text on each slide and change the font, or better, use [[path: Home > Clear All Formatting]] so it goes back to the theme font.
:::

## Check the contrast

The **Accessibility Checker** ([[path: Review > Check Accessibility]]) flags text that's hard to read against its background. You'll use it fully in Week 5. For now, if text looks faint, give it a darker theme color (or a lighter one on a dark background).

::: drill D2.07
title: Fix a cluttered deck
badge: Declutterer
starter: starters/week-02/P2-L4-cluttered.pptx
platform: web
other: N13
minutes: 10

- step: Replace **Comic Sans MS** everywhere with the theme's font (or any plain font).
  check: text.is notFont="Comic Sans MS"
  hint: Desktop: [[path: Home > Replace > Replace Fonts]], replace Comic Sans MS with a plain font. Web: select each text box's border and use [[path: Home > Clear All Formatting]], or pick a font.
- step: On **Hours**, the pale gray text is hard to read. Change it to a dark color.
  check: text.is title="Hours" says="Open daily" minContrast=4.5
  hint: Select the list's border, then [[path: Home > Font Color]] and pick black or a dark theme color.
- step: **Rules for renters** has ten bullets. Split it into two slides, **Rules for renters (1 of 2)** and **Rules for renters (2 of 2)**, with five rules each.
  check: slide.count title="Rules for renters" min=2
  check: list.is title="Rules for renters" maxItems=6
  hint: Duplicate the slide ([[key: Ctrl+D]]), delete the last five rules from the first copy and the first five from the second, then fix the titles.
:::

::: quiz Q24.1
question: A chart shows profits in green and losses in red. What should you add so everyone can read it?
type: single
- [ ] A brighter green
- [x] Labels or a legend that say which is which, so color isn't the only clue
- [ ] An animation
- [ ] A darker background
explain: About 1 in 12 men can't easily tell red from green. Words, labels or patterns make the meaning clear without relying on color.
:::

## Check your understanding

- What's wrong with a slide titled "Slide 4" or "Update"?
- Where should the extra detail go when you cut a list down to six lines?
- Why is a novelty font a problem even if you like it?
