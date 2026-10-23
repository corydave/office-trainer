---
type: project
id: PP2-P2
week: 2
order: 6
title: "Project 2: Brand Template and Deck"
objectives: [W2.1, W2.2, W2.3, W2.4, W2.5]
clos: [CLO1, CLO2, CLO3]
minutes: 210
status: draft
rubricBadge: Brand Builder
---

::: callout why
Every organization with more than a few people has a template, and someone had to build it. This project has you build one properly, in the master, so anyone can make a good-looking deck from it without breaking it.
:::

## What you'll make

Two files for a club, a small business or an organization of your own (real, with permission, or made up):

1. **A template**, `Lastname_P2_Brand.potx`, with:
   - a **theme** other than Office Theme, with colors and fonts that fit the organization,
   - the **logo** (yours, or a simple one you make with shapes or text) and the **background** set on the **slide master**,
   - at least **two custom layouts**, renamed (for example **Quote**, **Team Photo** or **Big Number**), each with a placeholder that suits it,
   - the **handout master** and **notes master** set up with the organization's name.
2. **A 6-slide presentation**, `Lastname_P2_Deck.pptx`, made **from** the template, that uses both custom layouts and follows the design checklist from Lesson 4.

::: callout privacy
Use a made-up organization or one you have permission to represent. Don't copy a real company's logo for a pretend brand.
:::

::: callout ai
Optional: ask an AI tool for a color palette and a font pair for your organization. Check the contrast of its colors yourself (text on background, at least 4.5 to 1) before you use them, and note what you changed.
:::

## Check before you submit

Run this on **Lastname_P2_Deck.pptx**.

::: preflight P2-check
title: Preflight for your brand deck
starter: none
platform: desktop

- step: The deck has at least 6 slides.
  check: slide.count min=6
- step: It uses a theme other than Office Theme.
  check: theme.is changed=true
- step: The logo is on the slide master, not pasted on each slide.
  check: master.is picture=true
- step: The handout master is set up.
  check: master.is kind=handout
- step: Every slide has a title, graphics have alt text, and the deck has a Title property.
  check: access.is rules="titles,uniqueTitles,alt,title"
:::

## How it's graded

::: rubric P2
- P2-a | 30 | CLO2 | The slide master does the work: theme, fonts, background and logo are set once, not slide by slide.
- P2-b | 20 | CLO2 | Custom layouts are named, useful and used in the deck.
- P2-c | 30 | CLO3 | The deck follows the design checklist: contrast, consistent fonts, short text, one idea per slide.
- P2-d | 20 | CLO1 | The handout and notes masters are set up, and the template opens as a new presentation.
:::
