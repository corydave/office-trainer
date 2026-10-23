---
type: lesson
id: PP3-L4
week: 3
order: 4
title: SmartArt
objectives: [W3.4]
clos: [CLO2, CLO3]
minutes: 30
status: draft
---

::: callout why
Some lists are really diagrams: steps in a process, a cycle that repeats, a team with a leader. SmartArt turns a list into that diagram, and keeps it editable as text.
:::

## Insert and convert

- **From scratch:** [[path: Insert > SmartArt]], pick a group (List, Process, Cycle, Hierarchy ...) and a layout, then type in the **Text Pane** (the arrow at the graphic's left edge opens it).
- **From a list:** select a bulleted list's placeholder, then [[path: Home > Convert to SmartArt]] and pick a layout. Each bullet becomes a shape; indented bullets become sub-points.
- **Back to a list:** select the graphic, then [[path: SmartArt Design > Convert > Convert to Text]].

## Change it

| To | Do this |
|---|---|
| Add a shape | [[path: SmartArt Design > Add Shape]], or press [[key: Enter]] at the end of a line in the Text Pane |
| Change the layout | [[path: SmartArt Design > Layouts]] (your text carries over) |
| Change the colors | [[path: SmartArt Design > Change Colors]] |
| Change the look | [[path: SmartArt Design > SmartArt Styles]] |

Give the whole graphic alt text that says what it shows, such as "Four safety steps, in order".

::: drill D3.07
title: Turn a list into a diagram
badge: SmartArt Builder
starter: starters/week-03/P3-L4-smartart.pptx
platform: desktop
mos: 4.3.1, 4.3.2, 4.3.3
minutes: 10

- step: On **Safety steps**, convert the list to a SmartArt **Basic Process**.
  check: object.is title="Safety steps" kind=smartart
  hint: Click in the list, then [[path: Home > Convert to SmartArt > More SmartArt Graphics > Process > Basic Process]].
- step: Add a fourth step: **Have fun**.
  check: object.is title="Safety steps" kind=smartart items=4
  hint: Select the last shape, then [[path: SmartArt Design > Add Shape > Add Shape After]], and type.
- step: The steps repeat every trip. Change the layout to any **Cycle** layout, and change the colors.
  check: object.is title="Safety steps" kind=smartart layout=cycle colorsChanged=true
  hint: [[path: SmartArt Design > Layouts]], select the **More** arrow, then **More Layouts > Cycle**, then [[path: SmartArt Design > Change Colors]].
- step: Give the graphic alt text.
  check: access.is rules=alt
- step: On **Our values**, convert the list to any SmartArt, then use **Convert to Text** to turn it back into a list.
  hint: [[path: SmartArt Design > Convert > Convert to Text]]. You end up where you started, which is the point: SmartArt and lists convert both ways.
:::

::: quiz Q34.1
question: Which SmartArt group fits "Reserve, check in, paddle, return"?
type: single
- [ ] Hierarchy
- [x] Process
- [ ] Relationship
- [ ] Matrix
explain: Steps that happen in order are a process. A hierarchy shows who reports to whom.
:::

## Check your understanding

- What happens to indented bullets when you convert a list to SmartArt?
- Where do you type to add a shape quickly?
- When would you convert SmartArt back to text?
