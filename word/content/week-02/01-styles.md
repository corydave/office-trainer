---
type: lesson
id: W2-L1
week: 2
order: 1
title: Styles and Style Sets
objectives: [W2.1]
clos: [CLO2]
minutes: 45
status: draft
---

::: callout why
A ten-page paper with twenty headings formatted by hand takes twenty changes to restyle. With styles, it takes one. Styles are also how screen readers, the Navigation pane and an automatic table of contents know what your headings are.
:::

## Your scenario

This week you're a student in a lab course at FLCC. Your field team wrote a report on litter and water clarity along a lake shore, and you've volunteered to turn the rough draft into a properly structured document. By the end of the week it will have real headings, working lists, page numbers, footnotes, a table of contents and citations, and you'll build a paper template you can reuse in every course.

## What a style is

A **style** is a saved set of formatting with a name: font, size, color, spacing, indents and more. When you apply **Heading 1** to a line, you're not just making it big and blue. You're telling Word "this is a main heading."

::: video
id: w2-l1-styles
url: TODO
title: Styles in two minutes
captions: none
duration: 2:30
:::

| Style | Use it for |
|---|---|
| **Normal** | Body text: every ordinary paragraph |
| **Title** | The document's title, once |
| **Heading 1** | Main sections (Introduction, Methods) |
| **Heading 2** | Sections inside a Heading 1 section |
| **Heading 3** | Sections inside a Heading 2 section, if you need them |

Find the styles in the **Styles gallery** on the Home tab ([[path: Home > Styles]]). Click anywhere in a paragraph and pick a style; you don't need to select the text.

::: callout tip
Keyboard shortcuts in desktop Word: [[key: Ctrl+Alt+1 / Cmd+Option+1]] for Heading 1, [[key: Ctrl+Alt+2 / Cmd+Option+2]] for Heading 2, and [[key: Ctrl+Shift+N / Cmd+Shift+N]] for Normal.
:::

::: quiz Q21.1
question: A heading looks right (bold, 14 pt) but doesn't appear in the Navigation pane. What's the most likely reason?
type: single
- [ ] The Navigation pane is turned off
- [x] It was formatted by hand instead of with a heading style
- [ ] It's too short to count as a heading
- [ ] The document hasn't been saved
explain: The Navigation pane, a table of contents and screen readers find headings by their style, not their look. Bold 14 pt Normal text is still body text to Word.
:::

::: drill D2.01
title: Apply heading styles to the report
badge: Style Starter
starter: starters/week-02/W2-L1-report.docx
platform: web
mos: 2.2.4
minutes: 8

The section titles in the draft were made bold and bigger by hand. Give them real styles.

- step: Apply the **Title** style to the report's title, **Keeping the Lake Clean**.
  check: style.is para="Keeping the Lake Clean" name="Title"
- step: Apply **Heading 1** to the main sections: **Introduction**, **Background**, **Methods**, **Results**, **Discussion**, **Limitations**, **Recommendations** and **Appendix: Sampling Data**.
  check: style.is para="Introduction" name="Heading 1"
  check: style.is para="Methods" name="Heading 1"
  check: style.is para="Recommendations" name="Heading 1"
  check: style.is para="Appendix: Sampling Data" name="Heading 1"
  hint: Click in a heading, then choose Heading 1 in [[path: Home > Styles]]. Repeat for each.
- step: Apply **Heading 2** to the four sub-sections: **Why the shoreline matters**, **What earlier surveys found**, **Water clarity** and **Litter counts**.
  check: style.is para="Why the shoreline matters" name="Heading 2"
  check: style.is para="Water clarity" name="Heading 2"
  check: style.is para="Litter counts" name="Heading 2"
- step: Open the Navigation pane ([[path: View > Navigation Pane]]) and check that every heading appears in the outline.
:::

## Change a style once, change it everywhere

To change how every Heading 1 looks, change the **style**, not each heading.

- Desktop Word: right-click the style in the gallery and choose **Modify**. The dialog has font and paragraph settings, and a **Format** button for everything else.
- Desktop Word also has a shortcut: format one heading the way you want, then right-click the style and choose **Update Heading 1 to Match Selection**.

**Keep with next** (in the Paragraph settings, on the **Line and Page Breaks** tab) stops a heading from sitting alone at the bottom of a page with its text on the next page. Heading styles have it on already; check that it stays on.

## Style sets

A **style set** changes the look of every built-in style at once: Normal, Title and every heading. Find them on the **Design** tab ([[path: Design > Document Formatting]]). A style set changes only the formatting. Your headings stay headings.

::: callout warning
Apply the style set **before** you modify individual styles. A style set replaces style settings, including changes you made yourself.
:::

::: drill D2.02
title: Modify styles and try a style set
badge: Style Shaper
starter: starters/week-02/W2-L1-report.docx
platform: desktop
mos: 1.2.2
other: E2.3, E2.2
minutes: 10

Many instructors want papers in a classic serif font. Set that up once, in the styles.

- step: Apply a different style set from [[path: Design > Document Formatting]] (any one you like), and look at what changes.
  screenshot: The Design tab with your chosen style set highlighted, and the report's headings in the new look.
  example: TODO
  describe: Name the style set you chose and one thing about the headings that changed.
- step: Modify the **Normal** style to **Times New Roman, 12 pt**.
  check: style.def name="Normal" font="Times New Roman" size=12
  hint: Right-click **Normal** in the Styles gallery, choose **Modify**, then set the font and size. Leave **Only in this document** selected.
- step: Modify **Heading 1**: **14 pt**, **bold**, **black** (Automatic is fine), **12 pt** space before, and **Keep with next** on.
  check: style.def name="Heading 1" size=14 bold=true color=#000000 spaceBefore=12pt keepNext=true
  hint: In the Modify Style dialog, set the font settings, then choose **Format > Paragraph** for the spacing and **Line and Page Breaks** for Keep with next.
:::

::: quiz Q21.2
question: You applied a style set, then modified Heading 1. Then you applied a different style set. What happens to your Heading 1 changes?
type: single
- [ ] They're kept, because you made them last
- [x] They're replaced by the new style set's Heading 1
- [ ] Word asks which version you want
- [ ] Only the font changes; the spacing stays
explain: A style set rewrites the built-in styles. That's why you choose the style set first and fine-tune afterwards.
:::

## Check your understanding

- Why does a heading formatted by hand cause problems later?
- What's the fastest way to make every Heading 2 in a 40-page report dark green?
- What does Keep with next prevent?
