---
type: lesson
id: W1-L3
week: 1
order: 3
title: Format Text and Paragraphs
objectives: [W1.3]
clos: [CLO1, CLO2]
minutes: 55
status: draft
---

::: callout why
Formatting tells the reader what matters before they read a word. Consistent fonts and spacing look professional. Random ones look like a document stitched together from three sources, which is usually what it is.
:::

Word formats at two levels:

- **Character formatting** changes selected letters: font, size, color, bold, italic, underline, text effects.
- **Paragraph formatting** changes a whole paragraph (everything up to the `¶`): alignment, line spacing, space before and after, indents.

You don't have to select a whole paragraph to change its paragraph formatting. Just click anywhere inside it.

## Character formatting

::: video
id: w1-l3-fonts-format-painter
url: TODO
title: Fonts, sizes, effects and the Format Painter
captions: none
duration: 4:00
:::

Most of what you need is in [[path: Home > Font]]:

| Command | Windows | Mac |
|---|---|---|
| Bold | [[key: Ctrl+B]] | [[key: Cmd+B]] |
| Italic | [[key: Ctrl+I]] | [[key: Cmd+I]] |
| Underline | [[key: Ctrl+U]] | [[key: Cmd+U]] |
| Bigger font | [[key: Ctrl+Shift+>]] | [[key: Cmd+Shift+>]] |
| Smaller font | [[key: Ctrl+Shift+<]] | [[key: Cmd+Shift+<]] |

**Aptos** is the default body font in current versions of Microsoft 365. Most business documents use 10 to 12 point body text.

**Text effects** ([[path: Home > Font > Text Effects and Typography]]) add a shadow, glow, reflection or outline. Use them for one thing, such as a letterhead. On body text they make reading harder.

::: platform web
Word for the web has a shorter Text Effects list than desktop Word. If an effect you need is missing, open the file in the desktop app with **Editing > Open in Desktop App**.
:::

### The Format Painter

The **Format Painter** ([[path: Home > Clipboard > Format Painter]]) copies formatting from one place to another:

1. Click in the text whose look you want to copy.
2. Select **Format Painter** (the paintbrush).
3. Drag across the text you want to change.

Double-click the paintbrush to keep it on for several places; press [[key: Esc]] to turn it off.

::: drill D1.06
title: Set fonts, bold, a text effect and use the Format Painter
badge: Font Stylist
starter: starters/week-01/W1-L3-letter.docx
platform: web
mos: 2.2.1, 2.2.2
other: N6
minutes: 8

- step: Change the body of the letter (from **Dear Volunteers,** through **Sincerely,**) to **Aptos, 11 pt**.
  check: font.is para="Thank you for" name="Aptos" size=11
  check: font.is para="Volunteer day starts" name="Aptos" size=11
  hint: Select from "Dear Volunteers," to "Sincerely,", then set the font and size in [[path: Home > Font]].
- step: Make the **Subject:** line bold.
  check: font.is para="Subject:" bold=true
- step: Apply any text effect (shadow, glow, reflection or outline) to the letterhead, **Lakeshore Paddle Company**, at the top.
  check: effect.present para="Lakeshore Paddle Company" kind=any
  hint: Select the letterhead, then [[path: Home > Font > Text Effects and Typography]].
- step: Use the Format Painter to copy the letterhead's formatting to the memo heading, **MEMORANDUM**, on page 2.
  check: format.same a="Lakeshore Paddle Company" b="MEMORANDUM"
  hint: Click in the letterhead, select the paintbrush once, then drag across MEMORANDUM.
:::

## Paragraph formatting

::: video
id: w1-l3-paragraph-spacing
url: TODO
title: Alignment, spacing and indents
captions: none
duration: 4:30
:::

| Setting | Where | Notes |
|---|---|---|
| **Alignment** | [[path: Home > Paragraph]] | Left is easiest to read. Justified (straight both edges) can leave uneven gaps between words |
| **Line spacing** | [[path: Home > Paragraph > Line and Paragraph Spacing]] | Space between the lines inside a paragraph: 1.0 (single), 1.15, 1.5, 2.0 (double) |
| **Space before / after** | [[path: Layout > Paragraph > Spacing]] | Space between paragraphs. This replaces pressing Enter twice |
| **Indents** | [[path: Layout > Paragraph > Indent]], or the Paragraph dialog | A **first-line indent** moves only the first line in. Never use spaces or tabs to indent |

For the full set, open the **Paragraph** dialog: select the small arrow in the corner of the Paragraph group, or right-click the paragraph and choose **Paragraph**.

::: callout tip
Space After is how professionals separate paragraphs. A common business setting is single line spacing with 8 pt after each paragraph.
:::

### Clear Formatting

Text pasted from email or the web often carries an odd font, color or highlight. **Clear All Formatting** ([[path: Home > Font > Clear All Formatting]], the letter A with an eraser) removes it and returns the text to its style's default look.

::: drill D1.07
title: Spacing, indents, alignment and Clear Formatting
badge: Paragraph Pro
starter: starters/week-01/W1-L3-letter.docx
platform: web
mos: 2.2.3, 2.2.5
other: N7
minutes: 8

- step: Give the letter's body paragraphs **single** line spacing with **8 pt** after.
  check: para.is para="Thank you for" lineSpacing=single spaceAfter=8pt
  check: para.is para="Volunteer day starts" lineSpacing=single spaceAfter=8pt
  hint: Select the body paragraphs, open the Paragraph dialog, set **Line spacing: Single** and **After: 8 pt**.
- step: The letter is justified. Change its body paragraphs to **left** alignment.
  check: para.is para="Thank you for" align=left
  check: para.is para="Volunteer day starts" align=left
  hint: [[key: Ctrl+L / Cmd+L]] aligns left.
- step: In the memo on page 2, give the paragraph that starts **Please review** a **0.5 in first-line indent**.
  check: para.is para="Please review" firstLineIndent=0.5in
  hint: Paragraph dialog > Indentation > Special: **First line**, By: **0.5"**.
- step: The paragraph that starts **Volunteers should bring** was pasted from an email. Clear its formatting.
  check: format.clean para="Volunteers should bring"
  hint: Select the whole paragraph, then [[path: Home > Font > Clear All Formatting]].
:::

::: quiz Q3.1
question: You click once inside a paragraph and change its line spacing. What changes?
type: single
- [ ] Only the word you clicked on
- [ ] Only the line you clicked on
- [x] The whole paragraph
- [ ] The whole document
explain: Line spacing is paragraph formatting, so it applies to the entire paragraph you are in, even with nothing selected.
:::

## Check your understanding

- What is the difference between character and paragraph formatting? Give one example of each.
- Why use Space After instead of pressing Enter twice?
- When is a text effect a good idea, and when is it a bad one?
