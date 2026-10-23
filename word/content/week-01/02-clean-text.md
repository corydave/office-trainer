---
type: lesson
id: W1-L2
week: 1
order: 2
title: Clean Up the Text
objectives: [W1.4, W1.6]
clos: [CLO2]
minutes: 50
status: draft
---

::: callout why
Most messy documents are not badly written. They are badly typed: extra spaces, blank lines used as spacing, text pasted from email with the wrong font. Five minutes of cleanup makes a document look professional, and it is the difference between a document that behaves and one that falls apart when someone edits it.
:::

Your manager pasted the volunteer letter together from an old template and an email. Open the starter and you'll see the problems.

## See what Word sees: formatting marks

Word hides the characters that control layout. **Show/Hide** makes them visible:

| Mark | Means |
|---|---|
| `¶` | You pressed [[key: Enter]] (end of a paragraph) |
| `·` | A space |
| `→` | A tab |
| `↵` | A line break ([[key: Shift+Enter]]) |

Turn it on with [[path: Home > Paragraph > Show/Hide]] or [[key: Ctrl+Shift+8 / Cmd+8]]. The marks never print.

::: video
id: w1-l2-show-hide
url: TODO
title: Using Show/Hide to find hidden problems
captions: none
duration: 2:30
:::

Two problems show up right away: **empty paragraphs** (a `¶` alone on a line, used to make space) and **double spaces** (`··`). Both cause trouble later. Empty paragraphs shift around when text changes, and double spaces make lines uneven. In Lesson 3 you'll add space the right way, with paragraph spacing.

::: drill D1.03
title: Remove extra paragraphs and spaces
badge: Space Cleaner
starter: starters/week-01/W1-L2-letter.docx
platform: web
mos: 1.1.4, 2.1.1
minutes: 6

- step: Turn on **Show/Hide** so you can see the marks.
  hint: [[path: Home > Paragraph > Show/Hide]] (the ¶ button). In Word for the web, if you don't see it, type "show/hide" in the search box.
- step: Delete every empty paragraph in the letter.
  check: para.noEmpty max=0
  hint: Click just after an empty ¶ and press [[key: Backspace]]. Don't delete the ¶ at the end of a line with text on it, or two paragraphs will join.
- step: Replace every double space with a single space.
  check: text.noDoubleSpaces
  hint: Open Replace ([[key: Ctrl+H]]), type two spaces in **Find what** and one space in **Replace with**, then select **Replace All**. Repeat until Word finds none.
:::

## Find and Replace

**Find** ([[key: Ctrl+F / Cmd+F]]) opens the Navigation pane and highlights every match. **Replace** ([[path: Home > Editing > Replace]], or [[key: Ctrl+H]] on Windows) changes them, one at a time or all at once.

::: callout warning
**Replace All** is fast and dangerous. Replacing "Acme" also changes "Acmeville". Use **Match case** and **Find whole words only** (under **More >>** in desktop Word) when the word could appear inside another word.
:::

::: quiz Q2.1
question: Use Find to count them. How many times does the word "deadline" appear in the letter?
type: number
answer: 3
explain: Find shows the count at the top of the Navigation pane ("3 results").
:::

::: drill D1.04
title: Replace the old company name
badge: Find and Replace
starter: starters/week-01/W1-L2-letter.docx
platform: web
mos: 1.1.1, 2.1.1
minutes: 4

The letter still uses the name of the company it was copied from.

- step: Replace every **Acme Outdoor Supply** with **Lakeshore Paddle Company**.
  check: text.absent value="Acme"
  check: text.present value="Lakeshore Paddle Company" min=4
  hint: Open Replace. Find what: `Acme Outdoor Supply`. Replace with: `Lakeshore Paddle Company`. Select **Replace All**.
:::

## Symbols and special characters

Some characters aren't on the keyboard. Three ways to type them:

- **Insert > Symbol** shows a short list; **More Symbols** shows everything.
- **AutoCorrect** turns `(c)` into © and `(r)` into ® as you type.
- **Accented letters:** on Windows press [[key: Ctrl+']] then the letter (é). On a Mac press [[key: Option+E]] then the letter.

::: drill D1.05
title: Add a copyright symbol and an accent
badge: Symbol Finder
starter: starters/week-01/W1-L2-letter.docx
platform: web
mos: 2.1.2
minutes: 4

- step: In the footer, type a **©** symbol and a space in front of "2026 Lakeshore Paddle Company", so it reads **© 2026 Lakeshore Paddle Company**.
  check: text.present value="© 2026" in=footer
  hint: Double-click the footer area to edit it, then use [[path: Insert > Symbol]] or type `(c)`.
  web: In Word for the web, open the footer with [[path: Insert > Header & Footer]].
- step: In the second paragraph, change **cafe** to **café**.
  check: text.present value="café"
  check: text.absent value="cafe"
  hint: Delete the e, then type the accented é ([[key: Ctrl+']] then [[key: E]] on Windows, [[key: Option+E]] then [[key: E]] on a Mac).
:::

## Cut, copy, paste and Undo

| Action | Windows | Mac |
|---|---|---|
| Cut | [[key: Ctrl+X]] | [[key: Cmd+X]] |
| Copy | [[key: Ctrl+C]] | [[key: Cmd+C]] |
| Paste | [[key: Ctrl+V]] | [[key: Cmd+V]] |
| Undo | [[key: Ctrl+Z]] | [[key: Cmd+Z]] |
| Redo | [[key: Ctrl+Y]] | [[key: Cmd+Y]] |

Text pasted from email or a website brings its own font and color along. **Paste Options** fix that. After pasting, select the small clipboard icon (or [[path: Home > Paste]] arrow) and pick:

- **Keep Source Formatting** - keeps the pasted look.
- **Merge Formatting** - keeps bold and italic, takes on your document's font.
- **Keep Text Only** - plain text that matches the paragraph you paste into. This is usually the one you want.

::: drill D1.08
title: Move a paragraph, and paste it clean
badge: Clean Paster
starter: starters/week-01/W1-L2-letter.docx
platform: web
other: E2.1, N9, N10
minutes: 5

The paragraph that starts "We look forward" ended up below the signature, and it still has the red email font.

- step: Cut the paragraph that starts **We look forward** and paste it just above **Sincerely,** using **Keep Text Only**.
  check: para.order first="We look forward" then="Sincerely,"
  check: format.clean para="We look forward"
  hint: Select the whole paragraph (triple-click it), press [[key: Ctrl+X]], click at the start of "Sincerely,", then use [[path: Home > Paste > Keep Text Only]].
- step: Delete the word **Sincerely,** by accident, then bring it back with Undo.
  hint: [[key: Ctrl+Z]] undoes your last action. You can press it several times.
:::

## Proofread with Editor

**Editor** checks spelling, grammar, clarity and more. Red underlines are spelling, blue double underlines are grammar. Open the full pane with [[path: Review > Editor]] in desktop Word or [[path: Home > Editor]] in Word for the web, or press [[key: F7]] on Windows. You'll use it on the letter in Lesson 4.

::: callout note
Editor is a second pair of eyes, not a replacement for yours. It misses correctly spelled wrong words ("their" for "there") and sometimes suggests changes that alter your meaning.
:::

## Check your understanding

- Why is an empty paragraph a bad way to add space?
- When would you choose **Keep Text Only** over **Keep Source Formatting**?
- What should you check before selecting **Replace All**?
