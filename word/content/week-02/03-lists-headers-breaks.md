---
type: lesson
id: W2-L3
week: 2
order: 3
title: Lists, Headers and Breaks
objectives: [W2.3, W2.4]
clos: [CLO1, CLO2]
minutes: 55
status: draft
---

::: callout why
Lists typed by hand ("1." then a space) don't renumber when you add a step. Page numbers typed by hand are wrong after the first edit. Word's own lists, headers and breaks keep up with the document, so you only set them once.
:::

## Bulleted and numbered lists

Select the lines, then choose **Bullets** or **Numbering** on the Home tab ([[path: Home > Paragraph]]). The small arrow next to each button opens a library of other symbols and number styles.

| To | Do this |
|---|---|
| Make a sub-item (level 2) | Click in the item and press [[key: Tab]], or **Increase Indent** |
| Move an item back up a level | [[key: Shift+Tab]], or **Decrease Indent** |
| Change the bullet symbol | Bullets arrow, then pick a symbol (or **Define New Bullet**) |
| Change the number style (1, a, i) | Numbering arrow, then pick a format |

::: video
id: w2-l3-lists
url: TODO
title: Lists that renumber themselves
captions: none
duration: 3:30
:::

::: drill D2.05
title: Make real lists
badge: List Maker
starter: starters/week-02/W2-L3-report.docx
platform: web
mos: 3.3.1, 3.3.2, 3.3.4
other: N8
minutes: 7

In the Methods section, the sampling steps are just separate lines. Make them a real list.

- step: Turn the six lines under **Sampling steps** (from **Walk to the marked sampling site** to **Store it in the cooler**) into a numbered list.
  check: list.is para="Walk to the marked sampling site" kind=number
  check: list.is para="Store it in the cooler" kind=number
  hint: Select all six lines, then select **Numbering** in [[path: Home > Paragraph]].
- step: **Label it with the site number** and **Rinse it twice with lake water** are part of preparing the bottle. Make them level 2 items.
  check: list.is para="Label it with the site number" level=2
  check: list.is para="Rinse it twice" level=2
  check: list.is para="Fill the bottle" level=1
  hint: Click in each one and press [[key: Tab]]. The numbers on the other steps update by themselves.
- step: Change the bullets in the **Equipment** list to check marks (✓).
  check: list.is para="Sample bottles" symbol="✓"
  check: list.is para="Cooler with ice packs" symbol="✓"
  hint: Select the four equipment lines, open the **Bullets** arrow and pick the check mark.
:::

## Custom numbers, restarting and starting values

Sometimes numbering needs help. Right-click any number in a list to see the options:

- **Restart at 1:** this list starts over, instead of carrying on from the list above it.
- **Continue Numbering:** this list carries on from the last one.
- **Set Numbering Value:** start at any number you choose.
- **Define New Number Format** ([[path: Home > Numbering > Define New Number Format]]): type text around the number, such as `Step 1.`

::: drill D2.06
title: Fix numbering that runs on
badge: Number Wrangler
starter: starters/week-02/W2-L3-report.docx
platform: desktop
mos: 3.3.3, 3.3.5, 3.3.6
minutes: 8

Do D2.05 first: this drill uses the sampling steps list you made.

- step: Give the sampling steps a custom number format, so they read **Step 1.**, **Step 2.** and so on.
  check: list.is para="Walk to the marked sampling site" pattern="Step %1"
  hint: Click in the list, then [[path: Home > Numbering > Define New Number Format]]. In **Number format**, type `Step ` in front of the gray number, and keep the period after it.
- step: The **Recommendations** list starts at 4, because it continues the Findings list. Restart it at 1.
  check: list.is para="Reduce single-use plastic" number=1
  hint: Right-click the 4 and choose **Restart at 1**.
- step: The **Lab follow-up** steps come after the four field steps, so they should be numbered 5, 6 and 7. Set the first one to start at 5.
  check: list.is para="Filter each sample" number=5
  check: list.is para="Enter the results" number=7
  hint: Right-click the 1 and choose **Set Numbering Value**, then type 5.
:::

## Headers, footers and page numbers

Headers and footers repeat on every page. Open one by double-clicking the top or bottom margin, or with [[path: Insert > Header & Footer]].

- **Page numbers:** [[path: Insert > Page Number]]. Word inserts a field that always shows the right number.
- **Different First Page** (in the Header & Footer tab): the first page gets its own header and footer, or none. Title pages and many style guides need this.

::: callout note
APA 7th edition puts the page number at the top right of every page, including the title page. MLA 9th edition puts your last name and the page number at the top right. Always check your instructor's directions.
:::

::: drill D2.07
title: Add a running header with page numbers
badge: Header Hero
starter: starters/week-02/W2-L3-report.docx
platform: web
mos: 1.2.3
minutes: 6

- step: Insert a page number at the top right of the page, after your last name and a space (for example **Rivera 2**).
  check: field.present code="PAGE" in=header
  hint: Double-click the top margin, type your last name and a space, then [[path: Insert > Page Number > Current Position]] (desktop) or [[path: Insert > Page Numbers]] (web).
  web: In Word for the web, [[path: Insert > Page Numbers]] adds the number; choose the top-right position, then type your name in front of it.
- step: This field-study report's style keeps the header off page 1, where the title is. Turn on **Different First Page**, so page 1 has no header. (For APA and MLA papers in Project 2, leave it off: both put the header on page 1.)
  check: section.is index=1 firstPageDifferent=true
  hint: In the header, select [[path: Header & Footer > Different First Page]] (desktop), or **Options > Different First Page** (web).
:::

## Page, section and column breaks

A **break** tells Word where something new starts.

| Break | Where it goes | Use it for |
|---|---|---|
| **Page break** ([[key: Ctrl+Enter / Cmd+Enter]]) | The next text starts on a new page | A new chapter or the references page |
| **Section break: Next Page** | A new page and a new **section** | Pages that need their own orientation, margins, headers or numbering |
| **Section break: Continuous** | A new section on the same page | Changing the number of columns partway down a page (Week 3) |
| **Column break** | The next text starts in the next column | Balancing a newsletter column |

Insert them from [[path: Layout > Breaks]]. Turn on **Show/Hide** (¶) to see them.

::: drill D2.08
title: A landscape appendix
badge: Section Setter
starter: starters/week-02/W2-L3-report.docx
platform: desktop
mos: 2.3.2, 2.3.3
minutes: 6

The sampling data table is too wide for a portrait page. Give the appendix its own landscape section.

- step: Click at the start of the **Appendix: Sampling Data** heading and insert a **Next Page** section break.
  check: section.is para="Appendix: Sampling Data" breakType=nextPage
  hint: [[path: Layout > Breaks > Section Breaks > Next Page]]. With Show/Hide on, you'll see "Section Break (Next Page)".
- step: With the cursor in the appendix, change only that section to **Landscape**.
  check: section.is para="Appendix: Sampling Data" orientation=landscape
  check: section.is para="Introduction" orientation=portrait
  hint: [[path: Layout > Orientation > Landscape]]. Because of the section break, the pages before it stay portrait. If every page turned, check that the cursor was after the break.
:::

::: quiz Q23.1
question: You want one landscape page in the middle of a portrait report. What do you need?
type: single
- [ ] A page break before the page
- [x] A section break before and after the page
- [ ] A column break
- [ ] Different First Page
explain: Orientation belongs to a section. One landscape page in the middle needs its own section, so a section break before it and another after it.
:::

## Check your understanding

- Why is a numbered list better than typing the numbers yourself?
- What's the difference between a page break and a Next Page section break?
- Where do you turn on Different First Page?
