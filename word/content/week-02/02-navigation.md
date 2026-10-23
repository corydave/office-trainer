---
type: lesson
id: W2-L2
week: 2
order: 2
title: Find Your Way Around Long Documents
objectives: [W2.2]
clos: [CLO2, CLO6]
minutes: 30
status: draft
---

::: callout why
Scrolling through a 20-page document to find one section wastes minutes every time. Long documents have shortcuts built in, and they all depend on the heading styles from the last lesson.
:::

## The Navigation pane

Turn it on with [[path: View > Navigation Pane]] or [[key: Ctrl+F / Cmd+F]]. It has three views:

- **Headings:** an outline of the document. Click a heading to jump to it. **Drag** a heading up or down to move the whole section, text and all.
- **Pages:** small pictures of each page (desktop Word).
- **Results:** every match for what you type in the search box.

::: video
id: w2-l2-navigation
url: TODO
title: Moving sections with the Navigation pane
captions: none
duration: 3:00
:::

::: callout tip
Right-click a heading in the Navigation pane for more: promote, demote, add a new heading before or after, or collapse every heading at that level.
:::

## Go To

**Go To** jumps straight to a page, section, table, comment or other object. Open it with [[key: Ctrl+G / Cmd+Option+G]] or [[path: Home > Find > Go To]] (desktop Word), choose what to go to, and type a number.

::: platform web
If your Word for the web has no **Go To** (look under the arrow next to **Find**), use the Navigation pane's headings instead, or scroll and watch the page number at the bottom of the window.
:::

::: drill D2.03
title: Move a section with the Navigation pane
badge: Navigator
starter: starters/week-02/W2-L2-report.docx
platform: web
mos: 1.1.3
minutes: 5

In this draft, Methods ended up after Results. Readers need to know how you collected the data before they see it.

- step: Open the Navigation pane and select **Headings**.
- step: Drag the **Methods** heading above **Results**. The whole section, with its steps and equipment list, moves with it.
  check: para.order first="Methods" then="Results"
  check: para.order first="We collected samples" then="Water clarity"
  hint: Drag the heading in the pane, not in the document. A line shows where it will land.
- step: In desktop Word, use **Go To** to jump to page 2. In Word for the web, click **Discussion** in the Navigation pane instead.
:::

## Bookmarks and links inside a document

A **bookmark** is a named, invisible marker. A **link** can jump to a bookmark, so a reader can click "see the methods" and land there.

1. Click in the place you want to jump to, then [[path: Insert > Links > Bookmark]]. Type a name with no spaces (for example `Methods`) and select **Add**.
2. Select the words that should become the link, then [[path: Insert > Link]] ([[key: Ctrl+K / Cmd+K]]).
3. Choose **Place in This Document** (desktop) or **Bookmark** (web), and pick the bookmark.

::: callout note
Link text should say where it goes. "how we collected our samples" is clear even read aloud by a screen reader; "click here" is not (WCAG 2.4.4).
:::

::: drill D2.04
title: Add a bookmark and link to it
badge: Bookmark Builder
starter: starters/week-02/W2-L2-report.docx
platform: web
mos: 1.1.2
minutes: 5

- step: Add a bookmark named **Methods** on the **Methods** heading.
  check: bookmark.is name="Methods" at="Methods"
  hint: Click in the Methods heading, then [[path: Insert > Links > Bookmark]]. Type Methods and select Add.
  web: In Word for the web, Bookmark is on the Insert tab. If your version doesn't show it, do this drill in desktop Word.
- step: In the Introduction, turn the words **how we collected our samples** into a link to that bookmark.
  check: link.present label="how we collected our samples" anchor="Methods"
  hint: Select the words, press [[key: Ctrl+K / Cmd+K]], choose **Place in This Document** (desktop) or **Bookmark** (web), then pick Methods.
- step: Hold [[key: Ctrl / Cmd]] and click the link to test it.
:::

::: quiz Q22.1
question: Which link text is best for a screen reader user?
type: single
- [ ] click here
- [ ] https://www.example.org/report.pdf
- [x] the full sampling data in Appendix A
- [ ] link
explain: Screen reader users often jump from link to link. Text that names the destination makes sense on its own; "click here" and raw addresses don't.
:::

## Check your understanding

- What moves when you drag a heading in the Navigation pane?
- Why must a bookmark name have no spaces?
- When is Go To faster than the Navigation pane?
