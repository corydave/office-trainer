---
type: lesson
id: W2-L4
week: 2
order: 4
title: Notes, Contents and Citations
objectives: [W2.5, W2.6]
clos: [CLO2]
minutes: 55
status: draft
---

::: callout why
Footnotes, a table of contents and a reference list are what make a long paper easy to trust and easy to use. Word builds all three for you and keeps them up to date, as long as you let it do the work instead of typing them by hand.
:::

## Footnotes and endnotes

A **footnote** puts a small number in the text and the note at the bottom of the same page. An **endnote** collects the notes at the end of the document.

- Click where the number should go, then [[path: References > Insert Footnote]] ([[key: Ctrl+Alt+F / Cmd+Option+F]]). Type the note.
- Word renumbers every note when you add, move or delete one.
- To change the numbering (1, 2, 3 to i, ii, iii, or symbols), open the **Footnote and Endnote** dialog: the small arrow in the corner of [[path: References > Footnotes]] (desktop Word).

::: video
id: w2-l4-footnotes
url: TODO
title: Footnotes and endnotes
captions: none
duration: 2:30
:::

::: drill D2.09
title: Add footnotes and change their numbers
badge: Footnote Finder
starter: starters/week-02/W2-L4-report.docx
platform: web
mos: 4.1.1, 4.1.2
minutes: 6

- step: Add a footnote at the end of the sentence that ends **Secchi disk readings at every visit.** Type: A Secchi disk is a black and white disk lowered until it disappears from view.
  check: notes.is kind=footnote min=1
  hint: Click just after the period, then [[path: References > Insert Footnote]].
- step: Add a second footnote after the words **Volunteers counted 1,284 pieces of litter**. Type: Counts include every item larger than a bottle cap.
  check: notes.is kind=footnote min=2
- step: Change the footnote numbers to **i, ii, iii**.
  check: notes.is kind=footnote min=2 format=lowerRoman
  platform: desktop
  hint: Open the **Footnote and Endnote** dialog (the arrow in the corner of the Footnotes group), choose **i, ii, iii** in **Number format**, then select **Apply**.
:::

## An automatic table of contents

Word builds a **table of contents** (TOC) from your headings: Heading 1 and Heading 2 lines become entries, with page numbers.

1. Click where it should go (usually after the title page), then [[path: References > Table of Contents]] and pick a style.
2. When headings change, click in the table and select **Update Table > Update entire table**. The table doesn't update by itself.
3. To change how many heading levels it shows, use [[path: References > Table of Contents > Custom Table of Contents]] and set **Show levels**.

::: callout tip
If a heading is missing from the table, it isn't using a heading style. Fix the heading, then update the table.
:::

::: drill D2.10
title: Insert and customize a table of contents
badge: Contents Creator
starter: starters/week-02/W2-L4-report.docx
platform: web
mos: 4.2.1, 4.2.2
minutes: 7

- step: Insert an automatic table of contents on the line after the practice-report note, before **Introduction**.
  check: toc.is
  hint: Click at the end of the practice-report note and press [[key: Enter]], then [[path: References > Table of Contents]]. (Pressing Enter at the start of a heading makes an empty heading instead.)
- step: **Limitations** is still plain text. Apply **Heading 1** to it.
  check: style.is para="Limitations" name="Heading 1"
- step: Update the table of contents, so **Limitations** appears in it.
  check: toc.is includes="Limitations"
  hint: Click in the table, then **Update Table > Update entire table**.
- step: Change the table so it shows only **2 levels** of headings.
  check: toc.is levels=2
  hint: [[path: References > Table of Contents > Custom Table of Contents]], then set **Show levels** to 2 and select **OK**. Choose **Yes** to replace the table.
  web: If your version of Word for the web has no Custom Table of Contents, open the file in desktop Word for this step.
:::

## Citations and a reference list

Word's **citation manager** keeps your sources and writes the in-text citations and the reference list for you, in APA, MLA or another style.

1. Choose the style first: [[path: References > Citations & Bibliography > Style]] (for example **APA**).
2. Click after the sentence that uses a source, then [[path: References > Insert Citation > Add New Source]]. Fill in the details and select **OK**.
3. To cite the same source again, choose it from **Insert Citation**.
4. At the end of the paper, insert the list: [[path: References > Bibliography]] and choose **References** (APA) or **Works Cited** (MLA).

::: callout ai
AI tools sometimes invent sources that look real, with believable titles and authors. Before any source goes into the citation manager, find it yourself in the library's databases or on the publisher's site. That habit protects your grade and your reputation.
:::

::: platform web
Use desktop Word for this drill. Word for the web shows citations made in desktop Word, but it doesn't have the full citation manager (Manage Sources, styles and a generated reference list).
:::

For this drill, use these three **practice sources**. They're made up for the course, so don't cite them anywhere else.

| Type | Author | Year | Title | Where |
|---|---|---|---|---|
| Journal article | Lin, M. | 2024 | Shoreline litter at small inland lakes | *Journal of Lake Stewardship*, 12(3), 44 to 58 |
| Report | North Shore Association | 2025 | Annual water clarity summary | North Shore Association |
| Web page | Okafor, T. | 2023 | Running a volunteer beach cleanup | Clean Shores Network |

::: drill D2.11
title: Cite three sources and build the reference list
badge: Citation Pro
starter: starters/week-02/W2-L4-report.docx
platform: desktop
other: E3.5
minutes: 12

- step: Make sure the citation style is **APA**. It's usually Word's default, but always check: a paper someone else started may be set to MLA or Chicago.
  check: citation.is citeStyle=APA
  hint: [[path: References > Style]] shows the current style. Pick **APA** if it shows anything else.
- step: Add the three practice sources in the citation manager.
  check: citation.is sources=3
  hint: [[path: References > Insert Citation > Add New Source]] adds a source and a citation at the same time. [[path: References > Manage Sources]] shows them all.
- step: Insert a citation after each of these sentences: the one ending **made up most of the items** (Lin), the one ending **farthest from the boat launch** (North Shore Association), and the one ending **first Saturday of each month** (Okafor).
  check: citation.is citations=3
- step: At the very end of the document, insert a **References** list.
  check: citation.is bibliography=true
  hint: Press [[key: Ctrl+End / Cmd+End]] (on a Mac keyboard without an End key, [[key: Cmd+Fn+Right Arrow]]), insert a page break, then [[path: References > Bibliography > References]].
:::

## Check your understanding

- What's the difference between a footnote and an endnote?
- You added a new Heading 1, but the table of contents doesn't show it. What two things could be wrong?
- Why should you check a source before you add it to the citation manager?
