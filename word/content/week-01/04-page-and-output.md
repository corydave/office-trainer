---
type: lesson
id: W1-L4
week: 1
order: 4
title: Page Setup, Properties and Output
objectives: [W1.5, W1.4]
clos: [CLO1]
minutes: 45
status: draft
---

::: callout why
The last ten percent of a document is the part people notice: the right paper size, a clean PDF, a file that says who wrote it. A letter printed on the wrong paper size, or a PDF titled "Document1", undoes all the work before it.
:::

## Page setup

::: video
id: w1-l4-page-setup
url: TODO
title: Margins, paper size and orientation
captions: none
duration: 3:00
:::

Everything about the shape of the page lives on the **Layout** tab:

| Setting | Where | US business default |
|---|---|---|
| **Margins** | [[path: Layout > Margins]] | **Normal**: 1 in on all four sides |
| **Size** | [[path: Layout > Size]] | **Letter** (8.5 x 11 in). Most of the world uses **A4** |
| **Orientation** | [[path: Layout > Orientation]] | **Portrait** (tall). **Landscape** (wide) suits wide tables and certificates |

::: callout note
A document made in another country, or from some online templates, may be set to A4. It looks almost right on screen, then prints with an odd bottom margin. Always check Size before printing.
:::

## Proofread with Editor

Run Editor ([[path: Review > Editor]] in desktop Word, [[path: Home > Editor]] in Word for the web, or [[key: F7]] on Windows) and work through each suggestion. Accept what is right, ignore what is not, and read the document once more yourself.

::: drill D1.09
title: Fix the page setup and proofread
badge: Page Setter
starter: starters/week-01/W1-L4-letter.docx
platform: web
mos: 1.2.1
other: N5, N11
minutes: 6

This version of the letter came from a template set up for another country.

- step: Set the margins to **Normal (1 in)**, the size to **Letter** and the orientation to **Portrait**.
  check: page.is margins=1in size=letter orientation=portrait
  hint: Three menus on the Layout tab: Margins, Orientation and Size.
- step: Run Editor and fix the three spelling mistakes.
  check: text.absent value="recieve" case=insensitive
  check: text.absent value="seperate" case=insensitive
  check: text.absent value="tommorow" case=insensitive
  hint: Red underlines mark spelling. Right-click one to see suggestions, or open the Editor pane.
:::

## Document properties and language

Every Word file carries **properties**: a title, an author, tags and more. They show up in search results, in file lists and, most importantly, as the title of a PDF. Screen readers announce the title when the file opens, so a real title is also an accessibility requirement.

- Desktop Word: [[path: File > Info]], then fill in **Title** and **Tags** on the right. **Author** is under **Related People**. Select **Show All Properties** for more.
- Word for the web shows fewer properties; set them in desktop Word.

The **proofing language** tells Editor which dictionary to use and tells screen readers how to pronounce the words. Set it with [[path: Review > Language > Set Proofing Language]].

::: drill D1.10
title: Set properties and the proofing language
badge: Properties and Language
starter: starters/week-01/W1-L4-letter.docx
platform: desktop
mos: 1.3.2
other: E1.3
minutes: 5

- step: Set the **Title** to **Volunteer Day Letter**.
  check: props.is title="Volunteer Day Letter"
- step: Add your name as **Author**, and add the tags **letter; volunteer**.
  check: props.is author="*" keywords="volunteer"
  hint: Author is under **Related People** on the File > Info page. Tags are in the Properties list on the right.
- step: The body text is marked as English (United Kingdom). Select all ([[key: Ctrl+A / Cmd+A]]) and set the proofing language to **English (United States)**.
  check: lang.is value=en-US
  hint: [[path: Review > Language > Set Proofing Language]].
:::

## Save as PDF and print

A **PDF** looks the same on every device, which makes it the right format for anything you send to be read rather than edited: applications, letters, final papers.

- Desktop Word: [[path: File > Export > Create PDF/XPS]], or [[path: File > Save As]] and choose **PDF** as the file type.
- Word for the web: [[path: File > Export > Download as PDF]].

To print, open [[path: File > Print]] ([[key: Ctrl+P / Cmd+P]]). Check the preview on the right, then set:

- **Copies**
- **Pages**: type `1` for page 1 only, `1-3` for a range, or `1,4` for specific pages
- **Printer** and, in desktop Word, one-sided or two-sided

::: platform web
Word for the web hands printing to your browser. It first builds a PDF, then opens your browser's print dialog, so the options you see depend on the browser.
:::

::: drill D1.11
title: Export a PDF and set up printing
badge: PDF and Print
starter: starters/week-01/W1-L4-letter.docx
platform: web
mos: 1.3.1, 1.3.3
other: N12, N13
minutes: 5

The trainer can't see files outside the document, so these steps use screenshots. You'll hand in a real PDF with Project 1.

- step: Export the letter as a PDF named **Lastname_W1_Letter.pdf** and open it to check that it looks right.
  screenshot: Your PDF open in a browser or PDF reader, with the file name Lastname_W1_Letter.pdf visible.
  example: TODO
  describe: Give the PDF's file name and how many pages it has.
- step: Open Print and set it to print **2 copies** of **page 1 only**. Look at the preview, then cancel (you don't need to print).
  screenshot: The Print screen showing Copies set to 2 and Pages set to 1, with the preview beside them.
  example: TODO
  describe: Say what Copies and Pages are set to.
- step: Close the document.
:::

::: quiz Q4.1
question: You're sending a cover letter with a job application. Which format should you send?
type: single
- [ ] .docx, so they can fix any mistakes
- [x] .pdf, so it looks the same on their computer
- [ ] .txt, so it opens anywhere
- [ ] .dotx, so they can reuse it
explain: A PDF keeps your layout and fonts exactly as you set them, on any device. Send .docx only when someone needs to edit the file.
:::

## Check your understanding

- Why does a PDF need a real Title property?
- What happens if a letter set to A4 prints on Letter paper?
- How do you print only pages 2 through 4?
