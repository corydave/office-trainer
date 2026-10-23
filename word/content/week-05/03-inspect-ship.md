---
type: lesson
id: W5-L3
week: 5
order: 3
title: Inspect, Protect and Ship
objectives: [W5.5]
clos: [CLO1, CLO8]
minutes: 40
status: draft
---

::: callout why
A document carries more than its text: your name as its author, old comments, hidden text, tracked changes nobody accepted. Sending it without checking can share things you never meant to. A five-minute check before you share protects you, and makes the file usable by everyone.
:::

## Inspect the document

**Inspect Document** ([[path: File > Info > Check for Issues > Inspect Document]], desktop Word) looks for hidden information and offers to remove it:

| It finds | Why it matters |
|---|---|
| **Comments and tracked changes** | Your reviewers' notes and every deleted sentence travel with the file |
| **Document properties and personal information** | Your name as Author and Last Modified By, and more |
| **Hidden text** | Text formatted as hidden is invisible on screen but still in the file |
| **Headers, footers, watermarks**, and more | Things you may have forgotten are there |

Select **Remove All** for each item you want gone. Some removals can't be undone, so inspect a **copy** if you might need the original.

**Check Compatibility** ([[path: File > Info > Check for Issues > Check Compatibility]]) lists features that won't work in older versions of Word. A document from an older version opens in **Compatibility Mode** (it says so in the title bar); [[path: File > Info > Convert]] updates it to the current format.

::: video
id: w5-l3-inspect
url: TODO
title: Inspect, check accessibility and protect
captions: none
duration: 4:00
:::

::: drill D5.06
title: Remove hidden data and update the format
badge: Document Inspector
starter: starters/week-05/W5-L3-ship.docx
platform: desktop
mos: 1.4.1, 1.4.3
minutes: 7

- step: Run **Inspect Document** and remove all **comments**.
  check: comment.count max=0
  hint: [[path: File > Info > Check for Issues > Inspect Document]], select **Inspect**, then **Remove All** next to Comments, Revisions and Versions. Careful: this also accepts every tracked change, so review changes before you inspect a real document.
- step: Remove **document properties and personal information**.
  check: doc.is personalInfo=removed
  hint: **Remove All** next to Document Properties and Personal Information. This clears the **Title** property too, so always inspect first and set the title afterward (you'll do that in the next drill).
- step: Remove the **hidden text**.
  check: text.hidden max=0
- step: The guide opens in Compatibility Mode. Run **Check Compatibility**, then **Convert** it to the current format.
  check: doc.is compatMode=15
  hint: [[path: File > Info > Convert]], then **OK**. "Compatibility Mode" disappears from the title bar.
:::

## Check accessibility

The **Accessibility Checker** ([[path: Review > Check Accessibility]], in Word for the web and desktop Word) lists **errors** (content some people can't use at all), **warnings** and **tips**, and explains how to fix each one. Common errors:

- **Missing alternative text** on pictures, charts and SmartArt
- **No header row** in a table
- **Headings that skip a level** (Heading 1 straight to Heading 3)
- **Link text that doesn't say where it goes**, like "click here"
- **No document title** (the Title property becomes the title a screen reader announces and the PDF's title)

::: callout note
Public colleges must meet WCAG 2.1 AA for what they post online, under the ADA Title II rule. Your documents are good practice for that standard, and they're easier for everyone to read.
:::

::: drill D5.07
title: Fix every accessibility error
badge: Access Champion
starter: starters/week-05/W5-L3-ship.docx
platform: web
mos: 1.4.2
minutes: 8

- step: Give the picture alt text.
  check: access.is rules="alt"
  hint: Select it, then [[path: Picture Format > Alt Text]] or right-click and choose **View Alt Text** (or **Edit Alt Text**).
- step: Make the top row of the Shifts table a header row.
  check: access.is rules="tableHeaders"
  hint: Click in the top row, then [[path: Table Layout > Repeat Header Rows]].
- step: **What to wear** skips from Heading 1 to Heading 3. Make it **Heading 2**.
  check: access.is rules="headings"
- step: Change the link text **click here** to words that say where it goes, such as **the Paddle Day sign-up page**.
  check: access.is rules="links"
  hint: Select the link text and type over it; the link stays. Or right-click it and choose **Edit Link**.
- step: Run the Accessibility Checker and fix anything else it reports.
  check: access.is rules="alt,tableHeaders,headings,links"
- step: Set the **Title** property to **Paddle Day Volunteer Guide**, so the checker's "no document title" error goes away.
  platform: desktop
  check: props.is title="Paddle Day Volunteer Guide"
  hint: [[path: File > Info]], then type in the **Title** box on the right. (Word for the web can't set the title, so this step needs desktop Word.)
:::

## Protect the document

**Restrict Editing** ([[path: Review > Restrict Editing]], desktop Word) limits what others can do: **No changes (Read only)**, **Comments** only, **Tracked changes** only, or **Filling in forms**. Select **Yes, Start Enforcing Protection** and add a password.

::: platform web
Word for the web opens a restricted document as read-only. To change the restriction, use desktop Word.
:::

::: drill D5.08
title: Allow comments only
badge: Protector
starter: starters/week-05/W5-L3-ship.docx
platform: desktop
other: E1.2
minutes: 4

- step: Restrict editing to **Comments** only, and start enforcing protection with a password.
  check: doc.is protection=comments password=true
  hint: [[path: Review > Restrict Editing]]. Under **Editing restrictions**, check the box and choose **Comments**. Then **Yes, Start Enforcing Protection**, and type a password twice.
:::

::: quiz Q53.1
question: You're sending a cover letter you wrote from a friend's template. What should you do first?
type: multi
- [x] Run Inspect Document and remove personal information and comments
- [x] Check that the Title and Author properties are yours
- [ ] Turn on Track Changes
- [x] Export a PDF and check that it looks right
explain: Templates and old files carry other people's names, comments and properties. Inspect, fix the properties, and send a PDF so it looks right on their screen.
:::

## Check your understanding

- Name three kinds of hidden information Inspect Document can find.
- What's the difference between an Accessibility Checker error and a warning?
- When would you restrict editing to comments only?
