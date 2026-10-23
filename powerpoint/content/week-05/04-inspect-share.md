---
type: lesson
id: PP5-L4
week: 5
order: 4
title: Inspect, Protect and Share
objectives: [W5.4]
clos: [CLO1, CLO7]
minutes: 45
status: draft
---

::: callout why
A presentation can carry more than its slides: your name, old comments, notes to yourself, a text box pushed off the edge. Before it leaves the team, inspect it, make it accessible, and lock it so nobody changes it by accident.
:::

## Inspect the presentation

[[path: File > Info > Check for Issues > Inspect Document]] (PowerPoint for Windows) looks for hidden things, then offers **Remove All** for each kind:

| Category | What it finds |
|---|---|
| **Comments** | Every comment thread, resolved or not |
| **Document Properties and Personal Information** | Your name as author, the last person who saved it, and other properties |
| **Off-Slide Content** | Objects dragged off the edge of a slide, where you can't see them in the show |
| **Invisible On-Slide Content** | Objects set to hidden in the Selection Pane |
| **Presentation Notes** | All speaker notes |

Remove only what you mean to. If you're sharing a handout of your notes, keep **Presentation Notes**.

## Accessibility

[[path: Review > Check Accessibility]] lists **errors** and **warnings** and how to fix each: missing alt text, slides with no title (or the same title twice), tables without a header row, hard-to-read text and a confusing reading order. Select an issue to jump to it. The checker also runs as you work: look for **Accessibility** on the status bar.

::: callout tip
A slide must have a title even if you don't want one to show. Type it in the title placeholder, then drag the placeholder off the slide, or use the **Selection Pane** to hide it. The title still appears in the slide list and to screen readers. If you later remove **Off-Slide Content** or **Invisible On-Slide Content** with Inspect Document, check that your hidden titles are still there.
:::

## Preserve and protect

| Goal | How (desktop) |
|---|---|
| **Keep fonts** looking right on other computers | [[path: File > Options > Save]] > **Embed fonts in the file** (Mac: [[path: PowerPoint > Preferences > Save]]) |
| **Warn people** not to edit | [[path: File > Info > Protect Presentation > Mark as Final]]. Anyone can select **Edit Anyway**, so it's a signal, not a lock |
| **Require a password to change it** | [[path: File > Save As]], **More options**, then **Tools > General Options > Password to modify** |
| **Require a password to open it** | [[path: File > Info > Protect Presentation > Encrypt with Password]]. Lose the password and the file is gone for good |

::: callout note
The trainer can't read a presentation that has a password to open; that's the point of the password. Do this drill with a **password to modify** only.
:::

::: drill D5.06
title: Get the deck ready to share
badge: Ready to Ship
starter: starters/week-05/P5-L4-final.pptx
platform: desktop
mos: 1.5.1, 1.5.2, 1.5.3, 1.5.5
minutes: 12

- step: Inspect the presentation and remove **Comments**, **Document Properties and Personal Information** and **Off-Slide Content**. Keep the notes.
  check: comment.count max=0
  check: doc.is personalInfo=removed
  check: doc.is offSlide=none
  check: notes.count min=1
  hint: [[path: File > Info > Check for Issues > Inspect Document]]. Clear the **Presentation Notes** check box (or just don't select Remove All next to it), then **Inspect**, and **Remove All** next to the three categories.
- step: Run the Accessibility Checker and fix every error: the photo needs alt text, the last slide needs the title **Thank you**, and the table needs a header row.
  check: access.is
  hint: [[path: Review > Check Accessibility]]. For the table, click in it and tick **Header Row** on the **Table Design** tab.
- step: Embed the fonts in the file.
  check: doc.is fontsEmbedded=true
  hint: [[path: File > Options > Save]], tick **Embed fonts in the file**, then **OK** and save. (Mac: [[path: PowerPoint > Preferences > Save]].)
- step: Add a **password to modify** (use **paddle** for practice; never reuse a real password).
  check: doc.is modifyPassword=true
  hint: [[path: File > Save As]], **More options**, **Tools > General Options**, type it in **Password to modify**, then **OK** and **Save**. On a Mac: [[path: File > Passwords]].
- step: Mark the presentation as final.
  check: doc.is final=true
:::

## Export to other formats

[[path: File > Export]] (desktop) makes copies in other formats. Your .pptx stays as it is.

| Choice | Makes |
|---|---|
| **Create PDF/XPS** | A PDF that looks the same everywhere (Options: slides, handouts or notes pages) |
| **Create a Video** | An .mp4 with your timings, narration and animations |
| **Create Handouts** | A Word document with the slides and notes, or lines for notes, that you can edit |
| **Change File Type** | A PowerPoint Show (.ppsx) that opens straight into the show, a template, or pictures of each slide (.png or .jpg) |

In PowerPoint for the web, [[path: File > Export]] offers **Download a Copy**, **Download as PDF**, **Download as Images** and **Download as ODP**. Video export needs desktop PowerPoint.

::: drill D5.07
title: Export for everyone
badge: Exporter
starter: starters/week-05/P5-L4-final.pptx
platform: web
mos: 1.5.6
minutes: 8

- step: Export the presentation as a **PDF** and save it in your course folder.
  hint: [[path: File > Export > Create PDF/XPS]] (desktop) or [[path: File > Export > Download as PDF]] (web).
- step: Export it as a **video**. Take a screenshot of the video file playing.
  platform: desktop
  screenshot: The exported video playing, with the file name showing.
  example: TODO
  describe: What resolution or quality did you choose, and how long is the video?
  hint: [[path: File > Export > Create a Video]], pick a quality, then **Create Video**. On a Mac: [[path: File > Export]], then File Format **MP4**. A short deck takes a minute or two.
- step: Create **handouts in Word** with the notes next to the slides (PowerPoint for Windows).
  platform: desktop
  screenshot: The Word document PowerPoint made, showing slides with their notes.
  example: TODO
  describe: Which page layout did you pick in the Send to Microsoft Word box?
  hint: [[path: File > Export > Create Handouts > Create Handouts]], then **Notes next to slides**.
:::

::: quiz Q54.1
question: What does Mark as Final do?
type: single
- [ ] Encrypts the file so it needs a password
- [x] Marks the file read-only as a signal, though anyone can still choose to edit it
- [ ] Removes all comments and notes
- [ ] Turns the file into a PDF
explain: Mark as Final is a polite "please don't edit". For real protection, use a password to modify or to open.
:::

## Check your understanding

- Why would you keep Presentation Notes when you inspect a deck?
- What's the difference between Mark as Final and a password to modify?
- When is a .ppsx file better than a .pptx?
