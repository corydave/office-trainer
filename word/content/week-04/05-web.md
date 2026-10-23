---
type: lesson
id: W4-L5
week: 4
order: 5
title: Publish for the Web
objectives: [W4.5]
clos: [CLO6]
minutes: 25
status: draft
---

::: callout why
Not every document is printed. A resume might be uploaded to a job site, linked from your LinkedIn profile, or sent as a PDF. Choosing the right format decides whether people can open it, whether it looks right, and whether they can change it.
:::

## Three ways to put a document online

| | **PDF** | **Web page** (.htm) | **Share link** (OneDrive) |
|---|---|---|---|
| Looks the same everywhere | Yes | Mostly; the browser decides some things | Yes, in Word for the web |
| Reader can change it | No (not easily) | Only by editing the file | Only if you give edit permission |
| Updates when you change your file | No; export again | No; save again | Yes |
| Best for | Applications, finished papers, anything official | A simple page on a website | Group work, drafts, anything you'll keep updating |

::: video
id: w4-l5-web
url: TODO
title: PDF, web page or share link?
captions: none
duration: 3:00
:::

## Links people can trust

Link text should say where the link goes: "my portfolio" or "LinkedIn profile", not a long web address or "click here". Select the words, then [[path: Insert > Link]] ([[key: Ctrl+K / Cmd+K]]) and paste the address.

::: drill D4.10
title: Add a descriptive link and export a PDF
badge: Link Linker
starter: starters/week-04/W4-L5-resume.docx
platform: web
mos: 1.3.1
other: N14
minutes: 6

- step: Replace **[add a link to your portfolio or LinkedIn profile here]** with words like **my portfolio**, and make those words a link to your portfolio or LinkedIn profile.
  check: link.present url="://" descriptive=true
  check: text.absent value="[add a link"
  hint: Type the words, select them, press [[key: Ctrl+K / Cmd+K]] and paste the address. No profile yet? Link to https://www.flcc.edu for now.
- step: Export the resume as a PDF and open it to check that the link works.
  hint: [[path: File > Export > Create PDF/XPS]] (desktop), or [[path: File > Save As > Download as PDF]] (web).
  screenshot: Your resume PDF open in a browser or PDF reader, with the link visible.
  example: TODO
  describe: Give the PDF's file name and say whether the link opened.
:::

## Save as a web page

Desktop Word can save a document as a web page: [[path: File > Save As]], then **Web Page, Filtered (*.htm)**. "Filtered" leaves out Word-only code, so the file is smaller and cleaner. Pictures are saved in a folder next to it; keep the two together.

Check it two ways: [[path: View > Web Layout]] in Word, and by opening the .htm file in a browser.

::: drill D4.11
title: Save and preview a web page
badge: Web Publisher
starter: starters/week-04/W4-L5-resume.docx
platform: desktop
mos: 1.3.1
minutes: 6

- step: Switch to **Web Layout** view and look at how the resume changes.
- step: Save the resume as **Web Page, Filtered (*.htm)** and open the .htm file in a browser.
  screenshot: The resume open in a web browser, with the .htm file name in the address bar.
  example: TODO
  describe: Name two things that look different in the browser than in Word.
:::

::: quiz Q45.1
question: A scholarship asks for your resume. It must look exactly as you designed it, and the committee shouldn't edit it. Which format do you send?
type: single
- [x] PDF
- [ ] Web page (.htm)
- [ ] A OneDrive share link with edit permission
- [ ] .docx
explain: A PDF keeps your layout on any device and isn't easy to change. A share link with edit permission lets others change your file.
:::

## Check your understanding

- When is a share link better than a PDF?
- What does "Filtered" mean in Web Page, Filtered?
