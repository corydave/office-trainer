---
type: lesson
id: W4-L2
week: 4
order: 2
title: "Two Resumes: Designed and ATS-Friendly"
objectives: [W4.2]
clos: [CLO1, CLO4, CLO5]
minutes: 40
status: draft
---

::: callout why
Most large employers run resumes through an **applicant tracking system** (ATS) before a person sees them. Hiring software reads plain, simple layouts reliably and can scramble fancy ones. So keep two versions: a designed one to hand to people, and a plain one to upload.
:::

## What makes a resume ATS-friendly

| Do | Avoid |
|---|---|
| Real headings (Experience, Education, Skills) | Headings made of pictures or WordArt |
| Text in the body of the document | Your name or contact details only in the header |
| Simple lines and tab stops | Text boxes and tables used for layout |
| Standard fonts | Icons in place of words (a phone icon instead of "Phone") |
| Save as .docx or PDF, as the posting asks | Columns that change the reading order |

::: video
id: w4-l2-resumes
url: TODO
title: One resume, two versions
captions: none
duration: 3:30
:::

## Tab stops

A **tab stop** sets exactly where [[key: Tab]] moves the cursor. A **right tab** lines text up against the stop, which is how dates line up at the right margin. A **leader** fills the gap with dots or a line, like a table of contents.

- Desktop Word: click in the line, then open [[path: Home > Paragraph]] settings and select **Tabs** (bottom left). Type the position, choose **Right**, pick a leader, and select **Set**.
- Or click the tab selector at the left end of the ruler until it shows the right-tab symbol, then click the ruler.

::: callout warning
Pressing the spacebar until things look lined up is the most common resume mistake. It never lines up on someone else's screen, and it reads badly in hiring software.
:::

::: drill D4.03
title: Line up the dates with a right tab
badge: Tab Stop Pro
starter: starters/week-04/W4-L2-resume-ats.docx
platform: desktop
minutes: 10

The dates in this resume were pushed to the right with spaces.

- step: In the **Office Assistant** line, delete the spaces before the dates and press [[key: Tab]] instead. Then add a **right tab stop at 6.5 in** with a **dot leader**.
  check: tabs.is para="Office Assistant" position=6.5in align=right leader=dot
  hint: Click in the line, then open the Paragraph settings, select **Tabs**, type 6.5, choose **Right** and leader **2 (dots)**, then **Set** and **OK**.
- step: Do the same for the **Volunteer Coordinator** and **Cashier** lines.
  check: tabs.is para="Volunteer Coordinator" position=6.5in align=right leader=dot
  check: tabs.is para="Cashier, Main Street Market" position=6.5in align=right leader=dot
  hint: Use the Format Painter to copy the tab stop from the first line, then replace the spaces with a tab.
- step: The **Skills** are laid out in a table, which hiring software can misread. Convert the table to text (separated by commas).
  check: table.count max=0
  check: text.present value="Scheduling"
  hint: Click in the table, then [[path: Table Layout > Convert to Text]] and choose **Commas**.
:::

::: quiz Q42.1
question: Which of these is most likely to confuse an applicant tracking system?
type: single
- [ ] A Heading 1 that says Experience
- [x] Your phone number and email in a text box at the top
- [ ] Dates lined up with a right tab stop
- [ ] A PDF saved from Word
explain: Text boxes, headers and layout tables often get skipped or read out of order. Contact details belong in the body of the document.
:::

## Check your understanding

- Why keep two versions of your resume?
- What's the difference between a left tab and a right tab?
