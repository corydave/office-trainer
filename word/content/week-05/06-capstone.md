---
type: capstone
id: W5-CAP
week: 5
order: 6
title: "Capstone: Practice Exam and Portfolio"
objectives: [W5.6]
clos: [CLO1, CLO2, CLO3, CLO4, CLO5, CLO6, CLO7, CLO8]
minutes: 120
rubricBadge: Course Portfolio
status: draft
---

::: callout why
This is where it comes together. The practice exam shows you how ready you are for the Microsoft Office Specialist: Word Associate exam (MO-110), and the portfolio gathers the five things you made into one place you can show a teacher, an employer or yourself.
:::

## Part 1: Practice exam

Like the real MO-110 exam, the practice exam is timed (50 minutes) and task-based, with no hints: twenty tasks from every part of the exam, in one document. (The real exam spreads its tasks across several small projects.) The real exam is pass or fail at a score of 700 out of 1000; here, aim for **70%** or better.

::: callout note
If FLCC gives you access to official Certiport practice tests, take one of those too. They feel closest to the real exam. Ask your instructor or the campus testing center about taking the real exam.
:::

::: skillcheck CAP
title: Annual report practice exam
badge: Practice Exam
starter: starters/week-05/W5-CAP-exam.docx
platform: web
mos: 1.1.1, 1.2.1, 1.2.3, 1.3.2, 2.1.1, 2.1.2, 2.2.1, 2.2.3, 2.2.4, 2.2.5, 2.3.1, 3.1.1, 3.2.6, 3.3.1, 4.1.1, 4.2.1, 5.2.2, 5.4.1, 5.4.2, 5.4.3, 6.1.2, 6.2.3
timeLimit: 50
attempts: 3
passing: 70

- task: Replace every **Lakeshore Paddle Co.** with **Lakeshore Paddle Company**.
  check: text.absent value="Lakeshore Paddle Co."
- task: Apply the **Title** style to the first line, and **Heading 1** to **Our season**, **Highlights**, **Rentals**, **Goals for next year**, **Staff** and **Next steps**.
  check: style.is para="Lakeshore Paddle Company Annual Report" name="Title"
  check: style.is para="Our season" name="Heading 1"
  check: style.is para="Goals for next year" name="Heading 1"
  check: style.is para="Next steps" name="Heading 1"
- task: Apply any text effect to the title.
  check: effect.present para="Lakeshore Paddle Company Annual Report" kind=any
- task: Give the paragraph under **Our season** 1.15 line spacing and 6 pt after.
  check: para.is para="had its busiest season" lineSpacing=1.15 spaceAfter=6pt
- task: Clear all formatting from the paragraph that starts **Thank you to every renter**.
  check: format.clean para="Thank you to every renter"
- task: Set 1 in margins and Letter size.
  check: page.is margins=1in size=letter
- task: Add a page number in the footer, after a © symbol and the year (for example **© 2026 1**).
  check: field.present code="PAGE" in=footer
  check: text.present value="©" in=footer
- task: Set the document's **Title** property to **Annual Report 2026**.
  check: props.is title="Annual Report 2026"
- task: Insert an automatic table of contents after the line that starts **Prepared for**.
  check: toc.is
- task: Turn the three lines under **Goals for next year** into a numbered list.
  check: list.is para="Add four kayaks" kind=number
  check: list.is para="Train every guide" kind=number
- task: Convert the four tab-separated lines under **Staff** into a table.
  check: table.is contains="Lead guide" cols=3 minRows=4
  platform: desktop
- task: Make the top row of the **Rentals** table repeat as a header row.
  check: table.is contains="Paddleboard" headerRepeat=true
- task: Put the paragraph under **Highlights** in two columns.
  check: section.is para="brought first-time paddlers" columns=2
  platform: desktop
- task: Add a footnote after **two new programs.** Type: The sunset series and Paddle Day.
  check: notes.is kind=footnote min=1
- task: Give the photo alt text.
  check: object.is kind=picture name="Report photo" alt=true
- task: Set the photo to **Square** wrapping, positioned at the **top right**.
  check: object.is kind=picture name="Report photo" wrap=square halign=right valign=top
- task: Apply any picture style to the photo.
  check: object.is kind=picture name="Report photo" pictureStyle=true
- task: Reply to Priya's comment about the **meeting date**.
  check: comment.is says="meeting date" state=replied
- task: Accept Priya's tracked change, so no tracked changes are left. (If you turned Track Changes on during the exam, accept your own changes too.)
  check: revision.count max=0
- task: Insert a ™ symbol right after **Paddle Day** in the Highlights paragraph.
  check: text.present value="Paddle Day™"
:::

## Part 2: Portfolio

Gather your five projects into one folder or one OneDrive share, with a **one-page reflection** (a Word document, of course) that answers:

- Which skill from this course will you use first, and where: in a class, at work or at home?
- Which project are you proudest of, and why?
- What would you still like to learn in Word?

::: callout privacy
Check each file with Inspect Document before it goes in your portfolio, especially if you share the folder with anyone. Inspecting clears the Title property, so set the title again afterward.
:::

::: rubric PORT
- PORT-a | 40 | CLO1, CLO2 | All five projects are included, in their final, clean versions.
- PORT-b | 20 | CLO6 | The portfolio is organized and shared safely (sensible names, view-only link or one folder).
- PORT-c | 40 | CLO1, CLO8 | The reflection answers all three questions with specific examples from your own work.
:::
