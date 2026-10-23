---
type: capstone
id: PP5-CAP
week: 5
order: 7
title: "Capstone: Practice Exam and Portfolio"
objectives: [W5.5]
clos: [CLO1, CLO2, CLO3, CLO4, CLO5, CLO6, CLO7]
minutes: 120
rubricBadge: Course Portfolio
status: draft
---

::: callout why
This is where it comes together. The practice exam shows you how ready you are for the Microsoft Office Specialist: PowerPoint Associate exam (MO-310), and the portfolio gathers the five presentations you made into one place you can show a teacher, an employer or yourself.
:::

## Part 1: Practice exam

Like the real MO-310 exam, the practice exam is timed (50 minutes) and task-based, with no hints: twenty tasks from every part of the exam, in one presentation. (The real exam spreads its tasks across several small projects.) The real exam is pass or fail at a score of 700 out of 1000; here, aim for **70%** or better.

::: callout note
If FLCC gives you access to official Certiport practice tests, take one of those too. They feel closest to the real exam. Ask your instructor or the campus testing center about taking the real exam.
:::

::: skillcheck CAP
title: Season review practice exam
badge: Practice Exam
starter: starters/week-05/P5-CAP-exam.pptx
platform: desktop
mos: 1.1.2, 1.2.1, 1.2.3, 1.4.1, 2.1.3, 2.2.1, 2.2.3, 2.3.1, 3.1.3, 3.2.1, 3.3.2, 3.4.1, 3.4.3, 3.4.7, 3.5.2, 4.1.2, 4.1.3, 4.2.2, 4.3.2, 5.1.1, 5.2.1, 5.2.3
timeLimit: 50
attempts: 3
passing: 70

- task: Change the slide size to **Widescreen (16:9)**.
  check: size.is ratio=16:9
- task: On the slide master, set the title font to **Georgia**.
  check: master.is titleFont=Georgia
- task: Set the **Title** property to **Lakeshore season review**.
  check: props.is title="Lakeshore season review"
- task: Right after **Highlights**, add a **Title and Content** slide titled **Next steps**.
  check: slide.is title="Next steps" layout="Title and Content" after="Highlights"
- task: Hide the **Draft numbers** slide.
  check: slide.is title="Draft numbers" hidden=true
- task: Add slide numbers to every slide except the title slide.
  check: footer.is slideNumbers=true skipTitle=true
- task: Add a section named **Results** that starts with the **Rentals** slide.
  check: section.is name="Results" starts="Rentals"
- task: On **Safety**, turn the steps into a numbered list.
  check: list.is title="Safety" kind=number
- task: On **Visit us**, link the words **our website** to `https://www.lakeshorepaddle.example/`.
  check: link.present title="Visit us" url="lakeshorepaddle.example" says="our website"
- task: On **Our beach**, apply a picture style to the photo.
  check: object.is title="Our beach" kind=picture styled=true
- task: Give that photo alt text.
  check: object.is title="Our beach" kind=picture alt=true
- task: On **Highlights**, insert any star or explosion shape with the text **Record year!**
  check: object.is title="Highlights" kind=shape says="Record year"
- task: On **What we offer**, align the three icons along their tops.
  check: align.is title="What we offer" kind=icon edge=top
- task: On **Prices**, add a row to the bottom of the table: **Paddleboard**, **$18**, **$60**.
  check: table.is title="Prices" rows=4 has="Paddleboard"
- task: Apply a different table style.
  check: table.is title="Prices" styleChanged=true
- task: On **Rentals**, change the chart to a **Line** chart.
  check: chart.is title="Rentals" type=line
- task: On **Our process**, convert the list to any SmartArt graphic.
  check: object.is title="Our process" kind=smartart
- task: Apply the **Fade** transition to every slide.
  check: transition.is slide=all type=fade
- task: On **Highlights**, give the bulleted list the **Wipe** entrance, one paragraph at a time, lasting **1 second**.
  check: anim.is title="Highlights" says="Busiest July" effect=Wipe byParagraph=true duration=1
- task: Create a custom show named **Short version** with these slides: **Lakeshore season review**, **Highlights**, **Rentals**, **Visit us**.
  check: show.is customShow="Short version" slides=4 includes="Highlights,Rentals,Visit us"
:::

## Part 2: Portfolio

Gather your five projects into one folder or one OneDrive share, with a **one-page reflection** (a Word document or a short PowerPoint) that answers:

- Which skill from this course will you use first, and where: in a class, at work or at home?
- Which project are you proudest of, and why?
- What would you still like to learn in PowerPoint?

::: callout privacy
Check each presentation with Inspect Document and the Accessibility Checker before it goes in your portfolio, especially if you share the folder with anyone. Removing document properties clears the Title too, so set it again afterward.
:::

::: rubric PORT
- PORT-a | 40 | CLO1, CLO2 | All five projects are included, in their final, clean versions.
- PORT-b | 20 | CLO7 | The portfolio is organized and shared safely (sensible names, a view-only link or one folder).
- PORT-c | 40 | CLO3, CLO4, CLO6 | The reflection answers all three questions with specific examples from your own work.
:::
