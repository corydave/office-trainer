---
type: capstone
id: XL5-CAP
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
This is where it comes together. The practice exam shows you how ready you are for the Microsoft Office Specialist: Excel Associate exam (MO-210), and the portfolio gathers the five workbooks you made into one place you can show a teacher, an employer or yourself.
:::

## Part 1: Practice exam

Like the real MO-210 exam, the practice exam is timed (50 minutes) and task-based, with no hints: twenty tasks from every part of the exam, in one workbook. (The real exam spreads its tasks across several small projects.) The real exam is pass or fail at a score of 700 out of 1000; here, aim for **70%** or better.

::: callout note
If FLCC gives you access to official Certiport practice tests, take one of those too. They feel closest to the real exam. Ask your instructor or the campus testing center about taking the real exam.
:::

On the **Summary** sheet, the labels in column A say what goes in column B.

::: skillcheck CAP
title: Season report practice exam
badge: Practice Exam
starter: starters/week-05/E5-CAP-exam.xlsx
platform: web
mos: 1.3.1, 1.3.3, 1.4.3, 1.4.5, 2.2.5, 2.3.1, 2.3.2, 2.4.2, 3.1.1, 3.1.2, 3.2.3, 3.3.2, 4.1.2, 4.2.1, 4.2.2, 4.2.3, 4.2.5, 4.3.2, 5.1.1, 5.2.3, 5.3.3
timeLimit: 50
attempts: 3
passing: 70

- task: On **Orders**, freeze the top row.
  check: view.is sheet=Orders frozen=A2
- task: Convert the orders (A1:F21) into a table named **Orders**.
  check: table.is name=Orders minRows=20
- task: Apply the table style **Medium 9** (Table Style Medium 9).
  check: table.is name=Orders style=TableStyleMedium9
- task: Add a total row that sums **Amount**.
  check: table.is name=Orders totalRow=true total=Amount totalFunction=sum
- task: Sort the table by **Boat** (A to Z), then by **Date** (oldest first).
  check: sort.is table=Orders by=Boat then=Date
- task: Format the **Amount** values (F2:F21) as **Currency** with 2 decimal places.
  check: cell.is sheet=Orders range=F2:F21 numFmt=currency decimals=2
- task: Highlight **Hours** values (D2:D21) **greater than 3**, with any format.
  check: cf.is sheet=Orders range=D2:D21 type=cellIs operator=greaterThan value=3
- task: On **Rates**, name cell B2 **TaxRate**.
  check: name.is name=TaxRate refersTo=Rates!$B$2
- task: On **Summary**, in B3, total the **Amount** column with `SUM` and a structured reference.
  check: formula.is sheet=Summary cell=B3 uses=SUM structured=true result=851
- task: In B4, count the orders with `COUNTA` on the **Order ID** column.
  check: formula.is sheet=Summary cell=B4 uses=COUNTA result=20
- task: In B5, use `IF` to show **Goal met** when B3 is at least 800, and **Keep going** otherwise.
  check: formula.is sheet=Summary cell=B5 uses=IF result="Goal met"
- task: In B6, work out the sales tax: B3 times the **TaxRate** name.
  check: formula.is sheet=Summary cell=B6 contains=TaxRate result=68.08
- task: In D3, list each boat once with `UNIQUE`.
  check: formula.is sheet=Summary range=D3:D6 uses=UNIQUE spill=true
- task: On **Rates**, in B5, show A5 in capital letters with `UPPER`.
  check: formula.is sheet=Rates cell=B5 uses=UPPER result="LPC-2027-SEASON"
- task: On **Summary**, insert a **Clustered Column** chart of A10:B14.
  check: chart.is sheet=Summary type=column grouping=clustered
- task: Give the chart the title **Rentals by boat**.
  check: chart.is sheet=Summary title="Rentals by boat"
- task: Give the chart alt text that states its point.
  check: chart.is sheet=Summary alt=true
- task: Set **Orders** to **landscape**, fitted to **one page wide**.
  check: page.is sheet=Orders orientation=landscape fitWide=1
  platform: desktop
- task: Add a page number to the center of the **Orders** footer.
  check: header.is sheet=Orders part=footer field=page
  platform: desktop
- task: Set the workbook's **Title** property to **Season report**.
  check: props.is title="Season report"
  platform: desktop
:::

## Part 2: Portfolio

Gather your five projects into one folder or one OneDrive share, with a **one-page reflection** (a Word document or an Excel sheet) that answers:

- Which skill from this course will you use first, and where: in a class, at work or at home?
- Which project are you proudest of, and why?
- What would you still like to learn in Excel?

::: callout privacy
Check each workbook with Inspect Document (or Check Accessibility and a look through the comments) before it goes in your portfolio, especially if you share the folder with anyone. Removing document properties can clear the Title too, so check it afterward.
:::

::: rubric PORT
- PORT-a | 40 | CLO1, CLO2 | All five projects are included, in their final, clean versions.
- PORT-b | 20 | CLO6, CLO7 | The portfolio is organized and shared safely (sensible names, view-only link or one folder).
- PORT-c | 40 | CLO1, CLO3, CLO5 | The reflection answers all three questions with specific examples from your own work.
:::
