---
type: skillcheck
id: XL1-SC1
week: 1
order: 5
title: "Skill Check 1: The Working Spreadsheet"
objectives: [W1.3, W1.4, W1.5]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

This skill check works like a MOS exam task: one starter workbook, a set of tasks and a time limit. There are no hints. Each task is worth one point, and partial credit is given task by task.

**Before you start:** open the starter workbook, then select **Start**. The timer starts then. You can do the tasks in any order and select **Check my work** as often as you like; your score is what the workbook shows when you select **Submit** or when time runs out. Save before you submit.

::: skillcheck SC1
title: Gear inventory
badge: Skill Check 1
starter: starters/week-01/E1-SC1-inventory.xlsx
platform: web
mos: 1.2.1, 1.3.1, 1.3.2, 1.3.3, 1.4.3, 1.4.5, 1.5.1, 2.1.1, 2.1.2, 2.1.3, 2.2.1, 2.2.2, 2.2.4, 2.2.5, 2.2.7
timeLimit: 30
attempts: 2
passing: 70

- task: Replace every **Lifejacket** in the workbook with **Life jacket** (two words).
  check: text.absent value="Lifejacket"
  check: text.present value="Life jacket" min=3
- task: On the **Inventory** sheet, Auto Fill the item numbers down column A so they run from **G-101** to **G-117**.
  check: cell.is sheet=Inventory cell=A20 value=G-117
- task: Insert a new column before **Cost** with the header **Supplier**.
  check: cell.is sheet=Inventory cell=D3 value=Supplier
  check: cell.is sheet=Inventory cell=E3 value=Cost
- task: Fill the **Reorder at** column with the numbers from the **Reorder at** column on the **Old Counts** sheet, pasted as values (no formulas).
  check: formula.is sheet=Inventory range=G4:G20 hasFormula=false
  check: cell.is sheet=Inventory cell=G4 number=15
  check: cell.is sheet=Inventory cell=G20 number=2
- task: Merge and center the title across the table's width (**A1:G1**).
  check: cell.is sheet=Inventory range=A1:G1 merged=true align=center
- task: Make the header row (**A3:G3**) wrap its text and center it.
  check: cell.is sheet=Inventory range=A3:G3 wrap=true align=center
- task: Format the **Cost** column's numbers as **Currency**.
  check: cell.is sheet=Inventory range=E4:E20 numFmt=currency
- task: Apply the **Title** cell style to the title.
  check: cell.is sheet=Inventory cell=A1 style=Title
- task: Set column **B** to a width of **30**.
  check: col.is sheet=Inventory col=B width=30
- task: Freeze rows 1 to 3 so the headers stay in view.
  check: view.is sheet=Inventory frozen=A4
- task: Print setup: **landscape**, fit to **1 page wide**, and a print area of **A1:G20**.
  platform: desktop
  check: page.is sheet=Inventory orientation=landscape fitWide=1 printArea=A1:G20
- task: Add a footer with the **page number** in the center, and set the workbook's **Title** property to **Gear Inventory**.
  platform: desktop
  check: header.is sheet=Inventory part=footer field=page position=center
  check: props.is title="Gear Inventory"
:::

::: callout note
MO-210 objective 1.5.2 (saving in other formats) is assessed with the PDF you hand in with Project 1.
:::
