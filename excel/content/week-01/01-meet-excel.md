---
type: lesson
id: XL1-L1
week: 1
order: 1
title: Meet Excel
objectives: [W1.1, W1.2]
clos: [CLO1, CLO2, CLO6]
minutes: 45
status: draft
---

::: callout why
A spreadsheet is a grid that remembers. Once you know your way around the grid, the sheets and the views, every other Excel skill is easier, and the workbooks you make are easier for other people to read.
:::

## Your scenario

All five weeks follow **Lakeshore Paddle Company**, a small kayak, canoe and paddleboard rental shop on a Finger Lakes beach. This week you'll organize its June rental records and turn them into a clean, printable report.

## The Excel window

| Part | What it's for |
|---|---|
| **Ribbon** | Commands, grouped on tabs: Home, Insert, Page Layout, Formulas, Data, Review, View |
| **Name Box** (left of the formula bar) | Shows the address of the selected cell, like **B4**. Type an address or a name here and press [[key: Enter]] to jump to it |
| **Formula bar** | Shows what's really in the cell: the number, the text or the formula |
| **Active cell** | The cell with the thick border. What you type goes here |
| **Column letters and row numbers** | Every cell's address is its column letter plus its row number |
| **Sheet tabs** (bottom) | A workbook holds several worksheets. Select a tab to switch |
| **Status bar** (bottom) | Shows the sum, average and count of the selected cells, plus the view buttons and zoom |
| **Quick Access Toolbar** | Your own one-click buttons (desktop Excel). If you don't see it, right-click the Ribbon and choose **Show Quick Access Toolbar** |

A file is a **workbook**; each tab in it is a **worksheet** (or sheet).

::: video
id: e1-l1-tour
url: TODO
title: A tour of the Excel window
captions: none
duration: 3:00
:::

::: drill D1.01
title: Save the workbook with a clear name
badge: File Namer
starter: starters/week-01/E1-L1-rentals.xlsx
platform: web
other: N1, N2
minutes: 3

- step: Open the practice file, then save it in your course folder as **Lastname_E1_Rentals.xlsx** (with your own last name).
  check: file.name pattern="*_E1_Rentals.xlsx"
  hint: [[path: File > Save As]], choose your course folder, and type the name. In Excel for the web, select the file name at the top to rename it.
:::

## Worksheets

- **Rename** a sheet: double-click its tab (or right-click it and choose **Rename**), type, press [[key: Enter]].
- **Add** a sheet: select **+** (New sheet) next to the tabs.
- **Move** a sheet: drag its tab left or right.
- **Color** a tab: right-click it and choose **Tab Color**.
- **Delete** a sheet: right-click its tab and choose **Delete**. There's no Undo for this one, so check first.

::: drill D1.02
title: Name, add, color and move sheets
badge: Sheet Shaper
starter: starters/week-01/E1-L1-rentals.xlsx
platform: web
other: N11
minutes: 4

- step: Rename **Sheet1** to **June**.
  check: sheet.is name=June
- step: Add a new sheet named **July**, and move it so it sits right after June.
  check: sheet.is name=July after=June
  hint: Select **+** to add the sheet, rename it, then drag its tab so the order is June, July, Summary.
- step: Give the June tab a color.
  check: sheet.is name=June tabColor=any
:::

## Your own buttons

The **Quick Access Toolbar** holds commands you use all the time, one click away on every tab. Select the arrow at its end (**Customize Quick Access Toolbar**) and check the commands you want, or right-click any Ribbon button and choose **Add to Quick Access Toolbar**.

::: platform web
Excel for the web doesn't have a Quick Access Toolbar to customize. Do this step in desktop Excel, or skip it if you only use the web.
:::

::: drill D1.03
title: Customize the Quick Access Toolbar
badge: Toolbar Tuner
starter: starters/week-01/E1-L1-rentals.xlsx
platform: desktop
mos: 1.4.1
other: N3
minutes: 4

- step: Add **Save As** and **Print Preview and Print** to the Quick Access Toolbar.
  hint: Print Preview and Print is in the Customize Quick Access Toolbar menu. Save As isn't: choose **More Commands**, set **Choose commands from** to **All Commands**, select **Save As**, then **Add**.
  screenshot: The Excel window with Save As and Print Preview showing on the Quick Access Toolbar.
  example: TODO
  describe: Say which buttons are on your Quick Access Toolbar now.
:::

::: quiz Q11.1
question: You select cell D7. Where can you see its address?
type: single
- [x] In the Name Box, at the left end of the formula bar
- [ ] On the sheet tab
- [ ] In the status bar
- [ ] In the Quick Access Toolbar
explain: The Name Box always shows the selected cell's address. You can also type an address there and press Enter to jump to it.
:::

## Views and frozen headings

Long lists are easier to read when the headings stay put as you scroll.

| To | Do this |
|---|---|
| Keep the top row in view | [[path: View > Freeze Panes > Freeze Top Row]] |
| Keep the top row **and** the first column | Select **B2**, then [[path: View > Freeze Panes > Freeze Panes]] |
| See page edges, headers and margins as you work | [[path: View > Page Layout]] (desktop Excel), or the view buttons on the status bar |
| See where pages break | [[path: View > Page Break Preview]] |
| Look at two parts of a workbook at once | [[path: View > New Window]], then [[path: View > Arrange All]] |

::: callout tip
Excel freezes everything **above and to the left of** the selected cell. That's why you select B2 to freeze row 1 and column A together.
:::

::: drill D1.04
title: Freeze the headings and change the view
badge: View Finder
starter: starters/week-01/E1-L1-rentals.xlsx
platform: desktop
mos: 1.4.2, 1.4.3, 1.4.4
other: N8
minutes: 6

- step: On the **June** sheet, freeze the top row and the first column together.
  check: view.is sheet=June frozen=B2
  hint: Select B2 first, then [[path: View > Freeze Panes > Freeze Panes]].
- step: Switch the **Summary** sheet to **Page Layout** view.
  check: view.is sheet=Summary view=pageLayout
- step: Open a second window of this workbook and arrange the two side by side. Save while both windows are open.
  check: window.count min=2
  hint: [[path: View > New Window]], then [[path: View > Arrange All > Vertical]]. The title bar shows ":1" and ":2".
:::

## Find and Replace

**Find** ([[key: Ctrl+F / Cmd+F]]) jumps to each match; **Find All** lists every match and how many there are. **Replace** ([[key: Ctrl+H]], also Control+H on a Mac) changes them. Select **Options** to search the whole workbook instead of one sheet (**Within: Workbook**), or to match the whole cell.

::: drill D1.05
title: Find and replace a boat name
badge: Find and Replace
starter: starters/week-01/E1-L1-rentals.xlsx
platform: web
mos: 1.2.1
other: N4
minutes: 5

- step: Use Find to count how many rentals on the **June** sheet were a **Kayak (single)**. Type the number here.
  answer: 12
  hint: [[key: Ctrl+F / Cmd+F]], type Kayak (single), then **Find All**. The count is at the bottom of the box.
- step: The shop renamed the boat. Replace every **Kayak (single)** in the whole workbook with **Solo kayak**.
  check: text.absent value="Kayak (single)"
  check: text.present value="Solo kayak" min=13
  hint: [[path: Home > Find & Select > Replace]], then **Options** and set **Within** to **Workbook** so the Summary sheet changes too.
:::

## Check your understanding

- What's the difference between a workbook and a worksheet?
- Which cell do you select to freeze the top two rows but no columns?
- Why would you search Within: Workbook instead of Within: Sheet?
