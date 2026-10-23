---
type: lesson
id: XL5-L3
week: 5
order: 3
title: Inspect, Check and Protect
objectives: [W5.3]
clos: [CLO1, CLO7]
minutes: 50
status: draft
---

::: callout why
A workbook can carry more than you see: old comments, a hidden sheet of cost notes, your name in the file properties. Before it goes outside your team, check what's in it, make sure everyone can read it, and lock the parts people shouldn't change.
:::

## Inspect Document

[[path: File > Info > Check for Issues > Inspect Document]] (Excel for Windows) looks for things you might not want to share and offers to remove them:

| It finds | Such as |
|---|---|
| **Comments and Annotations** | Every comment thread and note, with authors' names |
| **Document Properties and Personal Information** | The author, "last modified by", and other properties, including the **Title** |
| **Hidden Rows and Columns** | Rows or columns someone hid |
| **Hidden Worksheets** | A sheet someone hid but didn't delete |
| **Headers and Footers**, **Invisible Content** | Text in headers, objects set to invisible |

Choose **Inspect**, then **Remove All** beside each kind you want gone. Removing can't be undone, so **inspect a copy** if you might need what's removed. Be careful with **Hidden Rows and Columns**: Remove All deletes them and their data, which can break formulas that use them. Removing document properties clears the Title too, so set it again afterward.

::: callout note
On a Mac there's no Inspect Document. Use [[path: Excel > Settings > Security]] (Preferences on older versions) and tick **Remove personal information from this file on save**, then delete comments and notes (find them all with [[path: Review > Show Comments]]) and hidden sheets (right-click a tab > **Unhide**, then delete) yourself.
:::

[[path: File > Info > Check for Issues > Check Compatibility]] lists features that won't work in older versions of Excel, like dynamic array formulas, so you can warn people who use them.

## Check Accessibility

[[path: Review > Check Accessibility]] (web and desktop) lists **errors** and **warnings** in a pane. Select one to jump to it, then use the **Recommended Actions** to fix it. In Excel it flags:

- Charts, pictures and shapes with no **alt text** (an error)
- Sheets still called **Sheet1**, **Sheet2** (a warning): give each a name that says what's on it
- Merged cells, and tables without a header row

Microsoft's checker doesn't catch everything, so also fix these yourself (the trainer's checks in this lesson look for them):

- Tables still called **Table1**: rename them in [[path: Table Design > Table Name]]
- No **Title** property: add one in [[path: File > Info]]
- Color used alone to carry meaning, and blank cells used for layout

::: video
id: e5-l3-inspect
url: TODO
title: Inspect Document, Check Compatibility and Check Accessibility
captions: none
duration: 4:00
:::

::: drill D5.04
title: Clean up the price list before it goes out
badge: Inspector
starter: starters/week-05/E5-L3-inspect.xlsx
platform: desktop
mos: 1.5.4
minutes: 8

- step: Run **Inspect Document** and remove all **comments and notes**.
  check: comment.count kind=any max=0
  hint: [[path: File > Info > Check for Issues > Inspect Document]], **Inspect**, then **Remove All** beside Comments and Annotations. If Excel asks to save first, choose **Yes**.
- step: Remove the **hidden worksheet**.
  check: doc.is hiddenSheets=0
- step: Remove **document properties and personal information**.
  check: doc.is personalInfo=removed
- step: Run **Check Compatibility**.
  screenshot: The Compatibility Checker window for this workbook.
  example: TODO
  describe: Say what the Compatibility Checker reported (or that it found no problems).
:::

::: drill D5.05
title: Fix every accessibility error
badge: Accessibility Fixer
starter: starters/week-05/E5-L3-access.xlsx
platform: web
mos: 1.5.4
minutes: 8

Run [[path: Review > Check Accessibility]] first, and keep the pane open as you work.

- step: Rename **Sheet1** to **Prices** and **Sheet2** to **Chart**.
  check: sheet.is name=Prices
  check: access.is rules=sheetNames
- step: Rename the table from **Table1** to **PriceList**.
  check: table.is name=PriceList
  check: access.is rules=tableNames
  hint: Click in the table, then [[path: Table Design > Table Name]], type PriceList, press [[key: Enter]].
- step: Give the chart alt text that states its point.
  check: access.is rules=alt
- step: Set the workbook's **Title** property to **Rental prices**.
  platform: desktop
  check: props.is title="Rental prices"
  hint: [[path: File > Info]], then type in the **Title** box on the right.
:::

## Protecting sheets and workbooks

Every cell starts **locked**, but locking does nothing until the sheet is **protected**. So:

1. Select the cells people **should** type in, then [[path: Home > Format > Format Cells > Protection]] and clear **Locked**.
2. [[path: Review > Protect Sheet]]. A password is optional; if you set one, don't lose it. Now only the unlocked cells can be changed.
3. [[path: Review > Protect Workbook]] (structure) stops people adding, deleting, renaming, hiding or moving sheets.

**Data validation** controls what can be typed in a cell: [[path: Data > Data Validation]], then choose **Allow: List** (a drop-down), **Whole number**, **Date** and so on. Add an **Error Alert** message that says what's allowed.

::: callout privacy
Sheet protection stops mistakes, not determined people: it isn't encryption. To keep a workbook private, share it only with the right people, or use [[path: File > Info > Protect Workbook > Encrypt with Password]].
:::

::: video
id: e5-l3-protect
url: TODO
title: Unlock cells, protect a sheet, data validation
captions: none
duration: 4:00
:::

::: drill D5.06
title: A sign-up sheet people can't break
badge: Protector
starter: starters/week-05/E5-L3-signup.xlsx
platform: desktop
other: E1.2, E2.2
minutes: 10

- step: On **Sign-up**, unlock the entry cells, A4:C13.
  check: protect.is sheet=Sign-up unlocked=A4:C13
  hint: Select A4:C13, [[path: Home > Format > Format Cells]], **Protection** tab, clear **Locked**, **OK**.
- step: Make B4:B13 a drop-down list of **Solo kayak, Tandem kayak, Canoe, Paddleboard**.
  check: validation.is sheet=Sign-up range=B4:B13 type=list list="Solo kayak,Tandem kayak,Canoe,Paddleboard"
  hint: [[path: Data > Data Validation]], Allow **List**, Source: type the four boats with commas between them.
- step: Allow only whole numbers from **1 to 8** in C4:C13.
  check: validation.is sheet=Sign-up range=C4:C13 type=whole
  hint: [[path: Data > Data Validation]], Allow **Whole number**, between 1 and 8. On the **Error Alert** tab, type a message that says what's allowed.
- step: Protect the **Sign-up** sheet (no password, for practice).
  check: protect.is sheet=Sign-up sheetProtected=true
- step: Protect the workbook's **structure**, so nobody can delete the Rates sheet.
  check: protect.is structure=true
  hint: [[path: Review > Protect Workbook]], keep **Structure** ticked, **OK**.
:::

::: quiz Q53.1
question: You protect a sheet, but now nobody can type their name in the sign-up cells. What did you skip?
type: single
- [ ] Data validation
- [x] Unlocking the entry cells before protecting the sheet
- [ ] Protecting the workbook structure
- [ ] Setting a password
explain: Every cell is locked by default. Protecting the sheet enforces that, so unlock the cells people should type in first.
:::

## Check your understanding

- Why inspect a copy instead of the original?
- Which accessibility errors can you fix in under a minute?
- What does Protect Workbook stop that Protect Sheet doesn't?
