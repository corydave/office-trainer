---
type: lesson
id: XL2-L3
week: 2
order: 3
title: Names and Links
objectives: [W2.3]
clos: [CLO2]
minutes: 35
status: draft
---

::: callout why
`=D10*TaxRate` explains itself; `=D10*$F$2` doesn't. Names make formulas readable and give you a quick way to jump around a big workbook. Links do the same for readers who don't know where to look.
:::

## Named ranges

A **name** stands for a cell or a range. To make one, select the cells, type the name in the **Name Box** and press [[key: Enter]], or use [[path: Formulas > Define Name]].

- Names can't have spaces or start with a number: **TaxRate**, **Tax_Rate** and **Deposits2027** work; **Tax Rate** doesn't.
- Use a name anywhere you'd use the address: `=D10*TaxRate`, `=SUM(Deposits)`. A name is always absolute, so no $ signs are needed.
- [[path: Formulas > Name Manager]] lists, edits and deletes names.
- Pick a name from the **Name Box** list (or press [[key: F5]] for **Go To**; in Excel for the web, [[key: Ctrl+G]], because F5 reloads the page) to jump to it and select it.

::: video
id: e2-l3-names
url: TODO
title: Naming cells and jumping with the Name Box
captions: none
duration: 3:00
:::

::: drill D2.08
title: Name the tax rate and use it
badge: Name Giver
starter: starters/week-02/E2-L3-invoice.xlsx
platform: web
mos: 2.3.1, 2.3.2, 1.2.2
minutes: 6

- step: On the **Invoice** sheet, name the tax-rate cell (F2) **TaxRate**.
  check: name.is name=TaxRate refersTo=F2
  hint: Select F2, click in the Name Box (where it says F2), type TaxRate and press Enter.
- step: Change the formula in D11 so it uses the name instead of F2.
  check: formula.is sheet=Invoice cell=D11 contains=TaxRate result=22
- step: Use the Name Box to jump to the range named **Deposits**. Type the **Sum** the status bar shows.
  answer: 275
  hint: Open the Name Box list and choose Deposits. Excel selects the range on the Details sheet.
:::

## Hyperlinks

A hyperlink can go to a web page, an email address, a file, or a place in the same workbook. Select a cell, then [[path: Insert > Link]] ([[key: Ctrl+K / Cmd+K]]):

- **Place in This Document** (desktop) or the sheet list (web) links to a sheet and cell.
- **Existing File or Web Page** links to an address.
- **Text to display** is what readers see. Make it say where the link goes, not "click here".

To take a link off and keep the text, right-click it and choose **Remove Hyperlink** (or **Remove Link**).

::: drill D2.09
title: Link the summary to the details and the web
badge: Link Maker
starter: starters/week-02/E2-L3-invoice.xlsx
platform: web
mos: 1.2.3
minutes: 6

- step: On the **Summary** sheet, make A5 a link to cell A1 on the **Details** sheet, showing the text **See every deposit**.
  check: link.present sheet=Summary location=Details label="See every deposit"
- step: Make A6 a link to **https://lakeshorepaddle.example/rates** that shows **Current rental rates**.
  check: link.present sheet=Summary url=lakeshorepaddle.example/rates descriptive=true
- step: The link in A8 goes to a site that's gone. Remove it.
  check: link.present sheet=Summary url=oldsite max=0
  hint: Right-click A8 and choose **Remove Hyperlink**. Then delete the text too if you like.
:::

::: quiz Q23.1
question: Which of these can be a name in Excel?
type: single
- [ ] Tax Rate
- [ ] 2027Tax
- [x] Tax_Rate_2027
- [ ] B2
explain: Names can't have spaces, can't start with a number, and can't look like a cell address such as B2.
:::

## Check your understanding

- Why does a name never need $ signs?
- How can you find every name a workbook uses?
- What makes link text good for screen reader users?
