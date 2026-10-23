---
type: skillcheck
id: XL2-SC2
week: 2
order: 5
title: "Skill Check 2: Formulas That Work"
objectives: [W2.1, W2.2, W2.3, W2.4, W2.5]
clos: [CLO2]
minutes: 30
status: draft
---

One starter workbook, a set of tasks and a time limit, with no hints. Each task is worth one point. Calculated cells must hold formulas: a typed number that happens to be right doesn't count.

::: skillcheck SC2
title: June orders
badge: Skill Check 2
starter: starters/week-02/E2-SC2-orders.xlsx
platform: web
mos: 1.2.2, 1.2.3, 1.4.6, 2.1.5, 2.3.1, 2.3.2, 4.1.1, 4.2.1, 4.2.2, 4.2.3, 4.3.1, 4.3.2, 4.3.3
timeLimit: 30
attempts: 2
passing: 70

- task: Use the Name Box to select the range named **Rates**. Type the sum the status bar shows.
  answer: 315
- task: On **Orders**, fill **Amount** (E4:E11) with a formula: **Boats** times **Rate**.
  check: formula.is sheet=Orders range=E4:E11 filled=true refs=true
  check: formula.is sheet=Orders cell=E11 result=105
- task: Fill **Discount** (F4:F11) with the amount times the discount rate on the **Settings** sheet (B2), using an absolute reference.
  check: formula.is sheet=Orders range=F4:F11 contains=$B$2 filled=true
  check: formula.is sheet=Orders cell=F4 result=10.5
- task: Name the discount rate cell (**Settings** B2) **Discount**.
  check: name.is name=Discount refersTo=B2
- task: On **Settings**, in B4, calculate the discount on $100 using the name **Discount**.
  check: formula.is sheet=Settings cell=B4 contains=Discount result=10
- task: In E12, total the amounts with `SUM`; in E13, average them with `AVERAGE`.
  check: formula.is sheet=Orders cell=E12 uses=SUM result=705
  check: formula.is sheet=Orders cell=E13 uses=AVERAGE result=88.125
- task: In E14, count the customers (B4:B11) with `COUNTA`.
  check: formula.is sheet=Orders cell=E14 uses=COUNTA result=8
- task: Fill **Status** (G4:G11) with an `IF` formula: **Group** when the order has 3 or more boats, otherwise **Single**.
  check: formula.is sheet=Orders range=G4:G11 uses=IF filled=true
  check: formula.is sheet=Orders cell=G4 result=Group
  check: formula.is sheet=Orders cell=G5 result=Single
- task: Fill **Prefix** (H4:H11) with the first character of the order number, and **Customer (caps)** (I4:I11) with the customer's name in capitals.
  check: formula.is sheet=Orders range=H4:H11 uses=LEFT filled=true
  check: formula.is sheet=Orders range=I4:I11 uses=UPPER filled=true
  check: formula.is sheet=Orders cell=I4 result="AVERY CHEN"
- task: Fill **Label** (J4:J11) with `CONCAT`: the order number, a space, a hyphen, a space, and the customer (for example **O-301 - avery chen**).
  check: formula.is sheet=Orders range=J4:J11 uses=CONCAT filled=true
  check: formula.is sheet=Orders cell=J4 result="O-301 - avery chen"
- task: In L4, number the weeks 1 to 4 (L4:L7) with one `SEQUENCE` formula.
  check: formula.is sheet=Orders range=L4:L7 uses=SEQUENCE spill=true
- task: Make A15 a link to cell A1 on the **Settings** sheet, showing **Discount settings**.
  check: link.present sheet=Orders location=Settings label="Discount settings"
- task: Turn on **Show Formulas** on the Orders sheet, and save.
  check: view.is sheet=Orders showFormulas=true
:::
