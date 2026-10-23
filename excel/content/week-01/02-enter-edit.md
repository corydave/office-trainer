---
type: lesson
id: XL1-L2
week: 1
order: 2
title: Enter and Edit Data
objectives: [W1.3]
clos: [CLO1, CLO2]
minutes: 50
status: draft
---

::: callout why
Typing every month name or every rental number by hand is slow, and it's how mistakes creep in. Excel can fill a series for you, make room for new data without retyping, and paste exactly the part of a cell you want.
:::

## Entering data

- Type in a cell, then press [[key: Enter]] (move down) or [[key: Tab]] (move right). Press [[key: Esc]] to cancel.
- To change a cell, select it and type over it, or press [[key: F2]] to edit what's there.
- Text lines up on the left and numbers on the right. A number that lines up on the left is probably stored as text, and Excel won't add it up.

## Auto Fill

Select a cell (or two that show a pattern) and drag the **fill handle**, the small square at the bottom-right corner of the selection.

| Start with | Drag to get |
|---|---|
| **Jan** (or January, Mon, Q1) | Feb, Mar, Apr ... |
| **1001** and **1002** in two cells | 1003, 1004 ... (Excel follows the step) |
| **G-101** | G-102, G-103 ... (the number at the end counts up) |
| A formula | The same formula, adjusted for each row (Week 2) |

For more control, use [[path: Home > Fill > Series]] (desktop Excel) to set a step value, like counting by 5.

::: video
id: e1-l2-fill
url: TODO
title: Auto Fill, inserting and deleting
captions: none
duration: 3:30
:::

## Inserting and deleting

| To | Do this |
|---|---|
| Insert rows or columns | Select as many row numbers or column letters as you want to add, right-click, **Insert**. New ones go above or to the left |
| Delete rows or columns | Select them, right-click, **Delete** |
| Insert or delete just some cells | Select them, right-click, **Insert** or **Delete**, then choose which way the other cells shift |

::: callout warning
Deleting a cell isn't the same as clearing it. **Delete** ([[path: Home > Delete > Delete Cells]]) removes the cell and moves its neighbors to fill the gap. Pressing the [[key: Delete]] key only empties it.
:::

::: drill D1.06
title: Fill, insert and delete
badge: Grid Builder
starter: starters/week-01/E1-L2-log.xlsx
platform: web
mos: 2.1.2, 2.1.3, 2.1.4
other: N5, N9, N14, E2.1
minutes: 10

- step: On the **Fleet** sheet, Auto Fill the month names from **Jan** in B3 across to **Dec** in M3.
  check: cell.is sheet=Fleet cell=M3 value=Dec
  check: cell.is sheet=Fleet cell=G3 value=Jun
  hint: Select B3 and drag its fill handle to the right until the tip says Dec.
- step: On the **Log** sheet, insert a new row above the headers and type **June Rental Log** in A1.
  check: cell.is sheet=Log cell=A1 value="June Rental Log"
  check: cell.is sheet=Log cell=A2 value="Rental #"
  hint: Right-click the row number **1**, then **Insert**.
- step: Fill the rental numbers down column A, from **1001** to **1020**.
  check: cell.is sheet=Log cell=A22 number=1020
  check: cell.is sheet=Log cell=A12 number=1010
  hint: Select the first two numbers together, then drag the fill handle down to the last rental.
- step: Insert two columns before **Notes**, and give them the headers **Deposit** and **Paid?**.
  check: cell.is sheet=Log cell=E2 value=Deposit
  check: cell.is sheet=Log cell=F2 value="Paid?"
  check: cell.is sheet=Log cell=G2 value=Notes
  hint: Select the column letters **E** and **F** (Notes and the empty column after it), right-click, **Insert**. Two new columns appear before Notes.
- step: Someone typed an extra **Tandem kayak** in the Boat column next to rental 1010, which pushed every boat below it down one row. Delete just that cell and shift the cells below it up.
  check: cell.is sheet=Log cell=C12 value="Solo kayak"
  check: cell.is sheet=Log cell=C23 blank=true
  check: cell.is sheet=Log cell=A2 value="Rental #"
  hint: Right-click the extra cell, **Delete**, then **Shift cells up**.
:::

## Paste Special

Copying a cell copies everything: its value, its formula and its formatting. **Paste Special** pastes only part of it.

| Choose | You get |
|---|---|
| **Values** | The results only, without formulas. Use it to freeze numbers that shouldn't change |
| **Formats** | Only the look (like Format Painter) |
| **Transpose** | Rows become columns and columns become rows |

Copy, then [[path: Home > Paste > Paste Special]] ([[key: Ctrl+Alt+V / Ctrl+Cmd+V]]), or pick an icon from the **Paste** button's menu.

::: drill D1.07
title: Paste values and transpose
badge: Paste Pro
starter: starters/week-01/E1-L2-log.xlsx
platform: web
mos: 2.1.1
other: N6
minutes: 5

- step: On the **Prices** sheet, the **With tax** column is calculated. Copy C2:C7 and paste it into the **Quote** column (D2:D7) as **values only**, so the quotes won't change if the tax rate does.
  check: formula.is sheet=Prices range=D2:D7 hasFormula=false
  check: cell.is sheet=Prices cell=D3 number=54
  hint: Copy C2:C7, select D2, then [[path: Home > Paste > Paste Values]].
- step: Copy the price list (A1:B7) and paste it **transposed** starting at A10, so the boats run across row 10 and the rates across row 11.
  check: cell.is sheet=Prices cell=A10 value=Boat
  check: cell.is sheet=Prices cell=G10 value="Dry bag"
  check: cell.is sheet=Prices cell=A11 value="Daily rate"
  check: cell.is sheet=Prices cell=C11 number=50
  hint: Copy A1:B7, select A10, then [[path: Home > Paste > Transpose]].
:::

::: quiz Q12.1
question: You copy a cell that shows $54.00 because of a formula, then paste it with Paste Values. What's in the new cell?
type: single
- [ ] The formula, adjusted for the new cell
- [x] The number 54, with no formula
- [ ] Only the dollar-sign formatting
- [ ] Nothing until you press F9
explain: Paste Values keeps the result and drops the formula. The number won't change later, even if the cells the formula used do.
:::

## Check your understanding

- What do you need to select so that Auto Fill counts by 5?
- What's the difference between deleting a cell and pressing the Delete key?
- When would you paste values instead of the formula?
