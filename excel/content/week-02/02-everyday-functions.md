---
type: lesson
id: XL2-L2
week: 2
order: 2
title: Everyday Functions
objectives: [W2.2]
clos: [CLO2]
minutes: 50
status: draft
---

::: callout why
Adding forty cells with + signs is slow and easy to get wrong. A function does a whole job in one go: `=SUM(B4:B43)` adds them all, and keeps adding them if you insert a row in the middle.
:::

## How functions work

A function is a name, then parentheses with its **arguments** (what it works on): `=SUM(B4:E4)`. A colon means "through": B4:E4 is B4, C4, D4 and E4. Commas separate arguments: `=SUM(B4:E4,G4)`.

| Function | Gives you |
|---|---|
| `SUM` | The total |
| `AVERAGE` | The average (mean) |
| `MAX` / `MIN` | The largest / smallest value |
| `COUNT` | How many cells hold **numbers** |
| `COUNTA` | How many cells are **not empty** (numbers, text, anything) |
| `COUNTBLANK` | How many cells are **empty** |
| `IF` | One result if a test is true, another if it's false |

The fastest way in: select the cell under (or beside) your numbers, then [[path: Home > AutoSum]] ([[key: Alt+= / Cmd+Shift+T]]). The arrow next to AutoSum offers Average, Count Numbers, Max and Min.

::: callout tip
The **status bar** shows the sum, average and count of whatever you select, without any formula. It's a quick way to check a formula's answer. Unlike AutoSum, it doesn't put anything in the sheet.
:::

::: video
id: e2-l2-functions
url: TODO
title: SUM, AVERAGE, MIN, MAX and AutoSum
captions: none
duration: 3:30
:::

::: drill D2.05
title: Totals, averages, highs and lows
badge: Function Finder
starter: starters/week-02/E2-L2-functions.xlsx
platform: web
mos: 4.2.1
other: N14
minutes: 8

- step: On the **Weeks** sheet, total each boat's four weeks in column F (F4:F7) with `SUM`.
  check: formula.is sheet=Weeks range=F4:F7 uses=SUM filled=true
  check: formula.is sheet=Weeks cell=F4 result=97
  hint: Select F4, select **AutoSum**, check that it chose B4:E4, press Enter, then fill down.
- step: In row 8 (**Total**), total each week's column (B8:E8).
  check: formula.is sheet=Weeks range=B8:E8 uses=SUM filled=true
  check: formula.is sheet=Weeks cell=E8 result=86
- step: In row 9, find each week's **Average**; in row 10 the **Highest**; in row 11 the **Lowest**. Use only the four boat rows (rows 4 to 7), not the totals.
  check: formula.is sheet=Weeks range=B9:E9 uses=AVERAGE filled=true
  check: formula.is sheet=Weeks cell=E9 result=21.5
  check: formula.is sheet=Weeks range=B10:E10 uses=MAX filled=true
  check: formula.is sheet=Weeks cell=E10 result=31
  check: formula.is sheet=Weeks range=B11:E11 uses=MIN filled=true
  check: formula.is sheet=Weeks cell=B11 result=9
  hint: AutoSum guesses the range above it, which would include the Total row. Drag to select B4:B7 before you press Enter.
:::

## Counting

The three counting functions answer different questions about the same cells:

- `=COUNT(B4:B18)`: how many payments are recorded **as numbers**
- `=COUNTA(A4:A18)`: how many names were entered at all
- `=COUNTBLANK(B4:B18)`: how many payment cells are still empty

::: drill D2.06
title: Count the sign-ups
badge: Counter
starter: starters/week-02/E2-L2-functions.xlsx
platform: web
mos: 4.2.2
minutes: 5

- step: On the **Signups** sheet, in E4 count the payments that are numbers (B4:B18).
  check: formula.is sheet=Signups cell=E4 uses=COUNT result=8
- step: In E5, count the names entered (A4:A18).
  check: formula.is sheet=Signups cell=E5 uses=COUNTA result=12
- step: In E6, count the empty payment cells (B4:B18).
  check: formula.is sheet=Signups cell=E6 uses=COUNTBLANK result=6
  hint: The answer is 6: the 3 people who haven't paid, plus the 3 blank spacer rows between the groups. `COUNTBLANK` counts every empty cell in the range.
:::

::: quiz Q22.1
question: B4:B18 holds eight numbers, one cell that says "paid cash" and six empty cells. What does `=COUNTA(B4:B18)` give?
type: single
- [ ] 8
- [x] 9
- [ ] 6
- [ ] 15
explain: COUNTA counts every cell that isn't empty, text included: the eight numbers plus "paid cash". COUNT would give 8.
:::

## IF: a decision in a formula

`=IF(test, value if true, value if false)`

`=IF(C4>B4,"Late","On time")` reads: if the return time is greater than the due time, show Late; otherwise show On time. Text in a formula goes in double quotes. Tests use `=`, `>`, `<`, `>=`, `<=` and `<>` (not equal).

::: drill D2.07
title: Late or on time?
badge: Decision Maker
starter: starters/week-02/E2-L2-functions.xlsx
platform: web
mos: 4.2.3
minutes: 6

- step: On the **Returns** sheet, write an `IF` formula in D4 that shows **Late** when the return time (C4) is later than the due time (B4), and **On time** otherwise. Fill it down to D15.
  check: formula.is sheet=Returns range=D4:D15 uses=IF filled=true
  check: formula.is sheet=Returns cell=D5 result=Late
  check: formula.is sheet=Returns cell=D4 result="On time"
  hint: `=IF(C4>B4,"Late","On time")`. Check the quotes: they must be straight quotes around the text.
:::

## Check your understanding

- Why does AutoSum sometimes pick the wrong range, and how do you fix it?
- What's the difference between COUNT and COUNTA?
- In `=IF(C4>B4,"Late","On time")`, what happens when C4 equals B4?
