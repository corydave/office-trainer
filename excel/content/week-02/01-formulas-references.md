---
type: lesson
id: XL2-L1
week: 2
order: 1
title: Formulas and Cell References
objectives: [W2.1, W2.5]
clos: [CLO1, CLO2]
minutes: 55
status: draft
---

::: callout why
A typed total is right until someone changes a number. A formula stays right. Getting the references right means you write a formula once, fill it down a whole column, and trust every answer.
:::

## Your scenario

This week you're building the paperwork side of Lakeshore Paddle Company: order totals, tax, an invoice, a rental-status list and a member list. Every number that can be calculated, Excel should calculate.

## Writing a formula

Every formula starts with **=**. Point at cells instead of typing their numbers, so the answer updates when they change.

| Operator | Means | Example |
|---|---|---|
| `+` `-` | Add, subtract | `=D9-D10` |
| `*` `/` | Multiply, divide | `=B4*C4` |
| `^` | Power | `=2^3` |
| `( )` | Do this part first | `=(D9-D10)*1.08` |

Excel follows the usual order: parentheses first, then powers, then multiply and divide, then add and subtract. `=10-2*3` is 4; `=(10-2)*3` is 24.

::: callout tip
Type **=**, then **click** the cells you want instead of typing their addresses. Excel adds the address for you and colors it so you can see what the formula uses.
:::

::: video
id: e2-l1-formulas
url: TODO
title: Formulas, filling and $ signs
captions: none
duration: 4:00
:::

::: drill D2.01
title: Line totals, a subtotal and a total
badge: Formula Starter
starter: starters/week-02/E2-L1-orders.xlsx
platform: web
other: N12
minutes: 6

- step: On the **Order** sheet, in D4 write a formula for the line total (**Qty** times **Price**), then fill it down to D8.
  check: formula.is sheet=Order range=D4:D8 filled=true refs=true
  check: formula.is sheet=Order cell=D4 result=70
  hint: In D4 type `=B4*C4` and press Enter, then drag D4's fill handle down to D8.
- step: In D9, add up the five line totals with a formula.
  check: formula.is sheet=Order cell=D9 refs=true result=275
- step: In D11, take the coupon (D10) off the subtotal **first**, then add 8% tax by multiplying by 1.08.
  check: formula.is sheet=Order cell=D11 contains="(" result=286.2
  hint: Without parentheses Excel multiplies before it subtracts. You want `=(D9-D10)*1.08`.
:::

## Relative, absolute and mixed references

When you fill a formula, Excel moves its references along with it: `=B4*C4` filled down becomes `=B5*C5`, `=B6*C6` and so on. That's a **relative** reference, and it's usually what you want.

Sometimes one cell should never move, like a tax rate. A **$** locks it:

| Reference | When filled, the column... | ...and the row | Use it for |
|---|---|---|---|
| **B1** (relative) | changes | changes | Most formulas |
| **$B$1** (absolute) | stays | stays | One fixed input, like a rate |
| **$A4** (mixed) | stays | changes | Always read column A |
| **B$3** (mixed) | changes | stays | Always read row 3 |

Click in a reference in the formula bar and press [[key: F4]] (Windows) or [[key: Cmd+T]] (Mac) to cycle through B1, $B$1, B$1 and $B1.

::: drill D2.02
title: Tax from one rate cell
badge: Anchor Setter
starter: starters/week-02/E2-L1-orders.xlsx
platform: web
mos: 4.1.1
minutes: 5

- step: On the **Tax** sheet, write a formula in C4 that multiplies the price by the tax rate in **B1**, using an **absolute** reference to B1. Fill it down to C9.
  check: formula.is sheet=Tax range=C4:C9 contains=$B$1 filled=true
  check: formula.is sheet=Tax cell=C9 result=0.4
  hint: Type `=B4*B1`, press [[key: F4]] once so it reads `=B4*$B$1`, then fill down. Without the $ signs, the filled formulas point at B2, B3 and so on, giving 0s and #VALUE! errors.
:::

::: drill D2.03
title: A price grid with mixed references
badge: Grid Master
starter: starters/week-02/E2-L1-orders.xlsx
platform: web
mos: 4.1.1
minutes: 7

- step: On the **Grid** sheet, write **one** formula in B4 that multiplies the rate in column A by the number of boats in row 3. Use mixed references so it works when filled across and down.
  check: formula.is sheet=Grid cell=B4 contains="$A4,B$3"
- step: Fill it across to F4 and down to F8.
  check: formula.is sheet=Grid range=B4:F8 filled=true
  check: formula.is sheet=Grid cell=F8 result=125
  hint: `=$A4*B$3` keeps the column A part and the row 3 part fixed. Fill right first, then fill the whole row down.
:::

## Finding formula mistakes

- **Show Formulas** ([[path: Formulas > Show Formulas]], [[key: Ctrl+`]]) shows every formula instead of its result. Typed numbers stand out right away.
- A green triangle in a cell's corner means Excel's **error checking** found something odd. Select the cell and the warning button to see what.
- Common errors: **#DIV/0!** (dividing by an empty or zero cell), **#NAME?** (a misspelled function or name), **#VALUE!** (math on text), **#REF!** (a cell the formula used was deleted).
- [[path: Formulas > Trace Precedents]] (desktop Excel) draws arrows to the cells a formula uses.

::: drill D2.04
title: Find the typed total and fix the error
badge: Formula Detective
starter: starters/week-02/E2-L1-orders.xlsx
platform: web
mos: 1.4.6
other: E3.5
minutes: 6

- step: On the **Weekly** sheet, turn on **Show Formulas**. Leave it on for this drill.
  check: view.is sheet=Weekly showFormulas=true
- step: The **Total** in B9 was typed by hand, and it's wrong. Replace it with a formula that adds up B2:B8.
  check: formula.is sheet=Weekly cell=B9 refs=true result=166
- step: The average in B11 shows **#DIV/0!** because **Days open** (B10) is empty. Type **7** in B10.
  check: formula.is sheet=Weekly cell=B11 result=23.714
  check: cell.is sheet=Weekly cell=B10 number=7
:::

::: quiz Q21.1
question: C4 has `=B4*$B$1`. You fill it down to C5. What does C5 have?
type: single
- [ ] `=B4*$B$1`
- [x] `=B5*$B$1`
- [ ] `=B5*$B$2`
- [ ] `=C5*$C$1`
explain: B4 is relative, so it moves down one row. $B$1 is absolute, so it stays exactly the same.
:::

## Check your understanding

- Why is a formula better than a typed total, even when the typed total is right?
- When would you use a mixed reference instead of an absolute one?
- What does #DIV/0! tell you, and how do you usually fix it?
