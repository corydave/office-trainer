---
type: lesson
id: XL4-L5
week: 4
order: 5
title: What-If Analysis
objectives: [W4.5]
clos: [CLO3]
minutes: 50
status: draft
---

::: callout why
"What if the boat costs more?" "What if rentals drop?" "How much can we spend and keep the payment at $250?" Excel can answer these in seconds, without retyping numbers over and over, and keep each answer so you can compare them.
:::

## Loan payments with PMT

`PMT` works out the payment on a loan. It needs three things, all in the **same time unit** (months, for monthly payments):

`=PMT(rate per period, number of periods, loan amount)`

| Part | For a 5-year loan at 6% a year, paid monthly |
|---|---|
| Rate per period | `B5/12` (the yearly rate divided by 12) |
| Number of periods | `B6*12` (years times 12) |
| Loan amount | `B3-B4` (the price minus the down payment) |

`PMT` shows money you pay out as a **negative** number. Put a minus sign in front of the loan amount, `-(B3-B4)`, to show the payment as a positive number.

::: video
id: e4-l5-pmt-goalseek
url: TODO
title: PMT and Goal Seek
captions: none
duration: 4:00
:::

## Goal Seek

**Goal Seek** works backward: you name the answer you want, and Excel finds the input that gives it. [[path: Data > What-If Analysis > Goal Seek]] (desktop) asks for three things:

- **Set cell**: the formula whose answer you want (the payment)
- **To value**: the answer you want (250)
- **By changing cell**: the input Excel may change (the price). It must be a number, not a formula.

::: drill D4.08
title: What boat can we afford?
badge: Goal Seeker
starter: starters/week-04/E4-L5-whatif.xlsx
platform: desktop
other: E3.4
minutes: 8

- step: On **Loan**, in B8, use `PMT` to work out the monthly payment, shown as a positive number.
  check: formula.is sheet=Loan cell=B8 uses=PMT min=0
  hint: `=PMT(B5/12,B6*12,-(B3-B4))`. Before Goal Seek, it shows $193.33.
- step: Use **Goal Seek** to find the boat price (B3) that makes the payment exactly **$250**.
  check: cell.is sheet=Loan cell=B3 number=14931.39 tolerance=1
  check: formula.is sheet=Loan cell=B8 uses=PMT result=250
  hint: [[path: Data > What-If Analysis > Goal Seek]]. Set cell B8, To value 250, By changing cell B3, then **OK** twice.
:::

## Scenario Manager

A **scenario** is a named set of input values, like "Low season" or "High season". [[path: Data > What-If Analysis > Scenario Manager]] (desktop) saves several, switches between them with **Show**, and makes a **Summary** report that lists every scenario's result side by side on a new sheet.

::: callout tip
Before you add scenarios, save your real numbers as a scenario too (for example "Expected"), so you can get back to them after you've tried the others.
:::

::: drill D4.09
title: Low, expected and high seasons
badge: Scenario Planner
starter: starters/week-04/E4-L5-whatif.xlsx
platform: desktop
other: E3.4
minutes: 8

- step: On **Season**, add three scenarios that change **Rentals per week** (B4): **Low** = 35, **Expected** = 60, **High** = 90.
  check: scenario.is sheet=Season names="Low,Expected,High" changing=B4
  hint: [[path: Data > What-If Analysis > Scenario Manager > Add]]. Type the name, set Changing cells to B4, **OK**, type the value, then **Add** for the next one.
- step: Make a scenario **summary** report with **Season profit** (B8) as the result cell.
  check: sheet.is name="Scenario Summary"
  hint: In Scenario Manager, choose **Summary**, leave **Scenario summary** selected, and set Result cells to B8.
:::

## Data tables

A **one-variable data table** shows a formula's answer for a whole list of inputs at once, like the payment at five interest rates:

1. List the inputs down a column (D5:D9).
2. In the cell one row up and one column right of the list (E4), point to the formula: `=B8`.
3. Select the whole block (D4:E9), then [[path: Data > What-If Analysis > Data Table]].
4. The inputs run **down a column**, so put the rate cell (B5) in **Column input cell**, and choose **OK**.

Excel fills E5:E9 with a special `{=TABLE(,B5)}` formula. You can't edit one cell of it: select all of E5:E9 to change or delete it.

::: video
id: e4-l5-scenarios-tables
url: TODO
title: Scenario Manager and data tables
captions: none
duration: 4:30
:::

::: drill D4.10
title: The payment at five interest rates
badge: Table Builder
starter: starters/week-04/E4-L5-whatif.xlsx
platform: desktop
minutes: 6

Do D4.08 first, so B8 has the payment formula.

- step: In E4, point to the payment: `=B8`.
  check: formula.is sheet=Loan cell=E4 contains=B8
- step: Build a one-variable data table in D4:E9 that shows the payment for each rate in D5:D9.
  check: datatable.is sheet=Loan range=E5:E9 input=B5
  hint: Select D4:E9, [[path: Data > What-If Analysis > Data Table]], leave Row input cell empty, Column input cell B5, **OK**.
:::

::: quiz Q45.1
question: You know the answer you want (a $250 payment) but not the input that gives it. Which tool fits?
type: single
- [x] Goal Seek
- [ ] A data table
- [ ] Scenario Manager
- [ ] Conditional formatting
explain: Goal Seek works backward from the answer to one input. A data table and Scenario Manager show answers for inputs you choose.
:::

## Check your understanding

- Why divide the yearly rate by 12 in `PMT`?
- Why must Goal Seek's "By changing cell" hold a number and not a formula?
- When is a data table better than scenarios, and the other way around?
