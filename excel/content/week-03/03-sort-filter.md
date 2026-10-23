---
type: lesson
id: XL3-L3
week: 3
order: 3
title: Sort and Filter
objectives: [W3.3]
clos: [CLO2]
minutes: 40
status: draft
---

::: callout why
"Which boat do weekend customers rent most?" is a question you can answer in thirty seconds with a sort and a filter, and in an hour without them.
:::

## Sorting

- **Quick sort:** click in a column, then [[path: Data > Sort A to Z]] or **Sort Z to A** (for numbers: smallest to largest, or largest to smallest; for dates: oldest or newest first).
- **Several levels:** [[path: Data > Sort]], then **Add Level**. Excel sorts by the first level, and uses the second level only to order rows that tie on the first.

::: callout warning
Always sort the whole table, never one column on its own. If you select just one column and sort it, Excel asks whether to expand the selection. Say yes, or every row's data gets scrambled.
:::

## Filtering

Tables have filter buttons in their header row (on a normal range, use [[path: Data > Filter]]). Select a button and tick the values to show; the others are hidden, not deleted. Number and date columns also offer **Number Filters** like Greater Than and **Date Filters** like Next Week.

To show everything again, select **Clear Filter** in the button's menu, or [[path: Data > Clear]].

::: video
id: e3-l3-sort
url: TODO
title: Multi-level sorts and filters
captions: none
duration: 3:30
:::

::: drill D3.07
title: Sort by two columns and filter the weekend
badge: Sort and Filter Pro
starter: starters/week-03/E3-L3-sort.xlsx
platform: web
mos: 3.3.1, 3.3.2
other: N15
minutes: 7

- step: Sort the **Rentals** table by **Boat** (A to Z), then by **Date** (newest first).
  check: sort.is table=Rentals by=Boat then=Date thenOrder=desc
  hint: Click in the table, [[path: Data > Sort]], sort by Boat A to Z, then **Add Level**: Date, **Newest to Oldest**.
- step: Filter the table so only **Sat** and **Sun** rentals show.
  check: filter.is table=Rentals column=Day shows="Sat,Sun"
  hint: Select the Day column's filter button, untick **Select All**, then tick Sat and Sun.
:::

## SORT() and UNIQUE(): sorting with a formula

Sorting the table changes the table. The `SORT` and `UNIQUE` functions make a **sorted or de-duplicated copy** somewhere else, and keep it up to date:

- `=UNIQUE(Rentals[Boat])` lists each boat type once.
- `=SORT(Rentals[Hours],,-1)` lists the hours from largest to smallest (`-1` means descending; leave it out for ascending).

Both are dynamic arrays: type them once in the top cell and the results spill down.

::: drill D3.08
title: A list of boats and the longest rentals
badge: Dynamic Lister
starter: starters/week-03/E3-L3-sort.xlsx
platform: web
mos: 4.2.4, 4.2.5
minutes: 5

- step: On **Summary**, in A2, list each boat type once with `UNIQUE` and the Boat column of the Rentals table.
  check: formula.is sheet=Summary range=A2:A5 uses=UNIQUE spill=true
  check: cell.is sheet=Summary cell=A6 blank=true
- step: In C2, list every rental's hours from largest to smallest with `SORT`.
  check: formula.is sheet=Summary range=C2:C33 uses=SORT spill=true
  check: formula.is sheet=Summary cell=C2 result=5
  hint: `=SORT(Rentals[Hours],,-1)`. The two commas skip an argument you don't need here.
:::

::: quiz Q33.1
question: You filter a table to show only canoes. What happens to the other rows?
type: single
- [ ] They're deleted
- [x] They're hidden until you clear the filter
- [ ] They move to the bottom of the table
- [ ] They're copied to a new sheet
explain: Filters only hide rows. Clear the filter and every row comes back.
:::

## Check your understanding

- When does Excel use the second level of a sort?
- What's the difference between sorting a table and using `SORT()`?
- How can you tell that a column is filtered?
