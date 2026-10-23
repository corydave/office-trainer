---
type: lesson
id: XL2-L4
week: 2
order: 4
title: Text Functions and Generated Data
objectives: [W2.4]
clos: [CLO2]
minutes: 45
status: draft
---

::: callout why
Exports from other systems arrive with names in capitals, codes stuck together and emails in a column you have to retype. Text functions clean and combine them in one formula, filled down a thousand rows as easily as ten.
:::

## Pulling text apart

| Function | Gives you | `=...("LPC-27-4417")` |
|---|---|---|
| `LEFT(text, n)` | The first n characters | `=LEFT(C4,3)` → LPC |
| `RIGHT(text, n)` | The last n characters | `=RIGHT(C4,4)` → 4417 |
| `MID(text, start, n)` | n characters, starting at position start | `=MID(C4,5,2)` → 27 |
| `LEN(text)` | How many characters | `=LEN(C4)` → 11 |
| `UPPER(text)` / `LOWER(text)` | The same text in capitals / lowercase | `=UPPER(B4)` |

::: callout tip
**Flash Fill** ([[path: Data > Flash Fill]], [[key: Ctrl+E]]) does simple splits without a formula: type the first answer yourself, start the second, and Excel fills the rest. It's quick, but it doesn't update if the source changes. A formula does.
:::

::: video
id: e2-l4-text
url: TODO
title: LEFT, RIGHT, MID and friends
captions: none
duration: 3:30
:::

::: drill D2.10
title: Split member codes and fix capitals
badge: Text Surgeon
starter: starters/week-02/E2-L4-members.xlsx
platform: web
mos: 4.3.1, 4.3.2
other: E2.1
minutes: 9

Member codes look like **LPC-27-4417**: a prefix, the year they joined, and a number.

- step: In **Prefix** (E4), use `LEFT` to pull the first 3 characters of the member code. Fill it down to E11.
  check: formula.is sheet=Members range=E4:E11 uses=LEFT filled=true
  check: formula.is sheet=Members cell=E4 result=LPC
- step: In **Year** (F4), use `MID` to pull the 2 characters that start at position 5. Fill down.
  check: formula.is sheet=Members range=F4:F11 uses=MID filled=true
  check: formula.is sheet=Members cell=F4 result=27
- step: In **Last 4** (G4), use `RIGHT` to pull the last 4 characters. Fill down.
  check: formula.is sheet=Members range=G4:G11 uses=RIGHT filled=true
  check: formula.is sheet=Members cell=G4 result=4417
- step: In **Last name** (H4), show the last name in capitals with `UPPER`. Fill down.
  check: formula.is sheet=Members range=H4:H11 uses=UPPER filled=true
  check: formula.is sheet=Members cell=H5 result=CHEN
- step: In **Code length** (I4), use `LEN` to count the characters in the code, so you can spot a mistyped one. Fill down.
  check: formula.is sheet=Members range=I4:I11 uses=LEN filled=true
  check: formula.is sheet=Members cell=I4 result=11
- step: Some emails were typed in capitals. In K3 type the header **Email (lowercase)**, then in K4 use `LOWER` on the email. Fill down.
  check: formula.is sheet=Members range=K4:K11 uses=LOWER filled=true
  check: formula.is sheet=Members cell=K4 result=jrivera@example.com
:::

## Putting text together

- `=CONCAT(A4," ",B4)` joins its arguments: first name, a space, last name. (The older `&` does the same: `=A4&" "&B4`.)
- `=TEXTJOIN("; ",TRUE,D4:D11)` joins a whole range with a separator between each item, skipping empty cells (that's the TRUE). Perfect for an email list.

::: drill D2.11
title: Full names and an email list
badge: Text Joiner
starter: starters/week-02/E2-L4-members.xlsx
platform: web
mos: 4.3.3
minutes: 5

- step: In **Full name** (J4), use `CONCAT` to join the first name, a space and the last name. Fill down. (It keeps the original capitals, so J4 shows **jordan RIVERA**. That's expected.)
  check: formula.is sheet=Members range=J4:J11 uses=CONCAT filled=true
  check: formula.is sheet=Members cell=J4 result="jordan RIVERA"
- step: In B13, next to **Email list**, use `TEXTJOIN` to join every email in D4:D11, separated by a semicolon and a space.
  check: formula.is sheet=Members cell=B13 uses=TEXTJOIN contains="; "
:::

## Generating numbers

- `=SEQUENCE(10)` fills ten cells with 1 to 10. It's a **dynamic array**: you type it once, in the top cell, and the results **spill** into the cells below.
- `=RANDBETWEEN(1,4)` gives a random whole number from 1 to 4, and a new one every time the sheet recalculates. Use it for practice data, never for real records.

::: drill D2.12
title: Number lines and make practice data
badge: Number Generator
starter: starters/week-02/E2-L4-members.xlsx
platform: web
mos: 2.1.5
other: E3.3
minutes: 4

- step: On the **Practice** sheet, type one `SEQUENCE` formula in A2 that numbers the lines 1 to 10 (down to A11).
  check: formula.is sheet=Practice range=A2:A11 uses=SEQUENCE spill=true
  hint: `=SEQUENCE(10)` in A2 only. If you see #SPILL!, something is in the way below it.
- step: In B2:B11, fill in practice hours from 1 to 4 with `RANDBETWEEN`.
  check: formula.is sheet=Practice range=B2:B11 uses=RANDBETWEEN filled=true
:::

::: quiz Q24.1
question: A2 has `=SEQUENCE(10)`. You click A5. What does the formula bar show?
type: single
- [ ] `=SEQUENCE(10)`
- [x] The formula, grayed out: A5 is part of A2's spilled results
- [ ] Nothing: A5 is empty
- [ ] `=SEQUENCE(7)`
explain: Spilled cells show the formula in gray. Only A2 really holds it; to change or delete the results, change A2.
:::

## Check your understanding

- When is Flash Fill enough, and when do you need a formula?
- What does the TRUE in TEXTJOIN do?
- What does #SPILL! mean?
