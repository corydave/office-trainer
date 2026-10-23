---
type: project
id: XL3-P3
week: 3
order: 6
title: "Project 3: Data Report"
objectives: [W3.1, W3.2, W3.3, W3.4]
clos: [CLO1, CLO2]
minutes: 180
status: draft
rubricBadge: Data Reporter
---

::: callout why
Turning a raw export into one clear finding is the most useful thing you can do with Excel at work. This project is that whole process, start to finish.
:::

## What you'll make

One workbook, `Lastname_P3_DataReport.xlsx`, that turns a dataset into a one-page report:

1. **The data**, imported with Get Data into a named table, cleaned (no duplicates, consistent types, a real header row).
2. **A report sheet** with:
   - a one- or two-sentence **finding** at the top ("Weekend rentals are 40% of the total, mostly canoes")
   - a sorted, filtered or totaled view that supports it (a total row, a multi-level sort, a filter)
   - a `UNIQUE` list and a `SORT` list built from the table with structured references
   - conditional formatting and sparklines that show the pattern

**Choose your data:** the provided **lakeshore-season.csv** (download it from the preflight check below), data from a class, club or job (with permission, and nothing personal), or a public dataset such as a city or state open-data portal.

::: callout privacy
Don't use data with people's names, contact details or grades unless it's made up. Public datasets about places, weather, prices or sports work well.
:::

## Required features

- Imported with **Get Data** (the workbook keeps the connection)
- A named Excel table with a sensible style and a total row
- A multi-level sort and a filter (you can make a copy of the table on another sheet for one of them)
- At least one formula with a structured reference, one `UNIQUE` and one `SORT`
- Conditional formatting that doesn't rely on color alone, and sparklines
- Real sheet and table names, a Title property, and page setup so the report prints on one page

::: callout ai
Optional: ask an AI tool what questions your dataset could answer, then answer one yourself with the table. Check every number an AI tool gives you against your sheet.
:::

## Check before you submit

::: preflight P3-check
title: Preflight for your data report
starter: none
files: starters/week-03/lakeshore-season.csv
platform: web

- step: The workbook has at least one Excel table.
  check: table.count min=1
- step: Sheets and tables have real names, and the workbook has a Title.
  check: access.is rules="sheetNames,tableNames,title"
- step: The workbook keeps its data connection (Get Data).
  platform: desktop
  check: connection.count min=1
:::

## How it's graded

::: rubric P3
- P3-a | 25 | CLO2 | The data is imported and clean: no duplicates, consistent types, a real header row, a named table.
- P3-b | 20 | CLO2 | Sorting, filtering and the total row answer real questions about the data.
- P3-c | 20 | CLO2 | Formulas use structured references and dynamic array functions correctly.
- P3-d | 20 | CLO1, CLO2 | Conditional formatting and sparklines highlight the pattern without relying on color alone.
- P3-e | 15 | CLO1 | The written finding is accurate and supported by the sheet.
:::
