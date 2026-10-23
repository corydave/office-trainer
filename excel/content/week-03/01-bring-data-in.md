---
type: lesson
id: XL3-L1
week: 3
order: 1
title: Bring Data In
objectives: [W3.1]
clos: [CLO2]
minutes: 40
status: draft
---

::: callout why
Most data you'll work with wasn't typed into Excel: it's exported from a booking system, a survey tool, a bank or a website. Importing it properly, instead of copying and pasting, keeps the columns straight and lets you refresh it when the source changes.
:::

## Your scenario

Lakeshore Paddle Company's booking system exports the summer's rentals (July and the first week of August) as a **.csv** file. This week you'll bring that data into Excel, turn it into a table, sort and filter it, and make the busiest boats stand out.

## Text files: .csv and .txt

A **.csv** (comma-separated values) file is plain text with one record per line and commas between the fields. A **.txt** file might use tabs instead. Excel can open either directly, but **Get Data** gives you more control:

1. [[path: Data > Get Data > From File > From Text/CSV]] (desktop Excel), and pick the file.
2. Check the preview: the first row should be the headers, and each column should have the right type.
3. Select **Load** to put the data on a new sheet as a table, or **Transform Data** to clean it first in the Power Query Editor.

The workbook remembers where the data came from. When the file changes, [[path: Data > Refresh All]] brings in the new version.

::: callout tip
Exports often contain the same record twice. In the Power Query Editor, [[path: Home > Remove Rows > Remove Duplicates]] removes them every time the data refreshes. On a normal sheet, [[path: Data > Remove Duplicates]] does it once.
:::

::: platform web
Excel for the web can open .csv files, but Get Data imports and the Power Query Editor need desktop Excel (some newer web versions have a limited Get Data). Use desktop Excel for these two drills.
:::

::: video
id: e3-l1-import
url: TODO
title: Import a .csv file with Get Data
captions: none
duration: 4:00
:::

::: drill D3.01
title: Import the rental export and remove duplicates
badge: Data Importer
starter: starters/week-03/E3-L1-import.xlsx
files: starters/week-03/rentals-july.csv
platform: desktop
mos: 1.1.1
other: E2.2
minutes: 8

Download **rentals-july.csv** first (it's next to the practice file above) and save it in your course folder.

- step: Import **rentals-july.csv** into the practice workbook with **Get Data**, so it loads as a table on its own sheet.
  check: connection.count min=1
  check: data.is table=* headers="Rental #,Date,Day,Boat,Hours,Rate,Customer" minRows=32
  hint: [[path: Data > Get Data > From File > From Text/CSV]], choose the file, then **Load**.
- step: The export has 4 duplicate rentals. Remove them, so 32 rentals are left.
  check: data.is table=* headers="Rental #,Date,Day,Boat,Hours,Rate,Customer" rows=32 noDuplicates=true
  hint: Select **Transform Data** (or right-click the query in **Queries & Connections** and choose **Edit**), select **every** column (click the first header, then Shift+click the last), then [[path: Home > Remove Rows > Remove Duplicates]], then **Close & Load**. Power Query only compares the columns you select.
:::

## Online sources

[[path: Data > Get Data > From Other Sources > From Web]] (desktop Excel) imports a table from a web page. Paste the page's address, wait for the **Navigator**, pick the table you want, and select **Load**.

::: drill D3.02
title: Import a table from a web page
badge: Web Importer
starter: starters/week-03/E3-L1-web.xlsx
platform: desktop
mos: 1.1.2
minutes: 6

- step: Import a table of the Finger Lakes from **https://en.wikipedia.org/wiki/Finger_Lakes** into the practice workbook.
  check: connection.count min=1
  check: table.count min=1
  hint: Copy the address, then [[path: Data > Get Data > From Other Sources > From Web]], paste it, and select **OK**. In the Navigator, choose a table with the lakes' names (preview each one), then **Load**.
:::

::: quiz Q31.1
question: Why import a .csv file with Get Data instead of copying it from Notepad and pasting it in?
type: single
- [ ] Pasting is not allowed for .csv files
- [x] Get Data keeps a connection, so you can refresh the data and repeat the same cleanup steps
- [ ] Get Data makes the file smaller
- [ ] Copy and paste can't handle more than 100 rows
explain: A Get Data import remembers the source and the cleanup steps. When next month's export arrives, Refresh does the whole job again.
:::

## Check your understanding

- What does the "C" in .csv stand for, and what separates the fields?
- What's the difference between removing duplicates in the Power Query Editor and with Data > Remove Duplicates?
- What does Refresh All do?
