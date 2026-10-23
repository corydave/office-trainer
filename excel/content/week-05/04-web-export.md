---
type: lesson
id: XL5-L4
week: 5
order: 4
title: Share as a Web Page, PDF or CSV
objectives: [W5.4]
clos: [CLO6]
minutes: 30
status: draft
---

::: callout why
Not everyone who needs your numbers has Excel. A web page anyone can open in a browser, a PDF that prints the same everywhere, and a CSV that any program can import each reach a different reader.
:::

## Three ways out of Excel

| Format | Choose it when | How |
|---|---|---|
| **Web page** (.htm) | People will view it in a browser or on a website | [[path: File > Save As]], type **Web Page (*.htm; *.html)** (desktop) |
| **PDF** | People will read or print it and shouldn't change it | [[path: File > Export > Create PDF/XPS]] (desktop), or [[path: File > Export > Download as PDF]] (web) |
| **CSV** | Another program (a database, a mailing tool) will import the data | [[path: File > Save As]], type **CSV UTF-8 (Comma delimited)** (desktop), or [[path: File > Export > Download this sheet as CSV]] (web) |

A CSV keeps only the **values on one sheet**: no formatting, formulas, charts or other sheets. Excel warns you about this when you save. Always keep the .xlsx as your working copy.

::: callout tip
Set the **Title** property first: a PDF keeps it, and PDF readers and screen readers use it. For a web page, **Change Title** in the Save As dialog sets the page's own title.
:::

## Saving as a web page

In the **Save As** dialog, after you choose **Web Page**, you can save the **Entire Workbook** or just the **Selection: Sheet**, and **Change Title** to set the heading at the top of the page. Excel makes a .htm file and, for a workbook with more than one sheet or with charts, a folder of supporting files beside it. Keep them together.

Open the .htm file in a browser (double-click it, or drag it onto a browser window) to see how it looks: this is how it will look on a website.

::: callout note
Today, most people share a live worksheet on the web from OneDrive instead: in Excel for the web, [[path: File > Share > Embed]] gives you code for a web page that always shows the current numbers, and **Share** gives you a link. (With a college or work account, viewers may have to sign in.) A saved .htm page is a snapshot that doesn't update.
:::

::: video
id: e5-l4-web
url: TODO
title: Save as a web page, PDF and CSV
captions: none
duration: 3:30
:::

::: drill D5.07
title: Put the schedule on the web
badge: Web Publisher
starter: starters/week-05/E5-L4-schedule.xlsx
platform: web
mos: 1.5.2
minutes: 10

- step: Set the workbook's **Title** property to **Paddle Day schedule**, and save.
  platform: desktop
  check: props.is title="Paddle Day schedule"
  hint: [[path: File > Info]], then type in the **Title** box on the right.
- step: Save the **Schedule** sheet as a web page, then open it in a browser.
  platform: desktop
  screenshot: The schedule open in a web browser, with the file name or title in the tab.
  example: TODO
  describe: Name the browser you used and say what the browser tab's title shows.
  hint: [[path: File > Save As]], **Web Page (*.htm; *.html)**, **Selection: Sheet**, **Save**. Then open the .htm file in your browser.
- step: Save a PDF of the schedule.
  screenshot: The PDF of the schedule open in a PDF reader or browser.
  example: TODO
  describe: Say how many pages the PDF has and whether the whole schedule fits.
- step: Save a copy of the schedule as **CSV UTF-8**, then open the .csv file in Notepad or TextEdit.
  hint: Desktop: [[path: File > Save As]], type **CSV UTF-8 (Comma delimited)**. Web: [[path: File > Export > Download this sheet as CSV]].
  screenshot: The .csv file open in a text editor, showing the commas between the values.
  example: TODO
  describe: Say what happened to the title row's formatting and why.
:::

::: quiz Q54.1
question: A volunteer coordinator needs the sign-up list to import into a mailing program. Which format should you send?
type: single
- [ ] A web page
- [ ] A PDF
- [x] A CSV file
- [ ] A picture of the sheet
explain: Programs import CSV files easily because they're plain values separated by commas. A PDF or a web page is for people to read.
:::

## Check your understanding

- What does a CSV lose, and why is that OK for importing?
- When is a live, embedded worksheet better than a saved web page?
- Why set the Title property before you save as a web page or PDF?
