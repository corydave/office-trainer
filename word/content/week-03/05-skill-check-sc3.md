---
type: skillcheck
id: W3-SC3
week: 3
order: 5
title: "Skill Check 3: The Visual Document"
objectives: [W3.1, W3.2, W3.3, W3.5]
clos: [CLO3, CLO4, CLO5]
minutes: 30
status: draft
---

Like a MOS exam task: one starter file, a set of tasks and a time limit, with no hints. Each task is worth one point, and you can do them in any order.

**Before you start:** open the starter file, then select **Start**. Select **Check my work** as often as you like; your score is what the document shows when you select **Submit**, or when time runs out. Tasks marked for desktop Word aren't counted in Word for the web.

::: skillcheck SC3
title: Sunset paddle flyer
badge: Skill Check 3
starter: starters/week-03/W3-SC3-flyer.docx
platform: web
mos: 1.2.4, 2.3.1, 3.1.1, 3.1.3, 3.2.1, 3.2.3, 3.2.6, 5.1.2, 5.1.4, 5.1.6, 5.2.2, 5.3.2, 5.4.1, 5.4.2, 5.4.3
timeLimit: 30
attempts: 2
passing: 70

- task: Apply any picture style to the sunset photo.
  check: object.is kind=picture name="Sunset photo" pictureStyle=true
- task: Give the sunset photo alt text.
  check: object.is kind=picture name="Sunset photo" alt=true
- task: Insert a text box that says **Bring a friend for free**, with **Square** wrapping, positioned at the **top right** of the page.
  check: object.is kind=textbox contains="Bring a friend" wrap=square halign=right valign=top
- task: Insert any shape and type **Limited spots** in it.
  check: object.is kind=shape any=true says="Limited spots"
- task: Insert a SmartArt graphic with three items: **Book**, **Check in** and **Paddle**.
  check: object.is kind=smartart items=3
  platform: desktop
- task: Convert the five tab-separated lines under **Dates** into a table.
  check: table.is contains="Lighthouse loop" cols=3 minRows=5
  platform: desktop
- task: In the **Prices** table, merge the two cells of the top row.
  check: table.is contains="Sunset series prices" merged=true
- task: Sort the prices table by **Price**, lowest first.
  check: table.is contains="Sunset series prices" sortedBy=2
  platform: desktop
- task: Make the top two rows of the prices table repeat as header rows.
  check: table.is contains="Sunset series prices" headerRepeat=true
- task: Put the three paragraphs under **About the series** in **two columns**, leaving the rest of the flyer in one column.
  check: section.is para="Each tour lasts" columns=2
  check: section.is para="Sunset Paddle Series" columns=1
  platform: desktop
- task: Add a **SAMPLE** watermark.
  check: page.background watermark="SAMPLE"
  platform: desktop
:::
