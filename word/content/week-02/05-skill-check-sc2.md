---
type: skillcheck
id: W2-SC2
week: 2
order: 5
title: "Skill Check 2: The Long Document"
objectives: [W2.1, W2.2, W2.3, W2.4, W2.5]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

Like a MOS exam task: one starter file, a set of tasks and a time limit, with no hints. Each task is worth one point, and you can do them in any order.

**Before you start:** open the starter file, then select **Start**. Select **Check my work** as often as you like; your score is what the document shows when you select **Submit**, or when time runs out. Tasks marked for desktop Word aren't counted in Word for the web.

::: skillcheck SC2
title: Staff handbook
badge: Skill Check 2
starter: starters/week-02/W2-SC2-handbook.docx
platform: web
mos: 1.1.2, 1.1.3, 1.2.2, 1.2.3, 2.2.4, 2.3.2, 2.3.3, 3.3.1, 3.3.2, 3.3.4, 3.3.5, 4.1.1, 4.1.2, 4.2.1, 4.2.2
timeLimit: 30
attempts: 2
passing: 70

- task: Apply **Heading 1** to **Welcome**, **Your first week**, **Rentals**, **Safety on the water**, **Timeline**, **Budget** and **Appendix: Contact List**.
  check: style.is para="Welcome" name="Heading 1"
  check: style.is para="Your first week" name="Heading 1"
  check: style.is para="Safety on the water" name="Heading 1"
  check: style.is para="Appendix: Contact List" name="Heading 1"
- task: Apply **Heading 2** to **Opening the shop** and **Closing the shop**.
  check: style.is para="Opening the shop" name="Heading 2"
  check: style.is para="Closing the shop" name="Heading 2"
- task: Modify the **Heading 1** style so it's **16 pt**.
  check: style.def name="Heading 1" size=16
  platform: desktop
- task: Move the **Budget** section, with its paragraph, above **Timeline**.
  check: para.order first="Budget" then="Timeline"
  check: para.order first="Our seasonal budget" then="Training runs"
- task: Add a bookmark named **Budget** on the Budget heading, and make the words **see the budget** in the first paragraph a link to it.
  check: bookmark.is name="Budget" at="Budget"
  check: link.present label="see the budget" anchor="Budget"
- task: Turn the four lines under **Opening checklist** into a numbered list, and make **Count the cash drawer.** a level 2 item.
  check: list.is para="Unlock the boat racks" kind=number
  check: list.is para="Turn on the dock lights" kind=number
  check: list.is para="Count the cash drawer." level=2
- task: The **Closing checklist** starts at 5. Restart it at 1.
  check: list.is para="Lock the boat racks" number=1
- task: Add a page number in the footer, and turn on **Different First Page**.
  check: field.present code="PAGE" in=footer
  check: section.is index=1 firstPageDifferent=true
- task: Insert a **Next Page** section break before **Appendix: Contact List** and make only that section **landscape**.
  check: section.is para="Appendix: Contact List" orientation=landscape
  check: section.is para="Welcome" orientation=portrait
  platform: desktop
- task: Add a footnote after **at all times.** in the Safety section.
  check: notes.is kind=footnote min=1
- task: Change the footnote numbering to **i, ii, iii**.
  check: notes.is kind=footnote min=1 format=lowerRoman
  platform: desktop
- task: Insert an automatic table of contents after the first paragraph, showing **2 levels**.
  check: toc.is levels=2
  platform: desktop
:::
