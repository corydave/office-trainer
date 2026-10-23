---
type: skillcheck
id: W1-SC1
week: 1
order: 5
title: "Skill Check 1: The Professional Document"
objectives: [W1.3, W1.4, W1.5, W1.6]
clos: [CLO1, CLO2]
minutes: 25
status: draft
---

This skill check works like a MOS exam task: one starter file, a set of tasks and a time limit. There are no hints. Each task is worth one point, and partial credit is given task by task.

**Before you start:** open the starter file, then select **Start**. The timer starts then. You can do the tasks in any order and select **Check my work** as often as you like; your score is what the document shows when you select **Submit** or when time runs out.

::: skillcheck SC1
title: Spring rental policy memo
badge: Skill Check 1
starter: starters/week-01/W1-SC1-memo.docx
platform: web
mos: 1.1.1, 1.1.4, 1.2.1, 1.3.2, 2.1.1, 2.1.2, 2.2.1, 2.2.2, 2.2.3, 2.2.5
timeLimit: 25
attempts: 2
passing: 70

- task: Use Find to count how many times the word **refund** appears in the memo, and type the number here.
  answer: 4
- task: Remove every empty paragraph and every double space.
  check: para.noEmpty max=0
  check: text.noDoubleSpaces
- task: Replace every **Northside Rentals** with **Lakeshore Paddle Company**.
  check: text.absent value="Northside Rentals"
  check: text.present value="Lakeshore Paddle Company" min=3
- task: Add a registered trademark symbol (®) directly after **PaddlePass** the first time it appears.
  check: text.present value="PaddlePass®"
- task: Apply a text effect to the title **Spring Rental Policy**.
  check: effect.present para="Spring Rental Policy" kind=any
- task: Make the **Late returns** heading match the formatting of the **Pickup times** heading.
  check: format.same a="Pickup times" b="Late returns"
- task: Give the paragraph that starts **Rentals must be** 1.5 line spacing, 6 pt after and a 0.5 in first-line indent.
  check: para.is para="Rentals must be" lineSpacing=1.5 spaceAfter=6pt firstLineIndent=0.5in
- task: Remove all formatting from the paragraph that starts **Damage fees**.
  check: format.clean para="Damage fees"
- task: Set 1 in margins, Letter size and Portrait orientation.
  check: page.is margins=1in size=letter orientation=portrait
- task: Set the document Title property to **Spring Rental Policy** and add the tag **policy**.
  check: props.is title="Spring Rental Policy" keywords="policy"
  platform: desktop
:::

::: callout note
MO-110 objective 1.3.1 (save in other formats) is assessed through the PDF you submit with Project 1, because a skill check cannot see files outside the document.
:::
