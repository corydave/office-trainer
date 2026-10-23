---
type: skillcheck
id: W4-SC4
week: 4
order: 6
title: "Skill Check 4: Mail Merge"
objectives: [W4.3, W4.4]
clos: [CLO7]
minutes: 25
status: draft
---

A timed merge task: a renewal letter and a member list. Each task is worth one point, and you can do them in any order.

**Before you start:** download the member list (next to the practice file), open the starter file, then select **Start**. Mail merge needs desktop Word.

::: skillcheck SC4
title: Paddle club renewal letters
badge: Skill Check 4
starter: starters/week-04/W4-SC4-letter.docx
files: starters/week-04/W4-SC4-members.docx
platform: desktop
other: E4.3
timeLimit: 25
attempts: 2
passing: 70

- task: Start a **Letters** merge and connect the member list.
  check: merge.is type=letters source=true
- task: Replace **[Address Block]** with an Address Block field.
  check: field.present code="ADDRESSBLOCK"
  check: text.absent value="[Address Block]"
- task: Replace **[Greeting Line]** with a Greeting Line field.
  check: field.present code="GREETINGLINE"
  check: text.absent value="[Greeting Line]"
- task: Replace **[Level]** with the **Level** field and **[RenewBy]** with the **RenewBy** field.
  check: field.present code="MERGEFIELD Level"
  check: field.present code="MERGEFIELD RenewBy"
  check: text.absent value="[Level]"
  check: text.absent value="[RenewBy]"
- task: Replace **[Level sentence]** with a rule: if **Level** is **Gold**, "Gold members get a free sunset tour."; otherwise "Upgrade to Gold for a free sunset tour."
  check: field.present code="IF"
  check: text.absent value="[Level sentence]"
- task: Set the document's **Title** property to **Renewal Letter**.
  check: props.is title="Renewal Letter"
:::

::: callout note
Mail merge isn't part of the MO-110 exam (it's on the Expert exam), so this skill check maps to the course outcomes instead. MO-110 objective 1.3.1 (saving in other formats) is practiced in Lesson 5 and assessed with the PDF and web page you hand in with Project 4.
:::
