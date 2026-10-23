---
type: skillcheck
id: XL5-SC5
week: 5
order: 5
title: "Skill Check 5: Review, Protect and Share"
objectives: [W5.1, W5.3]
clos: [CLO1, CLO7]
minutes: 30
status: draft
---

Like a MOS exam task: one starter file, a set of tasks and a time limit, with no hints. Each task is worth one point. Do the tasks **in order**: the last few protect the workbook, which stops some of the earlier tasks from working.

::: skillcheck SC5
title: Volunteer roster
badge: Skill Check 5
starter: starters/week-05/E5-SC5-roster.xlsx
platform: desktop
mos: 1.4.5, 1.5.4, 1.5.5, 5.3.3
timeLimit: 30
attempts: 2
passing: 70

- task: Reply to Marcus's comment about **shirt sizes**.
  check: comment.is kind=comment says="shirt sizes" state=replied
- task: Resolve Marcus's thread.
  check: comment.is kind=comment says="shirt sizes" state=resolved
- task: Add a note to D9 that says **Jules can only stay until noon.**
  check: comment.is kind=note cell=D9 says="until noon"
- task: Delete the old note in B5.
  check: comment.is kind=note cell=B5 state=absent
- task: Rename the table on **Roster** to **Volunteers**.
  check: table.is name=Volunteers
- task: Rename **Sheet3** to **Lists**.
  check: sheet.is name=Lists
- task: Give the chart on **Hours** alt text that states its point.
  check: access.is rules=alt
- task: Remove the hidden worksheet. (If you use Inspect Document, don't remove the comments.)
  check: doc.is hiddenSheets=0
- task: Remove document properties and personal information from the workbook.
  check: doc.is personalInfo=removed
- task: Set the **Title** property to **Volunteer roster**.
  check: props.is title="Volunteer roster"
- task: Make C2:C9 (Shirt) a drop-down list of **S, M, L, XL**.
  check: validation.is sheet=Roster range=C2:C9 type=list list="S,M,L,XL"
- task: Unlock C2:D9 on **Roster** so they can still be changed after you protect the sheet.
  check: protect.is sheet=Roster unlocked=C2:D9
- task: Protect the **Roster** sheet (no password).
  check: protect.is sheet=Roster sheetProtected=true
- task: Protect the workbook's structure. (Do this last.)
  check: protect.is structure=true
:::
