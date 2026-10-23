---
type: project
id: XL5-P5
week: 5
order: 6
title: "Project 5: Team Event Workbook"
objectives: [W5.1, W5.2, W5.3, W5.4]
clos: [CLO6, CLO7]
minutes: 180
status: draft
rubricBadge: Team Collaborator
---

::: callout why
At work, almost every workbook is a team effort, and almost every one ends up in front of people outside the team. This project is that whole cycle: plan together, combine everyone's work, then check, protect and publish the result.
:::

## What you'll make

With one or two classmates, plan a small event (a club fundraiser, a study session, a team outing). You'll hand in **three files** and a short evidence document:

1. **The shared workbook** (a OneDrive link with **Can view** access for your instructor), where your team:
   - co-authored at least one sheet together, so Show Changes lists edits from each of you
   - left at least **three comment threads**, with replies and @mentions, and resolved the ones that are settled
2. **`Lastname_P5_Final.xlsx`**, your clean copy, with:
   - a **Combined** sheet that links to (external references) or **Consolidates** each teammate's part of the budget
   - a **sign-up** or **task** sheet with unlocked entry cells, data validation on at least one column, and sheet protection
   - workbook **structure** protection
   - no comments, notes or hidden sheets, personal information removed, and no accessibility errors (sheet and table names, alt text, a Title). Removing personal information clears the Title, so set it **after** you inspect
3. **A web page (.htm) and a PDF** of the main sheet, saved from the final copy.

Your **evidence document** has screenshots of the Show Changes pane and Version History in the shared workbook, and two or three sentences on how your team split the work.

::: callout privacy
Use first names only, or made-up names, in sign-up sheets. Share the workbook with your team and your instructor only, and use **Can view** for anyone who doesn't need to edit.
:::

::: callout ai
Optional: ask an AI tool for a checklist of things to plan for your kind of event. Decide together which items belong in your budget, and write one comment in the workbook about an item you added or dropped and why.
:::

## Check before you submit

Check your **final copy** (Lastname_P5_Final.xlsx):

::: preflight P5-check
title: Preflight for your final copy
starter: none
platform: desktop

- step: No comments or notes are left in the final copy.
  check: comment.count kind=any max=0
- step: No hidden sheets are left.
  check: doc.is hiddenSheets=0
- step: Personal information is removed.
  check: doc.is personalInfo=removed
- step: There are no accessibility errors: alt text, real sheet and table names, and a Title.
  check: access.is
- step: The workbook's structure is protected.
  check: protect.is structure=true
:::

## How it's graded

::: rubric P5
- P5-a | 25 | CLO7 | The shared workbook shows real co-authoring: edits from each teammate in Show Changes, comment threads with replies and @mentions, resolved where settled.
- P5-b | 20 | CLO7 | The Combined sheet correctly links to or consolidates each teammate's work, and the totals are right.
- P5-c | 20 | CLO1, CLO7 | The sign-up sheet works for others: entry cells unlocked, useful data validation, sheet and structure protected.
- P5-d | 20 | CLO1 | The final copy is inspected and accessible: no comments, hidden sheets or personal information, and no accessibility errors.
- P5-e | 15 | CLO6 | The web page and PDF look right in a browser and a PDF reader, with a sensible title.
:::
