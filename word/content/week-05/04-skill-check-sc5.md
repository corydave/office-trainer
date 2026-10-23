---
type: skillcheck
id: W5-SC5
week: 5
order: 4
title: "Skill Check 5: The Review Cycle"
objectives: [W5.1, W5.2, W5.5]
clos: [CLO1, CLO8]
minutes: 25
status: draft
---

Like a MOS exam task: one starter file, a set of tasks and a time limit, with no hints. Each task is worth one point, and you can do them in any order, except that the last task locks the document, so do it last.

**Before you start:** open the starter file, then select **Start**. Tasks marked for desktop Word aren't counted in Word for the web.

::: skillcheck SC5
title: Season report review
badge: Skill Check 5
starter: starters/week-05/W5-SC5-report.docx
platform: web
mos: 1.4.1, 1.4.2, 6.1.1, 6.1.2, 6.1.3, 6.1.4, 6.2.1, 6.2.2, 6.2.3, 6.2.4
timeLimit: 25
attempts: 2
passing: 70

- task: Reply to Sam's comment about the **rental numbers**.
  check: comment.is says="rental numbers" state=replied
- task: Resolve Priya's comment **Love this section**.
  check: comment.is says="Love this section" state=resolved
- task: Delete the comment that starts **Remove before sharing**.
  check: comment.is says="Remove before sharing" state=absent
- task: Add a comment on the **Next season** heading.
  check: comment.is anchor="Next season"
- task: Accept all **formatting** changes.
  check: revision.count kind=format max=0
- task: Sam replaced **thanks to our volunteers** with **helped by**. Reject both parts of that change.
  check: text.present value="thanks to our volunteers"
  check: text.absent value="helped by"
- task: Accept all the other changes, so none are left.
  check: revision.count max=0
- task: Remove the hidden text.
  check: text.hidden max=0
- task: Give the picture alt text.
  check: access.is rules="alt"
- task: Turn on Track Changes and lock tracking with a password. (Do this last.)
  check: doc.is tracking=true protection=trackedChanges password=true
  platform: desktop
:::
