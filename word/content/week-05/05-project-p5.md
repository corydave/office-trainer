---
type: project
id: W5-P5
week: 5
order: 5
title: "Project 5: Peer Review Exchange"
objectives: [W5.1, W5.2, W5.3, W5.4, W5.5]
clos: [CLO1, CLO8]
minutes: 150
status: draft
---

::: callout why
Giving useful feedback, and taking it well, is one of the most valued skills at work. This project puts the whole review cycle into practice on real work: review, respond, decide, then ship a clean file.
:::

## What you'll do

1. **Swap** your Project 4 cover letter or your Project 3 newsletter with a partner. (No partner? Use the Week 5 practice files instead.)
2. **Review** your partner's document with **Track Changes** on: at least three tracked edits and three comments that are specific and kind ("This sentence is long; try splitting it after 'staff'" rather than "Fix this").
3. **Respond** to the review of your own document: reply to or resolve every comment, and accept or reject every change on purpose.
4. **Compare and Combine:** get a second reviewer's copy (a classmate, a friend, or your own second pass), and combine the two reviews into one document.
5. **Ship it clean,** in this order: Inspect Document, then set the Title property again (inspecting clears it), then the Accessibility Checker with no errors, current file format, a PDF, and a **view-only** share link with an expiration date (if your account can't set one, share view-only with your instructor by name instead).
6. Write a **short reflection memo** (half a page): which feedback changed your document most, and one thing you'll do differently next time you review someone's work.

Submit: the combined review document (.docx), your final clean document (.docx and .pdf), the share link, and the memo.

::: callout privacy
Peer review shares real work. You choose which document to share, and you may use the practice files instead. Before you send anything, run Inspect Document: it shows you what hidden information travels with a file.
:::

::: callout ai
Optional: Copilot and similar tools can summarize comments or suggest edits. Weigh their suggestions the way you would a person's, and keep Track Changes on so the history stays honest about who changed what.
:::

## Check before you submit

Run this on your **final, clean document**.

::: preflight P5-check
title: Clean-and-ship check
starter: none
platform: web

- step: No comments or tracked changes left.
  check: comment.count max=0
  check: revision.count max=0
- step: No hidden text.
  check: text.hidden max=0
- step: No accessibility errors.
  check: access.is
- step: Personal information removed.
  check: doc.is personalInfo=removed
  platform: desktop
:::

## How it's graded

::: rubric P5
- P5-a | 25 | CLO8 | Feedback to your partner is specific and constructive, using comments and tracked edits.
- P5-b | 25 | CLO8 | Every comment you received gets a reply or resolution, and every change is accepted or rejected on purpose.
- P5-c | 20 | CLO8 | Compare and Combine are done correctly.
- P5-d | 30 | CLO8, CLO1 | The clean-and-ship checklist is complete: Document Inspector, Accessibility Checker, compatibility and permissions.
:::
