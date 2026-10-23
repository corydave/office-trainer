---
type: lesson
id: W5-L1
week: 5
order: 1
title: Comments and Track Changes
objectives: [W5.1, W5.2]
clos: [CLO8]
minutes: 45
status: draft
---

::: callout why
Instructors grade with comments. Group projects and workplace documents go through rounds of edits by several people. Comments and Track Changes keep every suggestion visible, so nothing changes without someone deciding it should.
:::

## Your scenario

Lakeshore Paddle Company is updating its **Rental Safety Policy**. The owner, Priya Shah, and the lead guide, Sam Ortiz, have reviewed the draft. Your job is to work through their comments and edits, make your own, and then get a clean, safe copy ready to share.

## Comments

| To | Do this |
|---|---|
| Add a comment | Select the text, then [[path: Review > New Comment]] ([[key: Ctrl+Alt+M / Cmd+Option+A]]) |
| Reply | Select **Reply** under the comment, type, and post it |
| Mention someone | Type @ and their name in a shared document; they get an email |
| Resolve | **Resolve** (or the check mark) marks the thread done but keeps it |
| Delete | [[path: Review > Delete]], or **Delete thread** in the comment's menu |

::: callout tip
Resolve when the conversation is finished but the record is useful. Delete when the comment was never meant to travel with the file, like a note to yourself.
:::

::: video
id: w5-l1-comments
url: TODO
title: Reply, resolve and delete comments
captions: none
duration: 3:00
:::

::: drill D5.01
title: Work through the comments
badge: Comment Keeper
starter: starters/week-05/W5-L1-policy.docx
platform: web
mos: 6.1.1, 6.1.2, 6.1.3, 6.1.4
minutes: 7

- step: Add a comment on the **Paddle Pass** paragraph (for example: "How far ahead can non-members book?").
  check: comment.is anchor="Paddle Pass holders"
  hint: Select some of the paragraph's text, then [[path: Review > New Comment]].
- step: Reply to Sam's comment about **48 hours**.
  check: comment.is says="Is 48 hours right" state=replied
- step: Priya's comment on the life jacket rule has been dealt with. **Resolve** it.
  check: comment.is says="must instead of should" state=resolved
- step: Priya's **Old note** comment shouldn't be shared. **Delete** it.
  check: comment.is says="Old note" state=absent
:::

## Track Changes

With **Track Changes** on ([[path: Review > Track Changes]], [[key: Ctrl+Shift+E / Cmd+Shift+E]]), every insertion, deletion and formatting change is marked with who made it and when, until someone accepts or rejects it.

- **Display for Review** sets how much markup you see: **Simple Markup**, **All Markup**, **No Markup** or **Original**. Changing the view never changes the document.
- **Lock Tracking** (desktop Word: [[path: Review > Track Changes > Lock Tracking]]) adds a password, so no one can turn tracking off without it.

::: drill D5.02
title: Track your own edits and lock tracking
badge: Change Tracker
starter: starters/week-05/W5-L1-draft.docx
platform: desktop
mos: 6.2.1, 6.2.4
minutes: 7

- step: Turn on **Track Changes**.
  check: doc.is tracking=true
- step: Make at least **four** edits to the checklist: add a word, delete a word, change a word, and bold something.
  check: revision.count min=4
  hint: Every edit shows in color with a line in the margin. If it doesn't, Track Changes isn't on.
- step: Lock tracking with a password you'll remember (for practice, try **paddle**).
  check: doc.is protection=trackedChanges password=true
  hint: [[path: Review > Track Changes > Lock Tracking]], type the password twice, then **OK**.
:::

## Accept and reject

Review changes one at a time with **Next**, then **Accept** or **Reject** each one ([[path: Review > Changes]]). The Accept and Reject menus also have options for all changes at once, or all changes shown, which you can narrow with **Show Markup** (for example, to formatting only).

::: callout warning
**Accept All Changes** is fast, and it's how mistakes get approved without anyone reading them. Read each change first, or at least filter the markup to the kind you're accepting.
:::

::: drill D5.03
title: Review Sam's changes
badge: Change Reviewer
starter: starters/week-05/W5-L1-policy.docx
platform: web
mos: 6.2.2, 6.2.3
minutes: 7

- step: Accept both of Sam's **formatting** changes (the bold text).
  check: revision.count kind=format max=0
  hint: Use [[path: Review > Next]] to step through. Or, in desktop Word, set **Show Markup** to Formatting only, then **Accept All Changes Shown**.
- step: In the Refunds paragraph, Sam replaced **fully refundable** with **refunded minus a $5 fee**. The original wording is correct: **reject** both his deletion and his insertion.
  check: text.present value="fully refundable"
  check: text.absent value="minus a $5 fee"
  hint: Click in the deleted words and select **Reject**, then do the same for the inserted words.
- step: Accept every other change, so no tracked changes are left.
  check: revision.count max=0
:::

::: quiz Q51.1
question: You switch Display for Review to No Markup. What happens to the tracked changes?
type: single
- [ ] They're all accepted
- [ ] They're all rejected
- [x] Nothing: they're hidden from view but still waiting for a decision
- [ ] They're deleted
explain: Display for Review only changes what you see. The changes stay in the file until someone accepts or rejects them, and they'll show up again for the next person who opens it.
:::

## Check your understanding

- When do you resolve a comment instead of deleting it?
- What does Lock Tracking stop someone from doing?
- Why is Accept All Changes risky?
