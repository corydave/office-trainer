---
type: lesson
id: XL5-L1
week: 5
order: 1
title: Comments and Notes
objectives: [W5.1]
clos: [CLO7]
minutes: 30
status: draft
---

::: callout why
Most workbooks at work have more than one author. Comments let people ask and answer questions right on the cell they're about, so nobody has to dig through email to find out why a number changed.
:::

## Your scenario

Paddle Day is coming. The budget, the schedule and the sign-up sheet are being reviewed by Priya, Marcus and Lena, so this week you'll handle their comments, combine their copies, check and protect the workbooks, and share the schedule on the web.

## Comments or notes?

Excel has two kinds of cell annotations:

| | **Comments** | **Notes** |
|---|---|---|
| For | A conversation: a question, a reply, a decision | A label or reminder that doesn't need an answer |
| Shows as | A purple corner; the comment opens in a card | A red corner; the note pops up when you point at the cell |
| Replies | Yes, and you can **@mention** someone to notify them | No |
| Insert with | [[path: Review > New Comment]] | [[path: Review > Notes > New Note]] |

When a question is answered, **resolve** the thread (the **...** menu on the comment > **Resolve thread**). It stays in the workbook, grayed out, and you can reopen it. **Delete** a comment or note when nobody needs it any more (right-click the cell > **Delete Comment** or **Delete Note**). [[path: Review > Show Comments]] opens the Comments pane with every thread in the workbook.

::: callout privacy
Comments and notes travel with the file, with the author's name on each one. Before you send a workbook outside your team, read them all, and delete the ones you don't want others to see (Week 5, Lesson 3 shows how to find them all at once).
:::

::: video
id: e5-l1-comments
url: TODO
title: Comments, replies, @mentions and notes
captions: none
duration: 3:30
:::

::: drill D5.01
title: Answer the budget reviewers
badge: Comment Handler
starter: starters/week-05/E5-L1-review.xlsx
platform: web
mos: 1.5.5
minutes: 8

- step: Add a **note** to B6 (the Signs estimate) that says **Includes two banners.**
  check: comment.is kind=note cell=B6 says="banners"
  hint: Select B6, then [[path: Review > Notes > New Note]].
- step: Add a **comment** to C5 (the Food actual), for example **The caterer's receipt is in the shared folder.**
  check: comment.is kind=comment cell=C5
  hint: Select C5, then [[path: Review > New Comment]], type, and select **Post** (or press [[key: Ctrl+Enter]]).
- step: Reply to Priya's comment on B4: the dock fee is included ($40 of it).
  check: comment.is kind=comment says="dock fee" state=replied
- step: Resolve Priya's thread.
  check: comment.is kind=comment says="dock fee" state=resolved
  hint: Open the comment, then the **...** menu > **Resolve thread**.
- step: Delete the old note in D9.
  check: comment.is kind=note cell=D9 state=absent
  hint: Right-click D9 > **Delete Note**.
:::

::: quiz Q51.1
question: You want Marcus to check a number and tell you if it's right. What should you add to the cell?
type: single
- [x] A comment that @mentions Marcus
- [ ] A note
- [ ] A text box beside the cell
- [ ] Bold red text in the cell
explain: A comment starts a conversation Marcus can reply to, and @mentioning him sends him a notification. A note is a one-way label.
:::

## Check your understanding

- When would you resolve a comment instead of deleting it?
- What does an @mention do that a plain comment doesn't?
- Why check the comments before you send a workbook to a customer?
