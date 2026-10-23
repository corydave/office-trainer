---
type: lesson
id: PP5-L1
week: 5
order: 1
title: Comments and Co-authoring
objectives: [W5.1]
clos: [CLO7]
minutes: 35
status: draft
---

::: callout why
Group presentations are built together. Comments keep the discussion attached to the slide it's about, and co-authoring lets everyone work in the same file at once instead of mailing copies around and losing edits.
:::

## Comments

Comments are threads: a note, replies to it, and a resolved state when it's handled.

| To | Do this |
|---|---|
| Add a comment | Select the slide (or an object on it), then [[path: Review > New Comment]] ([[path: Insert > Comment]] also works). Type, then **Post** ([[key: Ctrl+Enter]], or [[key: Cmd+Enter]] on a Mac) |
| See all comments | [[path: Review > Show Comments]] opens the Comments pane |
| Reply | Type in the **Reply** box under the comment |
| Resolve | Select **...** (More thread actions) > **Resolve thread**. It stays in the file, marked resolved, and you can reopen it |
| Delete | **...** > **Delete thread**, or [[path: Review > Delete]] |

In a shared file on a work or school account, type **@** and a name to mention someone; they get an email with the comment and a link to it.

::: drill D5.01
title: Work through the comments
badge: Commenter
starter: starters/week-05/P5-L1-review.pptx
platform: web
mos: 1.5.4
minutes: 7

- step: Reply to **Sam Lee**'s comment on **Sunset tours** with the price: **Yes: $35 a person.**
  check: comment.is says="sunset tour price" state=replied
  hint: [[path: Review > Show Comments]], then type in the Reply box under Sam's comment.
- step: Resolve Sam's thread.
  check: comment.is says="sunset tour price" state=resolved
  hint: Select **...** at the top of the thread > **Resolve thread**.
- step: Delete **Ana Ortiz**'s old comment on **Hours**.
  check: comment.is says="last year's hours" state=absent
- step: On **Staff picnic**, add a new comment: **Who brings the grill?**
  check: comment.is title="Staff picnic" says="grill"
:::

## Co-authoring and version history

Save the presentation to **OneDrive** or **SharePoint**, then select **Share** at the top right and invite people with **Can edit**. Everyone who opens it works in the same file at the same time:

- Each person's picture or initials appear at the top, and a flag on a slide thumbnail shows who's editing it.
- Changes save and appear for everyone within a few seconds. (In desktop PowerPoint, keep **AutoSave** on.)
- Two people shouldn't edit the same text box at once; split the slides between you.

**Version history** keeps earlier copies: [[path: File > Info > Version History]] (on a Mac, [[path: File > Browse Version History]]). Open a version to look at it, then select **Restore** to make it the current one (the version you replace is kept too, so nothing is lost).

::: callout privacy
Share with **Can view** unless people need to edit, and share with named people rather than "Anyone with the link" when the deck has anything internal in it.
:::

::: drill D5.02
title: Co-edit with a partner
badge: Co-author
starter: starters/week-05/P5-L1-review.pptx
platform: web
minutes: 12

Work with a classmate. If you're on your own, open the shared file in a second browser window signed in with a different Microsoft account (or as a guest through a **Can edit** link) to see how it works.

- step: Save the presentation to OneDrive, share it with your partner with **Can edit**, and both open it. Take a screenshot that shows both of you in the file.
  screenshot: The presentation open in PowerPoint with two people's pictures or initials at the top.
  example: TODO
  describe: Who is in the file with you, and how can you tell which slide they're on?
- step: Have your partner add a slide titled **Partner's ideas** with one idea on it, while you have the file open.
  check: slide.is title="Partner's ideas"
- step: Wait a few minutes after your partner adds the slide (versions are saved every so often, not on every edit). Then delete the **Rain policy** slide "by mistake". Then open **Version History**, find the version from just before you deleted it, and **Restore** it.
  check: slide.is title="Rain policy"
  check: slide.is title="Partner's ideas"
  hint: [[path: File > Info > Version History]] (Mac: [[path: File > Browse Version History]]). Pick a version from the list on the right, look at it, then select **Restore**.
:::

::: quiz Q51.1
question: A reviewer's comment has been dealt with, but you want a record that it was. What should you do?
type: single
- [ ] Delete the comment
- [x] Resolve the thread
- [ ] Reply "done" and leave it open
- [ ] Hide the slide
explain: A resolved thread stays in the file, marked as handled, so the team can see what was decided. Delete only comments nobody needs.
:::

## Check your understanding

- What's the difference between resolving and deleting a comment?
- How do you see who else is editing a shared presentation?
- What happens to the current version when you restore an older one?
