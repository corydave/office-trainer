---
type: lesson
id: W5-L2
week: 5
order: 2
title: Compare, Combine and Co-author
objectives: [W5.3, W5.4]
clos: [CLO8]
minutes: 40
status: draft
---

::: callout why
Not everyone uses Track Changes. When someone sends back an edited copy with no markup, or three reviewers each send their own copy, Word can work out what changed and put it all in one place. And when a document lives online, several people can work on it at once.
:::

## Compare and Combine

Both are in [[path: Review > Compare]] (desktop Word), and both make a **new document** with the differences shown as tracked changes.

| | **Compare** | **Combine** |
|---|---|---|
| You have | An original and one revised copy (no tracked changes needed) | Two reviewers' copies of the same original (combine again to add a third) |
| You get | One document showing every difference as a tracked change | One document with everyone's tracked changes, labeled by reviewer |
| Use it when | Someone edited a copy without Track Changes | Several people reviewed separately |

::: callout tip
Compare and Combine open the result in a **new window**. Save it right away with a clear name, then open **Office Trainer** in that window (Home tab) to check it. The trainer remembers the steps you've already passed.
:::

::: video
id: w5-l2-compare
url: TODO
title: Compare two versions and combine reviewers' copies
captions: none
duration: 4:00
:::

::: drill D5.04
title: Compare versions and combine reviewers' copies
badge: Version Detective
starter: starters/week-05/W5-L2-letter-original.docx
files: starters/week-05/W5-L2-letter-revised.docx, starters/week-05/W5-L2-review-priya.docx, starters/week-05/W5-L2-review-sam.docx
platform: desktop
other: E1.1
minutes: 10

Alex Morgan asked two people from Alex's old job at Lakeshore Paddle Company, Priya and Sam, to look over this cover letter before sending it. Download the three other copies of the letter (next to the practice file): Alex's own revised copy and the two reviewers' copies. Save all four in one folder.

- step: Compare the **original** with the **revised** letter, and save the result as **Lastname_W5_Compared.docx**. Then check this step in the new window.
  check: revision.count min=3
  check: file.name pattern="*_W5_Compared.docx"
  hint: [[path: Review > Compare > Compare]]. Original document: W5-L2-letter-original. Revised document: W5-L2-letter-revised.
- step: Combine **Priya's** copy and **Sam's** copy, and save the result as **Lastname_W5_Combined.docx**. Check this step in that window.
  check: revision.count author="Priya" min=1
  check: revision.count author="Sam" min=1
  check: file.name pattern="*_W5_Combined.docx"
  hint: [[path: Review > Compare > Combine]]. Pick Priya's copy first, then Sam's. Each reviewer's changes keep their name.
:::

## Share and co-author

Save a document to **OneDrive** (or SharePoint), then select **Share** at the top right:

- **Who can open it:** specific people, or anyone with the link. FLCC accounts may limit links to people at the college.
- **Can edit** or **Can view**. Desktop and web Word also offer **Can review**, which forces Track Changes on.
- **Expiration date** and **password** (in the link settings) limit how long the link works and who can open it. They're only offered for **Anyone with the link** links, and only if FLCC allows those.

When several people open a shared file, everyone's edits appear live, with a colored flag showing where each person is working. **Version history** ([[path: File > Info > Version History]]) lists earlier saved versions, so you can open or restore one.

::: callout privacy
A share link is easy to forward. For anything personal or graded, share with specific people and give view-only access unless they need to edit. If you have to use an Anyone link, set an expiration date.
:::

::: drill D5.05
title: Share, co-edit and restore a version
badge: Collaborator
starter: starters/week-05/W5-L2-letter-original.docx
platform: web
mos: 1.3.4
other: E1.1
minutes: 10

You need a partner (a classmate, or a second account of your own) for step 2.

- step: Save the letter to OneDrive and create a **view-only** link that **expires** in one week. If your account doesn't offer **Anyone with the link** (so there's no expiration date), share it view-only with one specific person instead.
  screenshot: The link settings showing view-only access, and the expiration date or the person you shared with.
  example: TODO
  describe: Say who can use the link, whether they can edit, and when it expires (or who it's shared with).
- step: Share the same file with your partner with **Can edit**, and both edit it at the same time.
  screenshot: The document with your partner's colored cursor flag or name showing where they're editing.
  example: TODO
  describe: Say what your partner changed while you watched.
- step: Open **Version History** and restore the version from before your partner's edits.
  screenshot: The Version History pane with the earlier version selected.
  example: TODO
  describe: Say which version you restored and what it no longer contains.
:::

## Check your understanding

- You sent a letter to a friend, who edited it without Track Changes. How do you see what they changed?
- What's the difference between Can view and Can edit?
- When would you restore an earlier version?
