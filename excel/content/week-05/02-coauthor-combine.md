---
type: lesson
id: XL5-L2
week: 5
order: 2
title: Co-author, Compare and Combine
objectives: [W5.2]
clos: [CLO7]
minutes: 50
status: draft
---

::: callout why
Emailing copies back and forth ("budget-final-v3-Priya.xlsx") is how numbers get lost. Today, people work in **one** shared workbook at the same time, and Excel keeps track of who changed what. When you do get separate copies back, Excel can link or add them together for you.
:::

## Co-authoring: one workbook, many people

When a workbook is saved in **OneDrive** or **SharePoint**, everyone you share it with can edit it at the same time, in Excel for the web or the desktop app:

1. Save the workbook to OneDrive (your FLCC Microsoft 365 account has it). **AutoSave** turns on in the top-left corner.
2. Select **Share** (top right), type your partner's email, choose **Can edit**, and send.
3. When your partner opens it, you'll see their picture or initials at the top, and a colored outline on the cell they've selected.

::: callout tip
Changes save as you type, for everyone. If you want to try something without others seeing it yet, use [[path: File > Save a Copy]] to make your own copy first.
:::

## Who changed what?

- **Show Changes** ([[path: Review > Show Changes]]) opens a pane listing recent edits: who changed which cell, when, and what the old value was. Select cells first to see only their changes.
- **Version history** keeps earlier saved versions of the whole workbook. Select the file name at the top of the window > **Version History** (or [[path: File > Info > Version History]] in the desktop app). Open an old version to look at it, then **Restore** it if you need to go back.

::: video
id: e5-l2-coauthor
url: TODO
title: Co-authoring, Show Changes and version history
captions: none
duration: 4:00
:::

::: drill D5.02
title: Edit the schedule with a partner
badge: Co-Author
starter: starters/week-05/E5-L2-schedule.xlsx
platform: web
other: E1.1
minutes: 12

Work with a classmate. (No partner? Open the shared workbook in a second browser, or on your phone, and be your own partner.)

- step: Save the schedule to OneDrive and share it with your partner with **Can edit**. Each of you changes one cell in the **Who** column.
  screenshot: The schedule open with your partner's initials or picture at the top, or their colored outline on a cell.
  example: TODO
  describe: Say who your partner was and which cell each of you changed.
- step: Open **Show Changes** and find your partner's edit.
  screenshot: The Show Changes pane listing your partner's change.
  example: TODO
  describe: Say what the pane listed for your partner's edit: the cell, the old value and the new value.
  hint: [[path: Review > Show Changes]].
- step: Open **Version History** and restore the version from before your partner's edit.
  screenshot: The Version History pane with the earlier version open, just before you select Restore.
  example: TODO
  describe: Say what time the version you restored was saved and what it changed back.
- step: Add a comment on C7 that **@mentions** your partner and asks them a question. (Do this after the restore: restoring an older version would remove a newer comment.)
  check: comment.count kind=comment min=1
:::

## Combining copies people sent back

Sometimes you get separate workbooks back anyway: a reviewer worked offline, or each department filled in its own copy. Two ways to bring them together:

### Link to a cell in another workbook

An **external reference** points at a cell in another file. Open both workbooks, type `=` in your cell, switch to the other workbook, click the cell, and press [[key: Enter]]. Excel writes something like `='[Review-Priya.xlsx]Budget'!$B$9`.

When the other workbook changes, your link can update. [[path: Data > Workbook Links]] (or **Edit Links** in older versions) lists every link, and lets you update, change or break them.

::: callout note
If you move or rename the other workbook, the link breaks. Keep linked workbooks in the same folder, and break links (turn them into plain numbers) before you send the file to someone who can't open the other one.
:::

### Consolidate

**Consolidate** adds up (or averages, counts...) matching ranges from several sheets or workbooks into one. With **labels** turned on, it matches rows by name, so it works even when the lists aren't in the same order or one has an extra row.

1. Click where the combined list should start.
2. [[path: Data > Consolidate]] (desktop). Choose the **Function** (Sum).
3. In **Reference**, select the first range, including its labels, and choose **Add**. Repeat for each range.
4. Under **Use labels in**, tick **Top row** and **Left column**, then **OK**.

::: video
id: e5-l2-combine
url: TODO
title: External references and Consolidate
captions: none
duration: 4:00
:::

### The legacy way: Compare and Merge

Older versions of Excel used **Shared Workbooks**, with a **Compare and Merge Workbooks** command that merged copies of one shared workbook. Microsoft has replaced both with co-authoring: they're hidden by default, only in Excel for Windows, and they don't work with tables, many newer features, or files in OneDrive. You may see them in older workplaces or exam guides, so you'll find the command once, then use the modern tools above.

::: drill D5.03
title: Combine the three budget reviews
badge: Combiner
starter: starters/week-05/E5-L2-combine.xlsx
files: starters/week-05/Review-Priya.xlsx
platform: desktop
other: E1.1, E3.4
minutes: 12

Download **Review-Priya.xlsx** (it's next to the practice file above) and save it in the **same folder** as the practice file.

- step: Open **Review-Priya.xlsx**. On **Summary**, in B3, link to Priya's total (her **Budget** sheet, B9).
  check: formula.is sheet=Summary cell=B3 external="Review-Priya"
  check: formula.is sheet=Summary cell=B3 result=1770
  hint: In B3 type `=`, switch to Review-Priya.xlsx ([[path: View > Switch Windows]]), click B9, and press [[key: Enter]].
- step: The **Priya**, **Marcus** and **Lena** sheets hold copies of each reviewer's estimate. On **Combined**, starting in A3, **Consolidate** all three with **Sum**, using the top row and left column as labels.
  check: consolidate.is sheet=Combined function=sum sources=3 labels=true
  check: text.present sheet=Combined value="Parking"
  hint: Click A3, [[path: Data > Consolidate]]. Add Priya!$A$1:$B$6, Marcus!$A$1:$B$6 and Lena!$A$1:$B$7, tick **Top row** and **Left column**, **OK**.
- step: (Windows) Add **Compare and Merge Workbooks** to the Quick Access Toolbar, then remove it again.
  platform: desktop
  screenshot: The Quick Access Toolbar with the Compare and Merge Workbooks button on it (it will be grayed out).
  example: TODO
  describe: Say where you found the command and why it's grayed out for this workbook. On a Mac, say that the command isn't available there.
  hint: [[path: File > Options > Quick Access Toolbar]], choose **All Commands**, find **Compare and Merge Workbooks**, **Add**, **OK**.
:::

::: quiz Q52.1
question: Three people need to update the same budget this week. What's the best way to work?
type: single
- [ ] Email each person a copy, then Compare and Merge them
- [x] Share one workbook in OneDrive and co-author it, using comments and Show Changes
- [ ] Take turns: each person edits, saves and emails it to the next
- [ ] Each person makes a copy, and you retype their numbers
explain: Co-authoring keeps one up-to-date version, saves every change, and shows who changed what. Separate copies are for when someone can't reach the shared file.
:::

## Check your understanding

- What's the difference between Show Changes and version history?
- When does a link to another workbook break, and how do you fix or remove it?
- Why does Consolidate need labels when the lists aren't in the same order?
