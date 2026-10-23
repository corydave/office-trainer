---
type: lesson
id: W4-L3
week: 4
order: 3
title: Mail Merge Letters
objectives: [W4.3]
clos: [CLO7]
minutes: 55
status: draft
---

::: callout why
A cover letter should name the company and the job. Typing that into five copies by hand invites mistakes, like sending one company a letter with another company's name. **Mail merge** writes the letter once and fills in each person's details for you.
:::

## How a merge works

A merge has two parts:

1. **The recipient list** (the data source): a table with one row per person and one column per detail. It can be an Excel sheet, an Outlook contact list, or a table in a Word document.
2. **The main document:** the letter, with **merge fields** where each person's details go, such as «Company» or «Position».

Then you **preview** the results and **finish** the merge: one letter per row, in a new document, or straight to the printer or email.

::: platform web
Mail merge needs desktop Word. Campus labs have it, and your FLCC Microsoft 365 account includes desktop Word for your own computer.
:::

::: video
id: w4-l3-merge
url: TODO
title: A mail merge from start to finish
captions: none
duration: 5:00
:::

::: drill D4.04
title: Finish the recipient list
badge: List Builder
starter: starters/week-04/W4-employers.docx
platform: desktop
other: E4.3
minutes: 6

The recipient list is a Word table with three employers. Add two more.

- step: Add two rows for two more employers (real postings, or made up). Fill in every column; use **Yes** or **No** under **Remote**.
  check: table.is contains="Company" rows=6 cols=9
  hint: Click in the last cell and press [[key: Tab]] to add a row.
- step: Save the file in your course folder as **Lastname_Employers.docx**. You'll pick it as the recipient list in the next drill.
  check: file.name pattern="*_Employers.docx"
:::

## Merge fields

On the **Mailings** tab, in order:

1. [[path: Mailings > Start Mail Merge > Letters]]
2. [[path: Mailings > Select Recipients > Use an Existing List]], then pick your list file.
3. Insert fields where they belong:
   - **Address Block** ([[path: Mailings > Address Block]]) writes the whole inside address.
   - **Greeting Line** ([[path: Mailings > Greeting Line]]) writes "Dear Ms. Lee". Choose the colon (**:**) as its punctuation: business letters use a colon, not a comma.
   - **Insert Merge Field** adds any single column, such as Position.

Merge fields look like «Position» in the letter. Select **Preview Results** to see real names in their place.

::: drill D4.05
title: Insert the merge fields
badge: Field Placer
starter: starters/week-04/W4-L3-letter.docx
platform: desktop
other: E4.3
minutes: 8

Do D4.04 first, so you have your recipient list.

- step: Start a **Letters** merge and select your recipient list.
  check: merge.is type=letters source=true
  hint: [[path: Mailings > Start Mail Merge > Letters]], then [[path: Mailings > Select Recipients > Use an Existing List]].
- step: Replace **[Address Block]** with an Address Block field, and **[Greeting Line]** with a Greeting Line field.
  check: field.present code="ADDRESSBLOCK"
  check: field.present code="GREETINGLINE"
  check: text.absent value="[Address Block]"
  hint: Select the bracketed text first, then [[path: Mailings > Address Block]] and **OK**. Do the same with Greeting Line.
- step: Replace **[Position]** with the **Position** merge field.
  check: field.present code="MERGEFIELD Position"
  check: text.absent value="[Position]"
  hint: Select [Position], then [[path: Mailings > Insert Merge Field > Position]].
:::

## Rules: If...Then...Else

A **rule** changes the text depending on a person's details. [[path: Mailings > Rules > If...Then...Else]] asks for a field, a comparison and two versions of the text:

> If **Remote** is equal to **Yes**, then insert "I'm experienced with remote work tools and keep a reliable home office." Otherwise insert "I'm glad to work on site, and I live within a short drive."

Then **preview every record** ([[path: Mailings > Preview Results]] and the arrows) to look for mistakes: blank lines from empty columns, a missing space, the wrong title.

::: drill D4.06
title: Add a rule and preview every letter
badge: Rule Writer
starter: starters/week-04/W4-L3-letter.docx
platform: desktop
other: E4.3, E4.1
minutes: 8

- step: Replace **[Remote sentence]** with an **If...Then...Else** rule on the **Remote** field: one sentence for Yes, another for No.
  check: field.present code="IF"
  check: text.absent value="[Remote sentence]"
  hint: Select [Remote sentence], then [[path: Mailings > Rules > If...Then...Else]].
- step: Preview every record and fix anything that looks wrong.
  screenshot: The letter in Preview Results, showing one of the employers you added, with the remote sentence filled in.
  example: TODO
  describe: Say which employer you're previewing and which version of the remote sentence it shows.
:::

## Finish the merge

[[path: Mailings > Finish & Merge > Edit Individual Documents]] creates a **new document** with one letter per row, each on its own page. Save it with a new name; it's an ordinary document now, so you can fix one letter without touching the others.

::: callout tip
The merged letters open in a new window. To check them, open **Office Trainer** in that window (Home tab) and go back to this lesson. The trainer remembers the steps you've already passed.
:::

::: drill D4.07
title: Merge to a new document
badge: Merge Master
starter: starters/week-04/W4-L3-letter.docx
platform: desktop
other: E4.3
minutes: 5

- step: Finish the merge with **Edit Individual Documents > All**, and save the result as **Lastname_W4_Merged.docx**.
  check: file.name pattern="*_W4_Merged.docx"
  hint: [[path: Mailings > Finish & Merge > Edit Individual Documents]], choose **All**, then **OK**.
- step: In the merged document, check that each employer has its own letter.
  check: field.present code="MERGEFIELD" max=0
  check: break.count kind=section min=4
  check: text.present value="Finger Lakes Outfitters"
  check: text.present value="Seneca Shore Inn"
:::

::: quiz Q43.1
question: You want one sentence to change depending on whether a job is remote. Which tool do you use?
type: single
- [ ] Address Block
- [ ] Greeting Line
- [x] Rules > If...Then...Else
- [ ] Insert Merge Field
explain: A rule picks between two versions of text based on a field's value. Merge fields only drop in the value itself.
:::

## Check your understanding

- What are the two parts of every mail merge?
- Why preview every record before finishing?
- What's in the document that Edit Individual Documents creates?
