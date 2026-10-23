---
type: lesson
id: W4-L4
week: 4
order: 4
title: Envelopes, Labels and Directories
objectives: [W4.4]
clos: [CLO7]
minutes: 35
status: draft
---

::: callout why
The same recipient list that made your letters can address their envelopes, print a sheet of mailing labels, or make a one-page contact list. Set the list up once, and every mailing after that takes minutes.
:::

## Labels and envelopes

Start each one in a **new, blank document**:

| Output | Start with | Then |
|---|---|---|
| **Envelopes** | [[path: Mailings > Start Mail Merge > Envelopes]] and pick a size (#10 is the standard US business envelope) | Select recipients, insert an Address Block in the address area |
| **Labels** | [[path: Mailings > Start Mail Merge > Labels]] and pick a product (**Avery US Letter 5160** is the most common address label) | Select recipients, insert an Address Block in the first label, then **Update Labels** to copy it to every label |

**Update Labels** adds a «Next Record» field to each label, so each one gets the next person on the list.

::: video
id: w4-l4-labels
url: TODO
title: Labels and envelopes from a recipient list
captions: none
duration: 4:00
:::

::: drill D4.08
title: A sheet of mailing labels
badge: Label Maker
starter: starters/week-04/W4-L4-blank.docx
platform: desktop
other: E4.3
minutes: 8

Use your recipient list from Lesson 3.

- step: In a blank document, start a **Labels** merge on **Avery US Letter 5160** labels, and select your recipient list.
  check: merge.is type=labels source=true
  hint: [[path: Mailings > Start Mail Merge > Labels]], choose **Avery US Letter** and **5160**, then [[path: Mailings > Select Recipients > Use an Existing List]].
- step: Insert an **Address Block** in the first label, then select **Update Labels**.
  check: field.present code="ADDRESSBLOCK" min=2
  check: field.present code="NEXT" min=1
  hint: Click in the first label, [[path: Mailings > Address Block]], then [[path: Mailings > Update Labels]].
- step: In another new document, set up an **envelope** merge with an Address Block, and preview one envelope.
  screenshot: The envelope in Preview Results, with one employer's address on it.
  example: TODO
  describe: Say which employer's envelope you're previewing.
:::

## Directories

A **directory** merge puts every record on the same page, one after another: a contact list, a price list, a phone tree. Start it with [[path: Mailings > Start Mail Merge > Directory]], type one line of fields (with tabs between them), then finish the merge.

::: drill D4.09
title: A contact directory
badge: Directory Builder
starter: starters/week-04/W4-L4-blank.docx
platform: desktop
other: E4.3
minutes: 6

- step: In a blank document, start a **Directory** merge and select your recipient list.
  check: merge.is type=directory source=true
- step: Type one line with the **Company**, **Contact** and **Position** fields, separated by tabs, then press [[key: Enter]] once.
  check: field.present code="MERGEFIELD Company"
  check: field.present code="MERGEFIELD Contact"
  check: field.present code="MERGEFIELD Position"
- step: Finish the merge to a new document and check that all five employers are listed on one page.
  screenshot: The finished directory, with all five employers on one page.
  example: TODO
  describe: Say how many employers are listed.
:::

## Check your understanding

- Why does a label merge need Update Labels?
- How is a directory merge different from a letters merge?
