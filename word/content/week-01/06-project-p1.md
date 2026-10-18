---
type: project
id: W1-P1
week: 1
order: 6
title: "Project 1: Business Correspondence Set"
objectives: [W1.2, W1.3, W1.4, W1.5]
clos: [CLO1, CLO2]
minutes: 180
status: draft
---

::: callout why
A formal letter is still how people appeal a decision, request information, ask for an interview or put something on the record. A memo is how organizations share decisions internally. You will write both for a real purpose of your own.
:::

## What you'll make

Two documents, each submitted as **.docx** and **.pdf** (four files in all):

1. **A block-format business letter** to a real or realistic recipient. Pick a purpose that matters to you, for example:
   - asking a local employer for a 20-minute informational interview,
   - asking the financial aid office to reconsider an award,
   - thanking someone who helped you, on the record.
2. **A one-page memo** to a group you belong to (a club, a team, your coworkers) announcing a decision or a change.

Name the files `Lastname_P1_Letter.docx`, `Lastname_P1_Letter.pdf`, `Lastname_P1_Memo.docx` and `Lastname_P1_Memo.pdf`.

::: callout privacy
You don't have to use your real address or phone number. A realistic made-up contact block is fine, and so is a made-up recipient. The skills are what's graded.
:::

## Block-format business letter

In block format everything starts at the left margin, and there are no first-line indents. Build the spacing with **Space After**, not empty paragraphs.

| Part | What goes there | Spacing tip |
|---|---|---|
| Your contact block | Name, street, city/state/ZIP, email or phone | Single spaced, 0 pt after inside the block |
| Date | Month day, year (October 20, 2026) | Space after the contact block |
| Inside address | Recipient's name, title, organization, address | |
| Salutation | Dear Ms. Rivera: (a colon in business letters) | |
| Body | Usually three short paragraphs: why you're writing, the details, what happens next | 8 to 12 pt after each paragraph |
| Complimentary close | Sincerely, | About 36 pt after, for a handwritten signature |
| Signature block | Your typed name (and title, if any) | |
| Enclosure line (optional) | Enclosure: Resume | |

## Memo

| Part | What goes there |
|---|---|
| Heading | MEMORANDUM (or MEMO), formatted to stand out |
| Guide words | TO:, FROM:, DATE:, SUBJECT: each on its own line, with a tab after the colon so the entries line up |
| Body | Most important information first; short paragraphs; a bulleted list is fine |

## Required features

Show these Week 1 skills somewhere in the two documents:

- Page setup: 1 in margins, Letter, Portrait
- Consistent fonts (one body font, 10 to 12 pt) and at least one deliberate character format
- Paragraph spacing done with Space After, not empty paragraphs
- At least one symbol or special character used correctly (©, ®, é, an em space - anything that fits)
- Document properties: Title, Author and Tags
- Proofing language set; Editor run with every issue resolved or deliberately ignored
- A PDF of each document that opens and matches the .docx

::: callout ai
Optional: ask an AI tool to draft the body of your letter. You make every formatting decision and check every fact. At the end of the letter's .docx (after the signature, on its own page), add one or two sentences on what you kept, changed or cut. There is no wrong way to try this, and not trying it is fine too.
:::

## Check before you submit

Run these automatic checks on each document. They don't grade anything; they catch the problems that cost the most points.

::: preflight P1-check
title: Preflight for your letter or memo
starter: none
platform: web

- step: No empty paragraphs used as spacing.
  check: para.noEmpty max=0
- step: No double spaces.
  check: text.noDoubleSpaces
- step: Page setup is 1 in margins, Letter, Portrait.
  check: page.is margins=1in size=letter orientation=portrait
- step: Title and Author are filled in.
  check: props.is title="*" author="*"
  platform: desktop
- step: Proofing language is English (United States).
  check: lang.is value=en-US
:::

## How it's graded

::: rubric P1
- P1-a | 25 | CLO1 | Block-letter structure is complete and correct (contact block, date, inside address, salutation, body, closing, signature block); memo heading and guide words are correct.
- P1-b | 25 | CLO1, CLO2 | Formatting is done with Word tools, not workarounds (no spaces for indents, no empty paragraphs for spacing).
- P1-c | 25 | CLO2 | Text is error-free (Editor run, Find and Replace used where needed).
- P1-d | 25 | CLO1 | Page setup, properties and language are set; each PDF opens and matches its .docx.
:::

This project is also the evidence for MO-110 objective 1.3.1 (saving in alternative formats).
