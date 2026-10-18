# Trainer Markdown - language spec (v0.1)

Trainer Markdown is plain Markdown plus a small set of `:::` directives. It is the ground truth for every lesson the Office Trainer shows. The editor, the add-in and any rebuild all read these files; nothing about a lesson lives anywhere else.

It reuses the `:::` fenced-directive style from MD to Lesson, so there is one dialect to learn.

## 1. Files and folders

```
word/
  course.config.md            course settings (fonts, video hosts, defaults)
  content/week-01/
    01-meet-word.md           one file per lesson
    02-clean-text.md
    03-format-text.md
    04-page-and-output.md
    05-skill-check-sc1.md     the week's timed skill check
    06-project-p1.md          the week's project brief and rubric
  starters/week-01/*.docx     starter documents named in drills
  spec/                       this spec and the checker library
  tools/                      editor and parser
```

File names start with a two-digit order number. The add-in shows lessons in that order.

## 2. Front matter (required)

Every file starts with a front matter block between two `---` lines. Values are `key: value`; lists use square brackets.

```
---
type: lesson
id: W1-L1
week: 1
order: 1
title: Meet Word
objectives: [W1.1, W1.2]
clos: [CLO1, CLO2, CLO6]
minutes: 40
---
```

| Key | Required | Meaning |
|---|---|---|
| `type` | yes | `lesson`, `skillcheck` or `project` |
| `id` | yes | Unique id, for example `W1-L1`, `W1-SC1`, `W1-P1` |
| `week` | yes | Week number |
| `order` | yes | Position in the week |
| `title` | yes | Becomes the page's only H1 |
| `objectives` | yes | Module objective ids from the course map |
| `clos` | no | Course learning outcomes supported |
| `minutes` | no | Estimated learner time |
| `status` | no | `draft` (default) or `ready`. A `ready` file must pass validation with zero errors |

## 3. Ordinary Markdown

Supported: `##` and `###` headings, paragraphs, `**bold**`, `*italic*`, `` `code` ``, links, images, bulleted and numbered lists, and simple pipe tables.

Rules that keep pages accessible and safe:

- **No `#` H1.** The title in front matter is the H1. Start sections at `##`.
- **Every image needs alt text**: `![Ruler showing a 0.5 in first-line indent](images/indent.png)`. Empty alt text is an error unless the image is marked decorative: `![](images/divider.png){decorative}`.
- **Links must be `https://`** (or a relative path to another lesson file).
- **Raw HTML is not allowed.** Anything that looks like a tag is shown as text. This keeps lesson files from injecting code into the add-in.

## 4. Inline extensions

| Syntax | Shows as | Example |
|---|---|---|
| `[[key: Ctrl+S]]` | Keyboard keys | `[[key: Ctrl+S]]` |
| `[[key: Ctrl+S / Cmd+S]]` | Windows / Mac pair | Windows first, Mac second |
| `[[path: Home > Paragraph > Show/Hide]]` | A ribbon path | Tab, group, command |

## 5. Block directives

A directive opens with `:::` plus its name (and sometimes an id) and closes with a line holding only `:::`. Directives do not nest, except that steps live inside a drill, skill check or preflight.

### 5.1 `::: video`

```
::: video
url: https://www.youtube-nocookie.com/embed/VIDEO_ID
title: Saving to OneDrive the first time
captions: verified
duration: 3:20
:::
```

| Key | Required | Notes |
|---|---|---|
| `url` | yes | Must be on an approved host in `course.config.md`. Use `TODO` while the video does not exist yet (a warning, not an error, in drafts) |
| `title` | yes | Shown above the player and used as the frame's accessible name |
| `captions` | yes | `verified` (human-checked), `auto` or `none`. Only `verified` is allowed in a `ready` file (ADA Title II / WCAG 1.2.2) |
| `duration` | no | `m:ss` |
| `transcript` | no | Link to a transcript |
| `start` | no | Start time in seconds |

### 5.2 `::: callout <kind>`

Kinds: `tip`, `note`, `warning`, `why` (why it matters), `ai` (optional AI moment), `privacy`.

```
::: callout why
Saving with a clear name in a known place is the habit that saves your grade the night before a deadline.
:::
```

### 5.3 `::: platform <name>`

Content shown only to learners on that platform: `web` (Word for the web), `desktop` (Windows or Mac), `windows`, or `mac`. The add-in detects the platform and hides the rest. The editor preview can switch between them.

```
::: platform web
In Word for the web, open **File > Save a Copy** to rename the file.
:::
```

### 5.4 `::: quiz <id>`

A short knowledge check. It is not a document check.

```
::: quiz Q1.1
question: Which of these files can Word open?
type: multi
- [x] report.docx
- [x] notes.txt
- [ ] budget.xlsx
- [x] form.pdf
explain: Word opens .docx, .doc, .txt, .rtf, .odt and .pdf (it converts PDFs for editing). .xlsx is an Excel file.
:::
```

`type` is `single`, `multi` or `number`. For `number`, use `answer: 4` instead of options.

### 5.5 `::: drill <id>`

A practice drill. The header keys come first, then a blank line, then optional intro Markdown, then the steps.

```
::: drill D1.04
title: Find and replace a company name
starter: starters/week-01/W1-L2-letter.docx
platform: web
mos: 1.1.1, 2.1.1
minutes: 6

The letter was copied from an old template. Fix the company name everywhere.

- step: Replace every **Acme** with **Finger Lakes Outfitters**.
  check: text.absent value="Acme"
  check: text.present value="Finger Lakes Outfitters" min=4
  hint: [[path: Home > Editing > Replace]] or [[key: Ctrl+H]].
- step: Select **Check my work**.
:::
```

Header keys:

| Key | Required | Notes |
|---|---|---|
| `title` | yes | |
| `starter` | yes | Path to the starter .docx. Checks are written against this file |
| `platform` | yes | `web` (works in both) or `desktop` (needs desktop Word) |
| `mos` | no | MO-110 codes practiced |
| `other` | no | MO-111 (`E1.1`), Northstar (`N3`) and similar codes |
| `minutes` | no | |

Step keys (indented two spaces under `- step:`):

| Key | Notes |
|---|---|
| `check:` | One check from the checker library. Repeat the line for several checks; all must pass. A step with no check is an instruction only and is never scored |
| `hint:` | Shown on request in practice. Never shown in a skill check |
| `video:` | A video URL for this step, optionally with `#t=SECONDS` |
| `web:` / `desktop:` | Platform-specific wording for this step |
| `points:` | Default 1 for a step with checks |
| `answer:` | The learner types an answer (a number or short text) instead of changing the document. Graded like a check. Used for "use Find to count..." tasks |
| `platform:` | `desktop` or `web` when only one platform can do this step. On the other platform the step is skipped and the score is scaled, so no learner is penalized for their computer |

Learners can see the wording, order, hints and values. They cannot change the check logic: `check:` names a tested function in the checker library and only passes it values.

### 5.6 `::: skillcheck <id>`

Same shape as a drill, with `- task:` instead of `- step:`. Hints and `self.attest` are not allowed. Extra header keys: `timeLimit` (minutes), `attempts`, `passing` (percent). The learner selects **Start** to begin the timer and **Submit** to finish; the score is what the document shows at that moment.

### 5.7 `::: preflight <id>`

Ungraded automatic checks a learner can run on their own project before submitting it. Same step syntax as a drill. Nothing here is scored.

### 5.8 `::: rubric <id>`

One criterion per line: `- id | points | CLOs | description`. Points must total 100.

```
::: rubric P1
- P1-a | 25 | CLO1 | Block-letter structure is complete and correct.
:::
```

## 6. The `check:` line

```
check: <check-name> <param>=<value> <param>="value with spaces"
```

- Values in double quotes may contain spaces. Use `\"` for a literal quote.
- Lengths take units: `1in`, `0.5in`, `36pt`, `2.54cm`. The engine works in points (72 pt = 1 in).
- Booleans are `true` or `false`.
- Where a check looks at part of the document, choose exactly one selector:

| Selector | Meaning |
|---|---|
| `in=body` | All body paragraphs (default) |
| `in=header` / `in=footer` | The primary header or footer |
| `para="text"` | The first paragraph that contains this text |
| `paras="text"` | Every paragraph that contains this text |
| `text="text"` | Exactly this text (the first search hit) |
| `style="Heading 1"` | Every paragraph using this style |

The full list of checks, their parameters and how each one works in Word for the web and desktop is in `checker-library.md`.

## 7. Validation

The editor (and `tools/validate.mjs`) reports:

- **Errors** (block a `ready` file): missing front matter keys, an H1 in the body, an unknown check or parameter, a bad unit, a missing starter file, a missing image alt text, a video host not on the approved list, a `ready` video without `captions: verified`, rubric points that do not total 100, a hint inside a skill check, raw HTML.
- **Warnings** (allowed in drafts): `url: TODO` videos, a drill with no checked steps, objective ids not found in the course map.

## 8. Security model

- Lesson files are data, never code. The only way a file affects grading is by naming a check from the library.
- All text is escaped when rendered. Raw HTML is shown as text.
- Videos load only from approved hosts, in sandboxed frames, using privacy-enhanced embed URLs where the host offers one (`youtube-nocookie.com`).
- Checks run in the learner's browser. That is fine for practice, but a determined learner could fake a result. Skill checks therefore need a separate integrity plan (see `checker-library.md`, section 4).
- The add-in stores progress only in the learner's own browser in v1. No server holds student data.
