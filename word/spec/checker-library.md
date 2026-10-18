# Checker library (v0.1)

The checker library is the only code that decides whether a step passed. Lesson files name a check and pass it values. They cannot add logic. Each check is written once, tested against known-good and known-bad documents, and then reused across lessons.

**Principle: check the outcome, not the clicks.** A check looks at the finished document, never at how the learner got there. If two routes give the same result, both pass.

## 1. How checks read the document

Every check uses the Word JavaScript API (Office.js) and works in Word for the web and desktop Word, using one of two strategies:

| Strategy | Used for | Requirement set |
|---|---|---|
| **Direct API** | Text search, paragraph properties, fonts, document properties, headers and footers | WordApi 1.1 to 1.3 (web and desktop) |
| **OOXML read** | Things the API does not expose everywhere: page setup on the web, text effects, direct-formatting detection, proofing language, exact line-spacing rules | `getOoxml()`, WordApi 1.1 (web and desktop) |

Page setup has a newer direct API (`Section.pageSetup`), but it is in the WordApiDesktop 1.3 requirement set. The engine uses it when the host supports it and reads the section properties (`w:sectPr`) from OOXML otherwise. Learners see the same result on both platforms.

Each check returns `{ pass, message, detail }`. The message is written for learners ("Top margin is 0.75 in; it should be 1 in"), never as an error code.

## 2. Week 1 checks

### `file.name`
The document is saved with a matching file name.
- `pattern` (required): a name with `*` wildcards, not case sensitive. Example: `pattern="*_W1_Letter.docx"`.
- How: `Office.context.document.url`, base name only. An unsaved document fails with "Save the file first."

### `text.present`
The text appears in the selected part of the document.
- `value` (required), `min` (default 1), `max`, `case` (`sensitive` by default, or `insensitive`), selector (`in=` default `body`).
- How: `body.search()` / `header.search()`.

### `text.absent`
The text does not appear anywhere in the selected part.
- `value` (required), `case`, selector.

### `text.noDoubleSpaces`
No run of two or more spaces in the selected part.
- Selector only.

### `para.noEmpty`
No empty paragraphs (Enter pressed twice to make space).
- `max`: how many empty paragraphs in a row are allowed (default 0). `after="text"` / `before="text"` limit the search to a span.
- How: paragraph text, trimmed.

### `para.order`
One paragraph comes before another.
- `first="text"`, `then="text"` (both required; each finds the first paragraph containing it).

### `font.is`
Text uses the stated character formatting. List only what matters; unlisted properties are ignored.
- Selector, plus any of `name`, `size` (pt), `bold`, `italic`, `underline`, `color` (`#RRGGBB`).
- How: `font` on the selected ranges. Mixed formatting fails with a message naming the first mismatch.

### `effect.present`
A text effect is applied.
- Selector, `kind` = `shadow`, `glow`, `reflection`, `outline` or `any`.
- How: OOXML, looking for the `w14:` effect elements on the selected runs.

### `format.same`
Two pieces of text have matching formatting (the Format Painter outcome).
- `a="text"`, `b="text"` (required), `props` (default `font.name,font.size,font.bold,font.italic,font.color,effect`).

### `format.clean`
Text carries no direct formatting beyond its style (the Clear Formatting and Keep Text Only outcome).
- Selector.
- How: OOXML. Passes when the runs' `w:rPr` holds nothing but the style reference and proofing marks.

### `para.is`
Paragraph formatting.
- Selector, plus any of `align` (`left`, `center`, `right`, `justify`), `lineSpacing` (`single`, `1.15`, `1.5`, `double`, or an exact value such as `14pt`), `spaceBefore`, `spaceAfter`, `firstLineIndent`, `leftIndent` (lengths).
- How: paragraph properties; line-spacing multiples come from OOXML (`w:spacing w:line="240" w:lineRule="auto"` is single). Lengths pass within 0.5 pt.

### `page.is`
Page setup for a section (default: every section).
- Any of `margins` (all four), `top`, `bottom`, `left`, `right`, `size` (`letter`, `a4`, `legal`), `orientation` (`portrait`, `landscape`). Lengths pass within 0.01 in.
- How: `Section.pageSetup` where supported, else OOXML `w:pgMar` / `w:pgSz`.

### `props.is`
Document properties (File > Info).
- Any of `title`, `author`, `keywords` (Tags). A value of `"*"` means "not empty". Matching is "contains", not case sensitive.
- How: `document.properties` (WordApi 1.3).

### `lang.is`
Proofing language for the body text.
- `value` (required), for example `en-US`.
- How: OOXML `w:lang` on runs, falling back to the style and document defaults.

### `self.attest`
The learner confirms they did something the document cannot show (printing, exporting a PDF, customizing the Quick Access Toolbar).
- `label` (required): the sentence the learner confirms.
- Counts toward drill completion only. **Not allowed in a skill check.**

## 3. Week 1 coverage: what can and cannot be auto-checked

| Drill | What the learner does | Auto-check | How |
|---|---|---|---|
| D1.01 | Open and save with a set name | Yes | `file.name` |
| D1.02 | Customize the Quick Access Toolbar; name window parts | Partly | Quiz plus `self.attest` (the API cannot read the toolbar) |
| D1.03 | Show formatting marks; remove extra paragraphs and spaces | Yes (result) | `para.noEmpty`, `text.noDoubleSpaces`. Show/Hide itself cannot be read |
| D1.04 | Find, then replace | Yes | Quiz (count the hits), `text.absent`, `text.present` |
| D1.05 | Symbols and accented letters | Yes | `text.present` with `in=footer` |
| D1.06 | Fonts, bold, text effect, Format Painter | Yes | `font.is`, `effect.present`, `format.same` |
| D1.07 | Spacing, indent, alignment, Clear Formatting | Yes | `para.is`, `format.clean` |
| D1.08 | Cut and paste with Keep Text Only; Undo | Yes (result) | `para.order`, `format.clean`. Undo cannot be read |
| D1.09 | Margins, size, orientation; run Editor | Yes | `page.is`, plus `text.absent` for each planted typo |
| D1.10 | Properties and proofing language | Yes | `props.is`, `lang.is` |
| D1.11 | Export PDF, print settings | No | `self.attest`. The PDF is submitted with Project 1 instead |
| D1.12 | Switch views | No | Quiz |

**Change to the course map:** skill check SC1 listed MO-110 1.3.1 (alternative file formats). A PDF export cannot be checked from inside the document, so 1.3.1 is assessed through the Project 1 PDF submission instead. Skill check SC1 keeps the other ten codes.

## 4. Integrity of skill checks

Checks run in the learner's browser, so a determined learner could fake a pass. Options, from lightest to strongest:

1. **Practice-grade only.** Auto-checks give feedback; the instructor grades the uploaded .docx with the same checks run in an instructor copy of the add-in. (Recommended for v1. It needs no server and stores no student data.)
2. **Signed report.** The add-in exports a report with a hash of the document. This only helps if the instructor re-runs the checks on the submitted file, so it is really option 1 with a receipt.
3. **Proctored session.** The skill check runs in a campus lab or during a live session.
4. **LTI 1.3 later.** The add-in or a small server re-runs the checks on the uploaded file and posts grades to Brightspace. This needs a server, an LTI registration and a FERPA review.

## 5. Adding a check

1. Write the check function with a learner-friendly message for every way it can fail.
2. Add a known-good and a known-bad .docx under `tools/test-docs/`.
3. Add the name and parameters to the validator's list (`trainer-md.js`, `CHECKS`).
4. Document it here.

Until all four steps are done, the validator rejects the check name.
