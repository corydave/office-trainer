# Checker library (v0.2, built: all five weeks)

The checker library is the only code that decides whether a step passed. Lesson files name a check and pass it values. They cannot add logic. Each check is written once, tested against known-good and known-bad documents, and then reused across lessons.

**Principle: check the outcome, not the clicks.** A check looks at the finished document, never at how the learner got there. If two routes give the same result, both pass.

## 1. How checks read the document

**Change from the first plan:** every check now reads the saved `.docx` file itself, not Office.js. The code is `tools/word-checker.js`.

- **Inside Word**, the add-in asks for a copy of the open document (`getFileAsync`, the full .docx). This works the same way in Word for the web and desktop Word.
- **In a browser** (a Chromebook without the add-in, or a phone), the learner uploads their saved file and gets the same result.
- The .docx holds everything the checks need, including what Office.js can't show on every platform: page setup on the web, text effects, direct formatting, the proofing language, the exact line-spacing rules, and the properties in `docProps/core.xml`.
- **Office.js is used for three things only:** getting the file, loading a starter into the open document, and "Show me where" (selecting the spot a failed check is about).

Formatting is looked up the way Word does it. Direct formatting comes first, then a character style, then the paragraph style and the styles it's based on, then the document defaults and the theme fonts. So a heading counts as bold whether the learner clicked **B** or used a bold style.

Each check returns `{ pass, message, where }`. The message is written for learners, for example "All four margins are 0.75 in, not 1 in; the paper is A4 (8.27 x 11.69 in), not Letter". `where` tells "Show me where" what to select.

## 2. Week 1 checks

### `file.name`
The document is saved with a matching file name.
- `pattern` (required): a name with `*` wildcards, not case sensitive. Example: `pattern="*_W1_Letter.docx"`.
- How: the file name Word reports (`Office.context.document.url`, last part only), or the uploaded file's name. An unsaved document fails with "Save the file first." A name that still starts with `Lastname_` fails with "put your own last name in place of Lastname".

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
- How: paragraph text, trimmed. Spaces and tabs don't count as content. Pictures, page breaks and section breaks do, and so does table cell text. One empty paragraph at the very end of the document is allowed, because Word always keeps a last paragraph mark that can't be deleted.

### `para.order`
One paragraph comes before another.
- `first="text"`, `then="text"` (both required; each finds the first paragraph containing it).

### `font.is`
Text uses the stated character formatting. List only what matters; unlisted properties are ignored.
- Selector, plus any of `name`, `size` (pt), `bold`, `italic`, `underline`, `color` (`#RRGGBB`).
- How: each run's formatting, resolved through styles and theme fonts. Mixed formatting fails with a message naming the first run that's wrong.

### `effect.present`
A text effect is applied.
- Selector, `kind` = `shadow`, `glow`, `reflection`, `outline` or `any`.
- How: looks for the `w14:` effect elements (and the older `w:shadow` / `w:outline`) on every run, including ones that come from a style. An outline with "no fill" doesn't count, because Word saves one with some presets.

### `format.same`
Two pieces of text have matching formatting (the Format Painter outcome).
- `a="text"`, `b="text"` (required), `props` (default `font.name,font.size,font.bold,font.italic,font.color,effect`).

### `format.clean`
Text carries no direct formatting beyond its style (the Clear Formatting and Keep Text Only outcome).
- Selector.
- How: passes when the runs' formatting holds nothing but a style, the language and proofing marks. The failure message names what's left, for example "(font, color, size)".

### `para.is`
Paragraph formatting.
- Selector, plus any of `align` (`left`, `center`, `right`, `justify`), `lineSpacing` (`single`, `1.15`, `1.5`, `double`, or an exact value such as `14pt`), `spaceBefore`, `spaceAfter`, `firstLineIndent`, `leftIndent` (lengths).
- How: paragraph settings, resolved through the style chain (`w:spacing w:line="240" w:lineRule="auto"` is single). Lengths pass within 0.5 pt.

### `page.is`
Page setup for a section (default: every section).
- Any of `margins` (all four), `top`, `bottom`, `left`, `right`, `size` (`letter`, `a4`, `legal`), `orientation` (`portrait`, `landscape`). Lengths pass within 0.01 in.
- How: `w:pgSz` and `w:pgMar` for each section. Orientation comes from the page's shape (wider than tall), not the `orient` flag alone. Every problem in a section is reported at once.

### `props.is`
Document properties (File > Info).
- Any of `title`, `author`, `keywords` (Tags). A value of `"*"` means "not empty". Matching is "contains", not case sensitive.
- How: `docProps/core.xml` (`dc:title`, `dc:creator`, `cp:keywords`).

### `lang.is`
Proofing language for the body text.
- `value` (required), for example `en-US`.
- How: `w:lang` on each run, then its styles, then the document defaults, then the document's own language (`w:themeFontLang`).

### `self.attest`
The learner confirms they did something the document cannot show (printing, exporting a PDF, customizing the Quick Access Toolbar).
- `label` (required): the sentence the learner confirms.
- Counts toward drill completion only. **Not allowed in a skill check.**

## 2b. Weeks 2-5 checks

Every check below reads the saved .docx, like the Week 1 checks. Selectors (`in`, `para`, `paras`, `text`, `style`) work the same way; parameters never reuse those names, which is why some are called `pictureStyle`, `says`, `label` or `citeStyle`. A count check given only `max` starts its `min` at 0.

| Check | Parameters | What it reads |
|---|---|---|
| `style.is` | `name` + selector | The paragraph style a paragraph uses |
| `style.def` | `name`, `font`, `size`, `bold`, `italic`, `color`, `spaceBefore`, `spaceAfter`, `keepNext`, `align`, `lineSpacing` | The style's own definition, resolved through its based-on chain (Automatic counts as black) |
| `bookmark.is` | `name`, `at` | A bookmark, and the paragraph it marks |
| `link.present` | `label`, `anchor`, `url`, `descriptive`, `min` | Hyperlinks: to a bookmark (`anchor`) or a web address; `descriptive` rejects "click here" and bare addresses |
| `field.present` | `code`, `min`, `max`, `in` | Fields (PAGE, TOC, SEQ, MERGEFIELD, IF, ADDRESSBLOCK...), simple or complex |
| `section.is` | `index`, `orientation`, `columns`, `breakType`, `firstPageDifferent`, `border` + selector | One section's page setup |
| `break.count` | `kind` (page, column, section), `min`, `max` | Manual breaks |
| `list.is` | `kind`, `level`, `format`, `symbol`, `pattern`, `number` + selector | Real Word lists: numbering format, bullet symbol, the number this item shows |
| `notes.is` | `kind`, `min`, `max`, `format` | Footnotes or endnotes and their number format |
| `toc.is` | `levels`, `includes` | A generated table of contents (its field switches and its entries) |
| `citation.is` | `citations`, `sources`, `bibliography`, `citeStyle` | The citation manager's source list, CITATION fields and the bibliography |
| `object.count` / `object.is` | `kind` (picture, icon, shape, textbox, smartart, chart, model3d), `index`, `name`, `any`, `contains`, `alt`, `decorative`, `wrap`, `halign`, `valign`, `pictureStyle`, `artistic`, `bgRemoved`, `fill`, `outline`, `says`, `layout`, `items`, `colors`, `smartStyle`, `chartType` | Drawings: pictures and their effects, shapes and their text, SmartArt parts, charts, 3D models |
| `table.count` / `table.is` | `index`, `contains`, `lacks`, `any`, `rows`, `minRows`, `cols`, `headerRepeat`, `merged`, `sortedBy`, `order`, `cellMargin`, `cellSpacing`, `col`, `width` | Tables. Sorting understands numbers, dates, month names and times |
| `theme.is` | `colors`, `accent1`-`accent6` | The theme's color set |
| `page.background` | `watermark`, `color` | Watermark text and page color |
| `tabs.is` | `position`, `align`, `leader` + selector | Tab stops on a paragraph that really uses a tab character |
| `controls.filled` | none | Content controls still showing their placeholder |
| `merge.is` | `type`, `source` | The mail merge settings Word saves in the main document |
| `comment.count` / `comment.is` | `says`, `anchor`, `author`, `state` (present, absent, replied, resolved, open) | Comments, replies (commentsExtended) and resolved threads |
| `revision.count` | `kind` (insert, delete, format, move), `author`, `min`, `max` | Tracked changes still waiting for a decision |
| `doc.is` | `tracking`, `protection`, `password`, `compatMode`, `personalInfo` | Settings: Track Changes, Restrict Editing, compatibility mode, Inspect Document's "remove personal information" |
| `text.hidden` | `max` | Runs formatted as hidden text |
| `access.is` | `rules` (alt, tableHeaders, title, headings, links) | The errors the Accessibility Checker reports. Without `rules`, all five |

`para.is` also gained `dropCap` and `keepNext`.

## 3. Coverage: what can and cannot be auto-checked

### Week 1

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

### Weeks 2-5: what is not auto-checked

Most drill steps in Weeks 2-5 are auto-checked. These use a screenshot (compared with an example in practice mode) or are instructions only, because the result isn't saved in the document:

| Drill | Step | Why |
|---|---|---|
| D2.02 | Try a style set | The style set's name isn't saved; its effect is checked through `style.def` in the next steps |
| D2.03 | Go To page 2 | Navigation leaves no trace |
| D3.06 | Screen clipping, rotating a 3D model | A clipping is just a picture; the model's angle is checked by eye |
| D4.02 | Save a Quick Part | Quick Parts live in Building Blocks.dotx, not the document. Using it (the placeholder is gone) is checked |
| D4.06, D4.08, D4.09 | Preview results; the merged labels and directory | Previews and some results open in a new window |
| D4.10, D4.11 | The PDF and the web page | Files outside the document; submitted with Project 4 |
| D5.05 | Sharing, co-authoring, version history | Live in OneDrive, not the file |

Mail merge results, Compare and Combine open in a **new window**. Learners save the result, open Office Trainer in that window and check there; passed steps are remembered.

## 4. Integrity of skill checks

Checks run in the learner's browser, so a determined learner could fake a pass. Options, from lightest to strongest:

1. **Practice-grade only.** Auto-checks give feedback; the instructor grades the uploaded .docx with the same checks run in an instructor copy of the add-in. (Recommended for v1. It needs no server and stores no student data.)
2. **Signed report.** The add-in exports a report with a hash of the document. This only helps if the instructor re-runs the checks on the submitted file, so it is really option 1 with a receipt.
3. **Proctored session.** The skill check runs in a campus lab or during a live session.
4. **LTI 1.3 later.** The add-in or a small server re-runs the checks on the uploaded file and posts grades to Brightspace. This needs a server, an LTI registration and a FERPA review.

## 5. Testing the checks

- `tools/test-docs/` holds answer keys (`*-solved.docx`, each starter with every drill done) and part-done files (`*-partial.docx`). `tools/make-test-docs.py` builds them from the starters.
- `tools/test-docs/expected.json` lists, for each file, whether each step should pass or fail.
- `tools/check-tests.html` runs every checked step in all five weeks against those files and the starters (74 documents), and shows a table. Open it from the course website, or any local web server (browsers block it from a `file://` address). Today: **558 of 558 results match** (checker 0.2.0).
- The same cases were also run end to end through the add-in's own buttons (Open it here, Check my work, skill check Start and Submit) in a browser with a stand-in for Word: 494 step results, 0 mismatches, and every badge and certificate, including the course certificate, was earned.
- These files are made by code, not saved from Word. Word saves some things differently (theme fonts, effect presets, Format Painter output). **Next:** do each drill in real Word for the web and desktop Word, save the results into `test-docs/`, and add them to `expected.json`.

## 6. Adding a check

1. Write the check function in `tools/word-checker.js` (PART 5), with a learner-friendly message for every way it can fail.
2. Add a known-good and a known-bad .docx to `tools/test-docs/`, and their expected results to `expected.json`. Run `check-tests.html`.
3. Add the name and parameters to the validator's list (`trainer-md.js`, `CHECKS`).
4. Document it here, then rebuild the add-in: `python3 tools/build-taskpane.py`.

Until steps 3 and 4 are done, the validator rejects the check name.
