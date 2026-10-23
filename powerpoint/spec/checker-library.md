# PowerPoint checker library (v0.1.0, all five weeks)

The checker is the only code that decides whether a step passed. Lesson files name a check and pass it values; they can't add logic. The code is `tools/ppt-checker.js` (one file; it needs only JSZip). The add-in, `tools/check-tests.html` and the command-line tests all share it.

**Principle: check the outcome, not the clicks.** A check looks at the saved presentation, never at how the learner got there. Bold from the Ribbon, from Ctrl+B or from the mini toolbar all pass "the title is bold".

## 1. How the checker reads a presentation

- **Inside PowerPoint**, the add-in asks for a copy of the open presentation (`getFileAsync`, the whole .pptx), saved or not. **In a browser**, the learner uploads the saved .pptx (or .potx).
- The checker opens the file with JSZip and reads the XML parts with the browser's DOMParser. Nothing in the file runs: media are never played, and only the first bytes of a picture are read (to get its size).
- It builds one plain model of the presentation: properties, slide size, the view it was saved in, slides in show order (with layout, title, notes, background, transition, animations and every shape), sections, custom shows, slide show settings, masters and layouts, the notes and handout masters, the theme, and modern comments.
- **Shapes** are sorted into kinds: placeholder, text box, shape, action button, line, picture, icon (a picture with an SVG copy), video, audio, table, chart, SmartArt, 3D model, Zoom, ink and group. Newer objects saved inside `mc:AlternateContent` (Zoom, 3D models, ink) are read from the newer choice.
- **Titles** come from the slide's title placeholder, so every check can point at a slide by its title.
- **A file name in the alt text box** (python-pptx and some tools put "photo.png" there) counts as no alt text.
- **Animations** are read from the slide's timing tree: class (entrance, emphasis, exit, motion path), preset, start (On Click, With Previous, After Previous), delay, duration, target shape and paragraph, motion path direction, and media play commands.
- **Transitions** are read from `p:transition` (and the newer `p14`/`p15` kinds inside AlternateContent). Newer kinds count as "3D", except Reveal, Flash, Morph and Clock reverse, which are subtle.
- **A presentation with a password to open** is encrypted into a different kind of file. The checker says so in plain words instead of failing oddly.
- **Office.js is used for three things only:** getting the file, opening a practice file (`PowerPoint.createPresentation`) and "Show me where" (going to the slide, and selecting the shape when PowerPointApi 1.5 is there).

Each check returns `{ pass, message, where }`. Messages are written for learners ("Slide 3 has no transition yet. Pick one on the Transitions tab."). `where` is `{ slide, slideId, shape }`.

## 2. Pointing at slides and shapes

- `slide=3`, `slide=first`, `slide=last` or `slide=all` picks slides by position. `title="Our prices"` picks slides whose title contains the text. With neither, a check looks at every slide.
- `says="Book now"` picks a shape by its text, `name=Logo` by its name (Selection Pane), `kind=picture` by kind. `kind=text` means any text holder, `kind=media` video or audio, `kind=image` picture or icon, `kind=graphic` anything but placeholders and groups.
- Lengths take units (`width=3in`); times are in seconds (`duration=2`, `advanceAfter=5`).

## 3. The checks

`*` = required.

### Files, text and slides

| Check | What it tests | Parameters |
|---|---|---|
| `file.name` | The file's name matches a pattern (`*` wildcards); "Lastname_" fails with a reminder | pattern* |
| `props.is` | Title, Author, Tags (keywords) or Subject property | title, author, keywords, subject |
| `self.attest` | Nothing: the learner ticks a box (not in skill checks) | label* |
| `text.present` / `text.absent` | Text on the slides (or `in=notes`), on at least `min` slides | value*, slide, title, in, min |
| `slide.count` | How many slides (or slides with a title, or visible ones) | is, min, max, title, visible |
| `slide.is` | One slide: position, right after another, layout, hidden, section, background (any, none, picture, gradient, solid, pattern), notes, title. `exists=false` for a deleted slide. `background=none` also fails when a picture went on the slide master through Apply to All | slide or title, number, after, layout, hidden, section, background, notes, hasNotes, titleIs, exists |
| `notes.count` | How many slides have speaker notes | min, max |
| `section.is` / `section.count` | A section's name, position, size, first slide and slides | name*, exists, index, slides, minSlides, starts, contains / min, max |
| `size.is` | Widescreen (16:9) or Standard (4:3) | ratio* |
| `view.is` | The view the file was saved in, and whether guides show | view (normal, outline, sorter, notes ...), guides |

### Text and lists

| Check | What it tests | Parameters |
|---|---|---|
| `text.is` | Formatting of a shape's text: bold, italic, underline, size, minimum size, font, a font that must not be used anywhere, WordArt, color, contrast with the background (WCAG ratio), columns, alignment | says, kind, phType, bold, italic, underline, size, minSize, font, notFont, wordart, color, colorChanged, minContrast, columns, align |
| `list.is` | A list: bullets or numbers, the level of some items, how many levels, at most so many items per slide. With no `says`, the longest list that isn't a title (and, across slides, one of the asked-for kind) | says, items, kind, level, levels, maxItems |
| `footer.is` | Slide numbers, footer text and date on every slide but the title slide, and none on the title slide | slideNumbers, footer, skipTitle, date |

### Design: themes, masters and layouts

| Check | What it tests | Parameters |
|---|---|---|
| `theme.is` | Theme name, changed from Office Theme, font set, colors, a variant | name, changed, fonts, fontsChanged, colors, variant |
| `master.is` | The slide, notes or handout master: background, title font, a picture on it, text, a placeholder's text (`phSays="hdr:Our name"`), a placeholder removed | kind, background, titleFont, picture, says, phSays, lacks |
| `layout.is` | A layout exists (or not), is custom, has placeholders (pic, body, chart ...), lacks one, is used by a slide | name*, exists, custom, placeholder, lacks, used |

### Graphics, tables and charts

| Check | What it tests | Parameters |
|---|---|---|
| `object.count` | How many objects of a kind | kind*, min, max |
| `object.is` | One object: alt text or decorative, exact or minimum size, cropped, square, not stretched, styled, effects, shape style changed, shape type (`geom=callout`), in front of anything it overlaps, grouped, rotated (or a 3D model turned), SmartArt items, layout and colors | kind, says, name, index, alt, decorative, width, height, minWidth, cropped, square, stretched, styled, effects, styleChanged, geom, front, grouped, rotated, items, layout, notLayout, colorsChanged |
| `align.is` | Objects share an edge (top, middle, left ...), or have even gaps | kind, says, name, edge, distributed |
| `table.is` | Rows, columns, style changed from the default, header row, text in or out of it | rows, cols, styleChanged, headerRow, has, lacks |
| `chart.is` | Chart type, grouping, number of series, chart title (`chartTitle`), data labels, alt text | type, notType, grouping, minSeries, chartTitle, dataLabels, alt |

### Motion, media and links

| Check | What it tests | Parameters |
|---|---|---|
| `media.is` | A video or audio clip: trimmed, fades in, starts automatically or on click, loops, plays across slides, alt text | kind, trimmed, fadeIn, start, loop, acrossSlides, alt |
| `transition.is` | A slide's (or every slide's) transition: type, 3D or basic, effect option changed from the default, duration, advance after, on click, sound; `exists=false` for none | type, kind, optionChanged, duration, advanceAfter, onClick, sound, exists |
| `anim.is` | An animation on a shape: class, effect name, by paragraph, start, minimum delay, duration, first in order, motion path direction | says, kind, name, class, effect, byParagraph, start, minDelay, duration, first, pathDir |
| `anim.count` | How many animations | min, max (and the same target filters) |
| `link.present` | A link to a web address, to a slide (by title), on certain words, or on an action button | url, toSlide, says, button |
| `zoom.is` | A Summary, Section or Slide Zoom, with Return to Zoom | kind, returnToZoom |

### Shows, review and sharing

| Check | What it tests | Parameters |
|---|---|---|
| `show.is` | Loop, kiosk, narration, use timings, a custom show's size and slides, timings on every visible slide | loop, kiosk, narration, useTimings, customShow, slides, includes, excludes, timings |
| `comment.is` / `comment.count` | A modern comment thread: there, gone, replied to, resolved; by text or author | says, author, state / min, max |
| `doc.is` | Personal information removed, fonts embedded, marked as final, a password to modify, no notes left, nothing off the slides | personalInfo, fontsEmbedded, final, modifyPassword, notes, offSlide |
| `access.is` | Accessibility rules: alt text, slide titles, unique titles, table header rows, link text, Title property, reading order | rules (default alt, titles, uniqueTitles, tableHeaders, links) |

## 4. What the checker can't see

- **Anything not saved in the file:** the Quick Access Toolbar, printing, Presenter View, a screen being shared, co-authors' presence. Those steps use screenshots.
- **What an animation looks like.** It reads the effect's preset id and timing, not the motion; a custom path's shape is reduced to a direction (left, right, up, down).
- **Media content:** whether a video has the right scene or captions. It sees trims, fades and playback settings.
- **Handout "slides per page".** PowerPoint doesn't save it in the handout master, so D2.05 checks the header text instead.
- **End Review after Compare.** The checker sees the accepted and rejected changes, not the review state.
- **Text colors with tints and shades.** The contrast check uses the base theme color and ignores brightness tweaks, and skips picture and gradient backgrounds.
- **Encrypted files** (password to open) can't be read at all, by design.

## 5. Tests

`tools/make-test-docs.py` builds answer keys and part-done copies with the same builders as the starters and writes `tools/test-docs/expected.json`. `tools/check-tests.html` runs every checked step against them. Every starter must fail every checked step it's meant to (a starter never passes a step by itself); the few deliberate exceptions are listed in `make-test-docs.py`.
