# Opportunities log

Ideas, open questions and fixes found while building and testing Weeks 2-5 (October 2026, add-in 0.5, checker 0.2.0). Newest pass first. Each item says what it is, why it matters, and a suggested next step. Priority: **High** (affects learners now), **Medium**, **Low**.

## 1. Fixed in this pass

Bugs that testing found, already corrected:

| What was wrong | Fix |
|---|---|
| The skill check start panel showed the word "null" (an empty optional line went straight into the page) | The add-in now adds page parts with `add()` and `fill()`, which skip empty pieces. Every `append`/`replaceChildren` in the add-in goes through them |
| The course certificate was never issued if the last badge was a skill check or the capstone (it only checked after a lesson certificate) | It's now checked after every badge, and once at start-up for learners who finished earlier |
| The course certificate's back overflowed the page (66 badges, every outcome, every MO-110 skill) | The course certificate lists badges by week, one line each, and the MO-110 skills as a compact list. Lesson certificates are unchanged |
| The home page said "badges available so far" and "more lessons are on the way" with the course complete | Shown only while some stops are still planned. A "Course complete" card with the certificate button appears on the home page |
| The Badges page said "21 of 20 lesson certificates" | The course certificate is counted separately |
| The roadmap counted 108 badges (it added the outline's planned drill list to the real drills) and kept the "Coming soon" legend | Planned drills only count for unwritten lessons; the legend and footer note show only while something is planned |
| The capstone's exam was labeled "Skill check CAP" | It's labeled "Practice exam CAP" |
| An automated accessibility scan (axe-core, WCAG 2.1 AA) found wide tables that scroll sideways but couldn't be reached with the keyboard | Lesson tables are now focusable, labeled regions. The scan now finds **0 violations** on the home page, all 33 stops, the roadmap and the certificate, in light and dark mode |
| The editor's **New > Lesson** template failed validation (no video id) | Fixed, and it now includes a badge name and a sample quiz |
| D4.02 and D4.03 were out of order (Lesson 1 had D4.01 and D4.03) | Renumbered: Quick Parts is D4.02, the ATS resume tabs drill is D4.03 |
| The python-docx default template saves in Compatibility Mode (Word 2010) | Starters are made in the current format (compat 15), except the one that teaches Convert |

Content fixes from a review of Weeks 2-5 (the details are in the lesson files):

- Share links: expiration dates and passwords only exist for *Anyone* links. D5.05 and Project 5 now have a fallback (share view-only with a named person).
- Inspect Document: **Remove All** also accepts tracked changes and clears the Title property. D5.06, D5.07, Project 5 and the capstone now say so, and put "set the title" after inspecting.
- D5.07's Title step is desktop-only (Week 1 says Word for the web can't set properties), and the "fix everything else" step no longer depends on it.
- D2.07 (Different First Page) now says why this report hides the page 1 header, and that APA and MLA papers don't.
- D2.10's hint no longer makes an empty Heading 1.
- D2.11's first step said "Set the style to APA", which passed on the starter because APA is Word's default. It now says "make sure", which is honest about that.
- Project 2: APA fonts (Times New Roman 12, Calibri 11, Arial 11; Aptos removed), and the title page includes affiliation.
- D3.09 and Project 3: a tip that a merged title row in a repeating header confuses screen readers; put the title above the table.
- Templates: Mac paths (File > New from Template, AutoText) and the lab-computer Building Blocks prompt.
- Mail merge: use a colon in the Greeting Line for business letters. Project 4: zip the .htm file with its folder.
- SC4 maps to MO-111 (mail merge isn't on MO-110), and now checks that [RenewBy] is gone. SC5 no longer lists two MO-110 skills it doesn't test.
- Smaller wording fixes: "View Alt Text (or Edit Alt Text)", page color printing, Combine takes two copies at a time, the capstone's description of the real exam, a Mac End-key shortcut.

## 2. Test in real Word first (High)

Everything so far ran in a browser with a stand-in for Word. Try these first, in both Word for the web and desktop Word:

1. **Loading starters.** `insertFileFromBase64` should bring over headers, footers, page setup, comments, tracked changes, content controls and the mail merge settings. Watch W5-L1-policy (comments and changes), W4-L3-letter (merge settings) and W3-L1-flyer (floating pictures).
2. **New-window results.** Mail merge (D4.07), Compare and Combine (D5.04) open a new document. Check that Office Trainer opens there from the Home tab and remembers passed steps, and that the result's file name reaches `file.name`.
3. **Show me where** for the new checks (tables, pictures, comments, list items).
4. **Files Word really writes.** The test documents are made by code. Save real results for at least: SmartArt (layout and color ids), charts (chart type and caption), 3D models, icons (SVG pictures), removed backgrounds and artistic effects, resolved comments (commentsExtended), Inspect Document's "remove personal information" setting, and Lock Tracking. Add them to `tools/test-docs/` and `expected.json`.
5. **Uploading a .dotx** (D4.01). The test used a .docx renamed to .dotx; a real template has a different content type inside.
6. **The skill check timer** across closing and reopening the pane in desktop Word.

## 3. UX opportunities

| Priority | Idea | Why |
|---|---|---|
| High | **A "this opens a new window" helper** for D4.07 and D5.04: a short step-by-step card with a screenshot of where the add-in button is | It's the one place learners are most likely to get lost |
| High | **Skill check answers in the page.** Quiz answers and `answer:` values are in the page source (`data-answer`). Fine for practice; for a class section, consider checking answers by a hash, or only after Submit | Integrity for graded sections |
| Medium | **Resume where you left off inside a lesson.** Reopening a lesson scrolls to the top. Jump to the first unfinished drill | Long lessons (Week 3 Lesson 1 has seven drills) |
| Medium | **"Start here" isn't required for the course certificate.** Stops with no badges count as done. Decide whether Start Here must be opened first | Policy choice |
| Medium | **Preflight results could be saved** like drills, so a learner sees "last checked: passed" next to the project | Projects take several sessions |
| Low | **Badges page by week** could show progress bars per week, like the home page | Motivation |

## 4. UI opportunities

| Priority | Idea |
|---|---|
| Medium | Week 3 Lesson 1 is long on a phone (seven drills). Consider splitting it into "Pictures and shapes" and "SmartArt, text boxes and alt text", or collapsing finished drills |
| Medium | The roadmap's station labels wrap to three lines in places (Week 2, Week 4). Shorter outline titles, or a smaller label size under 900 px |
| Low | The course certificate back could use a week seal row instead of a seal per line, if it ever feels busy |
| Low | Dark mode screenshots look right, but the badge art's week colors haven't been contrast-checked against the dark background |

## 5. Content opportunities

| Priority | Idea |
|---|---|
| High | **Videos and screenshots.** 21 videos and 16 example screenshots are still TODO. Lessons work without them, but the screenshot drills can't be compared until the examples exist |
| Medium | **Real-Word fixtures for Week 3** (see 2.4) also make good example screenshots |
| Medium | **Accessibility thread.** Alt text, header rows, headings and link text come up in Weeks 3-5. A one-page "Accessibility checklist" in Start Here would tie them together, with the ADA Title II note |
| Medium | **AI callouts.** Week 1 has a Copilot note; Weeks 2-5 have none. Candidates: Copilot drafting a cover letter (Week 4, with an honesty note), Copilot summarizing comments (Week 5) |
| Low | **CLO 7 wording.** The official outcome starts with a lowercase "use mail merge...". It's kept verbatim, but it shows that way on the course certificate. Worth fixing in the official course description when CLO 6 and CLO 2 are reworded |
| Low | **MO-110 map.** A table of every MO-110 objective and where it's practiced and assessed would help instructors (the data is already in the outline and lesson `mos:` lines) |
| Low | **Week 4 resume drills** use Alex Morgan throughout. A second persona (a career changer) would show a different resume shape |

## 6. Verify in current Word (facts that may have changed)

The lessons state these; check them against current Microsoft 365 and fix the wording if needed:

- Word for the web: inserting charts, Go To, watermarks, the citation manager, text box position presets (D3.05, SC3), **Repeat Header Rows** (D5.07 is a web drill), setting the Title property, and updating a style to match a selection (D2.01 text vs. SC2).
- Word's citation styles list: does it say "APA Seventh Edition" (Project 2 requires APA 7)?
- Mac: Quick Parts vs. AutoText naming in current Mac Word.
- "View Alt Text" vs. "Edit Alt Text" (the lessons now name both).

## 7. Housekeeping

- `tools/test-docs/expected.json` is long and hand-edited through scripts. A compact one-line-per-document format would be easier to review.
- The Playwright tests used for this pass live outside the repo. Moving them into `tools/tests/` (with the Word stand-in) would let anyone rerun them.
- Badge ids include the drill id, so renumbering a drill changes its badge number. Fine before launch; after launch, never renumber.
