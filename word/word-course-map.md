# Microsoft Word - 5-Week Course Map

**Finger Lakes Community College (FLCC)** · Draft 1 · October 17, 2026 · Dave Ghidiu (with Claude)

- **Format:** 5 weeks, online or hybrid, about 9 hours of learner work per week (assumes 1 credit - see Decisions)
- **Certification target:** Microsoft Office Specialist (MOS): Word Associate (Microsoft 365 Apps), exam MO-110
- **Coverage:** 8/8 course learning outcomes assessed · 7/7 outline topics · 71/71 MO-110 sub-skills practiced (53 also in weekly skill checks; all 71 in the capstone) · 5 MO-111 objectives fully and 7 partly covered
- **Companion:** `word-course-map.html` shows this map interactively. This Markdown file is the master copy. Both files are generated from the same data (`_build/data.py`), so they always agree; after changing the map, regenerate both with `python3 build.py`.

## How to read this map

- **CLO** = course learning outcome from the official course description (verbatim). **T** = topic outline item (verbatim).
- **I / R / A** = Introduced, Reinforced, Assessed (standard curriculum mapping).
- **Drills** are short, outcome-checked practice tasks for the Office Trainer (it checks the result, not the clicks). Each is tagged *Web + desktop* or *Desktop only*.
- **MO-110 codes** (for example 2.2.3) follow Microsoft's skills outline. **E** codes (for example E4.3) are MO-111 Expert objectives. **N** codes are Northstar Digital Literacy Word standards.

## Course at a glance

| Week | Theme | Project | CLOs (I/R/A) | Outline topics | MO-110 domains drilled |
|---|---|---|---|---|---|
| 1 | **The Professional Document** - Create, format, save and print documents people take seriously. | P1 Business Correspondence Set | CLO1 IA; CLO2 IA; CLO6 I | T1, T2, T3 | 1, 2 |
| 2 | **Structure and the College Paper** - Let styles do the work: long documents, citations and a reusable paper template. | P2 Research Paper Template and Paper | CLO1 R; CLO2 RA; CLO6 R | T3, T6 | 1, 2, 3, 4 |
| 3 | **Visual Communication** - Graphics, tables, charts and desktop publishing that stay accessible. | P3 Newsletter or Event Flyer | CLO3 IA; CLO4 IA; CLO5 IA; CLO2 R; CLO1 R | T4, T5, T6 | 1, 2, 3, 5 |
| 4 | **The Job Application Packet** - Templates, mail merge and web output, built around getting hired. | P4 Job Application Packet | CLO7 IA; CLO6 A; CLO1 RA; CLO2 R; CLO4 R; CLO5 R | T5, T6, T7 | 1 |
| 5 | **Working Together and Shipping Clean** - Review, collaborate, then share safely. Capstone and certification practice. | P5 Peer Review Exchange and Clean-and-Ship | CLO8 IA; CLO1 RA; CLO2 RA; CLO3 A; CLO4 A; CLO5 A; CLO6 R; CLO7 R | T7 | 1, 6 |

## Decisions and flags

- **[FLAG] Reword CLO 6.** "Create and preview the document as a Web page" reflects an older workflow. The plan meets it as written (Web Layout view plus Save as Web Page in Week 4), then spends most of the time on what people do now: PDF export and share links. Consider rewording to "Publish documents for screen and web audiences (web page, PDF, share link)."
- **[FLAG] Make CLO 2 measurable.** "Demonstrate features offered by the application" is hard to assess on its own. It is mapped to every week here. Consider rewording to "Use Word's structure and productivity features (styles, navigation, references, templates) to produce documents efficiently."
- **[NOTE] Some course outcomes go beyond MO-110.** Mail merge, templates, compare and combine, citations and restrict editing are MO-111 (Expert) skills. They stay because the course outcomes require them, and they are labeled as beyond-the-certification content.
- **[FLAG] Plan for desktop Word.** Weeks 2 to 5 include features that need desktop Word (mail merge, compare and combine, SmartArt, section breaks, citations). Chromebook users will need campus labs or a virtual desktop for those tasks. Each drill is tagged web or desktop so the Office Trainer can show the right path.
- **[NOTE] Accessibility is built in, not bolted on.** The ADA Title II web rule (WCAG 2.1 AA) covers documents public colleges post. A DOJ interim final rule (April 20, 2026) moved compliance to April 26, 2027 for entities of 50,000 or more and April 26, 2028 for smaller ones. Confirm FLCC's date with the college's accessibility office.
- **[FLAG] Certification access.** Check whether FLCC is a Certiport testing center and has practice-test licenses. If not, the capstone uses the Office Trainer's own MOS-style set and the exam is an optional add-on.
- **[NOTE] Privacy by design.** Fictional persona option for resumes, letters and peer review; merge lists deleted after grading; Document Inspector taught before any sharing; no personal data in AI prompts.
- **[NOTE] Northstar as a diagnostic.** Twelve of Northstar's 14 Word skills are covered in Week 1, lists in Week 2, and inserting images, shapes, hyperlinks and tables in Week 3. If FLCC has a Northstar site license, use it as a Day 1 pre-assessment and a Week 1 post-check.

## Grading (suggested)

| Component | Weight |
|---|---|
| Practice drills (completion, unlimited tries) | 15% |
| Weekly skill checks SC1-SC5 (4% each) | 20% |
| Projects P1-P4 (10% each) | 40% |
| Project P5 Peer Review Exchange | 10% |
| Capstone practice exam and portfolio | 15% |
| **Total** | **100%** |

## Weekly plan

### Week 1: The Professional Document

*Create, format, save and print documents people take seriously.*  
**Outline topics:** T1 Introduction to Microsoft Word application features, T2 Creating Documents with Microsoft Word, T3 Formatting and Organizing Text  
**CLOs:** CLO1 (Introduced, Assessed); CLO2 (Introduced, Assessed); CLO6 (Introduced)  
**Workload:** About 9 hours: 2 h lessons and videos, 3 h drills, 1 h skill check, 3 h project

**Why it matters**

- *College:* Every course expects clean, correctly formatted submissions saved where you can find them.
- *Work:* Letters and memos are still how organizations make requests and keep records.
- *Life:* Landlord letters, appeals, complaints and requests get better results when they look professional.

**Module objectives** - by the end of the week, learners can:

- **W1.1** Identify and use the parts of the Word window (Ribbon, Backstage view, status bar, Quick Access Toolbar) and switch among Print Layout, Read Mode and Web Layout views. *(CLO2, CLO6)*
- **W1.2** Create, name, save and reopen documents in a deliberate location (OneDrive or local), and save copies as .docx and .pdf. *(CLO1)*
- **W1.3** Format characters and paragraphs: font, size, color, effects, alignment, line and paragraph spacing, indentation, Format Painter and Clear Formatting. *(CLO1, CLO2)*
- **W1.4** Use Find and Replace, symbols and special characters, cut/copy/paste options, Undo and Editor (spelling and grammar) to produce error-free text. *(CLO2)*
- **W1.5** Set up pages (margins, orientation, paper size), set document properties and language, and use Print Preview and print settings. *(CLO1)*
- **W1.6** Show formatting marks to diagnose and fix spacing and layout problems. *(CLO2)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-110 | Other | Platform |
|---|---|---|---|---|
| D1.01 | Open the starter file and save it to your course folder as Lastname_W1_Letter.docx. | - | N1, N3, N4 | Web + desktop |
| D1.02 | Add Print Preview and Save As to the Quick Access Toolbar, then name three parts of the window in the drill quiz. | - | E1.1, N2 | Desktop only |
| D1.03 | Show formatting marks; remove the extra blank paragraphs and double spaces in the letter. | 1.1.4, 2.1.1 | - | Web + desktop |
| D1.04 | Use Find to count every mention of the deadline, then replace every Acme Outdoor Supply with Lakeshore Paddle Company. | 1.1.1, 2.1.1 | - | Web + desktop |
| D1.05 | Insert a copyright symbol in the footer and type cafe with an accented e. | 2.1.2 | - | Web + desktop |
| D1.06 | Set body text to 11 pt Aptos; bold the subject line and apply a text effect to the letterhead; copy the letterhead format to the memo heading with Format Painter. | 2.2.1, 2.2.2 | N6 | Web + desktop |
| D1.07 | Set single line spacing with 8 pt after; give the memo body a 0.5 in first-line indent; left-align the letter; clear formatting from the pasted paragraph. | 2.2.3, 2.2.5 | N7 | Web + desktop |
| D1.08 | Move the closing paragraph above the signature block with cut and paste (Keep Text Only); undo an accidental deletion. | - | E2.1, N9, N10 | Web + desktop |
| D1.09 | Set 1 in margins, Letter size and portrait orientation; run Editor and fix every flagged issue. | 1.2.1 | N5, N11 | Web + desktop |
| D1.10 | Fill in the Title, Author and Tags properties and set the proofing language to English (United States). | 1.3.2 | E1.3 | Desktop only |
| D1.11 | Export the letter to PDF; set print to 2 copies of page 1 only and capture the print preview. | 1.3.1, 1.3.3 | N12, N13 | Any (self-check) |
| D1.12 | Switch to Web Layout and Read Mode and describe in one sentence what changes (Word for the web has no Web Layout; compare Reading view instead). | - | - | Desktop only |

**Skill check SC1:** Timed, MOS-style: 10 tasks in a fresh starter file, auto-checked by the Office Trainer, 2 attempts. MO-110 1.3.1 is assessed through the Project 1 PDF instead, because a check cannot see files outside the document. Covers MO-110 1.1.1, 1.1.4, 1.2.1, 1.3.2, 2.1.1, 2.1.2, 2.2.1, 2.2.2, 2.2.3, 2.2.5.

**Project P1: Business Correspondence Set**

A block-format business letter (for example, a request to a financial aid office or an informational-interview request to a local employer) and a one-page memo. Submit .docx and .pdf.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P1-a | Block-letter structure is complete and correct (date, inside address, salutation, body, closing, signature block); memo heading is correct. | CLO1 | 25 |
| P1-b | Formatting is done with Word tools, not workarounds (no spaces for indents, no extra returns for spacing). | CLO1, CLO2 | 25 |
| P1-c | Text is error-free (Editor run, Find and Replace used where needed). | CLO2 | 25 |
| P1-d | Page setup, properties and language are set; PDF opens and matches the .docx. | CLO1 | 25 |
| | **Total** | | **100** |

**AI moment:** Optional: ask an AI tool to draft the letter's body. You make every formatting decision, check every fact, and add a one-line note on what you kept, changed or cut. There is no wrong way to try this.

**Platform:** Everything this week works in Word for the web and desktop Word. Quick Access Toolbar customizing and the proofing-language setting are easier in desktop Word. Print dialogs vary by browser.

**Standards:** NACE - Communication, Professionalism, Technology. DigComp 3.0 - Content creation, Information search, evaluation and management. WCAG 2.1 - 2.4.2, 3.1.1.

**Learner takeaway:** A formal letter and memo you can send with confidence.

### Week 2: Structure and the College Paper

*Let styles do the work: long documents, citations and a reusable paper template.*  
**Outline topics:** T3 Formatting and Organizing Text, T6 Working with Templates, Styles and Charts  
**CLOs:** CLO1 (Reinforced); CLO2 (Reinforced, Assessed); CLO6 (Reinforced)  
**Workload:** About 9 hours: 2 h lessons, 3 h drills, 1 h skill check, 3 h project

**Why it matters**

- *College:* A correct APA or MLA template saves time in every writing-heavy course you take.
- *Work:* Reports, manuals and proposals depend on styles, navigation and automatic tables of contents.
- *Life:* Long documents (a family history, a club handbook) stay manageable when the structure is real.

**Module objectives** - by the end of the week, learners can:

- **W2.1** Apply and modify built-in heading and body styles and switch style sets, so structure comes from styles rather than manual formatting. *(CLO2)*
- **W2.2** Navigate long documents with the Navigation pane, Go To, bookmarks and internal hyperlinks. *(CLO2, CLO6)*
- **W2.3** Create and control bulleted, numbered and multilevel lists (custom bullets, number formats, levels, restart and continue, start value). *(CLO1, CLO2)*
- **W2.4** Insert headers, footers and page numbers (including a different first page) and control layout with page and section breaks. *(CLO1, CLO2)*
- **W2.5** Insert and modify footnotes and endnotes, and generate and customize an automatic table of contents. *(CLO2)*
- **W2.6** Manage citation sources and generate a reference list in APA 7th or MLA 9th style, and save the setup as a reusable template. *(CLO2)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-110 | Other | Platform |
|---|---|---|---|---|
| D2.01 | Apply Heading 1 and Heading 2 to the report's section titles. | 2.2.4 | - | Web + desktop |
| D2.02 | Modify the Normal and Heading 1 styles (font, spacing, keep with next) and apply a different style set. | 1.2.2 | E2.3, E2.2 | Desktop only |
| D2.03 | Use the Navigation pane to move the Methods section above Results by dragging its heading; use Go To to jump to page 4. | 1.1.3 | - | Web + desktop |
| D2.04 | Add a bookmark at the Methods heading and a hyperlink to it from the introduction. | 1.1.2 | - | Web + desktop |
| D2.05 | Turn the plain list into a numbered list; demote two items to level 2; change the bullet character in the second list. | 3.3.1, 3.3.2, 3.3.4 | N8 | Web + desktop |
| D2.06 | Define a custom number format (Step 1, Step 2); restart numbering in the third list and set it to start at 5. | 3.3.3, 3.3.5, 3.3.6 | - | Desktop only |
| D2.07 | Insert a header with your last name and page number, using a different first page. | 1.2.3 | - | Web + desktop |
| D2.08 | Insert a Next Page section break and make only the appendix section landscape. | 2.3.2, 2.3.3 | - | Desktop only |
| D2.09 | Insert two footnotes and change the number format to i, ii, iii. | 4.1.1, 4.1.2 | - | Desktop only |
| D2.10 | Insert an automatic table of contents, add a new heading, then update the table and limit it to two levels. | 4.2.1, 4.2.2 | - | Web + desktop |
| D2.11 | Add three sources in the citation manager, insert in-text citations, and generate the references list in APA 7th. | - | E3.5 | Desktop only |

**Skill check SC2:** Timed, MOS-style: 10 tasks on a 6-page starter report, auto-checked, 2 attempts. Covers MO-110 1.1.2, 1.1.3, 1.2.2, 1.2.3, 2.2.4, 2.3.2, 2.3.3, 3.3.1, 3.3.2, 3.3.4, 3.3.5, 4.1.1, 4.1.2, 4.2.1, 4.2.2.

**Project P2: Research Paper Template and Paper**

Build a reusable APA 7th or MLA 9th template (.dotx) with modified styles, header and page numbers, then use it for a 3 to 5 page paper with at least three verified sources. Learners may reuse a paper from another course, which ties this skill to their real coursework.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P2-a | Headings use real heading styles; the Navigation pane shows the outline. | CLO2 | 20 |
| P2-b | Style-guide layout is correct (margins, spacing, header, page numbers, title page or heading). | CLO1 | 20 |
| P2-c | Citations and references are generated by Word and every source is real and checked. | CLO2 | 25 |
| P2-d | Lists, breaks and notes are used correctly. | CLO2 | 15 |
| P2-e | The .dotx template opens a new, correctly formatted document. | CLO2 | 20 |
| | **Total** | | **100** |

**AI moment:** Optional: AI tools sometimes invent sources that look real. If an AI tool suggests a source, find it in the library database before it goes into the citation manager. That habit protects your grade.

**Platform:** Styles, the Navigation pane, footnotes and tables of contents work in Word for the web. Section breaks, custom number formats and the citation manager need desktop Word, per Microsoft's feature comparison.

**Standards:** NACE - Communication, Critical Thinking, Technology. DigComp 3.0 - Information search, evaluation and management, Content creation. WCAG 2.1 - 1.3.1, 2.4.4, 2.4.6.

**Learner takeaway:** A paper template reused every semester, plus long-document skills for reports at work.

### Week 3: Visual Communication

*Graphics, tables, charts and desktop publishing that stay accessible.*  
**Outline topics:** T4 Using Graphics and Tables, T5 Special Document Formats, Columns and Mail Merge, T6 Working with Templates, Styles and Charts  
**CLOs:** CLO3 (Introduced, Assessed); CLO4 (Introduced, Assessed); CLO5 (Introduced, Assessed); CLO2 (Reinforced); CLO1 (Reinforced)  
**Workload:** About 9 hours: 2 h lessons, 3.5 h drills, 1 h skill check, 2.5 h project

**Why it matters**

- *College:* Lab reports, posters and club flyers mix text, tables and images.
- *Work:* Schedules, price lists, one-pagers and newsletters are everyday workplace documents.
- *Life:* Event flyers, small-business menus and community newsletters - and making them readable for everyone.

**Module objectives** - by the end of the week, learners can:

- **W3.1** Insert and format pictures, shapes, icons, screenshots, 3D models and SmartArt, including picture styles, artistic effects and background removal. *(CLO3)*
- **W3.2** Position graphics with text wrapping and alignment, add text to shapes and text boxes, and write meaningful alternative text. *(CLO3)*
- **W3.3** Create tables from scratch and from text; merge and split cells, resize, sort, set cell margins, split tables and repeat header rows. *(CLO4)*
- **W3.4** Insert and format a chart, and add captions to figures and tables. *(CLO3)*
- **W3.5** Apply desktop publishing features: multiple columns, column breaks, themes and custom theme colors, page color, borders, watermark and drop caps. *(CLO5)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-110 | Other | Platform |
|---|---|---|---|---|
| D3.01 | Insert the provided photo, apply a picture style and remove its background. | 5.1.2, 5.2.2, 5.2.3 | N14 | Desktop only |
| D3.02 | Apply an artistic effect to the second photo. | 5.2.1 | - | Desktop only |
| D3.03 | Insert an icon and a shape; type a label in the shape; change its fill and outline. | 5.1.1, 5.1.7, 5.3.2, 5.2.4 | N14 | Web + desktop |
| D3.04 | Insert a four-step SmartArt process, enter the steps, and change its colors and style. | 5.1.4, 5.3.3, 5.2.5 | - | Desktop only |
| D3.05 | Insert a text box pull quote; set Square wrapping and position it top right. | 5.1.6, 5.3.1, 5.4.1, 5.4.2 | - | Web + desktop |
| D3.06 | Insert a screen clipping of the course page and a 3D model; rotate the model. | 5.1.5, 5.1.3, 5.2.6 | - | Desktop only |
| D3.07 | Write alt text for every image and mark the divider line as decorative. | 5.4.3 | - | Web + desktop |
| D3.08 | Convert the tab-separated schedule to a table; convert the old price table back to text. | 3.1.1, 3.1.2 | - | Desktop only |
| D3.09 | Create a 4 x 6 table; merge the title row; resize columns; sort by date; set cell margins; repeat the header row; split the table after row 10. | 3.1.3, 3.2.1, 3.2.2, 3.2.3, 3.2.4, 3.2.5, 3.2.6 | N14 | Desktop only |
| D3.10 | Insert a column chart from the survey data and add the caption Figure 1. | - | E3.4 | Desktop only |
| D3.11 | Put the newsletter body in two columns with a column break; use a Continuous section break so the masthead stays one column. | 2.3.1, 2.3.2, 2.3.3 | - | Desktop only |
| D3.12 | Create custom theme colors for the club; add a page border, a DRAFT watermark and a drop cap. | 1.2.4 | E3.2 | Desktop only |

**Skill check SC3:** Timed, MOS-style: 10 tasks on a starter flyer, auto-checked, 2 attempts. Covers MO-110 1.2.4, 2.3.1, 3.1.1, 3.1.3, 3.2.1, 3.2.3, 3.2.6, 5.1.2, 5.1.4, 5.1.6, 5.2.2, 5.3.2, 5.4.1, 5.4.2, 5.4.3.

**Project P3: Newsletter or Event Flyer**

A two-page newsletter or event flyer for a real campus club, small business or community group: masthead, two-column body, at least three graphics with alt text, a SmartArt graphic, a schedule table with a header row, a captioned chart and a consistent theme. It must pass the Accessibility Checker with zero errors.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P3-a | Graphics are inserted, formatted and positioned on purpose, not by accident. | CLO3 | 20 |
| P3-b | The table has a real header row that repeats, is sorted and formatted, and does not use merged cells in ways that confuse reading order. | CLO4 | 20 |
| P3-c | Columns, sections, theme and page elements create a clear visual hierarchy. | CLO5 | 25 |
| P3-d | Accessibility: meaningful alt text, sufficient contrast, logical reading order, Accessibility Checker clean. | CLO3, CLO5 | 20 |
| P3-e | The chart supports the message and has a caption. | CLO3 | 15 |
| | **Total** | | **100** |

**AI moment:** Optional: AI image tools and automatic alt text are quick but not always right. If you use an AI image, say so, check the tool's terms, and rewrite any automatic alt text so it says what matters in context.

**Platform:** Pictures, shapes, icons, text boxes, simple tables, columns, page color, alt text and the Accessibility Checker work in Word for the web. SmartArt, charts, 3D models, watermarks and WordArt need desktop Word, per Microsoft's feature comparison.

**Standards:** NACE - Communication, Equity & Inclusion, Technology. DigComp 3.0 - Content creation. WCAG 2.1 - 1.1.1, 1.3.1, 1.3.2, 1.4.1, 1.4.3.

**Learner takeaway:** A flyer or newsletter for any event or group, built so everyone can read it.

### Week 4: The Job Application Packet

*Templates, mail merge and web output, built around getting hired.*  
**Outline topics:** T5 Special Document Formats, Columns and Mail Merge, T6 Working with Templates, Styles and Charts, T7 Creating Form Letters and Working in Groups (Collaboration)  
**CLOs:** CLO7 (Introduced, Assessed); CLO6 (Assessed); CLO1 (Reinforced, Assessed); CLO2 (Reinforced); CLO4 (Reinforced); CLO5 (Reinforced)  
**Workload:** About 9 hours: 2 h lessons, 3 h drills, 1 h skill check, 3 h project

**Why it matters**

- *College:* Internship, scholarship and transfer applications need a resume and tailored letters.
- *Work:* Mail merge powers customer letters, certificates, name badges and mailing labels in every office.
- *Life:* Job searches, holiday card labels and organization mailings - done once, reused for years.

**Module objectives** - by the end of the week, learners can:

- **W4.1** Customize a built-in or online template, save your own reusable template (.dotx) and store reusable content as a Quick Part. *(CLO2)*
- **W4.2** Build a resume in two versions: a designed version and a plain, applicant tracking system (ATS) friendly version. *(CLO1, CLO4, CLO5)*
- **W4.3** Create a recipient data source and run a mail merge for personalized form letters, with preview and an If...Then...Else rule. *(CLO7)*
- **W4.4** Use mail merge to produce envelopes, mailing labels and a directory. *(CLO7)*
- **W4.5** Save a document as a web page, preview it in Web Layout and a browser, and compare web page, PDF and share link as ways to publish online. *(CLO6)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-110 | Other | Platform |
|---|---|---|---|---|
| D4.01 | Open a resume template; replace every placeholder content control; save as Lastname_Resume.dotx. | - | E1.1, E4.1 | Desktop only |
| D4.02 | In the ATS version, align dates at the right margin with a right tab stop and dot leader. | - | - | Desktop only |
| D4.03 | Save your name and contact line as a Quick Part and insert it into the cover letter. | - | E3.1 | Desktop only |
| D4.04 | Create a recipient list of five employers (Company, Contact, Title, Address, City, State, ZIP, Position). | - | E4.3 | Desktop only |
| D4.05 | Insert an Address Block, a Greeting Line and the Position field in the letter body. | - | E4.3 | Desktop only |
| D4.06 | Add an If...Then...Else rule that changes one sentence for remote jobs; preview every record and fix any blank lines. | - | E4.3, E4.1 | Desktop only |
| D4.07 | Finish the merge to a new document with one letter per recipient. | - | E4.3 | Desktop only |
| D4.08 | Create envelopes and a sheet of Avery 5160 labels from the same list. | - | E4.3 | Desktop only |
| D4.09 | Create a directory merge that lists every employer and contact on one page. | - | E4.3 | Desktop only |
| D4.10 | Add a hyperlink to your portfolio or LinkedIn with descriptive link text; export the resume to PDF. | 1.3.1 | N14 | Web + desktop |
| D4.11 | Save the resume as Web Page, Filtered (.htm); view it in Web Layout and in a browser; note two things that changed. | 1.3.1 | - | Desktop only |

**Skill check SC4:** Timed merge check: given a data source and letter, produce merged letters, labels and a directory. Mostly MO-111 (Expert) skills. Covers MO-110 1.3.1.

**Project P4: Job Application Packet**

A designed resume and an ATS-friendly resume (.docx and .pdf), cover letters merged to at least three employers (real postings or a fictional persona), envelopes or labels, a directory of contacts, and a web page version of the resume with a short reflection comparing web page, PDF and share link.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P4-a | Resume content and layout are professional; a personal template is saved and reused. | CLO1, CLO2, CLO5 | 20 |
| P4-b | The ATS version uses simple structure (no text boxes or layout tables; key details not hidden in headers). | CLO1, CLO4 | 15 |
| P4-c | Mail merge is correct: data source, fields, rule, preview, no merge errors. | CLO7 | 25 |
| P4-d | Envelopes or labels and a directory are produced from the same data source. | CLO7 | 20 |
| P4-e | Web page version is created and previewed, and the reflection compares output formats. | CLO6 | 20 |
| | **Total** | | **100** |

**AI moment:** Optional: AI can tailor a cover letter to a job post quickly. Your judgment decides what is true about you - never claim a skill you do not have. Keep other people's personal data out of AI prompts.

**Privacy:** Learners may use a fictional persona instead of their own details. Recipient lists hold personal data: store them in the course folder only and delete them when the project is graded.

**Platform:** Mail merge, envelopes, labels, Quick Parts and Save as Web Page need desktop Word. Word for the web can export PDF and create share links.

**Standards:** NACE - Career & Self-Development, Professionalism, Technology. DigComp 3.0 - Content creation, Safety, wellbeing and responsible use. WCAG 2.1 - 2.4.4.

**Learner takeaway:** A complete job application packet, plus mail merge for any mailing.

### Week 5: Working Together and Shipping Clean

*Review, collaborate, then share safely. Capstone and certification practice.*  
**Outline topics:** T7 Creating Form Letters and Working in Groups (Collaboration)  
**CLOs:** CLO8 (Introduced, Assessed); CLO1 (Reinforced, Assessed); CLO2 (Reinforced, Assessed); CLO3 (Assessed); CLO4 (Assessed); CLO5 (Assessed); CLO6 (Reinforced); CLO7 (Reinforced)  
**Workload:** About 9 hours: 1.5 h lessons, 2 h drills, 1 h skill check, 2.5 h project, 2 h capstone

**Why it matters**

- *College:* Group projects and instructor feedback arrive as comments and tracked changes.
- *Work:* Contracts, policies and reports go through review cycles with several people editing at once.
- *Life:* Sharing a document safely - without hidden comments or personal data - protects you.

**Module objectives** - by the end of the week, learners can:

- **W5.1** Add, reply to, resolve and delete comments, including @mentions in shared documents. *(CLO8)*
- **W5.2** Track changes, review them, accept or reject them, and lock tracking with a password. *(CLO8)*
- **W5.3** Compare two versions of a document and combine edits from multiple reviewers. *(CLO8)*
- **W5.4** Share documents with appropriate permissions, co-author in real time and restore an earlier version from version history. *(CLO8)*
- **W5.5** Inspect a document before sharing: remove hidden data and personal information, fix accessibility and compatibility issues, and restrict editing. *(CLO1, CLO8)*
- **W5.6** Show exam readiness on a timed MO-110 style practice test and assemble a course portfolio. *(CLO1, CLO2, CLO3, CLO4, CLO5, CLO8)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-110 | Other | Platform |
|---|---|---|---|---|
| D5.01 | Add three comments; reply to one, resolve one and delete one. | 6.1.1, 6.1.2, 6.1.3, 6.1.4 | - | Web + desktop |
| D5.02 | Turn on Track Changes, make four edits, then lock tracking with a password. | 6.2.1, 6.2.4 | - | Desktop only |
| D5.03 | Review all tracked changes; accept the formatting changes and reject one deletion. | 6.2.2, 6.2.3 | - | Web + desktop |
| D5.04 | Compare the original and revised cover letter; combine two reviewers' copies into one. | - | E1.1 | Desktop only |
| D5.05 | Share a copy as view-only with an expiration date; co-edit with a partner; restore a previous version. | 1.3.4 | E1.1 | Any (self-check) |
| D5.06 | Run Inspect Document and remove comments, personal information and hidden text; run Check Compatibility. | 1.4.1, 1.4.3 | - | Desktop only |
| D5.07 | Run the Accessibility Checker and fix every error and warning. | 1.4.2 | - | Web + desktop |
| D5.08 | Restrict editing to comments only, protected with a password. | - | E1.2 | Desktop only |

**Skill check SC5:** Timed, MOS-style: 8 tasks on a reviewed document, auto-checked, 2 attempts. Covers MO-110 1.3.4, 1.4.1, 1.4.2, 1.4.3, 6.1.1, 6.1.2, 6.1.3, 6.1.4, 6.2.1, 6.2.2, 6.2.3, 6.2.4.

**Project P5: Peer Review Exchange and Clean-and-Ship**

Swap the cover letter or newsletter with a partner (or use the provided persona documents). Review with Track Changes and comments, respond to every comment, accept or reject each change, compare and combine two reviewers' copies, then run the clean-and-ship checklist: inspect, accessibility, compatibility, PDF export and a view-only share link. Finish with a short reflection memo.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P5-a | Feedback to the partner is specific and constructive, using comments and tracked edits. | CLO8 | 25 |
| P5-b | Every comment received gets a reply or resolution; every change is accepted or rejected on purpose. | CLO8 | 25 |
| P5-c | Compare and Combine are done correctly. | CLO8 | 20 |
| P5-d | Clean-and-ship checklist is complete: Document Inspector, Accessibility Checker, compatibility, permissions. | CLO8, CLO1 | 30 |
| | **Total** | | **100** |

**Capstone: Practice Exam and Portfolio**

A timed, 50-minute MO-110 style practice exam covering all six domains (Certiport practice tests if FLCC has access; otherwise the Office Trainer capstone set), plus a portfolio of the five projects with a short reflection on which skills you will use next and where. Evidence for CLO1, CLO2, CLO3, CLO4, CLO5, CLO6, CLO7, CLO8.

**AI moment:** Optional: Copilot and similar tools can summarize comments or suggest edits. Weigh those suggestions the way you would a human reviewer's, and keep Track Changes on so the history stays honest.

**Privacy:** Peer review shares real work. Learners choose which document to share and may use the persona documents instead. The Document Inspector step teaches what hidden data travels with a file.

**Platform:** Comments, Track Changes review, sharing, version history and the Accessibility Checker work in Word for the web. Compare and Combine, Document Inspector and password-protected tracking and restrictions need desktop Word (the web shows restricted documents read-only).

**Standards:** NACE - Teamwork, Communication, Leadership, Professionalism, Technology. DigComp 3.0 - Communication and collaboration, Safety, wellbeing and responsible use. WCAG 2.1 - 1.1.1, 1.3.1, 1.3.2, 1.4.3, 2.4.2, 3.1.1.

**Learner takeaway:** Safe sharing and feedback habits for any group document - plus a certification attempt if wanted.

## Alignment matrix

### Course learning outcomes by week

| CLO | Wk 1 | Wk 2 | Wk 3 | Wk 4 | Wk 5 |
|---|---|---|---|---|---|
| **CLO1** Create, save, format and print documents | I, A | R | R | R, A | R, A |
| **CLO2** Demonstrate features offered by the application | I, A | R, A | R | R | R, A |
| **CLO3** Insert and modify graphics |  |  | I, A |  | A |
| **CLO4** Create and format tables |  |  | I, A | R | A |
| **CLO5** Apply desktop publishing features to a document |  |  | I, A | R | A |
| **CLO6** Create and preview the document as a Web page | I | R |  | A | R |
| **CLO7** use mail merge to create form letters, directories, and envelopes |  |  |  | I, A | R |
| **CLO8** Collaborate on documents, Compare and Combine documents, Manage comments, Track document changes |  |  |  |  | I, A |

### Topic outline by week

| Topic | Wk 1 | Wk 2 | Wk 3 | Wk 4 | Wk 5 |
|---|---|---|---|---|---|
| **T1** Introduction to Microsoft Word application features | X |  |  |  |  |
| **T2** Creating Documents with Microsoft Word | X |  |  |  |  |
| **T3** Formatting and Organizing Text | X | X |  |  |  |
| **T4** Using Graphics and Tables |  |  | X |  |  |
| **T5** Special Document Formats, Columns and Mail Merge |  |  | X | X |  |
| **T6** Working with Templates, Styles and Charts |  | X | X | X |  |
| **T7** Creating Form Letters and Working in Groups (Collaboration) |  |  |  | X | X |

### Module objectives to course outcomes

| Objective | CLO1 | CLO2 | CLO3 | CLO4 | CLO5 | CLO6 | CLO7 | CLO8 |
|---|---|---|---|---|---|---|---|---|
| W1.1 |  | X |  |  |  | X |  |  |
| W1.2 | X |  |  |  |  |  |  |  |
| W1.3 | X | X |  |  |  |  |  |  |
| W1.4 |  | X |  |  |  |  |  |  |
| W1.5 | X |  |  |  |  |  |  |  |
| W1.6 |  | X |  |  |  |  |  |  |
| W2.1 |  | X |  |  |  |  |  |  |
| W2.2 |  | X |  |  |  | X |  |  |
| W2.3 | X | X |  |  |  |  |  |  |
| W2.4 | X | X |  |  |  |  |  |  |
| W2.5 |  | X |  |  |  |  |  |  |
| W2.6 |  | X |  |  |  |  |  |  |
| W3.1 |  |  | X |  |  |  |  |  |
| W3.2 |  |  | X |  |  |  |  |  |
| W3.3 |  |  |  | X |  |  |  |  |
| W3.4 |  |  | X |  |  |  |  |  |
| W3.5 |  |  |  |  | X |  |  |  |
| W4.1 |  | X |  |  |  |  |  |  |
| W4.2 | X |  |  | X | X |  |  |  |
| W4.3 |  |  |  |  |  |  | X |  |
| W4.4 |  |  |  |  |  |  | X |  |
| W4.5 |  |  |  |  |  | X |  |  |
| W5.1 |  |  |  |  |  |  |  | X |
| W5.2 |  |  |  |  |  |  |  | X |
| W5.3 |  |  |  |  |  |  |  | X |
| W5.4 |  |  |  |  |  |  |  | X |
| W5.5 | X |  |  |  |  |  |  | X |
| W5.6 | X | X | X | X | X |  |  | X |

## Evidence for each course learning outcome

### CLO1: Create, save, format and print documents

- **Bloom level:** Apply
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 3, Week 4, Week 5 · **Assessed:** Week 1, Week 4, Week 5
- **Taught through objectives:** W1.2, W1.3, W1.5, W2.3, W2.4, W4.2, W5.5, W5.6
- **Rubric criteria:** P1-a, P1-b, P1-d, P2-b, P4-a, P4-b, P5-d
- **Skill checks and capstone:** SC1, SC4, SC5, CAP
- **Evidence artifact:** Project 1 letter and memo (.docx and .pdf); Project 2 paper layout; Project 4 resumes; clean-and-ship checklist in Project 5.

### CLO2: Demonstrate features offered by the application

- **Bloom level:** Apply
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 3, Week 4, Week 5 · **Assessed:** Week 1, Week 2, Week 5
- **Taught through objectives:** W1.1, W1.3, W1.4, W1.6, W2.1, W2.2, W2.3, W2.4, W2.5, W2.6, W4.1, W5.6
- **Rubric criteria:** P1-b, P1-c, P2-a, P2-c, P2-d, P2-e, P4-a
- **Skill checks and capstone:** SC1, SC2, SC5, CAP
- **Evidence artifact:** Skill checks SC1 to SC5 (feature use under time); Project 2 styles, citations and template; capstone practice exam.

### CLO3: Insert and modify graphics

- **Bloom level:** Apply
- **Introduced:** Week 3 · **Reinforced:** - · **Assessed:** Week 3, Week 5
- **Taught through objectives:** W3.1, W3.2, W3.4, W5.6
- **Rubric criteria:** P3-a, P3-d, P3-e
- **Skill checks and capstone:** SC3, SC5, CAP
- **Evidence artifact:** Project 3 graphics, chart and alt text; SC3; capstone domain 5.

### CLO4: Create and format tables

- **Bloom level:** Apply
- **Introduced:** Week 3 · **Reinforced:** Week 4 · **Assessed:** Week 3, Week 5
- **Taught through objectives:** W3.3, W4.2, W5.6
- **Rubric criteria:** P3-b, P4-b
- **Skill checks and capstone:** SC3, SC5, CAP
- **Evidence artifact:** Project 3 schedule table; Project 4 ATS resume structure; SC3; capstone domain 3.

### CLO5: Apply desktop publishing features to a document

- **Bloom level:** Apply
- **Introduced:** Week 3 · **Reinforced:** Week 4 · **Assessed:** Week 3, Week 5
- **Taught through objectives:** W3.5, W4.2, W5.6
- **Rubric criteria:** P3-c, P3-d, P4-a
- **Skill checks and capstone:** SC3, SC5, CAP
- **Evidence artifact:** Project 3 newsletter layout (columns, sections, theme, page elements); Project 4 designed resume.

### CLO6: Create and preview the document as a Web page

- **Bloom level:** Apply
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 5 · **Assessed:** Week 4
- **Taught through objectives:** W1.1, W2.2, W4.5
- **Rubric criteria:** P4-e
- **Skill checks and capstone:** SC4, CAP
- **Evidence artifact:** Project 4 web page version of the resume (.htm) and the format-comparison reflection.

### CLO7: use mail merge to create form letters, directories, and envelopes

- **Bloom level:** Apply
- **Introduced:** Week 4 · **Reinforced:** Week 5 · **Assessed:** Week 4
- **Taught through objectives:** W4.3, W4.4
- **Rubric criteria:** P4-c, P4-d
- **Skill checks and capstone:** SC4, CAP
- **Evidence artifact:** Project 4 merged cover letters, envelopes or labels, and directory; SC4 merge check.

### CLO8: Collaborate on documents, Compare and Combine documents, Manage comments, Track document changes

- **Bloom level:** Apply / Evaluate
- **Introduced:** Week 5 · **Reinforced:** - · **Assessed:** Week 5
- **Taught through objectives:** W5.1, W5.2, W5.3, W5.4, W5.5, W5.6
- **Rubric criteria:** P5-a, P5-b, P5-c, P5-d
- **Skill checks and capstone:** SC5, CAP
- **Evidence artifact:** Project 5 tracked review, comment responses, compare and combine, and clean-and-ship; SC5.

### Topic outline evidence

- **T1 Introduction to Microsoft Word application features:** Week 1 tour of the Ribbon, Backstage view, status bar, Quick Access Toolbar and views (Print Layout, Read Mode, Web Layout).
- **T2 Creating Documents with Microsoft Word:** Week 1 create, name, save, reopen and export documents; Project 1 business letter and memo.
- **T3 Formatting and Organizing Text:** Week 1 character and paragraph formatting; Week 2 styles, lists, headers and footers, breaks and the Navigation pane.
- **T4 Using Graphics and Tables:** Week 3 pictures, shapes, icons, SmartArt, text boxes and tables; Project 3 newsletter.
- **T5 Special Document Formats, Columns and Mail Merge:** Week 3 columns and section-level layout; Week 4 envelopes, labels, directories, web page output and mail merge.
- **T6 Working with Templates, Styles and Charts:** Week 2 styles, style sets and a reusable paper template; Week 3 charts and captions; Week 4 resume template.
- **T7 Creating Form Letters and Working in Groups (Collaboration):** Week 4 merged cover letters (form letters); Week 5 comments, Track Changes, Compare and Combine, co-authoring.

## MOS crosswalk: MO-110 Word Associate (Microsoft 365 Apps)

All 71 sub-skills are practiced in at least one drill and appear in the capstone practice exam. Wording follows Microsoft's skills outline; check the current study guide before printing.

### Domain 1: Manage documents (20-25%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **1.1** | **Navigate within documents** | | | |
| 1.1.1 | Search for text | 1 | D1.04 | SC1 |
| 1.1.2 | Link to locations within documents | 2 | D2.04 | SC2 |
| 1.1.3 | Move to specific locations and objects in documents | 2 | D2.03 | SC2 |
| 1.1.4 | Show and hide formatting symbols and hidden text | 1 | D1.03 | SC1 |
| **1.2** | **Format documents** | | | |
| 1.2.1 | Set up document pages | 1 | D1.09 | SC1 |
| 1.2.2 | Apply style sets | 2 | D2.02 | SC2 |
| 1.2.3 | Insert and modify headers and footers | 2 | D2.07 | SC2 |
| 1.2.4 | Configure page background elements | 3 | D3.12 | SC3 |
| **1.3** | **Save and share documents** | | | |
| 1.3.1 | Save documents in alternative file formats | 1, 4 | D1.11, D4.10, D4.11 | SC4 |
| 1.3.2 | Modify basic document properties | 1 | D1.10 | SC1 |
| 1.3.3 | Modify print settings | 1 | D1.11 | Capstone |
| 1.3.4 | Share documents electronically | 5 | D5.05 | SC5 |
| **1.4** | **Inspect documents for issues** | | | |
| 1.4.1 | Locate and remove hidden properties and personal information | 5 | D5.06 | SC5 |
| 1.4.2 | Locate and correct accessibility issues | 5 | D5.07 | SC5 |
| 1.4.3 | Locate and correct compatibility issues | 5 | D5.06 | SC5 |

### Domain 2: Insert and format text, paragraphs, and sections (20-25%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **2.1** | **Insert text and paragraphs** | | | |
| 2.1.1 | Find and replace text | 1 | D1.03, D1.04 | SC1 |
| 2.1.2 | Insert symbols and special characters | 1 | D1.05 | SC1 |
| **2.2** | **Format text and paragraphs** | | | |
| 2.2.1 | Apply text effects | 1 | D1.06 | SC1 |
| 2.2.2 | Apply formatting by using Format Painter | 1 | D1.06 | SC1 |
| 2.2.3 | Set line and paragraph spacing and indentation | 1 | D1.07 | SC1 |
| 2.2.4 | Apply built-in styles to text | 2 | D2.01 | SC2 |
| 2.2.5 | Clear formatting | 1 | D1.07 | SC1 |
| **2.3** | **Create and configure document sections** | | | |
| 2.3.1 | Format text in multiple columns | 3 | D3.11 | SC3 |
| 2.3.2 | Insert page, section, and column breaks | 2, 3 | D2.08, D3.11 | SC2 |
| 2.3.3 | Change page setup options for a section | 2, 3 | D2.08, D3.11 | SC2 |

### Domain 3: Manage tables and lists (20-25%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **3.1** | **Create tables** | | | |
| 3.1.1 | Convert text to tables | 3 | D3.08 | SC3 |
| 3.1.2 | Convert tables to text | 3 | D3.08 | Capstone |
| 3.1.3 | Create tables by specifying rows and columns | 3 | D3.09 | SC3 |
| **3.2** | **Modify tables** | | | |
| 3.2.1 | Sort table data | 3 | D3.09 | SC3 |
| 3.2.2 | Configure cell margins and spacing | 3 | D3.09 | Capstone |
| 3.2.3 | Merge and split cells | 3 | D3.09 | SC3 |
| 3.2.4 | Resize tables, rows, and columns | 3 | D3.09 | Capstone |
| 3.2.5 | Split tables | 3 | D3.09 | Capstone |
| 3.2.6 | Configure a repeating row header | 3 | D3.09 | SC3 |
| **3.3** | **Create and modify lists** | | | |
| 3.3.1 | Format paragraphs as numbered and bulleted lists | 2 | D2.05 | SC2 |
| 3.3.2 | Change bullet characters and number formats | 2 | D2.05 | SC2 |
| 3.3.3 | Define custom bullet characters and number formats | 2 | D2.06 | Capstone |
| 3.3.4 | Increase and decrease list levels | 2 | D2.05 | SC2 |
| 3.3.5 | Restart and continue list numbering | 2 | D2.06 | SC2 |
| 3.3.6 | Set starting number values | 2 | D2.06 | Capstone |

### Domain 4: Create and manage references (5-10%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **4.1** | **Create and manage reference elements** | | | |
| 4.1.1 | Insert footnotes and endnotes | 2 | D2.09 | SC2 |
| 4.1.2 | Modify footnote and endnote properties | 2 | D2.09 | SC2 |
| **4.2** | **Create and manage reference tables** | | | |
| 4.2.1 | Insert tables of contents | 2 | D2.10 | SC2 |
| 4.2.2 | Customize tables of contents | 2 | D2.10 | SC2 |

### Domain 5: Insert and format graphic elements (15-20%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **5.1** | **Insert illustrations and text boxes** | | | |
| 5.1.1 | Insert shapes | 3 | D3.03 | Capstone |
| 5.1.2 | Insert pictures | 3 | D3.01 | SC3 |
| 5.1.3 | Insert 3D models | 3 | D3.06 | Capstone |
| 5.1.4 | Insert SmartArt graphics | 3 | D3.04 | SC3 |
| 5.1.5 | Insert screenshots and screen clippings | 3 | D3.06 | Capstone |
| 5.1.6 | Insert text boxes | 3 | D3.05 | SC3 |
| 5.1.7 | Insert icons | 3 | D3.03 | Capstone |
| **5.2** | **Format illustrations and text boxes** | | | |
| 5.2.1 | Apply artistic effects | 3 | D3.02 | Capstone |
| 5.2.2 | Apply picture effects and picture styles | 3 | D3.01 | SC3 |
| 5.2.3 | Remove picture backgrounds | 3 | D3.01 | Capstone |
| 5.2.4 | Format graphic elements | 3 | D3.03 | Capstone |
| 5.2.5 | Format SmartArt graphics | 3 | D3.04 | Capstone |
| 5.2.6 | Format 3D models | 3 | D3.06 | Capstone |
| **5.3** | **Add text to graphic elements** | | | |
| 5.3.1 | Add and modify text in text boxes | 3 | D3.05 | Capstone |
| 5.3.2 | Add and modify text in shapes | 3 | D3.03 | SC3 |
| 5.3.3 | Add and modify SmartArt graphic content | 3 | D3.04 | Capstone |
| **5.4** | **Modify graphic elements** | | | |
| 5.4.1 | Position objects | 3 | D3.05 | SC3 |
| 5.4.2 | Wrap text around objects | 3 | D3.05 | SC3 |
| 5.4.3 | Add alternative text to objects for accessibility | 3 | D3.07 | SC3 |

### Domain 6: Manage document collaboration (5-10%)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **6.1** | **Add and manage comments** | | | |
| 6.1.1 | Add comments | 5 | D5.01 | SC5 |
| 6.1.2 | Review and reply to comments | 5 | D5.01 | SC5 |
| 6.1.3 | Resolve comments | 5 | D5.01 | SC5 |
| 6.1.4 | Delete comments | 5 | D5.01 | SC5 |
| **6.2** | **Manage change tracking** | | | |
| 6.2.1 | Track changes | 5 | D5.02 | SC5 |
| 6.2.2 | Review tracked changes | 5 | D5.03 | SC5 |
| 6.2.3 | Accept and reject tracked changes | 5 | D5.03 | SC5 |
| 6.2.4 | Lock and unlock change tracking | 5 | D5.02 | SC5 |

## MOS crosswalk: MO-111 Word Expert (stretch content)

This course does not prepare learners for MO-111. It covers the Expert objectives that the official course outcomes require.

| Code | Objective | Status | Weeks | How |
|---|---|---|---|---|
| E1.1 | Manage documents and templates | Partly covered | 1, 4, 5 | Templates (W4), Quick Access Toolbar (W1), compare and combine plus version history (W5). Macros, external links and Normal template defaults are not taught. |
| E1.2 | Prepare documents for collaboration | Covered | 5 | Restrict editing and password protection in the W5 clean-and-ship checklist. |
| E1.3 | Use and configure language options | Partly covered | 1 | Set the document proofing language (also an accessibility requirement). Display language is not taught. |
| E2.1 | Find, replace, and paste document content | Partly covered | 1 | Paste options and special-character search (W1). Wildcards are optional enrichment. |
| E2.2 | Configure paragraph layout options | Partly covered | 2 | Keep with next and widow/orphan control for headings (W2). Hyphenation and line numbers are not taught. |
| E2.3 | Create and manage styles | Covered | 2 | Modify and create styles for the paper template (W2). |
| E3.1 | Create and modify building blocks | Covered | 4 | Save the contact heading as a Quick Part and reuse it (W4). |
| E3.2 | Create custom design elements | Partly covered | 3 | Custom theme colors for a club brand (W3). Custom font sets and saved themes are optional. |
| E3.3 | Create and manage indexes | Not covered | - | Not taught. Low value for most learners; optional enrichment. |
| E3.4 | Create and manage tables of figures | Partly covered | 3 | Captions on the chart and table (W3). Table of figures is optional. |
| E3.5 | Create and manage bibliographies | Covered | 2 | Citation manager and generated references list in Project 2 (W2). |
| E4.1 | Manage forms, fields, and controls | Partly covered | 4 | Template content controls and the If...Then...Else merge field (W4). |
| E4.2 | Create and modify macros | Not covered | - | Not taught. Macros are a security risk to explain, not a beginner skill to build. |
| E4.3 | Perform mail merges | Covered | 4 | All of Project 4's merge work (W4). |

## Other standards and conventions

### Northstar Digital Literacy: Microsoft Word

| Code | Standard | Week | Drill |
|---|---|---|---|
| N1 | Open a new or existing document. | 1 | D1.01 |
| N2 | Identify the parts of the Word window, including the Ribbon, Status Bar and Quick Access Toolbar. | 1 | D1.02 |
| N3 | Save a document, being intentional about name and location. | 1 | D1.01 |
| N4 | Identify file extensions that can be opened by Microsoft Word. | 1 | D1.01 |
| N5 | Use Spelling and Grammar check. | 1 | D1.09 |
| N6 | Format text: size, color and font type. | 1 | D1.06 |
| N7 | Set text spacing and alignment. | 1 | D1.07 |
| N8 | Apply bullets and automatic numbering. | 2 | D2.05 |
| N9 | Undo the previous action. | 1 | D1.08 |
| N10 | Cut, copy and paste. | 1 | D1.08 |
| N11 | Modify page layout, including margins and orientation. | 1 | D1.09 |
| N12 | Print. | 1 | D1.11 |
| N13 | Close a document. | 1 | D1.11 |
| N14 | Insert objects into a document, including images, shapes, hyperlinks, and tables. | 3 | D3.01, D3.03, D3.09, D4.10 |

### NACE Career Readiness Competencies

| Competency | Weeks | Where it shows up |
|---|---|---|
| Career & Self-Development | 4, 5 | Resume, cover letters and a portfolio reflection on next steps. |
| Communication | 1, 2, 3, 5 | Letters, memos, papers, newsletters and written peer feedback. |
| Critical Thinking | 2, 5 | Checking sources (including AI-suggested ones) and deciding which review edits to accept. |
| Equity & Inclusion | 3, 5 | Accessible documents: alt text, contrast, reading order and the Accessibility Checker. NACE lists this competency as under review; some campuses now label it Access & Opportunity. |
| Leadership | 5 | Giving constructive feedback and resolving a review cycle. |
| Professionalism | 1, 4, 5 | Business conventions, error-free work and safe sharing habits. |
| Teamwork | 5 | Co-authoring, comments, Track Changes and combining reviewers' work. |
| Technology | 1, 2, 3, 4, 5 | Every week: choosing the right Word feature and using it responsibly. |

### DigComp 3.0 (European Digital Competence Framework, December 2025)

| Competence area | Weeks | Where it shows up |
|---|---|---|
| Information search, evaluation and management | 1, 2 | Saving and organizing files; finding and verifying sources for citations. |
| Communication and collaboration | 5 | Sharing, co-authoring, comments and Track Changes. |
| Content creation | 1, 2, 3, 4 | Creating and formatting letters, papers, newsletters, resumes and web output. |
| Safety, wellbeing and responsible use | 4, 5 | Protecting personal data in merge lists and AI prompts; Document Inspector; permissions. |
| Problem identification and solving | 1, 3, 5 | Using formatting marks, the Accessibility Checker and Compatibility Checker to find and fix problems. |

### Accessibility: WCAG 2.1 AA success criteria practiced (ADA Title II)

| Criterion | Name | Level | In Word | Weeks |
|---|---|---|---|---|
| 1.1.1 | Non-text Content | A | Alt text on pictures, charts and SmartArt; decorative items marked decorative. | 3, 5 |
| 1.3.1 | Info and Relationships | A | Real heading styles, real lists and table header rows instead of visual formatting only. | 2, 3, 5 |
| 1.3.2 | Meaningful Sequence | A | Reading order stays logical when graphics float or text runs in columns. | 3, 5 |
| 1.4.1 | Use of Color | A | Color is never the only way information is shown. | 3 |
| 1.4.3 | Contrast (Minimum) | AA | Text contrast at least 4.5:1 (3:1 for large text) on flyers and newsletters. | 3, 5 |
| 2.4.2 | Page Titled | A | The document Title property is set (it becomes the PDF title). | 1, 5 |
| 2.4.4 | Link Purpose (In Context) | A | Hyperlinks use descriptive text, not raw URLs or 'click here'. | 2, 4 |
| 2.4.6 | Headings and Labels | AA | Headings describe the content that follows. | 2 |
| 3.1.1 | Language of Page | A | Document proofing language is set correctly. | 1, 5 |

### Word for the web vs desktop Word

From Microsoft's feature comparison table. Microsoft updates Word for the web often, so recheck each term.

| Feature | Word for the web | Desktop Word | Used in |
|---|---|---|---|
| Mail merge | No | Yes | W4 |
| Envelopes and labels | No | Yes | W4 |
| Compare and combine revisions | No | Yes | W5 |
| Citations and bibliography | No | Yes | W2 |
| Section breaks | No | Yes | W2, W3 |
| SmartArt, charts and 3D models | No | Yes | W3 |
| Macros | No | Yes | Not taught |
| Watermark | No (page color only) | Yes | W3 |
| WordArt | No (shapes and text boxes only) | Yes | W3 |
| Restricted editing | Read only | Yes | W5 |
| Track Changes | Yes | Yes | W5 |
| Styles | Yes | Yes | W2 |
| Table of contents | Yes | Yes | W2 |
| Footnotes and endnotes | Yes | Yes | W2 |
| Columns | Yes | Yes | W3 |
| Templates | Yes | Yes | W2, W4 |
| Export as PDF | Yes | Yes | W1, W4, W5 |
| Accessibility Checker | Yes | Yes | W3, W5 |

Drill platform split: 21 work on the web, 31 need desktop Word, 2 work anywhere with a self-check.

### Conventions and assumptions

| Item | Type | Note | Weeks |
|---|---|---|---|
| APA 7th edition / MLA 9th edition | Published style guide | Current editions for student papers; confirm with the FLCC Library before each term. | W2 |
| Block-format business letter | Industry convention | The most common US business letter layout: everything flush left, single spaced, a blank line between parts. | W1 |
| ATS-friendly resume | Industry convention | Applicant tracking systems parse simple layouts most reliably; keep a plain version alongside the designed one. | W4 |
| Avery 5160 labels | Industry convention | The most common US address label size, built into Word's label options. | W4 |
| Bloom's revised taxonomy | Assumption | Course outcome verbs sit mostly at Apply; peer review reaches Evaluate. | All |
| Credit hour (about 45 hours per credit) | Assumption | The 9 hours per week figure assumes 1 credit over 5 weeks. Scale if the course carries more credit. | All |
| I / R / A curriculum mapping | Common practice | Introduced, Reinforced, Assessed - the standard way to show where an outcome is taught and measured. | All |

## Sources

- [MO-110: Microsoft Word (Microsoft 365 Apps) - Microsoft Learn](https://learn.microsoft.com/en-us/credentials/exams/mo-110/)
- [MO-110 study guide (skills measured) - Microsoft](https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE5dvir)
- [MO-111 Word Expert study guide - Microsoft](https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RW16Vea)
- [Northstar Digital Literacy: Microsoft Word standards](https://www.digitalliteracyassessment.org/badges/ms-word/v2/proctored/)
- [NACE Career Readiness: Competencies Defined - NACE](https://www.naceweb.org/career-readiness/competencies/career-readiness-defined)
- [NACE Career Readiness Competencies - Rutgers-Newark summary](https://careers.newark.rutgers.edu/resources/nace-career-readiness-competencies/)
- [Equity & Inclusion competency under review - Inside Higher Ed](https://www.insidehighered.com/index%2Ephp/node/196261)
- [DigComp 3.0 launches (December 2025) - CEPIS](https://cepis.org/digcomp-3-0-launches/)
- [DigComp framework - European Commission Joint Research Centre](https://joint-research-centre.ec.europa.eu/digcomp/digcomp-framework_en)
- [Understanding WCAG 2.1 success criteria - W3C](https://www.w3.org/WAI/WCAG21/Understanding/)
- [ADA Title II compliance deadline extension (2026) - Kwall](https://www.kwallcompany.com/2026/06/04/ada-title-ii-website-compliance-deadline/)
- [ADA Title II 2026 update - Aberdeen](https://aberdeen.io/blog/2026/04/07/ada-title-iis-2026-update-what-changed-and-what-didnt/)
- [Word features comparison: web vs desktop - Microsoft Support](https://support.microsoft.com/en-us/Word/word-features-comparison-word-for-the-web-vs-desktop)
