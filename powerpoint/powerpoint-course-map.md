# Microsoft PowerPoint - 5-Week Course Map

**Finger Lakes Community College (FLCC)** · Draft 1 · October 22, 2026 · Dave Ghidiu (with Claude)

- **Format:** 5 weeks, online or hybrid, about 9 hours of learner work per week (CSC 136 PowerPoint is 1 credit)
- **Certification target:** Microsoft Office Specialist (MOS): PowerPoint Associate (Microsoft 365 Apps), exam MO-310
- **Coverage:** 7/7 course learning outcomes assessed · 9/9 outline topics · 73/73 MO-310 sub-skills practiced (54 also in weekly skill checks; 22 sampled across every domain in the capstone) · no Expert exam for PowerPoint
- **Companion:** `powerpoint-course-map.html` shows this map interactively. This Markdown file is the master copy. Both files are generated from the same data (`_build/data.py`), so they always agree; after changing the map, regenerate both with `python3 build.py`.

## How to read this map

- **CLO** = course learning outcome from the official course description (verbatim). **T** = topic outline item (verbatim).
- **I / R / A** = Introduced, Reinforced, Assessed (standard curriculum mapping).
- **Drills** are short, outcome-checked practice tasks for the Office Trainer (it checks the result, not the clicks). Each is tagged *Web + desktop* or *Desktop only*.
- **MO-310 codes** (for example 2.2.3) follow Microsoft's skills outline. **N** codes are Northstar Digital Literacy PowerPoint standards.

## Course at a glance

| Week | Theme | Project | CLOs (I/R/A) | Outline topics | MO-310 domains drilled |
|---|---|---|---|---|---|
| 1 | **The Clear Presentation** - Build, organize and print a presentation people can follow. | P1 Explainer Presentation | CLO1 IA; CLO2 IA; CLO3 I | T1, T2, T3, T9 | 1, 2, 3 |
| 2 | **Design That Works** - Themes, backgrounds, the slide master and templates: set the look once and reuse it. | P2 Brand Template and Deck | CLO1 R; CLO2 RA; CLO3 RA | T2, T3, T6 | 1, 2 |
| 3 | **Visual Slides** - Pictures, shapes, tables, charts and SmartArt that explain at a glance, with alt text. | P3 Data Story | CLO4 IA; CLO5 IA; CLO2 R; CLO3 R | T5, T6 | 3, 4 |
| 4 | **Motion, Media and Links** - Transitions, animations, video and navigation that help the story instead of distracting from it. | P4 Interactive Tour | CLO6 IA; CLO4 RA; CLO2 R; CLO3 R | T4, T5, T7 | 2, 3, 4, 5 |
| 5 | **Present and Collaborate** - Review with others, compare and combine their changes, then rehearse, deliver and share safely. Capstone and certification practice. | P5 Team Presentation | CLO7 IA; CLO3 A; CLO1 RA; CLO2 RA; CLO4 A; CLO5 A; CLO6 A | T8, T9 | 1 |

## Decisions and flags

- **[NOTE] CLO 7 can be taught as written.** Unlike Excel, PowerPoint still has a working Review > Compare that merges a reviewer's copy into yours (desktop PowerPoint for Windows and Mac). Week 5 teaches it alongside co-authoring and comments, which is how most teams work today.
- **[FLAG] Make CLO 2 measurable.** "Demonstrate features offered by the application" is hard to assess on its own. It is mapped to every week here. Consider rewording to "Use PowerPoint's design, media and delivery features to build and present slide shows."
- **[NOTE] CLO 3 is the heart of the course.** "Create a professional presentation slide show" is assessed in every project's rubric (organization, design, delivery), not just with auto-checks, because professional is a judgment.
- **[FLAG] Exam number.** Microsoft's current PowerPoint Associate exam for Microsoft 365 Apps is MO-310 (MO-300 was the Office 2019 version). There is no PowerPoint Expert exam. Check which exam FLCC's Certiport setup offers.
- **[FLAG] Plan for desktop PowerPoint.** Week 2 (slide master) and Weeks 4 and 5 include features that need desktop PowerPoint (layouts, media trimming, Zoom, Compare, custom shows, rehearsed timings, Inspect Document, video export). Chromebook users will need campus labs or a virtual desktop for those tasks. Each drill is tagged web or desktop.
- **[NOTE] Motion with restraint.** Animation (CLO 6) is taught with accessibility in mind: no flashing, sensible timing, and a way to pause self-running shows (WCAG 2.2.2, 2.3.1).
- **[NOTE] Accessibility is built in, not bolted on.** A unique title on every slide, real layouts and placeholders, alt text, reading order, contrast and captions are taught as normal practice, and the accessibility checker is part of the final check (ADA Title II, WCAG 2.1 AA).
- **[NOTE] Privacy by design.** Made-up or permitted topics, photos and data; no classmates' details on slides; screen recordings of your own screen only; Inspect Document before sharing.
- **[NOTE] Northstar as a diagnostic.** All 17 Northstar PowerPoint standards are covered, 11 of them in Week 1. If FLCC has a Northstar site license, use it as a Day 1 pre-assessment and a Week 1 post-check.

## Grading (suggested)

| Component | Weight |
|---|---|
| Practice drills (completion, unlimited tries) | 15% |
| Weekly skill checks SC1-SC5 (4% each) | 20% |
| Projects P1-P4 (10% each) | 40% |
| Project P5 Team Presentation | 10% |
| Capstone practice exam and portfolio | 15% |
| **Total** | **100%** |

## Weekly plan

### Week 1: The Clear Presentation

*Build, organize and print a presentation people can follow.*  
**Outline topics:** T1 Introduction to Microsoft PowerPoint application features, T2 Creating presentation slide shows, T3 Formatting and Organizing presentations, T9 Presenting and viewing a slide show  
**CLOs:** CLO1 (Introduced, Assessed); CLO2 (Introduced, Assessed); CLO3 (Introduced)  
**Workload:** About 9 hours: 2 h lessons and videos, 3 h drills, 1 h skill check, 3 h project

**Why it matters**

- *College:* Class presentations, group projects and capstone defenses are graded on clarity as much as content.
- *Work:* Team updates, trainings and pitches are built from slides someone else can pick up and reuse.
- *Life:* Club meetings, family events and community groups run on a few clear slides.

**Module objectives** - by the end of the week, learners can:

- **W1.1** Find your way around PowerPoint: the Ribbon, thumbnails, notes, views, zoom and the Quick Access Toolbar, and save presentations with clear names. *(CLO1, CLO2)*
- **W1.2** Insert slides with the right layouts, duplicate and reuse slides, and import a Word outline. *(CLO1, CLO2)*
- **W1.3** Organize slides into named sections, reorder them and hide slides you don't need. *(CLO2, CLO3)*
- **W1.4** Format text, columns and bulleted and numbered lists. *(CLO1)*
- **W1.5** Change the slide size, add slide numbers and footers, write speaker notes, set properties, and print or export slides, notes and handouts. *(CLO1)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-310 | Other | Platform |
|---|---|---|---|---|
| D1.01 | Open the starter presentation and save it to your course folder as Lastname_P1_Welcome.pptx. | - | N1 | Web + desktop |
| D1.02 | Switch between Normal, Outline, Slide Sorter, Notes Page and Reading views, zoom the slide, and leave the presentation in Slide Sorter view. | 1.2.2 | N2, N6, N9 | Web + desktop |
| D1.03 | Add New Slide and Start From Beginning to the Quick Access Toolbar. | - | N11 | Desktop only |
| D1.04 | Insert a Title and Content slide and a Two Content slide, change one slide's layout, and duplicate a slide. | 2.1.3, 2.1.5 | N3 | Web + desktop |
| D1.05 | Reuse two slides from another presentation, then build slides from a Word outline. | 2.1.1, 2.1.2 | - | Desktop only |
| D1.06 | Group the slides into named sections, rename one, move a section, reorder two slides and hide the old schedule slide. | 2.2.1, 2.3.1, 2.3.2, 2.3.3 | N10 | Web + desktop |
| D1.07 | Format the title text, apply a WordArt style, and set the gear list in two columns. | 3.1.1, 3.1.2 | N4 | Desktop only |
| D1.08 | Turn plain lines into a bulleted list with two levels and the steps into a numbered list. | 3.1.3 | - | Web + desktop |
| D1.09 | Change the slide size from 4:3 to Widescreen (16:9). | 1.2.1 | - | Web + desktop |
| D1.10 | Add slide numbers and a footer to every slide except the title slide. | 2.2.3 | - | Web + desktop |
| D1.11 | Write speaker notes, set the Title property, then print notes pages and a 3-per-page handout and export a PDF. | 1.2.3, 1.3.1, 1.3.2, 1.3.3, 1.5.6 | N15, N16, N17 | Any (self-check) |

**Skill check SC1:** Timed, MOS-style: 13 tasks on a club orientation deck, auto-checked by the Office Trainer, 2 attempts. Covers MO-310 1.2.1, 1.2.3, 2.1.3, 2.1.5, 2.2.1, 2.2.3, 2.3.1, 2.3.2, 2.3.3, 3.1.1, 3.1.2, 3.1.3.

**Project P1: Explainer Presentation**

An 8 to 10 slide presentation that teaches something you know well (a skill, a place, a process), built from a Word outline or your own: sections, a clear title on every slide, lists, slide numbers, speaker notes, and a PDF handout.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P1-a | The presentation is organized: sections in a logical order, one idea per slide, a clear title on every slide. | CLO3 | 25 |
| P1-b | Text is formatted with PowerPoint's tools on purpose: layouts and placeholders, lists and levels, readable sizes (no text boxes faking layouts). | CLO1, CLO2 | 25 |
| P1-c | Speaker notes carry the detail, so the slides stay short. | CLO3 | 20 |
| P1-d | Ready to share: slide numbers and footer, Title property, and a PDF handout that matches the slides. | CLO1 | 30 |
| | **Total** | | **100** |

**AI moment:** Optional: ask an AI tool to turn your topic into an outline, then cut it down to what you can explain yourself. Check every fact it adds.

**Privacy:** Use your own topic or a made-up one. Don't put classmates' names, photos or personal details on slides.

**Platform:** Most of this week works in PowerPoint for the web. Outline and Notes Page views, the Quick Access Toolbar, Reuse Slides, Slides from Outline, WordArt styles and text columns need desktop PowerPoint.

**Standards:** NACE - Communication, Technology, Professionalism. DigComp 3.0 - Content creation, Communication and collaboration. WCAG 2.1 - 2.4.2, 2.4.6, 1.3.1.

**Learner takeaway:** A reusable, clearly organized presentation and the habit of putting details in the notes.

### Week 2: Design That Works

*Themes, backgrounds, the slide master and templates: set the look once and reuse it.*  
**Outline topics:** T2 Creating presentation slide shows, T3 Formatting and Organizing presentations, T6 Working with Templates, Styles and Charts  
**CLOs:** CLO1 (Reinforced); CLO2 (Reinforced, Assessed); CLO3 (Reinforced, Assessed)  
**Workload:** About 9 hours: 2 h lessons, 2.5 h drills, 1 h skill check, 3.5 h project

**Why it matters**

- *College:* A consistent look makes a group project feel like one presentation, not five.
- *Work:* Companies hand out a branded template; knowing the slide master is how you use it without breaking it.
- *Life:* A good template turns every future slideshow (a toast, a fundraiser) into a fill-in-the-blanks job.

**Module objectives** - by the end of the week, learners can:

- **W2.1** Apply and vary a theme, and change the background of one slide. *(CLO1, CLO2)*
- **W2.2** Edit the slide master: its theme, background and content, and create and modify layouts. *(CLO2)*
- **W2.3** Modify the handout master and the notes master. *(CLO2)*
- **W2.4** Save and use a presentation template (.potx). *(CLO2, CLO3)*
- **W2.5** Clean up a cluttered deck with design principles: contrast, consistent fonts, less text per slide. *(CLO3)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-310 | Other | Platform |
|---|---|---|---|---|
| D2.01 | Apply a theme, pick one of its variants, and change the theme fonts. | - | N5 | Web + desktop |
| D2.02 | Give the title slide a picture background and one other slide a solid fill, without changing the rest. | 2.2.2 | - | Web + desktop |
| D2.03 | In the slide master, change the background, add the logo to every slide and set the title font. | 1.1.1, 1.1.2 | - | Desktop only |
| D2.04 | Create a Quote layout with a picture placeholder, rename it, and change an existing layout's title position. | 1.1.3, 1.1.4 | - | Desktop only |
| D2.05 | Set the handout master to 3 slides per page with a header, and add the date and footer to the notes master. | 1.1.5, 1.1.6 | - | Desktop only |
| D2.06 | Save the branded deck as a template, Lastname_Lakeshore.potx, and start a new presentation from it. | - | - | Desktop only |
| D2.07 | Fix a cluttered deck: replace a novelty font everywhere, fix low-contrast text, and split a wall-of-text slide in two. | - | N13 | Web + desktop |

**Skill check SC2:** Timed, MOS-style: 10 tasks on a company kickoff deck (themes, backgrounds, master, layouts, handout and notes masters), auto-checked, 2 attempts. Most tasks need desktop PowerPoint. Covers MO-310 1.1.1, 1.1.2, 1.1.3, 1.1.4, 1.1.5, 1.1.6, 2.2.2.

**Project P2: Brand Template and Deck**

A template (.potx) for a real or made-up organization: a slide master with a logo and theme, at least two custom layouts, handout and notes masters set up, and a 6-slide presentation built from it that follows the design checklist.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P2-a | The slide master does the work: theme, fonts, background and logo set once, not slide by slide. | CLO2 | 30 |
| P2-b | Custom layouts are named, useful and used in the deck. | CLO2 | 20 |
| P2-c | The deck follows the design checklist: contrast, consistent fonts, short text, one idea per slide. | CLO3 | 30 |
| P2-d | Handout and notes masters are set up, and the template opens as a new presentation. | CLO1 | 20 |
| | **Total** | | **100** |

**AI moment:** Optional: ask an AI tool for a color palette and font pair for your organization, then check the contrast yourself before you use it.

**Privacy:** Use a made-up organization or one you have permission to represent. Don't copy a real company's logo for a fake brand.

**Platform:** Themes, variants and backgrounds work in PowerPoint for the web. The slide master, layouts, handout and notes masters and saving templates need desktop PowerPoint.

**Standards:** NACE - Communication, Professionalism, Technology. DigComp 3.0 - Content creation. WCAG 2.1 - 1.4.3, 1.3.1.

**Learner takeaway:** A template you can reuse, and an eye for slides that are easy to read.

### Week 3: Visual Slides

*Pictures, shapes, tables, charts and SmartArt that explain at a glance, with alt text.*  
**Outline topics:** T5 Using graphics, videos and SmartArt, T6 Working with Templates, Styles and Charts  
**CLOs:** CLO4 (Introduced, Assessed); CLO5 (Introduced, Assessed); CLO2 (Reinforced); CLO3 (Reinforced)  
**Workload:** About 9 hours: 2.5 h lessons, 3 h drills, 1 h skill check, 2.5 h project

**Why it matters**

- *College:* A chart or a diagram often explains a result faster than a paragraph on a slide.
- *Work:* Status updates, budgets and process changes are shown as tables, charts and diagrams.
- *Life:* Photos and simple diagrams make event and family slideshows worth watching.

**Module objectives** - by the end of the week, learners can:

- **W3.1** Insert, resize, crop and style pictures and screenshots, and give every graphic alt text. *(CLO4)*
- **W3.2** Insert and format shapes, text boxes and ink, and order, align and group slide content. *(CLO4, CLO2)*
- **W3.3** Create and format tables: rows, columns and table styles. *(CLO5)*
- **W3.4** Create and modify charts, and build SmartArt from scratch or from a list. *(CLO2, CLO3)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-310 | Other | Platform |
|---|---|---|---|---|
| D3.01 | Insert the kayak photo, resize it without stretching, and crop it to a square. | 3.3.1 | N7 | Web + desktop |
| D3.02 | Apply a picture style and an effect, give the photos alt text, and insert a screen clipping. | 3.3.2, 3.3.3, 3.4.7 | - | Desktop only |
| D3.03 | Insert a callout shape and a text box, add text, size them exactly, and apply a shape style. | 3.4.1, 3.4.3, 3.4.4, 3.4.5, 3.4.6 | N8 | Web + desktop |
| D3.04 | Bring the logo to the front, align three icons and distribute them, group them, turn on guides, and draw an ink circle. | 3.4.2, 3.5.1, 3.5.2, 3.5.3, 3.5.4 | - | Web + desktop |
| D3.05 | Insert a 4-column price table, add a row, delete a column and apply a table style. | 4.1.1, 4.1.2, 4.1.3 | - | Web + desktop |
| D3.06 | Insert a clustered column chart of season rentals, change its type, and add a title and data labels. | 4.2.1, 4.2.2 | N7 | Desktop only |
| D3.07 | Convert the safety list to a SmartArt process, add a step, change its layout and colors, and convert another SmartArt back to text. | 4.3.1, 4.3.2, 4.3.3 | - | Desktop only |

**Skill check SC3:** Timed, MOS-style: 14 tasks on a season report deck (pictures, shapes, a table, a chart and SmartArt, with alt text), auto-checked, 2 attempts. Covers MO-310 3.3.1, 3.3.2, 3.4.1, 3.4.3, 3.4.4, 3.4.5, 3.4.6, 3.4.7, 3.5.1, 3.5.2, 3.5.3, 4.1.1, 4.1.2, 4.1.3, 4.2.1, 4.2.2, 4.3.1, 4.3.3.

**Project P3: Data Story**

A 6 to 8 slide presentation that tells one finding with data (from a class, a club, a job with permission, or public data): a table, a chart that makes the point, a SmartArt process or cycle, pictures, and alt text on every graphic.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P3-a | The table and chart fit the data and support one clear finding. | CLO5, CLO3 | 30 |
| P3-b | Pictures, shapes and SmartArt explain, rather than decorate, and are aligned and grouped neatly. | CLO4 | 25 |
| P3-c | Every graphic has useful alt text or is marked decorative. | CLO4 | 20 |
| P3-d | The data and sources are honest and cited on the slide or in the notes. | CLO3, CLO2 | 25 |
| | **Total** | | **100** |

**AI moment:** Optional: ask an AI tool to suggest which chart type fits your data, then explain in your notes why you agreed or chose another.

**Privacy:** Use public or made-up data. Don't show anyone's grades, health or contact details, and use photos you took or have the right to use.

**Platform:** Pictures, shapes, ink, tables and alignment work in PowerPoint for the web. Screen clippings, editing chart data and some SmartArt changes need desktop PowerPoint.

**Standards:** NACE - Communication, Critical Thinking, Technology. DigComp 3.0 - Content creation, Information search, evaluation and management. WCAG 2.1 - 1.1.1, 1.4.1, 1.3.1.

**Learner takeaway:** Charts and diagrams that make one point, and alt text as a habit.

### Week 4: Motion, Media and Links

*Transitions, animations, video and navigation that help the story instead of distracting from it.*  
**Outline topics:** T4 Using different slide transitions, T5 Using graphics, videos and SmartArt, T7 Creating Animations, using hyperlinks  
**CLOs:** CLO6 (Introduced, Assessed); CLO4 (Reinforced, Assessed); CLO2 (Reinforced); CLO3 (Reinforced)  
**Workload:** About 9 hours: 2.5 h lessons, 3 h drills, 1 h skill check, 2.5 h project

**Why it matters**

- *College:* Revealing one point at a time keeps an audience with you; a short video clip can replace a paragraph.
- *Work:* Training decks, kiosk loops and product demos use motion, media and links to let people explore.
- *Life:* A photo slideshow with music and smooth transitions is the classic party presentation.

**Module objectives** - by the end of the week, learners can:

- **W4.1** Apply basic and 3D transitions and set their effect options, duration and advance timing. *(CLO6)*
- **W4.2** Animate text and graphics with entrance, emphasis, exit and motion-path effects, and set their timing and order. *(CLO6)*
- **W4.3** Insert and trim video and audio, set playback options, insert a screen recording, and insert and animate 3D models. *(CLO4, CLO6)*
- **W4.4** Add hyperlinks, action buttons and Summary, Section and Slide Zoom links. *(CLO2, CLO3)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-310 | Other | Platform |
|---|---|---|---|---|
| D4.01 | Apply Fade to every slide, a 3D transition to the title slide, and change its effect option and duration. | 5.1.1, 5.1.2 | N12 | Web + desktop |
| D4.02 | Set the photo slides to advance automatically after 5 seconds, and add a transition sound to one slide. | 5.1.2 | N12 | Desktop only |
| D4.03 | Animate the bullet list to appear one paragraph at a time, add an emphasis to the logo and an exit to the sticky note. | 5.2.1, 5.2.3 | - | Web + desktop |
| D4.04 | Set the animations to start After Previous with a delay, change a duration, and move the logo animation first. | 5.2.3, 5.2.5 | - | Web + desktop |
| D4.05 | Add a motion path that moves the kayak across the slide and reverse its direction. | 5.2.4 | - | Desktop only |
| D4.06 | Insert the video clip, trim it, set it to start automatically and fade in, and add background music that plays across slides. | 4.5.1, 4.5.3 | N7 | Desktop only |
| D4.07 | Record a short screen recording of a web page and insert it on the last slide. | 4.5.2 | - | Desktop only |
| D4.08 | Insert a 3D model, turn it to a new view, and give it the Turntable animation. | 4.4.1, 4.4.2, 5.2.2 | - | Desktop only |
| D4.09 | Link text to the rentals web page and to the Prices slide, and add an action button back to the menu. | 3.2.1 | - | Web + desktop |
| D4.10 | Insert a Summary Zoom of the three sections, a Section Zoom and a Slide Zoom, and return to the zoom slide after each. | 2.1.4, 3.2.2 | - | Desktop only |

**Skill check SC4:** Timed, MOS-style: 13 tasks on a kiosk tour deck (transitions, animations, a video, links and Zoom), auto-checked, 2 attempts. Most tasks need desktop PowerPoint. Covers MO-310 5.1.1, 5.1.2, 5.2.1, 5.2.3, 5.2.4, 5.2.5, 3.2.1, 3.2.2, 2.1.4, 4.5.1, 4.5.3.

**Project P4: Interactive Tour**

A self-running or click-through tour of a place, a product or a process: a Summary Zoom menu, sections, transitions with automatic timing where it fits, animations that reveal one idea at a time, a video or audio clip, and links back to the menu.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P4-a | Animations and transitions serve the story: they reveal ideas in order, with sensible timing, and never distract. | CLO6 | 30 |
| P4-b | Navigation works: the Zoom menu, links and action buttons all go where they should. | CLO3, CLO2 | 25 |
| P4-c | Media is trimmed, set to play the right way, and captioned or described. | CLO4 | 25 |
| P4-d | The tour runs on its own or by clicking without dead ends, and has a clear start and end. | CLO3 | 20 |
| | **Total** | | **100** |

**AI moment:** Optional: ask an AI tool to write a 30-second script for your video clip, then record it in your own words.

**Privacy:** Record only your own screen and voice. Close email and messages before a screen recording, and get permission before you include anyone else's voice or face.

**Platform:** Transitions, most animations and hyperlinks work in PowerPoint for the web. Trimming media, playback options, screen recordings, 3D model animation, motion paths and Zoom need desktop PowerPoint.

**Standards:** NACE - Communication, Technology, Critical Thinking. DigComp 3.0 - Content creation, Safety, wellbeing and responsible use. WCAG 2.1 - 2.2.2, 2.3.1, 1.2.2, 2.4.4.

**Learner takeaway:** Motion and media used with restraint, and navigation people can follow.

### Week 5: Present and Collaborate

*Review with others, compare and combine their changes, then rehearse, deliver and share safely. Capstone and certification practice.*  
**Outline topics:** T8 Collaboration and sharing, T9 Presenting and viewing a slide show  
**CLOs:** CLO7 (Introduced, Assessed); CLO3 (Assessed); CLO1 (Reinforced, Assessed); CLO2 (Reinforced, Assessed); CLO4 (Assessed); CLO5 (Assessed); CLO6 (Assessed)  
**Workload:** About 9 hours: 1.5 h lessons, 2 h drills, 1 h skill check, 2.5 h project, 2 h capstone

**Why it matters**

- *College:* Group presentations are written together and rehearsed together; feedback arrives as comments.
- *Work:* A deck goes through reviewers before a client sees it, and someone has to merge their changes and run the show.
- *Life:* Presenting with confidence, notes in hand, is a skill that carries far beyond slides.

**Module objectives** - by the end of the week, learners can:

- **W5.1** Add, reply to, resolve and delete comments, and co-author a shared presentation. *(CLO7)*
- **W5.2** Compare a reviewer's copy with yours, accept and reject the differences, and combine the result. *(CLO7)*
- **W5.3** Create custom shows, set slide show options, rehearse timings, set recording options and present with Presenter View. *(CLO3, CLO1)*
- **W5.4** Inspect, fix accessibility, restrict editing, protect with a password, preserve fonts and export to other formats. *(CLO1, CLO7)*
- **W5.5** Show exam readiness on a timed MO-310 style practice test and assemble a course portfolio. *(CLO1, CLO2, CLO3, CLO4, CLO5, CLO6, CLO7)*

**Practice drills (Office Trainer)**

| ID | Task (checked outcome) | MO-310 | Other | Platform |
|---|---|---|---|---|
| D5.01 | Add a comment, reply to and resolve a reviewer's comment, and delete an old one. | 1.5.4 | - | Web + desktop |
| D5.02 | Co-edit a shared copy with a partner, then restore an earlier version from version history. | - | - | Any (self-check) |
| D5.03 | Compare the deck with a reviewer's copy, accept the title and price changes, reject the color change, and end the review. | - | - | Desktop only |
| D5.04 | Create a custom show of the short tour, and set the show to loop until Esc and browse at a kiosk. | 1.4.1, 1.4.2 | - | Desktop only |
| D5.05 | Rehearse timings for the tour, check the recording options, and practice with Presenter View. | 1.4.3, 1.4.4, 1.4.5 | N14 | Desktop only |
| D5.06 | Inspect the deck, fix every accessibility issue, embed the fonts, mark it as final and add a password. | 1.5.1, 1.5.2, 1.5.3, 1.5.5 | - | Desktop only |
| D5.07 | Export the deck as a PDF and a video, and create handouts in Word. | 1.5.6 | - | Any (self-check) |

**Skill check SC5:** Timed, MOS-style: 11 tasks on a team presentation (comments, a custom show, show options, timings, inspection and accessibility), auto-checked, 2 attempts. Covers MO-310 1.4.1, 1.4.2, 1.4.3, 1.5.3, 1.5.4, 1.5.5.

**Project P5: Team Presentation**

With one or two classmates, build a 10 to 12 slide presentation together in OneDrive: comment threads, a reviewer's copy compared and combined, rehearsed timings and speaker notes, a custom short version, and a clean final copy (inspected, accessible, fonts embedded) exported as a PDF and a video. Deliver it live or recorded.

| Criterion | Description | CLOs | Points |
|---|---|---|---|
| P5-a | The team collaborated for real: comment threads with replies, edits from each member, and a compared and combined reviewer's copy. | CLO7 | 25 |
| P5-b | The presentation is professional: clear story, consistent design, notes for every slide, a custom short version. | CLO3 | 25 |
| P5-c | The delivery is rehearsed: timings, Presenter View or a recording, within the time limit. | CLO3, CLO1 | 25 |
| P5-d | The final copy ships clean: inspected, accessible, fonts embedded, exported as PDF and video. | CLO1, CLO7 | 25 |
| | **Total** | | **100** |

**Capstone: Practice Exam and Portfolio**

A timed, 50-minute MO-310 style practice exam covering all five domains (Certiport practice tests if FLCC has access; otherwise the Office Trainer capstone set), plus a portfolio of the five projects with a short reflection on which skills you will use next and where. Evidence for CLO1, CLO2, CLO3, CLO4, CLO5, CLO6, CLO7.

**AI moment:** Optional: Copilot in PowerPoint can draft speaker notes or summarize a deck. Treat it like a reviewer's comment: check it, then decide.

**Privacy:** Shared decks carry hidden data: names in comments, notes and file properties. Inspect before you share outside the team, and share view-only unless people need to edit.

**Platform:** Comments, co-authoring, version history and the accessibility checker work in PowerPoint for the web. Compare, custom shows, show options, rehearsed timings, recording, Inspect Document, passwords, embedding fonts and video export need desktop PowerPoint.

**Standards:** NACE - Teamwork, Communication, Professionalism, Career & Self-Development. DigComp 3.0 - Communication and collaboration, Safety, wellbeing and responsible use. WCAG 2.1 - 1.1.1, 1.3.2, 2.4.2, 2.4.6.

**Learner takeaway:** Confident delivery and safe sharing habits - plus a certification attempt if wanted.

## Alignment matrix

### Course learning outcomes by week

| CLO | Wk 1 | Wk 2 | Wk 3 | Wk 4 | Wk 5 |
|---|---|---|---|---|---|
| **CLO1** Create, save, format and print presentations | I, A | R |  |  | R, A |
| **CLO2** Demonstrate features offered by the application | I, A | R, A | R | R | R, A |
| **CLO3** Create a professional presentation slide show | I | R, A | R | R | A |
| **CLO4** Insert and modify graphics and video clips |  |  | I, A | R, A | A |
| **CLO5** Create and format tables |  |  | I, A |  | A |
| **CLO6** Animate a slide show |  |  |  | I, A | A |
| **CLO7** Collaborate on presentations: Compare and Combine presentations |  |  |  |  | I, A |

### Topic outline by week

| Topic | Wk 1 | Wk 2 | Wk 3 | Wk 4 | Wk 5 |
|---|---|---|---|---|---|
| **T1** Introduction to Microsoft PowerPoint application features | X |  |  |  |  |
| **T2** Creating presentation slide shows | X | X |  |  |  |
| **T3** Formatting and Organizing presentations | X | X |  |  |  |
| **T4** Using different slide transitions |  |  |  | X |  |
| **T5** Using graphics, videos and SmartArt |  |  | X | X |  |
| **T6** Working with Templates, Styles and Charts |  | X | X |  |  |
| **T7** Creating Animations, using hyperlinks |  |  |  | X |  |
| **T8** Collaboration and sharing |  |  |  |  | X |
| **T9** Presenting and viewing a slide show | X |  |  |  | X |

### Module objectives to course outcomes

| Objective | CLO1 | CLO2 | CLO3 | CLO4 | CLO5 | CLO6 | CLO7 |
|---|---|---|---|---|---|---|---|
| W1.1 | X | X |  |  |  |  |  |
| W1.2 | X | X |  |  |  |  |  |
| W1.3 |  | X | X |  |  |  |  |
| W1.4 | X |  |  |  |  |  |  |
| W1.5 | X |  |  |  |  |  |  |
| W2.1 | X | X |  |  |  |  |  |
| W2.2 |  | X |  |  |  |  |  |
| W2.3 |  | X |  |  |  |  |  |
| W2.4 |  | X | X |  |  |  |  |
| W2.5 |  |  | X |  |  |  |  |
| W3.1 |  |  |  | X |  |  |  |
| W3.2 |  | X |  | X |  |  |  |
| W3.3 |  |  |  |  | X |  |  |
| W3.4 |  | X | X |  |  |  |  |
| W4.1 |  |  |  |  |  | X |  |
| W4.2 |  |  |  |  |  | X |  |
| W4.3 |  |  |  | X |  | X |  |
| W4.4 |  | X | X |  |  |  |  |
| W5.1 |  |  |  |  |  |  | X |
| W5.2 |  |  |  |  |  |  | X |
| W5.3 | X |  | X |  |  |  |  |
| W5.4 | X |  |  |  |  |  | X |
| W5.5 | X | X | X | X | X | X | X |

## Evidence for each course learning outcome

### CLO1: Create, save, format and print presentations

- **Bloom level:** Apply
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 5 · **Assessed:** Week 1, Week 5
- **Taught through objectives:** W1.1, W1.2, W1.4, W1.5, W2.1, W5.3, W5.4, W5.5
- **Rubric criteria:** P1-b, P1-d, P2-d, P5-c, P5-d
- **Skill checks and capstone:** SC1, SC5, CAP
- **Evidence artifact:** Project 1 explainer deck and PDF handout; printing and export in Week 1; clean final copy in Project 5.

### CLO2: Demonstrate features offered by the application

- **Bloom level:** Apply
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 3, Week 4, Week 5 · **Assessed:** Week 1, Week 2, Week 5
- **Taught through objectives:** W1.1, W1.2, W1.3, W2.1, W2.2, W2.3, W2.4, W3.2, W3.4, W4.4, W5.5
- **Rubric criteria:** P1-b, P2-a, P2-b, P3-d, P4-b
- **Skill checks and capstone:** SC1, SC2, SC5, CAP
- **Evidence artifact:** Skill checks SC1 to SC5 (feature use under time); Project 2 slide master and layouts; capstone practice exam.

### CLO3: Create a professional presentation slide show

- **Bloom level:** Create
- **Introduced:** Week 1 · **Reinforced:** Week 2, Week 3, Week 4 · **Assessed:** Week 2, Week 5
- **Taught through objectives:** W1.3, W2.4, W2.5, W3.4, W4.4, W5.3, W5.5
- **Rubric criteria:** P1-a, P1-c, P2-c, P3-a, P3-d, P4-b, P4-d, P5-b, P5-c
- **Skill checks and capstone:** SC2, SC5, CAP
- **Evidence artifact:** Every project's design and delivery criteria; Project 5 rehearsed team presentation.

### CLO4: Insert and modify graphics and video clips

- **Bloom level:** Apply
- **Introduced:** Week 3 · **Reinforced:** Week 4 · **Assessed:** Week 3, Week 4, Week 5
- **Taught through objectives:** W3.1, W3.2, W4.3, W5.5
- **Rubric criteria:** P3-b, P3-c, P4-c
- **Skill checks and capstone:** SC3, SC4, SC5, CAP
- **Evidence artifact:** Project 3 pictures, shapes and SmartArt with alt text; Project 4 video or audio; D3.01, D4.06.

### CLO5: Create and format tables

- **Bloom level:** Apply
- **Introduced:** Week 3 · **Reinforced:** - · **Assessed:** Week 3, Week 5
- **Taught through objectives:** W3.3, W5.5
- **Rubric criteria:** P3-a
- **Skill checks and capstone:** SC3, SC5, CAP
- **Evidence artifact:** Project 3 table; SC3; D3.05.

### CLO6: Animate a slide show

- **Bloom level:** Apply
- **Introduced:** Week 4 · **Reinforced:** - · **Assessed:** Week 4, Week 5
- **Taught through objectives:** W4.1, W4.2, W4.3, W5.5
- **Rubric criteria:** P4-a
- **Skill checks and capstone:** SC4, SC5, CAP
- **Evidence artifact:** Project 4 interactive tour animations and transitions; SC4.

### CLO7: Collaborate on presentations: Compare and Combine presentations

- **Bloom level:** Apply / Evaluate
- **Introduced:** Week 5 · **Reinforced:** - · **Assessed:** Week 5
- **Taught through objectives:** W5.1, W5.2, W5.4, W5.5
- **Rubric criteria:** P5-a, P5-d
- **Skill checks and capstone:** SC5, CAP
- **Evidence artifact:** Project 5 comment threads and compared and combined reviewer's copy; D5.03.

### Topic outline evidence

- **T1 Introduction to Microsoft PowerPoint application features:** Week 1 tour of the Ribbon, Backstage view, the slide thumbnails, notes, views, zoom and the Quick Access Toolbar.
- **T2 Creating presentation slide shows:** Week 1 slides and layouts, reusing slides and importing a Word outline; Week 2 themes, masters and templates.
- **T3 Formatting and Organizing presentations:** Week 1 text, lists, sections, slide size and headers and footers; Week 2 backgrounds, the slide master and design clean-up.
- **T4 Using different slide transitions:** Week 4 basic and 3D transitions, effect options, duration and automatic advance.
- **T5 Using graphics, videos and SmartArt:** Week 3 pictures, shapes, icons and SmartArt; Week 4 video, audio, screen recordings and 3D models.
- **T6 Working with Templates, Styles and Charts:** Week 2 themes, slide masters, layouts and a reusable template (.potx); Week 3 charts and table styles.
- **T7 Creating Animations, using hyperlinks:** Week 4 entrance, emphasis, exit and motion-path animations, timing and order; hyperlinks, action buttons and Zoom.
- **T8 Collaboration and sharing:** Week 5 comments, co-authoring, Compare and Combine, inspecting, protecting and exporting.
- **T9 Presenting and viewing a slide show:** Week 1 views and Reading view; Week 5 custom shows, show options, rehearsed timings, recording and Presenter View.

## MOS crosswalk: MO-310 PowerPoint Associate (Microsoft 365 Apps)

All 73 sub-skills are practiced in at least one drill; the capstone practice exam samples 22 of them from every domain. Wording follows Microsoft's skills outline; check the current study guide before printing.

### Domain 1: Manage presentations (see outline)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **1.1** | **Modify slide masters, handout masters, and note masters** | | | |
| 1.1.1 | Change slide master theme or background elements | 2 | D2.03 | SC2 |
| 1.1.2 | Modify slide master content | 2 | D2.03 | SC2 |
| 1.1.3 | Create slide layouts | 2 | D2.04 | SC2 |
| 1.1.4 | Modify slide layouts | 2 | D2.04 | SC2 |
| 1.1.5 | Modify the handout master | 2 | D2.05 | SC2 |
| 1.1.6 | Modify the notes master | 2 | D2.05 | SC2 |
| **1.2** | **Change presentation options and views** | | | |
| 1.2.1 | Change slide size | 1 | D1.09 | SC1 |
| 1.2.2 | Display presentations in different views | 1 | D1.02 | Capstone |
| 1.2.3 | Modify built-in presentation properties | 1 | D1.11 | SC1 |
| **1.3** | **Configure print settings for presentations** | | | |
| 1.3.1 | Print slides | 1 | D1.11 | Capstone |
| 1.3.2 | Print notes pages | 1 | D1.11 | Capstone |
| 1.3.3 | Print handouts | 1 | D1.11 | Capstone |
| **1.4** | **Configure and present slide shows** | | | |
| 1.4.1 | Create custom slide shows | 5 | D5.04 | SC5 |
| 1.4.2 | Configure slide show options | 5 | D5.04 | SC5 |
| 1.4.3 | Rehearse slide show timing | 5 | D5.05 | SC5 |
| 1.4.4 | Set up slide show recording options | 5 | D5.05 | Capstone |
| 1.4.5 | Present slide shows by using Presenter View | 5 | D5.05 | Capstone |
| **1.5** | **Prepare presentations for collaboration and distribution** | | | |
| 1.5.1 | Restrict editing | 5 | D5.06 | Capstone |
| 1.5.2 | Protect presentations by using passwords | 5 | D5.06 | Capstone |
| 1.5.3 | Inspect presentations and correct issues | 5 | D5.06 | SC5 |
| 1.5.4 | Manage comments | 5 | D5.01 | SC5 |
| 1.5.5 | Preserve presentation content | 5 | D5.06 | SC5 |
| 1.5.6 | Export presentations to other formats | 1, 5 | D1.11, D5.07 | Capstone |

### Domain 2: Manage slides (see outline)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **2.1** | **Insert slides** | | | |
| 2.1.1 | Import Word document outlines | 1 | D1.05 | Capstone |
| 2.1.2 | Insert slides from another presentation | 1 | D1.05 | Capstone |
| 2.1.3 | Insert slides and select slide layouts | 1 | D1.04 | SC1 |
| 2.1.4 | Insert Summary Zoom slides | 4 | D4.10 | SC4 |
| 2.1.5 | Duplicate slides | 1 | D1.04 | SC1 |
| **2.2** | **Modify slides** | | | |
| 2.2.1 | Hide and unhide slides | 1 | D1.06 | SC1 |
| 2.2.2 | Modify individual slide backgrounds | 2 | D2.02 | SC2 |
| 2.2.3 | Insert slide headers, footers, and page numbers | 1 | D1.10 | SC1 |
| **2.3** | **Order and group slides** | | | |
| 2.3.1 | Create sections | 1 | D1.06 | SC1 |
| 2.3.2 | Reorder slides and sections | 1 | D1.06 | SC1 |
| 2.3.3 | Rename sections | 1 | D1.06 | SC1 |

### Domain 3: Insert and format text, shapes, and images (see outline)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **3.1** | **Format text** | | | |
| 3.1.1 | Apply formatting and styles to text | 1 | D1.07 | SC1 |
| 3.1.2 | Format text in multiple columns | 1 | D1.07 | SC1 |
| 3.1.3 | Create bulleted and numbered lists | 1 | D1.08 | SC1 |
| **3.2** | **Insert links** | | | |
| 3.2.1 | Insert hyperlinks | 4 | D4.09 | SC4 |
| 3.2.2 | Insert Section Zoom links and Slide Zoom links | 4 | D4.10 | SC4 |
| **3.3** | **Insert and format images** | | | |
| 3.3.1 | Resize and crop images | 3 | D3.01 | SC3 |
| 3.3.2 | Apply built-in styles and effects to images | 3 | D3.02 | SC3 |
| 3.3.3 | Insert screenshots and screen clippings | 3 | D3.02 | Capstone |
| **3.4** | **Insert and format graphic elements** | | | |
| 3.4.1 | Insert graphic elements | 3 | D3.03 | SC3 |
| 3.4.2 | Draw by using digital ink | 3 | D3.04 | Capstone |
| 3.4.3 | Add text to graphic elements | 3 | D3.03 | SC3 |
| 3.4.4 | Resize graphic elements | 3 | D3.03 | SC3 |
| 3.4.5 | Format graphic elements | 3 | D3.03 | SC3 |
| 3.4.6 | Apply built-in styles to graphic elements | 3 | D3.03 | SC3 |
| 3.4.7 | Add alt text to graphic elements for accessibility | 3 | D3.02 | SC3 |
| **3.5** | **Order, align, and group slide content** | | | |
| 3.5.1 | Order slide content | 3 | D3.04 | SC3 |
| 3.5.2 | Align slide content | 3 | D3.04 | SC3 |
| 3.5.3 | Group slide content | 3 | D3.04 | SC3 |
| 3.5.4 | Display alignment tools | 3 | D3.04 | Capstone |

### Domain 4: Insert tables, charts, SmartArt, 3D models, and media (see outline)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **4.1** | **Insert and format tables** | | | |
| 4.1.1 | Create and insert tables | 3 | D3.05 | SC3 |
| 4.1.2 | Insert and delete table rows and columns | 3 | D3.05 | SC3 |
| 4.1.3 | Apply built-in table styles | 3 | D3.05 | SC3 |
| **4.2** | **Insert and modify charts** | | | |
| 4.2.1 | Create and insert charts | 3 | D3.06 | SC3 |
| 4.2.2 | Modify charts | 3 | D3.06 | SC3 |
| **4.3** | **Insert and format SmartArt graphics** | | | |
| 4.3.1 | Insert SmartArt graphics | 3 | D3.07 | SC3 |
| 4.3.2 | Convert between SmartArt graphics and lists | 3 | D3.07 | Capstone |
| 4.3.3 | Add and modify SmartArt graphic content | 3 | D3.07 | SC3 |
| **4.4** | **Insert and modify 3D models** | | | |
| 4.4.1 | Insert 3D models | 4 | D4.08 | Capstone |
| 4.4.2 | Modify 3D model appearance | 4 | D4.08 | Capstone |
| **4.5** | **Insert and manage media** | | | |
| 4.5.1 | Insert audio and video clips | 4 | D4.06 | SC4 |
| 4.5.2 | Create and insert screen recordings | 4 | D4.07 | Capstone |
| 4.5.3 | Configure media playback options | 4 | D4.06 | SC4 |

### Domain 5: Apply transitions and animations (see outline)

| Code | Skill | Week | Drills | Skill check |
|---|---|---|---|---|
| **5.1** | **Apply and configure slide transitions** | | | |
| 5.1.1 | Apply basic and 3D slide transitions | 4 | D4.01 | SC4 |
| 5.1.2 | Configure transition effects and timing | 4 | D4.01, D4.02 | SC4 |
| **5.2** | **Animate slide content** | | | |
| 5.2.1 | Animate text and graphic elements | 4 | D4.03 | SC4 |
| 5.2.2 | Animate 3D models | 4 | D4.08 | Capstone |
| 5.2.3 | Configure animation effects and timing | 4 | D4.03, D4.04 | SC4 |
| 5.2.4 | Configure animation paths | 4 | D4.05 | SC4 |
| 5.2.5 | Reorder animations on a slide | 4 | D4.04 | SC4 |

## Other standards and conventions

### Northstar Digital Literacy: Microsoft PowerPoint

| Code | Standard | Week | Drill |
|---|---|---|---|
| N1 | Open a new or existing PowerPoint presentation. | 1 | D1.01 |
| N2 | Identify parts of the PowerPoint screen (slide navigation pane, slide pane, notes, the ribbon, quick access toolbar, and scroll bars). | 1 | D1.02 |
| N3 | Insert new slides, duplicate, or reuse slides. | 1 | D1.04 |
| N4 | Manage text (insert, delete, copy, cut and paste, drag and drop, format, and use spellcheck). | 1 | D1.07 |
| N5 | Apply or change a theme. | 2 | D2.01 |
| N6 | Use zoom control. | 1 | D1.02 |
| N7 | Insert items into a presentation, resize, and adjust them (video, chart, pictures, clip art, screenshots). | 3 | D3.01, D3.06, D4.06 |
| N8 | Add a textbox, adjust it, resize it, or delete it. | 3 | D3.03 |
| N9 | Change the view (normal view, slide sorter, reading view, slideshow view). | 1 | D1.02 |
| N10 | Insert, delete, and move slides using slide navigation pane. | 1 | D1.06 |
| N11 | Use the quick access toolbar. | 1 | D1.03 |
| N12 | Apply and customize slide transitions (select, preview, add sound, automatic advance). | 4 | D4.01, D4.02 |
| N13 | Understand the basics of PowerPoint etiquette (limited text, text that stands out on background, clear titles). | 2 | D2.07 |
| N14 | Play a slideshow, advance through the slides, and end slideshow (using screen toolbar features). | 5 | D5.05 |
| N15 | Save a presentation as a .ppt, .pdf, .png, etc. | 1 | D1.11 |
| N16 | Create handouts. | 1 | D1.11 |
| N17 | Print a presentation. | 1 | D1.11 |

### NACE Career Readiness Competencies

| Competency | Weeks | Where it shows up |
|---|---|---|
| Career & Self-Development | 5 | A portfolio reflection on where presentation skills will be used next. |
| Communication | 1, 2, 3, 4, 5 | Every week: slides that make one point, notes that carry the detail, delivery. |
| Critical Thinking | 3, 4 | Choosing the right chart, and deciding when motion helps and when it distracts (including AI suggestions). |
| Professionalism | 1, 2, 5 | Consistent design, templates and clean, inspected files. |
| Teamwork | 5 | Comments, co-authoring, comparing and combining reviewers' copies. |
| Technology | 1, 2, 3, 4, 5 | Every week: choosing the right PowerPoint feature and using it responsibly. |

### DigComp 3.0 (European Digital Competence Framework, December 2025)

| Competence area | Weeks | Where it shows up |
|---|---|---|
| Information search, evaluation and management | 3 | Finding, checking and citing data for a data story. |
| Communication and collaboration | 1, 5 | Presenting, sharing, co-authoring and comments. |
| Content creation | 1, 2, 3, 4 | Building decks, templates, charts, diagrams and media. |
| Safety, wellbeing and responsible use | 4, 5 | Privacy in screen recordings and photos; Inspect Document; passwords. |
| Problem identification and solving | 2, 4 | Fixing cluttered slides and broken navigation. |

### Accessibility: WCAG 2.1 AA success criteria practiced (ADA Title II)

| Criterion | Name | Level | In PowerPoint | Weeks |
|---|---|---|---|---|
| 1.1.1 | Non-text Content | A | Alt text on pictures, charts, SmartArt, video and 3D models; decorative items marked decorative. | 3, 5 |
| 1.2.2 | Captions (Prerecorded) | A | Videos in the deck have captions or a described alternative. | 4 |
| 1.3.1 | Info and Relationships | A | Real layouts and placeholders, lists and table header rows instead of look-alike text boxes. | 1, 2, 3 |
| 1.3.2 | Meaningful Sequence | A | The reading order (Selection pane) matches the order a sighted reader follows. | 5 |
| 1.4.1 | Use of Color | A | Charts and diagrams don't rely on color alone. | 3 |
| 1.4.3 | Contrast (Minimum) | AA | Text keeps enough contrast with its background, including over pictures. | 2 |
| 2.2.2 | Pause, Stop, Hide | A | Self-running shows and looping media can be paused or stopped. | 4 |
| 2.3.1 | Three Flashes or Below Threshold | A | No flashing animations or transitions. | 4 |
| 2.4.2 | Page Titled | A | Every slide has a unique title, and the file has a Title property. | 1, 5 |
| 2.4.4 | Link Purpose (In Context) | A | Hyperlinks and action buttons say where they go. | 4 |
| 2.4.6 | Headings and Labels | AA | Slide titles and section names describe their contents. | 1, 5 |

### PowerPoint for the web vs desktop PowerPoint

From Microsoft's feature comparison table. Microsoft updates PowerPoint for the web often, so recheck each term.

| Feature | PowerPoint for the web | Desktop PowerPoint | Used in |
|---|---|---|---|
| Slides, layouts, sections, hide slides | Yes | Yes | W1 |
| Reuse Slides from a file, Slides from Outline | No | Yes | W1 |
| Header and Footer, slide numbers | Check (varies) | Yes | W1 |
| Themes, variants, slide backgrounds | Yes | Yes | W2 |
| Slide master, layouts, handout and notes masters | No | Yes | W2 |
| Save as template (.potx) | No | Yes | W2 |
| Pictures, shapes, icons, ink, alignment | Yes | Yes | W3 |
| Screenshots and screen clippings | No | Yes | W3 |
| Tables and table styles | Yes | Yes | W3 |
| Charts (edit data) | Limited | Yes | W3 |
| SmartArt | Limited | Yes | W3 |
| Transitions and most animations | Yes | Yes | W4 |
| Motion paths, animation pane details | Limited | Yes | W4 |
| Trim media, playback options | No | Yes | W4 |
| Screen recording, 3D model animations | No | Yes | W4 |
| Summary, Section and Slide Zoom | Play only | Yes | W4 |
| Comments and co-authoring | Yes | Yes | W5 |
| Compare (combine a reviewer's copy) | No | Yes | W5 |
| Custom shows, show options, rehearse timings, recording | No | Yes | W5 |
| Inspect Document, passwords, embed fonts, Mark as Final | No | Yes | W5 |
| Export as PDF | Yes | Yes | W1, W5 |
| Export as video, handouts in Word | No | Yes | W5 |
| Accessibility checker | Yes | Yes | W5 |

Drill platform split: 19 work on the web, 20 need desktop PowerPoint, 3 work anywhere with a self-check.

### Conventions and assumptions

| Item | Type | Note | Weeks |
|---|---|---|---|
| Presentation design conventions | Industry convention | One idea per slide, a real title on every slide, readable sizes (about 24 pt and up for body text), high contrast, details in the notes. | W1, W2 |
| Restraint with motion | Industry convention | Transitions and animations serve the story; at most one or two kinds per deck; nothing that flashes. | W4 |
| Bloom's revised taxonomy | Assumption | Outcome verbs sit mostly at Apply; building a professional slide show (CLO 3) reaches Create, and peer review reaches Evaluate. | All |
| Credit hour (about 45 hours per credit) | Assumption | CSC 136 is 1 credit; 9 hours a week for 5 weeks. | All |
| I / R / A curriculum mapping | Common practice | Introduced, Reinforced, Assessed - the standard way to show where an outcome is taught and measured. | All |

## Sources

- [CSC 136 PowerPoint - course description (FLCC catalog)](https://flcc.smartcatalogiq.com/en/current/catalog/course-descriptions/csc-computer-science/100/csc-136)
- [CSC 136 PowerPoint - course outline with outcomes (FLCC)](https://cyber.flcc.edu/pdf/syllabi/CSC136.pdf)
- [MOS PowerPoint Associate (Microsoft 365 Apps), exam MO-310 - Microsoft Learn](https://learn.microsoft.com/en-us/credentials/certifications/mos-powerpoint-associate-m365-apps/)
- [PowerPoint (Microsoft 365 Apps) Exam MO-310 objective domains (CCI Learning copy)](https://products.ccilearning.com/wp-content/uploads/2023/12/MOS-365-PPT-2022-OD.pdf)
- [Northstar Digital Literacy - Microsoft PowerPoint criteria](https://www.digitalliteracyassessment.org/badges/ms-ppt/v2/proctored/)
- [NACE Career Readiness: Competencies Defined - NACE](https://www.naceweb.org/career-readiness/competencies/career-readiness-defined)
- [DigComp framework - European Commission Joint Research Centre](https://joint-research-centre.ec.europa.eu/digcomp/digcomp-framework_en)
- [Understanding WCAG 2.1 success criteria - W3C](https://www.w3.org/WAI/WCAG21/Understanding/)
