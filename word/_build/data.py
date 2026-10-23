# -*- coding: utf-8 -*-
"""
Single source of data for the 5-week Microsoft Word course map.
build.py turns this into the Markdown course map and the interactive HTML page,
and runs coverage checks so nothing in the official outline gets dropped.
Style rule: hyphens, never em dashes.
"""

COURSE = {
    "title": "Microsoft Word - 5-Week Course Map",
    "short": "Word Course Map",
    "college": "Finger Lakes Community College (FLCC)",
    "version": "Draft 1",
    "date": "October 17, 2026",
    "author": "Dave Ghidiu (with Claude)",
    "format": "5 weeks, online or hybrid, about 9 hours of learner work per week (assumes 1 credit - see Decisions)",
    "cert": "Microsoft Office Specialist (MOS): Word Associate (Microsoft 365 Apps), exam MO-110",
}

# ---------------------------------------------------------------------------
# Official course description - carried through verbatim
# ---------------------------------------------------------------------------
CLOS = [
    {"id": "CLO1", "text": "Create, save, format and print documents", "bloom": "Apply"},
    {"id": "CLO2", "text": "Demonstrate features offered by the application", "bloom": "Apply"},
    {"id": "CLO3", "text": "Insert and modify graphics", "bloom": "Apply"},
    {"id": "CLO4", "text": "Create and format tables", "bloom": "Apply"},
    {"id": "CLO5", "text": "Apply desktop publishing features to a document", "bloom": "Apply"},
    {"id": "CLO6", "text": "Create and preview the document as a Web page", "bloom": "Apply"},
    {"id": "CLO7", "text": "use mail merge to create form letters, directories, and envelopes", "bloom": "Apply"},
    {"id": "CLO8", "text": "Collaborate on documents, Compare and Combine documents, Manage comments, Track document changes", "bloom": "Apply / Evaluate"},
]

TOPICS = [
    {"id": "T1", "text": "Introduction to Microsoft Word application features", "weeks": ["W1"]},
    {"id": "T2", "text": "Creating Documents with Microsoft Word", "weeks": ["W1"]},
    {"id": "T3", "text": "Formatting and Organizing Text", "weeks": ["W1", "W2"]},
    {"id": "T4", "text": "Using Graphics and Tables", "weeks": ["W3"]},
    {"id": "T5", "text": "Special Document Formats, Columns and Mail Merge", "weeks": ["W3", "W4"]},
    {"id": "T6", "text": "Working with Templates, Styles and Charts", "weeks": ["W2", "W3", "W4"]},
    {"id": "T7", "text": "Creating Form Letters and Working in Groups (Collaboration)", "weeks": ["W4", "W5"]},
]

# How each topic is met, in plain words (for the evidence section)
TOPIC_HOW = {
    "T1": "Week 1 tour of the Ribbon, Backstage view, status bar, Quick Access Toolbar and views (Print Layout, Read Mode, Web Layout).",
    "T2": "Week 1 create, name, save, reopen and export documents; Project 1 business letter and memo.",
    "T3": "Week 1 character and paragraph formatting; Week 2 styles, lists, headers and footers, breaks and the Navigation pane.",
    "T4": "Week 3 pictures, shapes, icons, SmartArt, text boxes and tables; Project 3 newsletter.",
    "T5": "Week 3 columns and section-level layout; Week 4 envelopes, labels, directories, web page output and mail merge.",
    "T6": "Week 2 styles, style sets and a reusable paper template; Week 3 charts and captions; Week 4 resume template.",
    "T7": "Week 4 merged cover letters (form letters); Week 5 comments, Track Changes, Compare and Combine, co-authoring.",
}

# ---------------------------------------------------------------------------
# MO-110 skills outline (Microsoft 365 Apps). Sub-skill wording follows
# Microsoft's study guide closely; verify against the current PDF before print.
# ---------------------------------------------------------------------------
MO110 = [
    {"id": "1", "name": "Manage documents", "weight": "20-25%", "objectives": [
        {"id": "1.1", "name": "Navigate within documents", "skills": [
            ("1.1.1", "Search for text"),
            ("1.1.2", "Link to locations within documents"),
            ("1.1.3", "Move to specific locations and objects in documents"),
            ("1.1.4", "Show and hide formatting symbols and hidden text")]},
        {"id": "1.2", "name": "Format documents", "skills": [
            ("1.2.1", "Set up document pages"),
            ("1.2.2", "Apply style sets"),
            ("1.2.3", "Insert and modify headers and footers"),
            ("1.2.4", "Configure page background elements")]},
        {"id": "1.3", "name": "Save and share documents", "skills": [
            ("1.3.1", "Save documents in alternative file formats"),
            ("1.3.2", "Modify basic document properties"),
            ("1.3.3", "Modify print settings"),
            ("1.3.4", "Share documents electronically")]},
        {"id": "1.4", "name": "Inspect documents for issues", "skills": [
            ("1.4.1", "Locate and remove hidden properties and personal information"),
            ("1.4.2", "Locate and correct accessibility issues"),
            ("1.4.3", "Locate and correct compatibility issues")]},
    ]},
    {"id": "2", "name": "Insert and format text, paragraphs, and sections", "weight": "20-25%", "objectives": [
        {"id": "2.1", "name": "Insert text and paragraphs", "skills": [
            ("2.1.1", "Find and replace text"),
            ("2.1.2", "Insert symbols and special characters")]},
        {"id": "2.2", "name": "Format text and paragraphs", "skills": [
            ("2.2.1", "Apply text effects"),
            ("2.2.2", "Apply formatting by using Format Painter"),
            ("2.2.3", "Set line and paragraph spacing and indentation"),
            ("2.2.4", "Apply built-in styles to text"),
            ("2.2.5", "Clear formatting")]},
        {"id": "2.3", "name": "Create and configure document sections", "skills": [
            ("2.3.1", "Format text in multiple columns"),
            ("2.3.2", "Insert page, section, and column breaks"),
            ("2.3.3", "Change page setup options for a section")]},
    ]},
    {"id": "3", "name": "Manage tables and lists", "weight": "20-25%", "objectives": [
        {"id": "3.1", "name": "Create tables", "skills": [
            ("3.1.1", "Convert text to tables"),
            ("3.1.2", "Convert tables to text"),
            ("3.1.3", "Create tables by specifying rows and columns")]},
        {"id": "3.2", "name": "Modify tables", "skills": [
            ("3.2.1", "Sort table data"),
            ("3.2.2", "Configure cell margins and spacing"),
            ("3.2.3", "Merge and split cells"),
            ("3.2.4", "Resize tables, rows, and columns"),
            ("3.2.5", "Split tables"),
            ("3.2.6", "Configure a repeating row header")]},
        {"id": "3.3", "name": "Create and modify lists", "skills": [
            ("3.3.1", "Format paragraphs as numbered and bulleted lists"),
            ("3.3.2", "Change bullet characters and number formats"),
            ("3.3.3", "Define custom bullet characters and number formats"),
            ("3.3.4", "Increase and decrease list levels"),
            ("3.3.5", "Restart and continue list numbering"),
            ("3.3.6", "Set starting number values")]},
    ]},
    {"id": "4", "name": "Create and manage references", "weight": "5-10%", "objectives": [
        {"id": "4.1", "name": "Create and manage reference elements", "skills": [
            ("4.1.1", "Insert footnotes and endnotes"),
            ("4.1.2", "Modify footnote and endnote properties")]},
        {"id": "4.2", "name": "Create and manage reference tables", "skills": [
            ("4.2.1", "Insert tables of contents"),
            ("4.2.2", "Customize tables of contents")]},
    ]},
    {"id": "5", "name": "Insert and format graphic elements", "weight": "15-20%", "objectives": [
        {"id": "5.1", "name": "Insert illustrations and text boxes", "skills": [
            ("5.1.1", "Insert shapes"),
            ("5.1.2", "Insert pictures"),
            ("5.1.3", "Insert 3D models"),
            ("5.1.4", "Insert SmartArt graphics"),
            ("5.1.5", "Insert screenshots and screen clippings"),
            ("5.1.6", "Insert text boxes"),
            ("5.1.7", "Insert icons")]},
        {"id": "5.2", "name": "Format illustrations and text boxes", "skills": [
            ("5.2.1", "Apply artistic effects"),
            ("5.2.2", "Apply picture effects and picture styles"),
            ("5.2.3", "Remove picture backgrounds"),
            ("5.2.4", "Format graphic elements"),
            ("5.2.5", "Format SmartArt graphics"),
            ("5.2.6", "Format 3D models")]},
        {"id": "5.3", "name": "Add text to graphic elements", "skills": [
            ("5.3.1", "Add and modify text in text boxes"),
            ("5.3.2", "Add and modify text in shapes"),
            ("5.3.3", "Add and modify SmartArt graphic content")]},
        {"id": "5.4", "name": "Modify graphic elements", "skills": [
            ("5.4.1", "Position objects"),
            ("5.4.2", "Wrap text around objects"),
            ("5.4.3", "Add alternative text to objects for accessibility")]},
    ]},
    {"id": "6", "name": "Manage document collaboration", "weight": "5-10%", "objectives": [
        {"id": "6.1", "name": "Add and manage comments", "skills": [
            ("6.1.1", "Add comments"),
            ("6.1.2", "Review and reply to comments"),
            ("6.1.3", "Resolve comments"),
            ("6.1.4", "Delete comments")]},
        {"id": "6.2", "name": "Manage change tracking", "skills": [
            ("6.2.1", "Track changes"),
            ("6.2.2", "Review tracked changes"),
            ("6.2.3", "Accept and reject tracked changes"),
            ("6.2.4", "Lock and unlock change tracking")]},
    ]},
]

# MO-111 Word Expert - used as "beyond the certification" stretch content.
# status: full = taught and assessed, partial = some skills taught, none = not in course
MO111 = [
    {"id": "1.1", "domain": "Manage document options and settings", "name": "Manage documents and templates",
     "skills": "Modify templates, manage versions, compare and combine, link to external content, enable macros, customize the Quick Access Toolbar, show hidden ribbon tabs, change the Normal template default font",
     "status": "partial", "weeks": ["W1", "W4", "W5"],
     "how": "Templates (W4), Quick Access Toolbar (W1), compare and combine plus version history (W5). Macros, external links and Normal template defaults are not taught."},
    {"id": "1.2", "domain": "Manage document options and settings", "name": "Prepare documents for collaboration",
     "skills": "Restrict editing, protect documents with passwords",
     "status": "full", "weeks": ["W5"], "how": "Restrict editing and password protection in the W5 clean-and-ship checklist."},
    {"id": "1.3", "domain": "Manage document options and settings", "name": "Use and configure language options",
     "skills": "Editing and display languages, language-specific features",
     "status": "partial", "weeks": ["W1"], "how": "Set the document proofing language (also an accessibility requirement). Display language is not taught."},
    {"id": "2.1", "domain": "Use advanced editing and formatting features", "name": "Find, replace, and paste document content",
     "skills": "Wildcards and special characters, find and replace formatting and styles, paste options",
     "status": "partial", "weeks": ["W1"], "how": "Paste options and special-character search (W1). Wildcards are optional enrichment."},
    {"id": "2.2", "domain": "Use advanced editing and formatting features", "name": "Configure paragraph layout options",
     "skills": "Hyphenation, line numbers, pagination options",
     "status": "partial", "weeks": ["W2"], "how": "Keep with next and widow/orphan control for headings (W2). Hyphenation and line numbers are not taught."},
    {"id": "2.3", "domain": "Use advanced editing and formatting features", "name": "Create and manage styles",
     "skills": "Create, modify and copy paragraph, character and table styles",
     "status": "full", "weeks": ["W2"], "how": "Modify and create styles for the paper template (W2)."},
    {"id": "3.1", "domain": "Create custom document elements", "name": "Create and modify building blocks",
     "skills": "Quick Parts, manage building blocks",
     "status": "full", "weeks": ["W4"], "how": "Save the contact heading as a Quick Part and reuse it (W4)."},
    {"id": "3.2", "domain": "Create custom document elements", "name": "Create custom design elements",
     "skills": "Custom color sets, font sets, themes, style sets",
     "status": "partial", "weeks": ["W3"], "how": "Custom theme colors for a club brand (W3). Custom font sets and saved themes are optional."},
    {"id": "3.3", "domain": "Create custom document elements", "name": "Create and manage indexes",
     "skills": "Mark entries, create and update indexes",
     "status": "none", "weeks": [], "how": "Not taught. Low value for most learners; optional enrichment."},
    {"id": "3.4", "domain": "Create custom document elements", "name": "Create and manage tables of figures",
     "skills": "Captions, caption properties, tables of figures",
     "status": "partial", "weeks": ["W3"], "how": "Captions on the chart and table (W3). Table of figures is optional."},
    {"id": "3.5", "domain": "Create custom document elements", "name": "Create and manage bibliographies",
     "skills": "Citation sources, citations, bibliographies",
     "status": "full", "weeks": ["W2"], "how": "Citation manager and generated references list in Project 2 (W2)."},
    {"id": "4.1", "domain": "Use advanced Word features", "name": "Manage forms, fields, and controls",
     "skills": "Custom fields, field properties, content controls",
     "status": "partial", "weeks": ["W4"], "how": "Template content controls and the If...Then...Else merge field (W4)."},
    {"id": "4.2", "domain": "Use advanced Word features", "name": "Create and modify macros",
     "skills": "Record, name, edit and copy macros",
     "status": "none", "weeks": [], "how": "Not taught. Macros are a security risk to explain, not a beginner skill to build."},
    {"id": "4.3", "domain": "Use advanced Word features", "name": "Perform mail merges",
     "skills": "Recipient lists, merge fields, preview, merged documents, labels, envelopes",
     "status": "full", "weeks": ["W4"], "how": "All of Project 4's merge work (W4)."},
]

# ---------------------------------------------------------------------------
# Weeks
# drill fields: id, task (outcome the Office Trainer can check), mos (MO-110 codes),
#   x (other tags: MO-111 ids as "E1.1", Northstar as "N3"), where ("web" = works in
#   Word for the web, "desktop" = needs desktop Word, "both-check" = works anywhere but
#   needs a screenshot or self-check because the result is not saved in the file)
# ---------------------------------------------------------------------------
WEEKS = [
  {
    "id": "W1", "n": 1, "title": "The Professional Document",
    "tagline": "Create, format, save and print documents people take seriously.",
    "clos": {"CLO1": "IA", "CLO2": "IA", "CLO6": "I"},
    "topics": ["T1", "T2", "T3"],
    "why": {
        "college": "Every course expects clean, correctly formatted submissions saved where you can find them.",
        "work": "Letters and memos are still how organizations make requests and keep records.",
        "life": "Landlord letters, appeals, complaints and requests get better results when they look professional.",
    },
    "objectives": [
        ("W1.1", "Identify and use the parts of the Word window (Ribbon, Backstage view, status bar, Quick Access Toolbar) and switch among Print Layout, Read Mode and Web Layout views.", ["CLO2", "CLO6"]),
        ("W1.2", "Create, name, save and reopen documents in a deliberate location (OneDrive or local), and save copies as .docx and .pdf.", ["CLO1"]),
        ("W1.3", "Format characters and paragraphs: font, size, color, effects, alignment, line and paragraph spacing, indentation, Format Painter and Clear Formatting.", ["CLO1", "CLO2"]),
        ("W1.4", "Use Find and Replace, symbols and special characters, cut/copy/paste options, Undo and Editor (spelling and grammar) to produce error-free text.", ["CLO2"]),
        ("W1.5", "Set up pages (margins, orientation, paper size), set document properties and language, and use Print Preview and print settings.", ["CLO1"]),
        ("W1.6", "Show formatting marks to diagnose and fix spacing and layout problems.", ["CLO2"]),
    ],
    "drills": [
        ("D1.01", "Open the starter file and save it to your course folder as Lastname_W1_Letter.docx.", [], ["N1", "N3", "N4"], "web"),
        ("D1.02", "Add Print Preview and Save As to the Quick Access Toolbar, then name three parts of the window in the drill quiz.", [], ["E1.1", "N2"], "desktop"),
        ("D1.03", "Show formatting marks; remove the extra blank paragraphs and double spaces in the letter.", ["1.1.4", "2.1.1"], [], "web"),
        ("D1.04", "Use Find to count every mention of the deadline, then replace every Acme Outdoor Supply with Lakeshore Paddle Company.", ["1.1.1", "2.1.1"], [], "web"),
        ("D1.05", "Insert a copyright symbol in the footer and type cafe with an accented e.", ["2.1.2"], [], "web"),
        ("D1.06", "Set body text to 11 pt Aptos; bold the subject line and apply a text effect to the letterhead; copy the letterhead format to the memo heading with Format Painter.", ["2.2.1", "2.2.2"], ["N6"], "web"),
        ("D1.07", "Set single line spacing with 8 pt after; give the memo body a 0.5 in first-line indent; left-align the letter; clear formatting from the pasted paragraph.", ["2.2.3", "2.2.5"], ["N7"], "web"),
        ("D1.08", "Move the closing paragraph above the signature block with cut and paste (Keep Text Only); undo an accidental deletion.", [], ["E2.1", "N9", "N10"], "web"),
        ("D1.09", "Set 1 in margins, Letter size and portrait orientation; run Editor and fix every flagged issue.", ["1.2.1"], ["N5", "N11"], "web"),
        ("D1.10", "Fill in the Title, Author and Tags properties and set the proofing language to English (United States).", ["1.3.2"], ["E1.3"], "desktop"),
        ("D1.11", "Export the letter to PDF; set print to 2 copies of page 1 only and capture the print preview.", ["1.3.1", "1.3.3"], ["N12", "N13"], "both-check"),
        ("D1.12", "Switch to Web Layout and Read Mode and describe in one sentence what changes (Word for the web has no Web Layout; compare Reading view instead).", [], [], "desktop"),
    ],
    "skillcheck": {"id": "SC1", "covers": ["1.1.1", "1.1.4", "1.2.1", "1.3.2", "2.1.1", "2.1.2", "2.2.1", "2.2.2", "2.2.3", "2.2.5"],
                   "desc": "Timed, MOS-style: 10 tasks in a fresh starter file, auto-checked by the Office Trainer, 2 attempts. MO-110 1.3.1 is assessed through the Project 1 PDF instead, because a check cannot see files outside the document."},
    "project": {
        "id": "P1", "title": "Business Correspondence Set",
        "desc": "A block-format business letter (for example, a request to a financial aid office or an informational-interview request to a local employer) and a one-page memo. Submit .docx and .pdf.",
        "criteria": [
            ("P1-a", "Block-letter structure is complete and correct (date, inside address, salutation, body, closing, signature block); memo heading is correct.", ["CLO1"], 25),
            ("P1-b", "Formatting is done with Word tools, not workarounds (no spaces for indents, no extra returns for spacing).", ["CLO1", "CLO2"], 25),
            ("P1-c", "Text is error-free (Editor run, Find and Replace used where needed).", ["CLO2"], 25),
            ("P1-d", "Page setup, properties and language are set; PDF opens and matches the .docx.", ["CLO1"], 25),
        ],
    },
    "ai": "Optional: ask an AI tool to draft the letter's body. You make every formatting decision, check every fact, and add a one-line note on what you kept, changed or cut. There is no wrong way to try this.",
    "platform": "Everything this week works in Word for the web and desktop Word. Quick Access Toolbar customizing and the proofing-language setting are easier in desktop Word. Print dialogs vary by browser.",
    "nace": ["Communication", "Professionalism", "Technology"],
    "digcomp": ["Content creation", "Information search, evaluation and management"],
    "wcag": ["2.4.2", "3.1.1"],
    "life": "A formal letter and memo you can send with confidence.",
    "hours": "About 9 hours: 2 h lessons and videos, 3 h drills, 1 h skill check, 3 h project",
  },
  {
    "id": "W2", "n": 2, "title": "Structure and the College Paper",
    "tagline": "Let styles do the work: long documents, citations and a reusable paper template.",
    "clos": {"CLO1": "R", "CLO2": "RA", "CLO6": "R"},
    "topics": ["T3", "T6"],
    "why": {
        "college": "A correct APA or MLA template saves time in every writing-heavy course you take.",
        "work": "Reports, manuals and proposals depend on styles, navigation and automatic tables of contents.",
        "life": "Long documents (a family history, a club handbook) stay manageable when the structure is real.",
    },
    "objectives": [
        ("W2.1", "Apply and modify built-in heading and body styles and switch style sets, so structure comes from styles rather than manual formatting.", ["CLO2"]),
        ("W2.2", "Navigate long documents with the Navigation pane, Go To, bookmarks and internal hyperlinks.", ["CLO2", "CLO6"]),
        ("W2.3", "Create and control bulleted, numbered and multilevel lists (custom bullets, number formats, levels, restart and continue, start value).", ["CLO1", "CLO2"]),
        ("W2.4", "Insert headers, footers and page numbers (including a different first page) and control layout with page and section breaks.", ["CLO1", "CLO2"]),
        ("W2.5", "Insert and modify footnotes and endnotes, and generate and customize an automatic table of contents.", ["CLO2"]),
        ("W2.6", "Manage citation sources and generate a reference list in APA 7th or MLA 9th style, and save the setup as a reusable template.", ["CLO2"]),
    ],
    "drills": [
        ("D2.01", "Apply Heading 1 and Heading 2 to the report's section titles.", ["2.2.4"], [], "web"),
        ("D2.02", "Modify the Normal and Heading 1 styles (font, spacing, keep with next) and apply a different style set.", ["1.2.2"], ["E2.3", "E2.2"], "desktop"),
        ("D2.03", "Use the Navigation pane to move the Methods section above Results by dragging its heading; use Go To to jump to page 4.", ["1.1.3"], [], "web"),
        ("D2.04", "Add a bookmark at the Methods heading and a hyperlink to it from the introduction.", ["1.1.2"], [], "web"),
        ("D2.05", "Turn the plain list into a numbered list; demote two items to level 2; change the bullet character in the second list.", ["3.3.1", "3.3.2", "3.3.4"], ["N8"], "web"),
        ("D2.06", "Define a custom number format (Step 1, Step 2); restart numbering in the third list and set it to start at 5.", ["3.3.3", "3.3.5", "3.3.6"], [], "desktop"),
        ("D2.07", "Insert a header with your last name and page number, using a different first page.", ["1.2.3"], [], "web"),
        ("D2.08", "Insert a Next Page section break and make only the appendix section landscape.", ["2.3.2", "2.3.3"], [], "desktop"),
        ("D2.09", "Insert two footnotes and change the number format to i, ii, iii.", ["4.1.1", "4.1.2"], [], "desktop"),
        ("D2.10", "Insert an automatic table of contents, add a new heading, then update the table and limit it to two levels.", ["4.2.1", "4.2.2"], [], "web"),
        ("D2.11", "Add three sources in the citation manager, insert in-text citations, and generate the references list in APA 7th.", [], ["E3.5"], "desktop"),
    ],
    "skillcheck": {"id": "SC2", "covers": ["1.1.2", "1.1.3", "1.2.2", "1.2.3", "2.2.4", "2.3.2", "2.3.3", "3.3.1", "3.3.2", "3.3.4", "3.3.5", "4.1.1", "4.1.2", "4.2.1", "4.2.2"],
                   "desc": "Timed, MOS-style: 10 tasks on a 6-page starter report, auto-checked, 2 attempts."},
    "project": {
        "id": "P2", "title": "Research Paper Template and Paper",
        "desc": "Build a reusable APA 7th or MLA 9th template (.dotx) with modified styles, header and page numbers, then use it for a 3 to 5 page paper with at least three verified sources. Learners may reuse a paper from another course, which ties this skill to their real coursework.",
        "criteria": [
            ("P2-a", "Headings use real heading styles; the Navigation pane shows the outline.", ["CLO2"], 20),
            ("P2-b", "Style-guide layout is correct (margins, spacing, header, page numbers, title page or heading).", ["CLO1"], 20),
            ("P2-c", "Citations and references are generated by Word and every source is real and checked.", ["CLO2"], 25),
            ("P2-d", "Lists, breaks and notes are used correctly.", ["CLO2"], 15),
            ("P2-e", "The .dotx template opens a new, correctly formatted document.", ["CLO2"], 20),
        ],
    },
    "ai": "Optional: AI tools sometimes invent sources that look real. If an AI tool suggests a source, find it in the library database before it goes into the citation manager. That habit protects your grade.",
    "platform": "Styles, the Navigation pane, footnotes and tables of contents work in Word for the web. Section breaks, custom number formats and the citation manager need desktop Word, per Microsoft's feature comparison.",
    "nace": ["Communication", "Critical Thinking", "Technology"],
    "digcomp": ["Information search, evaluation and management", "Content creation"],
    "wcag": ["1.3.1", "2.4.4", "2.4.6"],
    "life": "A paper template reused every semester, plus long-document skills for reports at work.",
    "hours": "About 9 hours: 2 h lessons, 3 h drills, 1 h skill check, 3 h project",
  },
  {
    "id": "W3", "n": 3, "title": "Visual Communication",
    "tagline": "Graphics, tables, charts and desktop publishing that stay accessible.",
    "clos": {"CLO3": "IA", "CLO4": "IA", "CLO5": "IA", "CLO2": "R", "CLO1": "R"},
    "topics": ["T4", "T5", "T6"],
    "why": {
        "college": "Lab reports, posters and club flyers mix text, tables and images.",
        "work": "Schedules, price lists, one-pagers and newsletters are everyday workplace documents.",
        "life": "Event flyers, small-business menus and community newsletters - and making them readable for everyone.",
    },
    "objectives": [
        ("W3.1", "Insert and format pictures, shapes, icons, screenshots, 3D models and SmartArt, including picture styles, artistic effects and background removal.", ["CLO3"]),
        ("W3.2", "Position graphics with text wrapping and alignment, add text to shapes and text boxes, and write meaningful alternative text.", ["CLO3"]),
        ("W3.3", "Create tables from scratch and from text; merge and split cells, resize, sort, set cell margins, split tables and repeat header rows.", ["CLO4"]),
        ("W3.4", "Insert and format a chart, and add captions to figures and tables.", ["CLO3"]),
        ("W3.5", "Apply desktop publishing features: multiple columns, column breaks, themes and custom theme colors, page color, borders, watermark and drop caps.", ["CLO5"]),
    ],
    "drills": [
        ("D3.01", "Insert the provided photo, apply a picture style and remove its background.", ["5.1.2", "5.2.2", "5.2.3"], ["N14"], "desktop"),
        ("D3.02", "Apply an artistic effect to the second photo.", ["5.2.1"], [], "desktop"),
        ("D3.03", "Insert an icon and a shape; type a label in the shape; change its fill and outline.", ["5.1.1", "5.1.7", "5.3.2", "5.2.4"], ["N14"], "web"),
        ("D3.04", "Insert a four-step SmartArt process, enter the steps, and change its colors and style.", ["5.1.4", "5.3.3", "5.2.5"], [], "desktop"),
        ("D3.05", "Insert a text box pull quote; set Square wrapping and position it top right.", ["5.1.6", "5.3.1", "5.4.1", "5.4.2"], [], "web"),
        ("D3.06", "Insert a screen clipping of the course page and a 3D model; rotate the model.", ["5.1.5", "5.1.3", "5.2.6"], [], "desktop"),
        ("D3.07", "Write alt text for every image and mark the divider line as decorative.", ["5.4.3"], [], "web"),
        ("D3.08", "Convert the tab-separated schedule to a table; convert the old price table back to text.", ["3.1.1", "3.1.2"], [], "desktop"),
        ("D3.09", "Create a 4 x 6 table; merge the title row; resize columns; sort by date; set cell margins; repeat the header row; split the table after row 10.", ["3.1.3", "3.2.1", "3.2.2", "3.2.3", "3.2.4", "3.2.5", "3.2.6"], ["N14"], "desktop"),
        ("D3.10", "Insert a column chart from the survey data and add the caption Figure 1.", [], ["E3.4"], "desktop"),
        ("D3.11", "Put the newsletter body in two columns with a column break; use a Continuous section break so the masthead stays one column.", ["2.3.1", "2.3.2", "2.3.3"], [], "desktop"),
        ("D3.12", "Create custom theme colors for the club; add a page border, a DRAFT watermark and a drop cap.", ["1.2.4"], ["E3.2"], "desktop"),
    ],
    "skillcheck": {"id": "SC3", "covers": ["1.2.4", "2.3.1", "3.1.1", "3.1.3", "3.2.1", "3.2.3", "3.2.6", "5.1.2", "5.1.4", "5.1.6", "5.2.2", "5.3.2", "5.4.1", "5.4.2", "5.4.3"],
                   "desc": "Timed, MOS-style: 10 tasks on a starter flyer, auto-checked, 2 attempts."},
    "project": {
        "id": "P3", "title": "Newsletter or Event Flyer",
        "desc": "A two-page newsletter or event flyer for a real campus club, small business or community group: masthead, two-column body, at least three graphics with alt text, a SmartArt graphic, a schedule table with a header row, a captioned chart and a consistent theme. It must pass the Accessibility Checker with zero errors.",
        "criteria": [
            ("P3-a", "Graphics are inserted, formatted and positioned on purpose, not by accident.", ["CLO3"], 20),
            ("P3-b", "The table has a real header row that repeats, is sorted and formatted, and does not use merged cells in ways that confuse reading order.", ["CLO4"], 20),
            ("P3-c", "Columns, sections, theme and page elements create a clear visual hierarchy.", ["CLO5"], 25),
            ("P3-d", "Accessibility: meaningful alt text, sufficient contrast, logical reading order, Accessibility Checker clean.", ["CLO3", "CLO5"], 20),
            ("P3-e", "The chart supports the message and has a caption.", ["CLO3"], 15),
        ],
    },
    "ai": "Optional: AI image tools and automatic alt text are quick but not always right. If you use an AI image, say so, check the tool's terms, and rewrite any automatic alt text so it says what matters in context.",
    "platform": "Pictures, shapes, icons, text boxes, simple tables, columns, page color, alt text and the Accessibility Checker work in Word for the web. SmartArt, charts, 3D models, watermarks and WordArt need desktop Word, per Microsoft's feature comparison.",
    "nace": ["Communication", "Equity & Inclusion", "Technology"],
    "digcomp": ["Content creation"],
    "wcag": ["1.1.1", "1.3.1", "1.3.2", "1.4.1", "1.4.3"],
    "life": "A flyer or newsletter for any event or group, built so everyone can read it.",
    "hours": "About 9 hours: 2 h lessons, 3.5 h drills, 1 h skill check, 2.5 h project",
  },
  {
    "id": "W4", "n": 4, "title": "The Job Application Packet",
    "tagline": "Templates, mail merge and web output, built around getting hired.",
    "clos": {"CLO7": "IA", "CLO6": "A", "CLO1": "RA", "CLO2": "R", "CLO4": "R", "CLO5": "R"},
    "topics": ["T5", "T6", "T7"],
    "why": {
        "college": "Internship, scholarship and transfer applications need a resume and tailored letters.",
        "work": "Mail merge powers customer letters, certificates, name badges and mailing labels in every office.",
        "life": "Job searches, holiday card labels and organization mailings - done once, reused for years.",
    },
    "objectives": [
        ("W4.1", "Customize a built-in or online template, save your own reusable template (.dotx) and store reusable content as a Quick Part.", ["CLO2"]),
        ("W4.2", "Build a resume in two versions: a designed version and a plain, applicant tracking system (ATS) friendly version.", ["CLO1", "CLO4", "CLO5"]),
        ("W4.3", "Create a recipient data source and run a mail merge for personalized form letters, with preview and an If...Then...Else rule.", ["CLO7"]),
        ("W4.4", "Use mail merge to produce envelopes, mailing labels and a directory.", ["CLO7"]),
        ("W4.5", "Save a document as a web page, preview it in Web Layout and a browser, and compare web page, PDF and share link as ways to publish online.", ["CLO6"]),
    ],
    "drills": [
        ("D4.01", "Open a resume template; replace every placeholder content control; save as Lastname_Resume.dotx.", [], ["E1.1", "E4.1"], "desktop"),
        ("D4.03", "In the ATS version, align dates at the right margin with a right tab stop and dot leader.", [], [], "desktop"),
        ("D4.02", "Save your name and contact line as a Quick Part and insert it into the cover letter.", [], ["E3.1"], "desktop"),
        ("D4.04", "Create a recipient list of five employers (Company, Contact, Title, Address, City, State, ZIP, Position).", [], ["E4.3"], "desktop"),
        ("D4.05", "Insert an Address Block, a Greeting Line and the Position field in the letter body.", [], ["E4.3"], "desktop"),
        ("D4.06", "Add an If...Then...Else rule that changes one sentence for remote jobs; preview every record and fix any blank lines.", [], ["E4.3", "E4.1"], "desktop"),
        ("D4.07", "Finish the merge to a new document with one letter per recipient.", [], ["E4.3"], "desktop"),
        ("D4.08", "Create envelopes and a sheet of Avery 5160 labels from the same list.", [], ["E4.3"], "desktop"),
        ("D4.09", "Create a directory merge that lists every employer and contact on one page.", [], ["E4.3"], "desktop"),
        ("D4.10", "Add a hyperlink to your portfolio or LinkedIn with descriptive link text; export the resume to PDF.", ["1.3.1"], ["N14"], "web"),
        ("D4.11", "Save the resume as Web Page, Filtered (.htm); view it in Web Layout and in a browser; note two things that changed.", ["1.3.1"], [], "desktop"),
    ],
    "skillcheck": {"id": "SC4", "covers": ["1.3.1"],
                   "desc": "Timed merge check: given a data source and letter, produce merged letters, labels and a directory. Mostly MO-111 (Expert) skills."},
    "project": {
        "id": "P4", "title": "Job Application Packet",
        "desc": "A designed resume and an ATS-friendly resume (.docx and .pdf), cover letters merged to at least three employers (real postings or a fictional persona), envelopes or labels, a directory of contacts, and a web page version of the resume with a short reflection comparing web page, PDF and share link.",
        "criteria": [
            ("P4-a", "Resume content and layout are professional; a personal template is saved and reused.", ["CLO1", "CLO2", "CLO5"], 20),
            ("P4-b", "The ATS version uses simple structure (no text boxes or layout tables; key details not hidden in headers).", ["CLO1", "CLO4"], 15),
            ("P4-c", "Mail merge is correct: data source, fields, rule, preview, no merge errors.", ["CLO7"], 25),
            ("P4-d", "Envelopes or labels and a directory are produced from the same data source.", ["CLO7"], 20),
            ("P4-e", "Web page version is created and previewed, and the reflection compares output formats.", ["CLO6"], 20),
        ],
    },
    "ai": "Optional: AI can tailor a cover letter to a job post quickly. Your judgment decides what is true about you - never claim a skill you do not have. Keep other people's personal data out of AI prompts.",
    "privacy": "Learners may use a fictional persona instead of their own details. Recipient lists hold personal data: store them in the course folder only and delete them when the project is graded.",
    "platform": "Mail merge, envelopes, labels, Quick Parts and Save as Web Page need desktop Word. Word for the web can export PDF and create share links.",
    "nace": ["Career & Self-Development", "Professionalism", "Technology"],
    "digcomp": ["Content creation", "Safety, wellbeing and responsible use"],
    "wcag": ["2.4.4"],
    "life": "A complete job application packet, plus mail merge for any mailing.",
    "hours": "About 9 hours: 2 h lessons, 3 h drills, 1 h skill check, 3 h project",
  },
  {
    "id": "W5", "n": 5, "title": "Working Together and Shipping Clean",
    "tagline": "Review, collaborate, then share safely. Capstone and certification practice.",
    "clos": {"CLO8": "IA", "CLO1": "RA", "CLO2": "RA", "CLO3": "A", "CLO4": "A", "CLO5": "A", "CLO6": "R", "CLO7": "R"},
    "topics": ["T7"],
    "why": {
        "college": "Group projects and instructor feedback arrive as comments and tracked changes.",
        "work": "Contracts, policies and reports go through review cycles with several people editing at once.",
        "life": "Sharing a document safely - without hidden comments or personal data - protects you.",
    },
    "objectives": [
        ("W5.1", "Add, reply to, resolve and delete comments, including @mentions in shared documents.", ["CLO8"]),
        ("W5.2", "Track changes, review them, accept or reject them, and lock tracking with a password.", ["CLO8"]),
        ("W5.3", "Compare two versions of a document and combine edits from multiple reviewers.", ["CLO8"]),
        ("W5.4", "Share documents with appropriate permissions, co-author in real time and restore an earlier version from version history.", ["CLO8"]),
        ("W5.5", "Inspect a document before sharing: remove hidden data and personal information, fix accessibility and compatibility issues, and restrict editing.", ["CLO1", "CLO8"]),
        ("W5.6", "Show exam readiness on a timed MO-110 style practice test and assemble a course portfolio.", ["CLO1", "CLO2", "CLO3", "CLO4", "CLO5", "CLO8"]),
    ],
    "drills": [
        ("D5.01", "Add three comments; reply to one, resolve one and delete one.", ["6.1.1", "6.1.2", "6.1.3", "6.1.4"], [], "web"),
        ("D5.02", "Turn on Track Changes, make four edits, then lock tracking with a password.", ["6.2.1", "6.2.4"], [], "desktop"),
        ("D5.03", "Review all tracked changes; accept the formatting changes and reject one deletion.", ["6.2.2", "6.2.3"], [], "web"),
        ("D5.04", "Compare the original and revised cover letter; combine two reviewers' copies into one.", [], ["E1.1"], "desktop"),
        ("D5.05", "Share a copy as view-only with an expiration date; co-edit with a partner; restore a previous version.", ["1.3.4"], ["E1.1"], "both-check"),
        ("D5.06", "Run Inspect Document and remove comments, personal information and hidden text; run Check Compatibility.", ["1.4.1", "1.4.3"], [], "desktop"),
        ("D5.07", "Run the Accessibility Checker and fix every error and warning.", ["1.4.2"], [], "web"),
        ("D5.08", "Restrict editing to comments only, protected with a password.", [], ["E1.2"], "desktop"),
    ],
    "skillcheck": {"id": "SC5", "covers": ["1.3.4", "1.4.1", "1.4.2", "1.4.3", "6.1.1", "6.1.2", "6.1.3", "6.1.4", "6.2.1", "6.2.2", "6.2.3", "6.2.4"],
                   "desc": "Timed, MOS-style: 8 tasks on a reviewed document, auto-checked, 2 attempts."},
    "project": {
        "id": "P5", "title": "Peer Review Exchange and Clean-and-Ship",
        "desc": "Swap the cover letter or newsletter with a partner (or use the provided persona documents). Review with Track Changes and comments, respond to every comment, accept or reject each change, compare and combine two reviewers' copies, then run the clean-and-ship checklist: inspect, accessibility, compatibility, PDF export and a view-only share link. Finish with a short reflection memo.",
        "criteria": [
            ("P5-a", "Feedback to the partner is specific and constructive, using comments and tracked edits.", ["CLO8"], 25),
            ("P5-b", "Every comment received gets a reply or resolution; every change is accepted or rejected on purpose.", ["CLO8"], 25),
            ("P5-c", "Compare and Combine are done correctly.", ["CLO8"], 20),
            ("P5-d", "Clean-and-ship checklist is complete: Document Inspector, Accessibility Checker, compatibility, permissions.", ["CLO8", "CLO1"], 30),
        ],
    },
    "capstone": {
        "id": "CAP", "title": "Capstone: Practice Exam and Portfolio",
        "desc": "A timed, 50-minute MO-110 style practice exam covering all six domains (Certiport practice tests if FLCC has access; otherwise the Office Trainer capstone set), plus a portfolio of the five projects with a short reflection on which skills you will use next and where.",
        "clos": ["CLO1", "CLO2", "CLO3", "CLO4", "CLO5", "CLO6", "CLO7", "CLO8"],
    },
    "ai": "Optional: Copilot and similar tools can summarize comments or suggest edits. Weigh those suggestions the way you would a human reviewer's, and keep Track Changes on so the history stays honest.",
    "privacy": "Peer review shares real work. Learners choose which document to share and may use the persona documents instead. The Document Inspector step teaches what hidden data travels with a file.",
    "platform": "Comments, Track Changes review, sharing, version history and the Accessibility Checker work in Word for the web. Compare and Combine, Document Inspector and password-protected tracking and restrictions need desktop Word (the web shows restricted documents read-only).",
    "nace": ["Teamwork", "Communication", "Leadership", "Professionalism", "Technology"],
    "digcomp": ["Communication and collaboration", "Safety, wellbeing and responsible use"],
    "wcag": ["1.1.1", "1.3.1", "1.3.2", "1.4.3", "2.4.2", "3.1.1"],
    "life": "Safe sharing and feedback habits for any group document - plus a certification attempt if wanted.",
    "hours": "About 9 hours: 1.5 h lessons, 2 h drills, 1 h skill check, 2.5 h project, 2 h capstone",
  },
]

# ---------------------------------------------------------------------------
# Grading (suggested, adjustable)
# ---------------------------------------------------------------------------
GRADING = [
    ("Practice drills (completion, unlimited tries)", 15),
    ("Weekly skill checks SC1-SC5 (4% each)", 20),
    ("Projects P1-P4 (10% each)", 40),
    ("Project P5 Peer Review Exchange", 10),
    ("Capstone practice exam and portfolio", 15),
]

# ---------------------------------------------------------------------------
# Other standards
# ---------------------------------------------------------------------------
NORTHSTAR = [
    ("N1", "Open a new or existing document.", "W1", "D1.01"),
    ("N2", "Identify the parts of the Word window, including the Ribbon, Status Bar and Quick Access Toolbar.", "W1", "D1.02"),
    ("N3", "Save a document, being intentional about name and location.", "W1", "D1.01"),
    ("N4", "Identify file extensions that can be opened by Microsoft Word.", "W1", "D1.01"),
    ("N5", "Use Spelling and Grammar check.", "W1", "D1.09"),
    ("N6", "Format text: size, color and font type.", "W1", "D1.06"),
    ("N7", "Set text spacing and alignment.", "W1", "D1.07"),
    ("N8", "Apply bullets and automatic numbering.", "W2", "D2.05"),
    ("N9", "Undo the previous action.", "W1", "D1.08"),
    ("N10", "Cut, copy and paste.", "W1", "D1.08"),
    ("N11", "Modify page layout, including margins and orientation.", "W1", "D1.09"),
    ("N12", "Print.", "W1", "D1.11"),
    ("N13", "Close a document.", "W1", "D1.11"),
    ("N14", "Insert objects into a document, including images, shapes, hyperlinks, and tables.", "W3", "D3.01, D3.03, D3.09, D4.10"),
]

NACE = [
    ("Career & Self-Development", ["W4", "W5"], "Resume, cover letters and a portfolio reflection on next steps."),
    ("Communication", ["W1", "W2", "W3", "W5"], "Letters, memos, papers, newsletters and written peer feedback."),
    ("Critical Thinking", ["W2", "W5"], "Checking sources (including AI-suggested ones) and deciding which review edits to accept."),
    ("Equity & Inclusion", ["W3", "W5"], "Accessible documents: alt text, contrast, reading order and the Accessibility Checker. NACE lists this competency as under review; some campuses now label it Access & Opportunity."),
    ("Leadership", ["W5"], "Giving constructive feedback and resolving a review cycle."),
    ("Professionalism", ["W1", "W4", "W5"], "Business conventions, error-free work and safe sharing habits."),
    ("Teamwork", ["W5"], "Co-authoring, comments, Track Changes and combining reviewers' work."),
    ("Technology", ["W1", "W2", "W3", "W4", "W5"], "Every week: choosing the right Word feature and using it responsibly."),
]

DIGCOMP = [
    ("Information search, evaluation and management", ["W1", "W2"], "Saving and organizing files; finding and verifying sources for citations."),
    ("Communication and collaboration", ["W5"], "Sharing, co-authoring, comments and Track Changes."),
    ("Content creation", ["W1", "W2", "W3", "W4"], "Creating and formatting letters, papers, newsletters, resumes and web output."),
    ("Safety, wellbeing and responsible use", ["W4", "W5"], "Protecting personal data in merge lists and AI prompts; Document Inspector; permissions."),
    ("Problem identification and solving", ["W1", "W3", "W5"], "Using formatting marks, the Accessibility Checker and Compatibility Checker to find and fix problems."),
]

WCAG = [
    ("1.1.1", "Non-text Content", "A", "Alt text on pictures, charts and SmartArt; decorative items marked decorative.", ["W3", "W5"]),
    ("1.3.1", "Info and Relationships", "A", "Real heading styles, real lists and table header rows instead of visual formatting only.", ["W2", "W3", "W5"]),
    ("1.3.2", "Meaningful Sequence", "A", "Reading order stays logical when graphics float or text runs in columns.", ["W3", "W5"]),
    ("1.4.1", "Use of Color", "A", "Color is never the only way information is shown.", ["W3"]),
    ("1.4.3", "Contrast (Minimum)", "AA", "Text contrast at least 4.5:1 (3:1 for large text) on flyers and newsletters.", ["W3", "W5"]),
    ("2.4.2", "Page Titled", "A", "The document Title property is set (it becomes the PDF title).", ["W1", "W5"]),
    ("2.4.4", "Link Purpose (In Context)", "A", "Hyperlinks use descriptive text, not raw URLs or 'click here'.", ["W2", "W4"]),
    ("2.4.6", "Headings and Labels", "AA", "Headings describe the content that follows.", ["W2"]),
    ("3.1.1", "Language of Page", "A", "Document proofing language is set correctly.", ["W1", "W5"]),
]

# Word for the web vs desktop (from Microsoft's feature comparison table; it changes often)
PLATFORM = [
    ("Mail merge", "No", "Yes", "W4"),
    ("Envelopes and labels", "No", "Yes", "W4"),
    ("Compare and combine revisions", "No", "Yes", "W5"),
    ("Citations and bibliography", "No", "Yes", "W2"),
    ("Section breaks", "No", "Yes", "W2, W3"),
    ("SmartArt, charts and 3D models", "No", "Yes", "W3"),
    ("Macros", "No", "Yes", "Not taught"),
    ("Watermark", "No (page color only)", "Yes", "W3"),
    ("WordArt", "No (shapes and text boxes only)", "Yes", "W3"),
    ("Restricted editing", "Read only", "Yes", "W5"),
    ("Track Changes", "Yes", "Yes", "W5"),
    ("Styles", "Yes", "Yes", "W2"),
    ("Table of contents", "Yes", "Yes", "W2"),
    ("Footnotes and endnotes", "Yes", "Yes", "W2"),
    ("Columns", "Yes", "Yes", "W3"),
    ("Templates", "Yes", "Yes", "W2, W4"),
    ("Export as PDF", "Yes", "Yes", "W1, W4, W5"),
    ("Accessibility Checker", "Yes", "Yes", "W3, W5"),
]

# Conventions and assumptions (clearly labeled as such)
CONVENTIONS = [
    ("APA 7th edition / MLA 9th edition", "Published style guide", "Current editions for student papers; confirm with the FLCC Library before each term.", ["W2"]),
    ("Block-format business letter", "Industry convention", "The most common US business letter layout: everything flush left, single spaced, a blank line between parts.", ["W1"]),
    ("ATS-friendly resume", "Industry convention", "Applicant tracking systems parse simple layouts most reliably; keep a plain version alongside the designed one.", ["W4"]),
    ("Avery 5160 labels", "Industry convention", "The most common US address label size, built into Word's label options.", ["W4"]),
    ("Bloom's revised taxonomy", "Assumption", "Course outcome verbs sit mostly at Apply; peer review reaches Evaluate.", ["All"]),
    ("Credit hour (about 45 hours per credit)", "Assumption", "The 9 hours per week figure assumes 1 credit over 5 weeks. Scale if the course carries more credit.", ["All"]),
    ("I / R / A curriculum mapping", "Common practice", "Introduced, Reinforced, Assessed - the standard way to show where an outcome is taught and measured.", ["All"]),
]

# Decisions and flags for Dave
DECISIONS = [
    ("Reword CLO 6", "flag", "\"Create and preview the document as a Web page\" reflects an older workflow. The plan meets it as written (Web Layout view plus Save as Web Page in Week 4), then spends most of the time on what people do now: PDF export and share links. Consider rewording to \"Publish documents for screen and web audiences (web page, PDF, share link).\""),
    ("Make CLO 2 measurable", "flag", "\"Demonstrate features offered by the application\" is hard to assess on its own. It is mapped to every week here. Consider rewording to \"Use Word's structure and productivity features (styles, navigation, references, templates) to produce documents efficiently.\""),
    ("Some course outcomes go beyond MO-110", "info", "Mail merge, templates, compare and combine, citations and restrict editing are MO-111 (Expert) skills. They stay because the course outcomes require them, and they are labeled as beyond-the-certification content."),
    ("Plan for desktop Word", "flag", "Weeks 2 to 5 include features that need desktop Word (mail merge, compare and combine, SmartArt, section breaks, citations). Chromebook users will need campus labs or a virtual desktop for those tasks. Each drill is tagged web or desktop so the Office Trainer can show the right path."),
    ("Accessibility is built in, not bolted on", "info", "The ADA Title II web rule (WCAG 2.1 AA) covers documents public colleges post. A DOJ interim final rule (April 20, 2026) moved compliance to April 26, 2027 for entities of 50,000 or more and April 26, 2028 for smaller ones. Confirm FLCC's date with the college's accessibility office."),
    ("Certification access", "flag", "Check whether FLCC is a Certiport testing center and has practice-test licenses. If not, the capstone uses the Office Trainer's own MOS-style set and the exam is an optional add-on."),
    ("Privacy by design", "info", "Fictional persona option for resumes, letters and peer review; merge lists deleted after grading; Document Inspector taught before any sharing; no personal data in AI prompts."),
    ("Northstar as a diagnostic", "info", "Twelve of Northstar's 14 Word skills are covered in Week 1, lists in Week 2, and inserting images, shapes, hyperlinks and tables in Week 3. If FLCC has a Northstar site license, use it as a Day 1 pre-assessment and a Week 1 post-check."),
]

SOURCES = [
    ("MO-110: Microsoft Word (Microsoft 365 Apps) - Microsoft Learn", "https://learn.microsoft.com/en-us/credentials/exams/mo-110/"),
    ("MO-110 study guide (skills measured) - Microsoft", "https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE5dvir"),
    ("MO-111 Word Expert study guide - Microsoft", "https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RW16Vea"),
    ("Northstar Digital Literacy: Microsoft Word standards", "https://www.digitalliteracyassessment.org/badges/ms-word/v2/proctored/"),
    ("NACE Career Readiness: Competencies Defined - NACE", "https://www.naceweb.org/career-readiness/competencies/career-readiness-defined"),
    ("NACE Career Readiness Competencies - Rutgers-Newark summary", "https://careers.newark.rutgers.edu/resources/nace-career-readiness-competencies/"),
    ("Equity & Inclusion competency under review - Inside Higher Ed", "https://www.insidehighered.com/index%2Ephp/node/196261"),
    ("DigComp 3.0 launches (December 2025) - CEPIS", "https://cepis.org/digcomp-3-0-launches/"),
    ("DigComp framework - European Commission Joint Research Centre", "https://joint-research-centre.ec.europa.eu/digcomp/digcomp-framework_en"),
    ("Understanding WCAG 2.1 success criteria - W3C", "https://www.w3.org/WAI/WCAG21/Understanding/"),
    ("ADA Title II compliance deadline extension (2026) - Kwall", "https://www.kwallcompany.com/2026/06/04/ada-title-ii-website-compliance-deadline/"),
    ("ADA Title II 2026 update - Aberdeen", "https://aberdeen.io/blog/2026/04/07/ada-title-iis-2026-update-what-changed-and-what-didnt/"),
    ("Word features comparison: web vs desktop - Microsoft Support", "https://support.microsoft.com/en-us/Word/word-features-comparison-word-for-the-web-vs-desktop"),
]
