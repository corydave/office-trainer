# -*- coding: utf-8 -*-
"""
Single source of data for the 5-week Microsoft Excel course map (CSC 135 Core Excel).
build.py turns this into the Markdown course map and the interactive HTML page,
and runs coverage checks so nothing in the official outline gets dropped.
Style rule: hyphens, never em dashes.
The variable names MO110 / MO111 are kept from the Word build script; here they
hold the Excel exams (MO-210 Associate and MO-211 Expert).
"""

COURSE = {
    "title": "Microsoft Excel - 5-Week Course Map",
    "short": "Excel Course Map",
    "college": "Finger Lakes Community College (FLCC)",
    "version": "Draft 1",
    "date": "October 22, 2026",
    "author": "Dave Ghidiu (with Claude)",
    "format": "5 weeks, online or hybrid, about 9 hours of learner work per week (CSC 135 Core Excel is 1 credit)",
    "cert": "Microsoft Office Specialist (MOS): Excel Associate (Microsoft 365 Apps), exam MO-210",
}

# ---------------------------------------------------------------------------
# Official course description (CSC 135 Core Excel) - carried through verbatim
# ---------------------------------------------------------------------------
CLOS = [
    {"id": "CLO1", "text": "Create, save, format and print worksheets and workbooks", "bloom": "Apply"},
    {"id": "CLO2", "text": "Demonstrate features offered by the application", "bloom": "Apply"},
    {"id": "CLO3", "text": "Create a \"what-if\" scenario", "bloom": "Apply / Analyze"},
    {"id": "CLO4", "text": "Insert and modify graphics", "bloom": "Apply"},
    {"id": "CLO5", "text": "Create and format charts", "bloom": "Apply"},
    {"id": "CLO6", "text": "Create and preview the worksheet as a Web page", "bloom": "Apply"},
    {"id": "CLO7", "text": "Collaborate on workbooks: Compare and Combine workbooks", "bloom": "Apply / Evaluate"},
]

TOPICS = [
    {"id": "T1", "text": "Introduction to Microsoft Excel application features", "weeks": ["W1"]},
    {"id": "T2", "text": "Creating Worksheets and Workbooks with Microsoft Excel", "weeks": ["W1"]},
    {"id": "T3", "text": "Formatting and Organizing worksheets", "weeks": ["W1", "W3"]},
    {"id": "T4", "text": "Using Functions and Formulas", "weeks": ["W2", "W3"]},
    {"id": "T5", "text": "Special Worksheet Formats", "weeks": ["W3", "W5"]},
    {"id": "T6", "text": "Working with Templates, Styles and Charts", "weeks": ["W1", "W4"]},
    {"id": "T7", "text": "Collaboration and sharing", "weeks": ["W5"]},
]

TOPIC_HOW = {
    "T1": "Week 1 tour of the Ribbon, Backstage view, formula bar, Name Box, sheet tabs, status bar, Quick Access Toolbar and views.",
    "T2": "Week 1 create, name and save workbooks; add, rename, color and move worksheets; enter data with Auto Fill; Project 1 budget.",
    "T3": "Week 1 number formats, alignment, cell styles, column widths and page setup; Week 3 tables, sorting, filtering and freeze panes on real data.",
    "T4": "Week 2 formulas, cell references, everyday functions, names and text functions; Week 3 structured references, SORT() and UNIQUE().",
    "T5": "Week 3 Excel tables, conditional formatting and sparklines; Week 5 notes and comments, protected sheets and web page output.",
    "T6": "Week 1 cell styles; Week 4 charts, chart sheets, chart styles, themes and a reusable workbook template (.xltx).",
    "T7": "Week 5 comments, co-authoring, version history, combining reviewers' workbooks, inspecting and sharing.",
}

# ---------------------------------------------------------------------------
# MO-210 skills outline (Excel Associate, Microsoft 365 Apps). Wording follows
# Microsoft's objective domain document. Microsoft's outline doesn't number the
# sub-skills; the numbers here follow its order. Verify against the current PDF.
# ---------------------------------------------------------------------------
MO110 = [
    {"id": "1", "name": "Manage worksheets and workbooks", "weight": "see outline", "objectives": [
        {"id": "1.1", "name": "Import data into workbooks", "skills": [
            ("1.1.1", "Import data from text files"),
            ("1.1.2", "Import data from online sources")]},
        {"id": "1.2", "name": "Navigate within workbooks", "skills": [
            ("1.2.1", "Search for data within a workbook"),
            ("1.2.2", "Navigate to named cells, ranges, or workbook elements"),
            ("1.2.3", "Insert and remove hyperlinks")]},
        {"id": "1.3", "name": "Format worksheets and workbooks", "skills": [
            ("1.3.1", "Modify page setup"),
            ("1.3.2", "Adjust row height and column width"),
            ("1.3.3", "Customize headers and footers")]},
        {"id": "1.4", "name": "Customize options and views", "skills": [
            ("1.4.1", "Manage the Quick Access toolbar"),
            ("1.4.2", "Display and modify worksheets in different views"),
            ("1.4.3", "Freeze worksheet rows and columns"),
            ("1.4.4", "Change window views"),
            ("1.4.5", "Modify built-in workbook properties"),
            ("1.4.6", "Display formulas")]},
        {"id": "1.5", "name": "Prepare workbooks for collaboration and distribution", "skills": [
            ("1.5.1", "Set a print area"),
            ("1.5.2", "Save and export workbooks in alternative file formats"),
            ("1.5.3", "Configure print settings"),
            ("1.5.4", "Inspect workbooks and correct issues"),
            ("1.5.5", "Manage comments and notes")]},
    ]},
    {"id": "2", "name": "Manage data cells and ranges", "weight": "see outline", "objectives": [
        {"id": "2.1", "name": "Manipulate data in worksheets", "skills": [
            ("2.1.1", "Paste data by using special paste options"),
            ("2.1.2", "Fill cells by using Auto Fill"),
            ("2.1.3", "Insert and delete multiple columns or rows"),
            ("2.1.4", "Insert and delete cells"),
            ("2.1.5", "Generate numeric data by using RANDBETWEEN() and SEQUENCE()")]},
        {"id": "2.2", "name": "Format cells and ranges", "skills": [
            ("2.2.1", "Merge and unmerge cells"),
            ("2.2.2", "Modify cell alignment, orientation, and indentation"),
            ("2.2.3", "Format cells by using Format Painter"),
            ("2.2.4", "Wrap text within cells"),
            ("2.2.5", "Apply number formats"),
            ("2.2.6", "Apply cell formats from the Format Cells dialog box"),
            ("2.2.7", "Apply cell styles"),
            ("2.2.8", "Clear cell formatting"),
            ("2.2.9", "Format multiple worksheets by grouping")]},
        {"id": "2.3", "name": "Define and reference named ranges", "skills": [
            ("2.3.1", "Define a named range"),
            ("2.3.2", "Reference a named range")]},
        {"id": "2.4", "name": "Summarize data visually", "skills": [
            ("2.4.1", "Insert Sparklines"),
            ("2.4.2", "Apply built-in conditional formatting"),
            ("2.4.3", "Remove conditional formatting")]},
    ]},
    {"id": "3", "name": "Manage tables and table data", "weight": "see outline", "objectives": [
        {"id": "3.1", "name": "Create and format tables", "skills": [
            ("3.1.1", "Create Excel tables from cell ranges"),
            ("3.1.2", "Apply table styles"),
            ("3.1.3", "Convert tables to cell ranges")]},
        {"id": "3.2", "name": "Modify tables", "skills": [
            ("3.2.1", "Add or remove table rows and columns"),
            ("3.2.2", "Configure table style options"),
            ("3.2.3", "Insert and configure total rows")]},
        {"id": "3.3", "name": "Filter and sort table data", "skills": [
            ("3.3.1", "Filter records"),
            ("3.3.2", "Sort data by multiple columns")]},
    ]},
    {"id": "4", "name": "Perform operations by using formulas and functions", "weight": "see outline", "objectives": [
        {"id": "4.1", "name": "Insert references", "skills": [
            ("4.1.1", "Insert relative, absolute, and mixed references"),
            ("4.1.2", "Use structured references in formulas")]},
        {"id": "4.2", "name": "Calculate and transform data", "skills": [
            ("4.2.1", "Perform calculations by using the AVERAGE(), MAX(), MIN(), and SUM() functions"),
            ("4.2.2", "Count cells by using the COUNT(), COUNTA(), and COUNTBLANK() functions"),
            ("4.2.3", "Perform conditional operations by using the IF() function"),
            ("4.2.4", "Sort data by using the SORT() function"),
            ("4.2.5", "Get unique values by using the UNIQUE() function")]},
        {"id": "4.3", "name": "Format and modify text", "skills": [
            ("4.3.1", "Format text by using RIGHT(), LEFT(), and MID() functions"),
            ("4.3.2", "Format text by using UPPER(), LOWER(), and LEN() functions"),
            ("4.3.3", "Format text by using the CONCAT() and TEXTJOIN() functions")]},
    ]},
    {"id": "5", "name": "Manage charts", "weight": "see outline", "objectives": [
        {"id": "5.1", "name": "Create charts", "skills": [
            ("5.1.1", "Create charts"),
            ("5.1.2", "Create chart sheets")]},
        {"id": "5.2", "name": "Modify charts", "skills": [
            ("5.2.1", "Add data series to charts"),
            ("5.2.2", "Switch between rows and columns in source data"),
            ("5.2.3", "Add and modify chart elements")]},
        {"id": "5.3", "name": "Format charts", "skills": [
            ("5.3.1", "Apply chart layouts"),
            ("5.3.2", "Apply chart styles"),
            ("5.3.3", "Add alternative text to charts for accessibility")]},
    ]},
]

# MO-211 Excel Expert - stretch content the official outcomes pull in.
MO111 = [
    {"id": "1.1", "domain": "Manage workbook options and settings", "name": "Manage workbooks",
     "skills": "Copy macros between workbooks, reference data in other workbooks, enable macros, manage workbook versions",
     "status": "partial", "weeks": ["W5"],
     "how": "Reference data in other workbooks (combining reviewers' copies) and version history (W5). Macros are not taught."},
    {"id": "1.2", "domain": "Manage workbook options and settings", "name": "Prepare workbooks for collaboration",
     "skills": "Restrict editing, protect worksheets and cell ranges, protect workbook structure, configure formula calculation options",
     "status": "partial", "weeks": ["W5"],
     "how": "Protect a sheet with unlocked input cells, and protect the workbook structure (W5). Calculation options are not taught."},
    {"id": "2.1", "domain": "Manage and format data", "name": "Fill cells based on existing data",
     "skills": "Flash Fill, advanced Fill Series options, RANDARRAY()",
     "status": "partial", "weeks": ["W1", "W2"], "how": "Fill Series with a step value (W1) and Flash Fill to split names (W2)."},
    {"id": "2.2", "domain": "Manage and format data", "name": "Format and validate data",
     "skills": "Custom number formats, data validation, group and ungroup data, subtotals, remove duplicates",
     "status": "partial", "weeks": ["W3", "W5"], "how": "Remove duplicates after an import (W3) and data validation on a shared sign-up sheet (W5)."},
    {"id": "2.3", "domain": "Manage and format data", "name": "Apply advanced conditional formatting and filtering",
     "skills": "Custom rules, formula-based rules, managing rules",
     "status": "none", "weeks": [], "how": "Not taught. Built-in rules (MO-210) cover the course's needs; optional enrichment."},
    {"id": "3.1", "domain": "Create advanced formulas and macros", "name": "Perform logical operations in formulas",
     "skills": "Nested IF(), IFS(), SWITCH(), SUMIF(), AVERAGEIF(), COUNTIF() and similar",
     "status": "none", "weeks": [], "how": "Not taught; IF() alone is MO-210. COUNTIF() appears once as an optional extension in Project 3."},
    {"id": "3.2", "domain": "Create advanced formulas and macros", "name": "Look up data by using functions",
     "skills": "XLOOKUP(), VLOOKUP(), HLOOKUP(), MATCH(), INDEX()",
     "status": "none", "weeks": [], "how": "Not taught. A strong candidate for enrichment; CSC 212 covers it."},
    {"id": "3.3", "domain": "Create advanced formulas and macros", "name": "Use advanced date and time functions",
     "skills": "NOW(), TODAY(), WEEKDAY(), WORKDAY()",
     "status": "partial", "weeks": ["W2"], "how": "TODAY() on an invoice date (W2)."},
    {"id": "3.4", "domain": "Create advanced formulas and macros", "name": "Perform data analysis",
     "skills": "Consolidate, Goal Seek and Scenario Manager, NPER(), PMT(), FILTER(), SORTBY()",
     "status": "partial", "weeks": ["W4", "W5"], "how": "PMT(), Goal Seek and Scenario Manager for the what-if outcome (W4); Consolidate to combine reviewers' workbooks (W5)."},
    {"id": "3.5", "domain": "Create advanced formulas and macros", "name": "Troubleshoot formulas",
     "skills": "Trace precedents and dependents, Watch Window, error checking, Evaluate Formula",
     "status": "partial", "weeks": ["W2"], "how": "Error checking and Trace Precedents when a formula goes wrong (W2)."},
    {"id": "3.6", "domain": "Create advanced formulas and macros", "name": "Create and modify simple macros",
     "skills": "Record, name and edit macros",
     "status": "none", "weeks": [], "how": "Not taught. Macros are a security risk to explain, not a beginner skill to build."},
    {"id": "4.1", "domain": "Manage advanced charts and tables", "name": "Create and modify advanced charts",
     "skills": "Dual-axis, Box & Whisker, Combo, Funnel, Histogram, Sunburst and Waterfall charts",
     "status": "none", "weeks": [], "how": "Not taught; the course stays with column, bar, line and pie charts."},
    {"id": "4.2", "domain": "Manage advanced charts and tables", "name": "Create and modify PivotTables",
     "skills": "PivotTables, fields, slicers, grouping, calculated fields, value field settings",
     "status": "none", "weeks": [], "how": "Not taught in this 1-credit course; covered in CSC 212. Flagged as possible enrichment."},
    {"id": "4.3", "domain": "Manage advanced charts and tables", "name": "Create and modify PivotCharts",
     "skills": "PivotCharts, options, styles, drill down",
     "status": "none", "weeks": [], "how": "Not taught; see PivotTables."},
]

# ---------------------------------------------------------------------------
# Weeks
# drill fields: id, task, mos (MO-210 codes), x (MO-211 as "E3.4", Northstar "N3"),
#   where ("web" = works in Excel for the web, "desktop" = needs desktop Excel,
#   "both-check" = works anywhere but needs a screenshot or self-check)
# ---------------------------------------------------------------------------
WEEKS = [
  {
    "id": "W1", "n": 1, "title": "The Working Spreadsheet",
    "tagline": "Build, format and print a clean, readable workbook.",
    "clos": {"CLO1": "IA", "CLO2": "IA", "CLO6": "I"},
    "topics": ["T1", "T2", "T3", "T6"],
    "why": {
        "college": "Lab data, club budgets and group schedules start as a spreadsheet someone else has to read.",
        "work": "Nearly every job posting that lists software lists Excel; clean, printable sheets are the baseline.",
        "life": "A monthly budget, a move checklist or a fantasy league is easier to keep up when it's organized.",
    },
    "objectives": [
        ("W1.1", "Identify the parts of the Excel window (Ribbon, formula bar, Name Box, active cell, sheet tabs, status bar, Quick Access Toolbar) and switch among Normal, Page Layout and Page Break Preview views.", ["CLO2", "CLO6"]),
        ("W1.2", "Create, name and save workbooks, and add, rename, color, move and delete worksheets.", ["CLO1"]),
        ("W1.3", "Enter and edit data with Auto Fill, insert and delete rows, columns and cells, use Paste Special, and search with Find and Replace.", ["CLO1", "CLO2"]),
        ("W1.4", "Format cells: number formats, alignment, orientation, indent, wrap, merge, borders and fill, cell styles, Format Painter and Clear Formats, including several sheets at once.", ["CLO1", "CLO2"]),
        ("W1.5", "Freeze panes, adjust rows and columns, set up the page (orientation, scaling, margins, print area, headers and footers), set properties and print or export a PDF.", ["CLO1"]),
    ],
    "drills": [
        ("D1.01", "Open the starter workbook and save it to your course folder as Lastname_E1_Rentals.xlsx.", [], ["N1", "N2"], "web"),
        ("D1.02", "Rename Sheet1 to June, add a sheet named July, give June a tab color and move July after June.", [], ["N11"], "web"),
        ("D1.03", "Add Save As and Print Preview to the Quick Access Toolbar, then name four parts of the window in the drill quiz.", ["1.4.1"], ["N3"], "desktop"),
        ("D1.04", "Freeze the top row and first column, switch to Page Layout view, then open a second window of the workbook and arrange them side by side.", ["1.4.2", "1.4.3", "1.4.4"], ["N8"], "desktop"),
        ("D1.05", "Use Find to count every single-kayak rental, then replace every Kayak (single) with Solo kayak.", ["1.2.1"], ["N4"], "web"),
        ("D1.06", "Auto Fill the month names and a numbered series; insert two columns and a row; delete a misplaced cell and shift cells up.", ["2.1.2", "2.1.3", "2.1.4"], ["N5", "N9", "N14", "E2.1"], "web"),
        ("D1.07", "Paste a column as values only and transpose the price list with Paste Special.", ["2.1.1"], ["N6"], "web"),
        ("D1.08", "Apply Currency, Percentage and Short Date formats; add borders and a fill from the Format Cells dialog box.", ["2.2.5", "2.2.6"], ["N7"], "web"),
        ("D1.09", "Merge and center the title, wrap the long headers, angle the month labels and indent the category rows.", ["2.2.1", "2.2.2", "2.2.4"], ["N7"], "web"),
        ("D1.10", "Apply the Title and Heading 1 cell styles, copy the header format with Format Painter, and clear the stray formatting in the notes area.", ["2.2.3", "2.2.7", "2.2.8"], [], "web"),
        ("D1.11", "Set exact column widths and row heights, then group June and July and format both header rows at once.", ["1.3.2", "2.2.9"], ["N10"], "desktop"),
        ("D1.12", "Set landscape orientation, fit to one page wide, a print area, and a footer with the sheet name and page number.", ["1.3.1", "1.3.3", "1.5.1", "1.5.3"], ["N17"], "desktop"),
        ("D1.13", "Fill in the Title, Author and Tags properties, then export the sheet to PDF and check it.", ["1.4.5", "1.5.2"], ["N13"], "both-check"),
    ],
    "skillcheck": {"id": "SC1", "covers": ["1.2.1", "1.3.1", "1.3.2", "1.3.3", "1.4.3", "1.4.5", "1.5.1", "2.1.1", "2.1.2", "2.1.3", "2.2.1", "2.2.2", "2.2.4", "2.2.5", "2.2.7"],
                   "desc": "Timed, MOS-style: 12 tasks in a fresh starter workbook, auto-checked by the Office Trainer, 2 attempts. The PDF export (1.5.2) is assessed through the Project 1 PDF instead, because a check cannot see files outside the workbook."},
    "project": {
        "id": "P1", "title": "Monthly Budget Workbook",
        "desc": "A one-sheet monthly budget for yourself, a household, a club or a small business (made-up numbers are fine): income and expense sections, consistent number formats and styles, frozen headings, and a print-ready page with a footer. Submit .xlsx and .pdf.",
        "criteria": [
            ("P1-a", "The budget is organized: clear sections, labeled rows and columns, sensible sheet name, no blank spacer rows inside the data.", ["CLO1"], 25),
            ("P1-b", "Formatting uses Excel tools on purpose: number formats, cell styles, alignment, column widths (no spaces or merged cells as workarounds).", ["CLO1", "CLO2"], 25),
            ("P1-c", "Navigation and views help the reader: frozen headings and a second sheet for notes or a second month.", ["CLO2"], 20),
            ("P1-d", "Print-ready: page setup, print area, footer and properties are set; the PDF matches the workbook.", ["CLO1"], 30),
        ],
    },
    "ai": "Optional: ask an AI tool to suggest budget categories for your situation, then decide which ones fit. Never paste real account numbers or balances into an AI tool.",
    "privacy": "Budgets are personal. Made-up or rounded numbers are always fine, and nothing in the course asks for real account details.",
    "platform": "Most of this week works in Excel for the web. The Quick Access Toolbar, Page Layout view, side-by-side windows, grouping sheets and full page setup are easier or only available in desktop Excel.",
    "nace": ["Professionalism", "Technology", "Communication"],
    "digcomp": ["Content creation", "Information search, evaluation and management"],
    "wcag": ["1.3.1", "2.4.2"],
    "life": "A budget and a clean, printable spreadsheet you can reuse every month.",
    "hours": "About 9 hours: 2 h lessons and videos, 3 h drills, 1 h skill check, 3 h project",
  },
  {
    "id": "W2", "n": 2, "title": "Formulas That Work",
    "tagline": "Let Excel do the math: formulas, references, everyday functions and names.",
    "clos": {"CLO2": "RA", "CLO1": "R"},
    "topics": ["T4"],
    "why": {
        "college": "Grade calculators, lab averages and statistics homework are formulas you can check.",
        "work": "Invoices, timesheets and price lists depend on formulas that stay right when the numbers change.",
        "life": "Splitting bills, comparing phone plans and tracking savings goals take a few good formulas.",
    },
    "objectives": [
        ("W2.1", "Write formulas with operators and parentheses, and use relative, absolute and mixed references so formulas fill correctly.", ["CLO2"]),
        ("W2.2", "Summarize data with SUM(), AVERAGE(), MIN(), MAX(), COUNT(), COUNTA() and COUNTBLANK(), and make decisions with IF().", ["CLO2"]),
        ("W2.3", "Define and use named ranges, navigate with the Name Box, and insert and remove hyperlinks.", ["CLO2"]),
        ("W2.4", "Clean and combine text with LEFT(), RIGHT(), MID(), UPPER(), LOWER(), LEN(), CONCAT() and TEXTJOIN(), and generate data with SEQUENCE() and RANDBETWEEN().", ["CLO2"]),
        ("W2.5", "Display formulas and use error checking to find and fix formula mistakes.", ["CLO1", "CLO2"]),
    ],
    "drills": [
        ("D2.01", "Write line-total and subtotal formulas with +, -, * and /, using parentheses where the order matters.", [], ["N12"], "web"),
        ("D2.02", "Calculate sales tax from one tax-rate cell with an absolute reference, and fill it down.", ["4.1.1"], [], "web"),
        ("D2.03", "Build a price-by-quantity grid with mixed references in one formula filled across and down.", ["4.1.1"], [], "web"),
        ("D2.04", "Display formulas to find the hard-typed total, then fix the #DIV/0! error the error checker flags.", ["1.4.6"], ["E3.5"], "web"),
        ("D2.05", "Total, average and find the highest and lowest weekly rentals with AutoSum and the functions.", ["4.2.1"], ["N14"], "web"),
        ("D2.06", "Count the numeric entries, the filled cells and the blank cells in the sign-up list.", ["4.2.2"], [], "web"),
        ("D2.07", "Use IF() to label each rental Late or On time.", ["4.2.3"], [], "web"),
        ("D2.08", "Name the tax-rate cell TaxRate, use the name in a formula, and jump to a named range with the Name Box.", ["2.3.1", "2.3.2", "1.2.2"], [], "web"),
        ("D2.09", "Link the summary to the Details sheet and to the company website; remove the broken link.", ["1.2.3"], [], "web"),
        ("D2.10", "Pull the first initial and last four digits out of member codes, and fix name capitalization and lengths.", ["4.3.1", "4.3.2"], ["E2.1"], "web"),
        ("D2.11", "Build full names and an email list with CONCAT() and TEXTJOIN().", ["4.3.3"], [], "web"),
        ("D2.12", "Number the invoice lines with SEQUENCE() and make practice data with RANDBETWEEN().", ["2.1.5"], ["E3.3"], "web"),
    ],
    "skillcheck": {"id": "SC2", "covers": ["1.2.2", "1.2.3", "1.4.6", "2.1.5", "2.3.1", "2.3.2", "4.1.1", "4.2.1", "4.2.2", "4.2.3", "4.3.1", "4.3.2", "4.3.3"],
                   "desc": "Timed, MOS-style: 13 tasks on an order-tracking workbook, auto-checked, 2 attempts."},
    "project": {
        "id": "P2", "title": "Invoice and Grade Calculator",
        "desc": "Two sheets: a reusable invoice (line totals, subtotal, tax from a named rate, total, invoice numbers built with text functions) and a course grade calculator with weighted categories and an IF() status. Use made-up names and scores.",
        "criteria": [
            ("P2-a", "Every calculated cell is a formula; nothing is typed that Excel should calculate.", ["CLO2"], 25),
            ("P2-b", "References are right: absolute or named for fixed values, relative for filled formulas; the sheets still work when inputs change.", ["CLO2"], 25),
            ("P2-c", "Functions fit the job: SUM, AVERAGE, COUNT family, IF and text functions used correctly.", ["CLO2"], 25),
            ("P2-d", "The workbook is readable and printable: labels, number formats, a named input area and page setup.", ["CLO1"], 25),
        ],
    },
    "ai": "Optional: ask an AI tool to explain a formula you don't understand, then test its explanation by changing an input. AI tools often get cell references subtly wrong, so check every formula they suggest.",
    "privacy": "The grade calculator uses made-up students. Never put real classmates' names or grades into a practice file or an AI tool.",
    "platform": "Everything this week works in Excel for the web and desktop Excel.",
    "nace": ["Critical Thinking", "Technology", "Professionalism"],
    "digcomp": ["Content creation", "Problem identification and solving"],
    "wcag": ["2.4.4"],
    "life": "An invoice and a grade calculator you'll actually use, and formulas you can trust.",
    "hours": "About 9 hours: 2 h lessons, 3.5 h drills, 1 h skill check, 2.5 h project",
  },
  {
    "id": "W3", "n": 3, "title": "Tables and Data",
    "tagline": "Bring data in, organize it as a table, and make the pattern visible.",
    "clos": {"CLO2": "R", "CLO1": "R"},
    "topics": ["T3", "T4", "T5"],
    "why": {
        "college": "Research, lab and survey data arrive as files you have to clean, sort and summarize.",
        "work": "Exports from other systems (sales, inventory, attendance) land in Excel every day.",
        "life": "Comparing apartments, tracking workouts or a club roster is easier as a sortable table.",
    },
    "objectives": [
        ("W3.1", "Import data from text (.csv and .txt) files and online sources, and remove duplicate records.", ["CLO2"]),
        ("W3.2", "Create, name and style Excel tables; add and remove rows and columns; set style options and a total row; convert a table back to a range.", ["CLO1", "CLO2"]),
        ("W3.3", "Use structured references, and sort and filter data, including with the SORT() and UNIQUE() functions.", ["CLO2"]),
        ("W3.4", "Summarize data visually with built-in conditional formatting and sparklines, and remove formatting rules.", ["CLO1", "CLO2"]),
    ],
    "drills": [
        ("D3.01", "Import the rentals .csv file into a new sheet, then remove the duplicate rows.", ["1.1.1"], ["E2.2"], "desktop"),
        ("D3.02", "Import a table of lake data from a web page with Get Data From Web.", ["1.1.2"], [], "desktop"),
        ("D3.03", "Turn the rental list into a table named Rentals with a medium table style.", ["3.1.1", "3.1.2"], [], "web"),
        ("D3.04", "Turn on banded columns and a total row (sum and average), add a Deposit column and delete the empty row.", ["3.2.1", "3.2.2", "3.2.3"], [], "web"),
        ("D3.05", "Calculate each rental's cost with a structured reference such as [@Hours]*[@Rate].", ["4.1.2"], [], "web"),
        ("D3.06", "Convert the old price table back to a normal range.", ["3.1.3"], [], "web"),
        ("D3.07", "Sort by Boat then by Date (newest first), and filter to show only the weekend rentals.", ["3.3.1", "3.3.2"], ["N15"], "web"),
        ("D3.08", "List each boat type once with UNIQUE() and the top rentals in order with SORT().", ["4.2.4", "4.2.5"], [], "web"),
        ("D3.09", "Add data bars to the hours, highlight costs over $100, then remove the leftover rule from the old report.", ["2.4.2", "2.4.3"], [], "web"),
        ("D3.10", "Insert line sparklines that show each boat's monthly trend and mark the high points.", ["2.4.1"], [], "desktop"),
    ],
    "skillcheck": {"id": "SC3", "covers": ["1.1.1", "2.4.1", "2.4.2", "2.4.3", "3.1.1", "3.1.2", "3.1.3", "3.2.1", "3.2.2", "3.2.3", "3.3.1", "3.3.2", "4.1.2", "4.2.4", "4.2.5"],
                   "desc": "Timed, MOS-style: 14 tasks on a gear inventory workbook, auto-checked, 2 attempts."},
    "project": {
        "id": "P3", "title": "Data Report",
        "desc": "Import a provided public-style dataset (or one from a class, club or job), clean it, turn it into a named table, and build a one-page report: sorted and filtered views, a total row, UNIQUE() and SORT() summaries, conditional formatting and sparklines, with a short written finding at the top.",
        "criteria": [
            ("P3-a", "The data is imported and clean: no duplicates, consistent types, a real header row, a named table.", ["CLO2"], 25),
            ("P3-b", "Sorting, filtering and the total row answer real questions about the data.", ["CLO2"], 20),
            ("P3-c", "Formulas use structured references and dynamic array functions correctly.", ["CLO2"], 20),
            ("P3-d", "Conditional formatting and sparklines highlight the pattern without relying on color alone.", ["CLO1", "CLO2"], 20),
            ("P3-e", "The written finding is accurate and supported by the sheet.", ["CLO1"], 15),
        ],
    },
    "ai": "Optional: ask an AI tool what questions a dataset like yours could answer, then answer one yourself with a sort, a filter or a formula. Check any number an AI tool gives you against the sheet.",
    "platform": "Tables, sorting, filtering, conditional formatting and the dynamic array functions work in Excel for the web. Get Data (Power Query) imports, removing duplicates from imported data and sparklines are easier or only available in desktop Excel.",
    "nace": ["Critical Thinking", "Technology", "Communication"],
    "digcomp": ["Information search, evaluation and management", "Problem identification and solving"],
    "wcag": ["1.3.1", "1.4.1"],
    "life": "The habit of turning a messy export into a clean table and one clear finding.",
    "hours": "About 9 hours: 2 h lessons, 3 h drills, 1 h skill check, 3 h project",
  },
  {
    "id": "W4", "n": 4, "title": "Charts, Graphics and What-If",
    "tagline": "Show the story in a chart, and ask Excel what happens if the numbers change.",
    "clos": {"CLO5": "IA", "CLO4": "IA", "CLO3": "IA", "CLO2": "R", "CLO1": "R"},
    "topics": ["T6"],
    "why": {
        "college": "Charts carry the results section of lab reports, posters and presentations.",
        "work": "Managers decide from a chart and a what-if: if we charge more, how many rentals do we need?",
        "life": "Car loans, rent increases and savings plans are what-if questions with real money attached.",
    },
    "objectives": [
        ("W4.1", "Create charts and chart sheets, add data series and switch rows and columns.", ["CLO5"]),
        ("W4.2", "Add and modify chart elements, apply chart layouts and styles, and add alternative text.", ["CLO5"]),
        ("W4.3", "Insert and modify pictures, shapes, icons and SmartArt in a worksheet, with alternative text.", ["CLO4"]),
        ("W4.4", "Apply a theme, create a cell style and save a reusable workbook template (.xltx).", ["CLO1", "CLO2"]),
        ("W4.5", "Answer what-if questions with Goal Seek, Scenario Manager and a one-variable data table, using PMT() for loan payments.", ["CLO3"]),
    ],
    "drills": [
        ("D4.01", "Insert a clustered column chart of monthly rentals by boat.", ["5.1.1"], ["N16"], "web"),
        ("D4.02", "Move the chart to its own chart sheet named Rentals Chart.", ["5.1.2"], [], "desktop"),
        ("D4.03", "Add the August series to the chart, then switch rows and columns so months are on the axis.", ["5.2.1", "5.2.2"], [], "web"),
        ("D4.04", "Add a chart title, axis titles and data labels, and move the legend to the bottom.", ["5.2.3"], [], "web"),
        ("D4.05", "Apply a quick layout and a chart style, then add alternative text that states the chart's point.", ["5.3.1", "5.3.2", "5.3.3"], [], "desktop"),
        ("D4.06", "Insert the logo picture, a callout shape with text, an icon and a SmartArt process; give each alt text or mark it decorative.", [], [], "desktop"),
        ("D4.07", "Apply a theme, create a cell style named Entry, and save the workbook as a template.", ["2.2.7"], [], "desktop"),
        ("D4.08", "Use PMT() for a boat loan, then Goal Seek the price that keeps the payment at $250.", [], ["E3.4"], "desktop"),
        ("D4.09", "Create Low, Expected and High season scenarios and a scenario summary report.", [], ["E3.4"], "desktop"),
        ("D4.10", "Build a one-variable data table that shows the payment at five interest rates.", [], [], "desktop"),
    ],
    "skillcheck": {"id": "SC4", "covers": ["5.1.1", "5.1.2", "5.2.1", "5.2.2", "5.2.3", "5.3.1", "5.3.2", "5.3.3", "2.2.7"],
                   "desc": "Timed, MOS-style: 11 tasks on a season sales workbook (a chart sheet, series, elements, layout, style and alt text, a cell style, Goal Seek and two scenarios), auto-checked, 2 attempts."},
    "project": {
        "id": "P4", "title": "Business Plan Workbook",
        "desc": "A plan for a small business or club idea of your own (a food truck, tutoring, a fundraiser): an Inputs sheet formatted with your own cell style, a Plan sheet with formulas and a PMT() loan payment, Goal Seek for a break-even or target price, at least three scenarios with a summary report, a one-variable data table, a chart that makes the plan's point, and a graphic, all in a theme other than Office.",
        "criteria": [
            ("P4-a", "Goal Seek, the scenarios with a summary, and the data table are set up correctly and answer real questions about the plan.", ["CLO3"], 25),
            ("P4-b", "The chart fits the data, shows the main point, and has a title, useful elements and alt text that states the point.", ["CLO5"], 25),
            ("P4-c", "The graphic is placed and sized well and has alt text or is marked decorative.", ["CLO4"], 15),
            ("P4-d", "The formulas, including PMT(), are correct and use cell references, so the what-if tools work.", ["CLO3"], 20),
            ("P4-e", "The workbook uses a theme and a custom cell style, has real sheet names and a Title, and the written findings are accurate.", ["CLO1", "CLO2"], 15),
        ],
    },
    "ai": "Optional: ask an AI tool what assumptions a loan or price model should include, and add the ones that fit. Treat any interest rate it gives you as an example, not a quote.",
    "privacy": "Use made-up or rounded figures for anything personal, like income or debt.",
    "platform": "Charts, chart elements, styles, alt text and pictures work in Excel for the web. Chart sheets, quick layouts, SmartArt, Scenario Manager, data tables and templates need desktop Excel; Goal Seek is in newer Excel for the web but not every version.",
    "nace": ["Critical Thinking", "Communication", "Technology"],
    "digcomp": ["Content creation", "Problem identification and solving"],
    "wcag": ["1.1.1", "1.4.1", "1.4.3"],
    "life": "A planner for your next big decision, and charts that make a point.",
    "hours": "About 9 hours: 2.5 h lessons, 3 h drills, 1 h skill check, 2.5 h project",
  },
  {
    "id": "W5", "n": 5, "title": "Share, Publish and Work Together",
    "tagline": "Review with others, combine their work, then share and publish safely. Capstone and certification practice.",
    "clos": {"CLO7": "IA", "CLO6": "A", "CLO1": "RA", "CLO2": "RA", "CLO3": "A", "CLO4": "A", "CLO5": "A"},
    "topics": ["T5", "T7"],
    "why": {
        "college": "Group budgets, shared data and instructor feedback arrive as comments in a shared workbook.",
        "work": "Several people edit the same budget or schedule; someone has to combine their versions and protect the formulas.",
        "life": "Sharing a family or club spreadsheet safely - without hidden data or broken formulas - saves arguments.",
    },
    "objectives": [
        ("W5.1", "Add, reply to, resolve and delete comments, and add, edit and delete notes.", ["CLO7"]),
        ("W5.2", "Co-author a shared workbook, review version history and Show Changes, and combine reviewers' workbooks with external references and Consolidate (the modern way to compare and combine).", ["CLO7"]),
        ("W5.3", "Inspect a workbook and fix accessibility and compatibility issues, and protect worksheets and workbook structure.", ["CLO1", "CLO7"]),
        ("W5.4", "Save and export workbooks as a web page, PDF and CSV, and preview the web page in a browser.", ["CLO6"]),
        ("W5.5", "Show exam readiness on a timed MO-210 style practice test and assemble a course portfolio.", ["CLO1", "CLO2", "CLO3", "CLO4", "CLO5", "CLO7"]),
    ],
    "drills": [
        ("D5.01", "Add a note and a comment, reply to and resolve one comment, and delete an old note.", ["1.5.5"], [], "web"),
        ("D5.02", "Co-edit a shared copy with a partner, open Show Changes, and restore an earlier version from version history.", [], ["E1.1"], "both-check"),
        ("D5.03", "Combine three reviewers' budget workbooks: link a total to another workbook, then Consolidate the three sheets into one.", [], ["E1.1", "E3.4"], "desktop"),
        ("D5.04", "Run Inspect Document to remove comments, hidden sheets and personal information, then run Check Compatibility.", ["1.5.4"], [], "desktop"),
        ("D5.05", "Run Check Accessibility and fix every error: alt text, default sheet names, a table with a real name.", ["1.5.4"], [], "web"),
        ("D5.06", "Unlock the input cells, protect the sheet, add data validation to the sign-up column, and protect the workbook structure.", [], ["E1.2", "E2.2"], "desktop"),
        ("D5.07", "Save the schedule as a web page, open it in a browser, and export a CSV copy of the data.", ["1.5.2"], [], "both-check"),
    ],
    "skillcheck": {"id": "SC5", "covers": ["1.4.5", "1.5.4", "1.5.5", "5.3.3"],
                   "desc": "Timed, MOS-style: 14 tasks on a volunteer roster (comments, notes, accessibility, inspection, data validation, sheet and structure protection), auto-checked, 2 attempts."},
    "project": {
        "id": "P5", "title": "Team Event Workbook",
        "desc": "With one or two classmates, plan a small event in one co-authored OneDrive workbook: comment threads with replies and @mentions, edits from each teammate in Show Changes, and Version History. Then make a clean final copy with a Combined sheet (external references or Consolidate), a protected sign-up sheet with data validation, structure protection, no comments, hidden sheets or personal information, and no accessibility errors, and save a web page and a PDF of the main sheet.",
        "criteria": [
            ("P5-a", "The shared workbook shows real co-authoring: edits from each teammate in Show Changes, comment threads with replies and @mentions, resolved where settled.", ["CLO7"], 25),
            ("P5-b", "The Combined sheet correctly links to or consolidates each teammate's work, and the totals are right.", ["CLO7"], 20),
            ("P5-c", "The sign-up sheet works for others: entry cells unlocked, useful data validation, sheet and structure protected.", ["CLO1", "CLO7"], 20),
            ("P5-d", "The final copy is inspected and accessible: no comments, hidden sheets or personal information, and no accessibility errors.", ["CLO1"], 20),
            ("P5-e", "The web page and PDF look right in a browser and a PDF reader, with a sensible title.", ["CLO6"], 15),
        ],
    },
    "capstone": {
        "id": "CAP", "title": "Capstone: Practice Exam and Portfolio",
        "desc": "A timed, 50-minute MO-210 style practice exam covering all five domains (Certiport practice tests if FLCC has access; otherwise the Office Trainer capstone set), plus a portfolio of the five projects with a short reflection on which skills you will use next and where.",
        "clos": ["CLO1", "CLO2", "CLO3", "CLO4", "CLO5", "CLO6", "CLO7"],
    },
    "ai": "Optional: Copilot in Excel can summarize a sheet or suggest a formula. Treat its suggestions like a reviewer's comment: check them, then decide.",
    "privacy": "Shared workbooks carry hidden data: names in comments, hidden sheets, file properties. Inspect before you share, and share view-only unless people need to edit.",
    "platform": "Comments, notes, co-authoring, version history, Show Changes and the accessibility checker work in Excel for the web. Consolidate, Inspect Document, Check Compatibility, protecting workbook structure and Save as Web Page need desktop Excel.",
    "nace": ["Teamwork", "Communication", "Professionalism", "Technology"],
    "digcomp": ["Communication and collaboration", "Safety, wellbeing and responsible use"],
    "wcag": ["1.1.1", "1.3.1", "2.4.2", "2.4.6"],
    "life": "Safe sharing habits for any shared spreadsheet - plus a certification attempt if wanted.",
    "hours": "About 9 hours: 1.5 h lessons, 2 h drills, 1 h skill check, 2.5 h project, 2 h capstone",
  },
]

GRADING = [
    ("Practice drills (completion, unlimited tries)", 15),
    ("Weekly skill checks SC1-SC5 (4% each)", 20),
    ("Projects P1-P4 (10% each)", 40),
    ("Project P5 Team Budget Review", 10),
    ("Capstone practice exam and portfolio", 15),
]

# Northstar Digital Literacy: Microsoft Excel (17 standards)
NORTHSTAR = [
    ("N1", "Open and close a workbook.", "W1", "D1.01"),
    ("N2", "Save a workbook, being intentional about name and location.", "W1", "D1.01"),
    ("N3", "Identify parts of Excel Screen: ribbon, formula bar, active cell, name box, column letter, row number, Quick Access Toolbar.", "W1", "D1.03"),
    ("N4", "Locate a specific cell.", "W1", "D1.05"),
    ("N5", "Enter data in a cell.", "W1", "D1.06"),
    ("N6", "Copy and move cell entries.", "W1", "D1.07"),
    ("N7", "Format cells and text: bold, underline, size, merge and center, wrap text, number (currency, time, percentages, etc.)", "W1", "D1.08, D1.09"),
    ("N8", "Create headings and freeze them.", "W1", "D1.04"),
    ("N9", "Insert and delete rows and columns.", "W1", "D1.06"),
    ("N10", "Adjust row and column size.", "W1", "D1.11"),
    ("N11", "Identify worksheet tabs, create a new tab, rename tabs, and rearrange tabs.", "W1", "D1.02"),
    ("N12", "Write a formula in the formula bar (-, +, *, /).", "W2", "D2.01"),
    ("N13", "Select a range.", "W1", "D1.13"),
    ("N14", "Use Auto Fill and AutoSum (Sum, Average, etc.) and understand the differences between them.", "W1", "D1.06, D2.05"),
    ("N15", "Sort (least to greatest, alphabetically, etc.) and filter data.", "W3", "D3.07"),
    ("N16", "Insert a chart to display data.", "W4", "D4.01"),
    ("N17", "Select a print area, choose page orientation, and print.", "W1", "D1.12"),
]

NACE = [
    ("Career & Self-Development", ["W5"], "A portfolio reflection on where Excel skills will be used next."),
    ("Communication", ["W1", "W3", "W4", "W5"], "Readable sheets, charts that make one point, written findings and review comments."),
    ("Critical Thinking", ["W2", "W3", "W4"], "Checking formulas, questioning data and testing what-if assumptions (including AI suggestions)."),
    ("Professionalism", ["W1", "W2", "W5"], "Clean, print-ready workbooks and safe sharing habits."),
    ("Teamwork", ["W5"], "Co-authoring, comments and combining reviewers' workbooks."),
    ("Technology", ["W1", "W2", "W3", "W4", "W5"], "Every week: choosing the right Excel feature and using it responsibly."),
]

DIGCOMP = [
    ("Information search, evaluation and management", ["W1", "W3"], "Organizing workbooks; importing, cleaning and filtering data."),
    ("Communication and collaboration", ["W5"], "Sharing, co-authoring, comments and notes."),
    ("Content creation", ["W1", "W2", "W4"], "Building budgets, calculators, charts and templates."),
    ("Safety, wellbeing and responsible use", ["W1", "W2", "W5"], "Keeping personal and financial data out of practice files and AI prompts; Inspect Document; protection."),
    ("Problem identification and solving", ["W2", "W3", "W4"], "Error checking, data cleanup and what-if analysis."),
]

WCAG = [
    ("1.1.1", "Non-text Content", "A", "Alt text on charts, pictures, shapes and SmartArt; decorative items marked decorative.", ["W4", "W5"]),
    ("1.3.1", "Info and Relationships", "A", "Real header rows in Excel tables; no merged cells or blank rows inside data.", ["W1", "W3", "W5"]),
    ("1.4.1", "Use of Color", "A", "Conditional formatting and charts don't rely on color alone (labels, icons, data labels).", ["W3", "W4"]),
    ("1.4.3", "Contrast (Minimum)", "AA", "Cell fills, chart colors and text keep enough contrast.", ["W4"]),
    ("2.4.2", "Page Titled", "A", "The workbook Title property is set (it becomes the PDF title).", ["W1", "W5"]),
    ("2.4.4", "Link Purpose (In Context)", "A", "Hyperlinks use descriptive text, not raw addresses.", ["W2"]),
    ("2.4.6", "Headings and Labels", "AA", "Sheet names and column headers describe their contents (no Sheet1, Table1).", ["W5"]),
]

PLATFORM = [
    ("Formulas, functions, names and tables", "Yes", "Yes", "W2, W3"),
    ("Sort, filter and conditional formatting", "Yes", "Yes", "W3"),
    ("Get Data (Power Query) imports", "Limited", "Yes", "W3"),
    ("Sparklines", "Check (varies)", "Yes", "W3"),
    ("Charts, chart elements and styles", "Yes", "Yes", "W4"),
    ("Chart sheets and quick layouts", "No", "Yes", "W4"),
    ("SmartArt", "No", "Yes", "W4"),
    ("Goal Seek", "Newer versions", "Yes", "W4"),
    ("Scenario Manager and data tables", "No", "Yes", "W4"),
    ("Templates (.xltx)", "Open only", "Yes", "W4"),
    ("Comments, notes and co-authoring", "Yes", "Yes", "W5"),
    ("Version history and Show Changes", "Yes", "Yes (OneDrive)", "W5"),
    ("Consolidate and external references", "No", "Yes", "W5"),
    ("Inspect Document and Check Compatibility", "No", "Yes", "W5"),
    ("Protect sheet", "Yes", "Yes", "W5"),
    ("Save as Web Page", "No", "Yes", "W5"),
    ("Export as PDF and CSV", "Yes", "Yes", "W1, W5"),
    ("Accessibility checker", "Yes", "Yes", "W5"),
]

CONVENTIONS = [
    ("Spreadsheet layout conventions", "Industry convention", "One header row, one record per row, no blank rows or merged cells inside data; inputs separated from results.", ["W1", "W3"]),
    ("Loan math (PMT)", "Assumption", "Practice loans use monthly payments and simple example rates; not financial advice.", ["W4"]),
    ("Bloom's revised taxonomy", "Assumption", "Course outcome verbs sit mostly at Apply; what-if analysis reaches Analyze and peer review reaches Evaluate.", ["All"]),
    ("Credit hour (about 45 hours per credit)", "Assumption", "CSC 135 is 1 credit; 9 hours a week for 5 weeks.", ["All"]),
    ("I / R / A curriculum mapping", "Common practice", "Introduced, Reinforced, Assessed - the standard way to show where an outcome is taught and measured.", ["All"]),
]

DECISIONS = [
    ("Reword CLO 7", "flag", "\"Collaborate on workbooks: Compare and Combine workbooks\" describes Excel's legacy Shared Workbooks with Compare and Merge, which Microsoft hides (it has to be added to the Quick Access Toolbar), works only in Windows desktop Excel and conflicts with co-authoring. The plan meets the outcome the modern way - co-authoring, comments, version history and Show Changes, then combining reviewers' workbooks with external references and Consolidate - and shows the legacy command once. Consider rewording to \"Collaborate on workbooks: share, co-author, review and combine workbooks.\""),
    ("Reword CLO 6", "flag", "\"Create and preview the worksheet as a Web page\" reflects an older workflow. The plan meets it as written (Save as Web Page and a browser preview in Week 5), alongside what people do now: PDF, CSV and share links. Consider rewording to \"Publish worksheets for screen and web audiences (web page, PDF, share link).\""),
    ("Make CLO 2 measurable", "flag", "\"Demonstrate features offered by the application\" is hard to assess on its own. It is mapped to every week here. Consider rewording to \"Use Excel's data, formula and table features to organize and analyze information.\""),
    ("What-if analysis is beyond MO-210", "info", "CLO 3 (what-if scenario) is MO-211 Expert content: PMT(), Goal Seek and Scenario Manager (MO-211 3.4) plus a data table. It stays because the course outcome requires it, and it is labeled as beyond-the-certification content."),
    ("Exam number", "flag", "Microsoft's current Excel Associate exam for Microsoft 365 Apps is MO-210 (MO-200 was the Office 2019 version). Check which exam FLCC's Certiport setup offers."),
    ("Plan for desktop Excel", "flag", "Weeks 3 to 5 include features that need desktop Excel (Get Data imports, chart sheets, SmartArt, Scenario Manager, data tables, Consolidate, Inspect Document, Save as Web Page). Chromebook users will need campus labs or a virtual desktop for those tasks. Each drill is tagged web or desktop."),
    ("PivotTables are not included", "info", "PivotTables and lookups (XLOOKUP) are MO-211 skills and are taught in CSC 212. They are the most-requested Excel skills at work, so consider them as optional enrichment if time allows."),
    ("Accessibility is built in, not bolted on", "info", "Alt text on charts, real table headers, named sheets and tables, and no merged cells inside data are taught as normal practice, and the accessibility checker is part of the clean-and-ship checklist (ADA Title II, WCAG 2.1 AA)."),
    ("Privacy by design", "info", "Budgets, loans and grades use made-up numbers; no real account details, classmates or grades in practice files or AI prompts; Inspect Document before sharing."),
    ("Northstar as a diagnostic", "info", "All 17 Northstar Excel standards are covered, 14 of them in Week 1. If FLCC has a Northstar site license, use it as a Day 1 pre-assessment and a Week 1 post-check."),
]

SOURCES = [
    ("CSC 135 Core Excel - course description (FLCC)", "https://www.flcc.edu/courses/csc-135/"),
    ("CSC 135 Core Excel - course outline with outcomes (FLCC)", "https://cyber.flcc.edu/pdf/syllabi/CSC135.pdf"),
    ("MOS Excel Associate (Microsoft 365 Apps), exam MO-210 - Microsoft Learn", "https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-associate-m365-apps/"),
    ("Excel (Microsoft 365 Apps) Exam MO-210 objective domains (Appalachian State copy)", "https://business.appstate.edu/sites/default/files/mos_365_excel_od_2022.pdf"),
    ("MO-211 Excel Expert (Microsoft 365 Apps) skills outline - Microsoft", "https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RW10j3G"),
    ("Northstar Digital Literacy standards - Microsoft Excel", "https://www.digitalliteracyassessment.org/features"),
    ("NACE Career Readiness: Competencies Defined - NACE", "https://www.naceweb.org/career-readiness/competencies/career-readiness-defined"),
    ("DigComp framework - European Commission Joint Research Centre", "https://joint-research-centre.ec.europa.eu/digcomp/digcomp-framework_en"),
    ("Understanding WCAG 2.1 success criteria - W3C", "https://www.w3.org/WAI/WCAG21/Understanding/"),
]

CLO_ARTIFACT = {
    "CLO1": "Project 1 budget (.xlsx and .pdf); print and page setup in every project; clean-and-ship checks in Project 5.",
    "CLO2": "Skill checks SC1 to SC5 (feature use under time); Project 2 formulas; Project 3 data report; capstone practice exam.",
    "CLO3": "Project 4 what-if planner (Goal Seek, scenarios with a summary report, data table); SC4.",
    "CLO4": "Project 4 planner graphic with alt text; D4.06.",
    "CLO5": "Project 4 chart sheet; SC4; capstone domain 5.",
    "CLO6": "Project 5 web page version of the workbook and the format-comparison reflection.",
    "CLO7": "Project 5 review comments, combined workbook and protected, inspected copy; D5.03.",
}
