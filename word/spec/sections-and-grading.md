# Sections and grading without a server (v0.1)

Office Trainer has no server, no accounts and no database. This file explains how instructors still get sections, their own videos, and grades in Brightspace, and what a FLCC-only server could add later.

## 1. The parts

| Part | Who uses it | Where | What it does |
|---|---|---|---|
| Section builder | Instructors | `instructor/sections.html` | Pick a course, untick the stops the section leaves out, add or swap videos, and get a **section key** (or a section file for the website) |
| Section key or code | Students | Panel: **Settings > Your section** | Changes which stops count and which videos show. Nothing is sent anywhere |
| Badges page | Students | `badges.html` | Every badge and certificate from all three courses on one page (see `badges-and-certificates.md`, 3a) |
| Grader | Instructors | `instructor/grader.html` | Checks a class's files in the browser and makes a Brightspace grade import, a results spreadsheet and feedback files |

All three pages are built by `word/tools/build-site-pages.py` from `word/tools/src/*.src.html`. Each has a Content Security Policy that only allows fetching this site's own course files, so nothing a page opens can be sent anywhere.

## 2. Section keys

A key is a whole section in one line of text: `OTK1.` followed by the section as JSON, encoded as base64url (letters, digits, `-` and `_`, so it survives copy and paste). Spaces and line breaks are ignored, so a key wrapped by an email still works.

```json
{ "app": "powerpoint", "title": "CSC 136-01, Spring 2027", "instructor": "Professor Rivera", "term": "Spring 2027",
  "skip": ["PP1-L3"],
  "videos": [
    { "replace": "p1-l1-tour", "url": "https://www.youtube-nocookie.com/embed/VIDEO_ID", "title": "Our lab tour", "captions": "verified" },
    { "after": "D1.06", "url": "https://player.vimeo.com/video/123", "title": "Sections tip", "captions": "auto", "note": "Watch this if..." }
  ] }
```

- **A key is not secret and not trusted.** Anyone can make one, so the panel checks every key with `TrainerMD.cleanSection` (the same code the builder uses):
  - Lengths are capped: a 20,000-character key, an 80-character title, a 60-character name and 240-character notes.
  - Control, invisible and text-direction characters are removed, so a name can't be disguised.
  - Lesson, video and activity ids must exist in the course.
  - Videos must be https links on the course's approved video sites (`videoHosts`). YouTube and Vimeo links must name one video, and links with a user name or port are refused. Links are stored as player links (YouTube in privacy-enhanced mode).
  - Notes are plain text: no links or pictures.
  - A key can't add a video site or run code. The panel's policy also only lets approved video sites play (`frame-src`).
- **Short codes** (`PP-7K3Q`) name a file, `sections/PP-7K3Q.md`, on the website (5.13 in `trainer-markdown.md`). They're easier to type, but someone who can update the site has to add the file. The builder can download that file.
- **What a section changes:**
  - Left-out stops show **Not in your section**. Learners can still open them for practice; they don't count toward progress or the course certificate.
  - Previous / Next skip the left-out stops.
  - Instructor videos say "From <instructor>", with a link to the standard video.
  - Progress, badges and certificates stay on the device. Leaving the section shows the whole course again; badges stay.
- **Course certificates in a section:**
  - A section that leaves out more than half of the stops with badges can't award one, so a key that skips everything can't hand out a free certificate.
  - A certificate earned in a section keeps the section's name and skip list. It's titled "Course Complete: <section>", the certificate says "the course as set for <section>", and the back lists only the badges that counted, even after the learner leaves the section.
- **Join links:** `taskpane.html?app=powerpoint&join=<key or code>` offers the section on the home screen and joins only if the learner says yes. The builder gives this link for students working in a browser.
- **Editing:** paste the old key into the builder's **Edit a section you made before**, change it, and make a new key. Students paste the new key; their badges stay.

## 3. The grader

1. **Choose the course and the activity.** The choices are drills, skill checks and project preflights: anything with steps the checker can see in the file.
   - Each checked step can be ticked off ("Counts"). Screenshot, typed-answer and self-check steps are marked **By hand**.
   - Skill checks use their steps' `points:`; everything else counts 1 point a step.
2. **Add the students' files:** the zip from Brightspace's **Download** on an assignment, or the files themselves.
   - Brightspace makes one folder per submission, named `<numbers> - <student name> - <date>`, with the student's file inside. Older versions put the same details in front of each file name, and both are read.
   - Files go together when the names match and they share one of Brightspace's id numbers, so two students with the same name are never mixed up.
   - When a student sent more than one, the newest (by the date in the folder name) is checked.
3. **Optional: a Brightspace grades export (CSV)** with Org Defined ID, Username, First Name, Last Name and the grade item.
   - Students are matched by name: accents ignored, any alphabet, either order.
   - Unmatched or ambiguous students (for example two Chris Lees and one roster row) are left for the instructor to pick from a list.
   - The grade item's MaxPoints fills in "points possible".
4. **Check every file.** Each file is opened and checked in the browser with the same checker as the panel. The results show:
   - each student's score, points and step-by-step notes;
   - how the class did on each step.
5. **Downloads:**
   - **Grades for Brightspace import (.csv):** `OrgDefinedId`, `Username`, the grade item's exact header from the export, and `End-of-Line Indicator` = `#`. Import with **Grades > Import**, which previews before changing anything. Unmatched students are left out.
   - **All results (.csv):** every student, every counted step. Names and file names that start with `= + - @` get a `'` in front, so they can't become formulas in Excel.
   - **Feedback files (.zip):** a plain-text file per student, in that student's original Brightspace folder, for the assignment's **Add Feedback Files**.

**Safety limits.** The grader checks:

- at most 600 files at once;
- 1 GB for the zip;
- 150 MB for a student file, counted while it unzips (a zip can lie about its size);
- 400 MB unzipped inside a student file, counted before the checker opens it, so a "zip bomb" shows as an error instead of freezing the tab.

**Why it's harder to fake than a score from the panel:** the check runs on the file the student handed in. (A determined student could still edit a document's XML by hand to pass a check without doing the skill, as with any file-based grading.)

### To confirm with a real Brightspace course (not tested yet)

- The exact folder names in **Download** (the grader was tested with made-up zips in the format above).
- That **Grades > Import** accepts the file as written (the `#` in front of ids is accepted, per D2L's documentation).
- That **Add Feedback Files** matches the feedback zip's folders to students.

## 4. Later: a FLCC-only server

The design above works anywhere with no server. If FLCC hosts a small server (FLCC-run hosting, approved by IT; FLCC students only, so FLCC's own agreements cover it), it would add, without changing anything above:

- **LTI 1.3 from FLCC's Brightspace only.** Launch identifies the student and section, Assignment and Grade Services posts scores, and Names and Roles gives the roster. The server refuses any other platform.
- **Pairing codes:** the Brightspace link shows a one-time code, and the student types it into the panel. That ties the panel to that student in that section.
- **An instructor dashboard:** live progress, the section settings above (stored instead of in a key), grades.
- **Store as little as possible:** LTI's anonymous ids, progress, scores and dates; no student files kept.
- **Checking files on the server** needs Node.js (the checkers are JavaScript); without it, results reported by the panel can be faked, and the grader page stays the strong option.

The panel must never depend on the server: everything above keeps working without it, and other colleges never contact it.
