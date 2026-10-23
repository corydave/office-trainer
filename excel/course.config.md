---
type: config
id: EXCEL
title: Microsoft Excel
app: excel
college: Finger Lakes Community College
license: CC BY 4.0
headingFont: Comfortaa
bodyFont: Atkinson Hyperlegible
learnerCanChangeFonts: true
videoHosts: [www.youtube-nocookie.com, www.youtube.com, player.vimeo.com]
defaultPlatform: web
mode: practice
certificateEach: lesson
courseCertificate: true
issuer: Finger Lakes Community College, Computing Sciences
---

## Course settings

These settings apply to every lesson in the Excel course (CSC 135 Core Excel). `app: excel` tells the validator and the add-in to use the Excel checks (`excel/tools/excel-checker.js`). Change a value above and rebuild; lesson files never repeat these settings.

## Modes: practice or section

`mode` sets what learners see when they have not joined a section.

- **practice** (the default): the course is a set of exercises. Auto-checks give instant feedback, screenshot steps are compared with an example, and project rubrics become self-checks. Nothing leaves the learner's computer, and the "How this course works" page says so before they start.
- **section**: a graded course. Auto-checks count toward grades, screenshots go to the instructor, and projects are graded with the rubric.

Once sections exist (Phase 2), joining a section with a code switches that learner to section mode. Everyone else stays in practice mode.

## Badges and certificates

- Every drill, skill check and project awards a **badge** when it is finished. Its name is the activity's title unless the activity sets `badge:`.
- `certificateEach: lesson` issues a **certificate** at the end of each lesson (chapter), with the lesson's badges and outcomes on the back. Other choices: `week`, `course` or `none`.
- `courseCertificate: true` adds one more certificate when the whole course is finished.
- `issuer` is printed on every certificate.
- In practice mode, certificates are marked **Practice record**: each has a unique number, but nothing outside the learner's device records it. See `spec/badges-and-certificates.md`.

## Fonts

`headingFont` and `bodyFont` set the default fonts for the course. When `learnerCanChangeFonts` is `true`, learners can pick their own pair in the add-in's settings (useful for dyslexia, low vision, or a simple preference). The first time a learner opens this course, their name and display settings from the Word course are copied over. Their choice stays in their own browser.

Each choice comes with a fallback chain so text still looks right on computers that do not have the font. Chromebooks and Linux usually lack Microsoft's fonts but include look-alikes with the same letter widths (Arimo for Arial, Tinos for Times New Roman, Carlito for Calibri).

| Choice | Fallback chain |
|---|---|
| Comfortaa (default headings) | Comfortaa, Atkinson Hyperlegible, Arial, sans-serif |
| Atkinson Hyperlegible (default body) | Atkinson Hyperlegible, Arial, Arimo, sans-serif |
| Arial | Arial, Arimo, Liberation Sans, Helvetica, sans-serif |
| Helvetica | Helvetica, Helvetica Neue, Arial, Arimo, sans-serif |
| Verdana | Verdana, DejaVu Sans, Arial, sans-serif |
| Calibri | Calibri, Carlito, Arial, sans-serif |
| Aptos | Aptos, Calibri, Carlito, Arial, sans-serif |
| Times New Roman | Times New Roman, Tinos, Liberation Serif, Times, serif |
| Georgia | Georgia, Gelasio, Times New Roman, Tinos, serif |
| OpenDyslexic | OpenDyslexic, Atkinson Hyperlegible, Arial, sans-serif |
| System default | system-ui, -apple-system, Segoe UI, Roboto, sans-serif |

Comfortaa and Atkinson Hyperlegible are hosted with the course (assets/fonts), so the add-in, roadmap and certificate send nothing to a font service (only the editor loads them from Google Fonts). Every other choice uses fonts already on the learner's computer, so nothing else is downloaded. OpenDyslexic is used only if the learner has installed it.

## Video hosts

Videos load only from the hosts in `videoHosts`. YouTube links are rewritten to `www.youtube-nocookie.com` (privacy-enhanced mode), which does not set tracking cookies until the learner presses play. Add FLCC's own video platform here (for example its Panopto, Kaltura or YuJa domain) once it is confirmed.

## License

Lesson text is shared under `CC BY 4.0`, so other colleges can reuse and adapt it with credit. Change this before publishing if FLCC prefers another license.
