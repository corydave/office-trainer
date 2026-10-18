# Word course - Office Trainer content

Everything learners see in the Word add-in comes from the Markdown files in this folder. To change a lesson, a video, a step or a check, edit the Markdown and rebuild. Nothing about a lesson lives anywhere else.

## What's here

| Path | What it is |
|---|---|
| `course.config.md` | Course settings: default fonts, approved video hosts, license |
| `content/week-01/` | Week 1: four lessons, Skill Check 1, Project 1 (one file each) |
| `starters/week-01/` | The starter .docx files the drills open. Checks are written against these |
| `spec/trainer-markdown.md` | The Trainer Markdown language: directives, steps, checks, validation rules |
| `spec/checker-library.md` | Every check, its parameters, how it works on web and desktop, and what Week 1 can't auto-check |
| `tools/editor.html` | The editor. Open it in Chrome or Edge, then **Open folder** and pick this folder |
| `tools/trainer-md.js` | The parser, validator and renderer (one file, no dependencies). The editor and the add-in share it |
| `tools/validate.mjs` | Command-line validator: `node tools/validate.mjs content/week-01` |
| `tools/make-starters.py` | Rebuilds the starter .docx files: `python3 tools/make-starters.py` (needs python-docx) |
| `tools/build-editor.py` | Rebuilds `editor.html` after a change to `trainer-md.js` or `tools/src/editor.src.html` |
| `word-course-map.md` / `.html` | The 5-week course map (`_build/` regenerates both) |
| `starters/poc/`, `fixtures/` | The proof-of-concept add-in's practice files and test files |

## Editing a lesson

1. Open `tools/editor.html` in Chrome or Edge and choose **Open folder** (this folder). Safari and Firefox can use **Open file** instead, but can't check that starter files exist.
2. Pick a lesson in the file list. Type on the left; the right side previews the add-in panel at its real width.
3. Use **Insert** for videos, callouts, quizzes, drills, steps and checks. **Check** lists every check in the library with an example.
4. Fix anything under **Problems**. Select a problem to jump to its line. Errors block a file from being marked `status: ready`.
5. Save with **Save** or [Ctrl/Cmd+S].

## Status of Week 1

- All six files validate with **0 errors**. The 11 warnings are expected for a draft: 6 videos still to be recorded (with checked captions), and one drill (D1.12, views) that is quiz-only by design.
- Starter files exist for every drill and the skill check. Their planted problems match the checks (for example, exactly 4 "Acme" and 3 "deadline" in the Lesson 2 letter).
- **Not built yet:** the add-in code that runs these checks (`checker-library.md` describes each one). It goes into the add-in at the top of this folder (`../taskpane.html`), which still runs the three proof-of-concept lessons from `starters/poc/`.

## Next steps

1. Build the checker library into `../taskpane.html` and have it load lessons from `content/` instead of its built-in lesson list.
2. Record the Week 1 videos and replace each `url: TODO`; set `captions: verified` once the captions have been checked.
3. Build answer keys (finished versions of each starter) and the known-good and known-bad test documents for each check.
4. Mark each file `status: ready` once it validates with 0 errors.
