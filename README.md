# Office Trainer

A free practice-and-check sidebar for Word, Excel and PowerPoint. Students do real tasks in the
real app, and the checker grades the **result** (is the title bold? is the total a
formula?) rather than the clicks used to get there (ribbon, Ctrl+B, or a style all count).
One add-in serves all three courses: it shows the Word course in Word, the Excel course in Excel and the PowerPoint course in PowerPoint.

## What's in this folder

| File | What it is |
| --- | --- |
| `taskpane.html` | The whole add-in in one file: course home, lessons, all three checkers, badges. **Built** from `word/tools/src/` by `word/tools/build-taskpane.py`; don't edit it by hand. |
| `manifest.xml` | Tells Word, Excel and PowerPoint the add-in exists and adds a button to each app's Home tab. |
| `word/` | Everything for the Word course: lessons in Trainer Markdown, starter documents, the course map, the editor and specs. See `word/README.md`. |
| `word/content/`, `word/starters/` | The lessons (Markdown) and practice documents the add-in loads when it opens. |
| `excel/` | Everything for the Excel course, in the same layout as `word/`. See `excel/README.md`. |
| `powerpoint/` | Everything for the PowerPoint course, in the same layout. See `powerpoint/README.md`. |
| `.htaccess` | HostGator/Apache only: makes updates show up right away while testing. |
| `commands.html` | Empty page the manifest format requires. Leave it alone. |
| `index.html` | The product page (the site's home page): pick a course, how it works, installing, privacy, instructors, open license. One file, no outside code. After editing its script or style block, run `python3 word/tools/build-index-csp.py` to update its security fingerprints. |
| `assets/` | Icons for the ribbon button, and `assets/fonts/` (Comfortaa and Atkinson Hyperlegible, SIL Open Font License), hosted here so no font service is contacted. |
| `word/tools/test-docs/`, `excel/tools/test-docs/`, `powerpoint/tools/test-docs/` | Answer keys and test files for the checks. Not needed for students. |

## 1. Put it on GitHub Pages

1. The repository is `corydave/office-trainer` (it must be **public** for free GitHub Pages).
2. Upload everything in this folder (you can skip `word/fixtures/`, `word/_build/` and `word/starters/poc/`). Keep `word/tools/` if you want `check-tests.html` online.
3. Go to **Settings > Pages**. Under "Build and deployment," choose **Deploy from a branch**, then **main** and **/ (root)**. Save.
4. After a minute or two, open https://corydave.github.io/office-trainer/ - you should see the product page, with a button for each course.

If you ever host it somewhere else, find-and-replace `https://corydave.github.io/office-trainer/` in `manifest.xml`.

## 2. Try it in Word on the web (works on a Chromebook)

1. Open a **blank** document in Word on the web (office.com).
2. Go to **Home > Add-ins**, then **More Settings** (older menus say **More Add-ins** or **Advanced**, then **My Add-ins**), then **Upload My Add-in**. (Microsoft moves this menu around; it's always under Add-ins.)
3. **Browse** to `manifest.xml` on your computer and select **Upload**.
4. An **Office Trainer** button appears on the Home tab. Click it to open the sidebar.
5. Pick a lesson. In a drill, select **Open it here** to load the practice document, do the steps, and select **Check my work**.

Keyboard users: **F6** moves between the document, the ribbon, and the sidebar (**Ctrl+F6** in the web apps).

**Excel works the same way:** open a blank workbook in Excel on the web, upload the same `manifest.xml`
(Home > Add-ins > More Settings > Upload My Add-in), and select **Office Trainer** on the Home tab.
In a drill, **Open it in Excel** opens the practice workbook in a new tab; open Office Trainer there to check your work.
**PowerPoint is the same:** upload the same `manifest.xml` in PowerPoint on the web; **Open it in PowerPoint** opens the practice presentation in a new tab.

## 3. Give it to everyone (later)

A Microsoft 365 admin can deploy it to students and faculty so nobody has to upload
anything: **Microsoft 365 admin center > Settings > Integrated apps > Upload custom apps**,
choose **Office Add-in** as the app type, upload `manifest.xml` and pick who gets it. It can take up to a day to appear.

## 4. Students without the sidebar (phones, etc.)

Send them to https://corydave.github.io/office-trainer/ and they pick a course, or link straight to
https://corydave.github.io/office-trainer/taskpane.html?app=word, `?app=excel` or `?app=powerpoint`. They download each practice file,
do the work in any version of the app (even the phone app), save, and select **Upload my file and check**.
Same lessons, same checker, same results. The trainer's home screen has a **Course: Word / Excel / PowerPoint** switch.

## Updating it (three kinds of updates)

**Lesson update** - a lesson's Markdown or a starter document: just upload the changed file.
The add-in reads the lessons fresh each time it opens.

**Add-in update** - anything in `word/tools/src/`, `trainer-md.js`, `word-checker.js`, `excel/tools/excel-checker.js` or `powerpoint/tools/ppt-checker.js`:

1. Bump `APP_VERSION` near the top of `word/tools/src/taskpane.app.js`, then run `python3 word/tools/build-taskpane.py`.
2. Upload `taskpane.html` to the server.
3. In Word, close the Office Trainer pane and open it again. Check the version number at the
   bottom of the pane. If it still shows the old number, reload the Word tab.

No manifest changes and nothing for students to reinstall.

**Manifest update** - only when `manifest.xml` itself changes (button text, icons, web addresses):

1. Bump `<Version>` in `manifest.xml`.
2. Upload the new manifest again (Home > Add-ins > My Add-ins > Upload My Add-in), or update it
   in the Microsoft 365 admin center once it's deployed there.

## Lessons and links

Each lesson can be opened directly, which is handy for Brightspace links:
`taskpane.html?lesson=W1-L2` (Lesson 2), or `taskpane.html?lesson=W1-L2&activity=D1.04` to jump to a drill.
Excel and PowerPoint lessons work the same way (`taskpane.html?lesson=XL1-L2`, `taskpane.html?lesson=PP1-L2`); in a browser, add `&app=excel` or `&app=powerpoint`.
With no `?lesson=`, the course home opens.

## How to add or change a lesson

Lessons are Markdown files in `word/content/`, `excel/content/` and `powerpoint/content/`, each listed in that course's `content/outline.md`. Edit them with
`word/tools/editor.html` (it previews exactly what the add-in shows; open the `word`, `excel` or `powerpoint` folder). See each course's README.

## Privacy and security

- The document is read **inside the student's browser**. No server, no database, no accounts.
- The page only loads code from Microsoft (office.js) and cdnjs (JSZip, locked with an integrity hash). Fonts are hosted here.
- Progress (badges, which steps passed, the name typed for certificates) stays in the browser, one record per course. Nothing from the document, workbook or presentation is stored.
- Text from documents is only ever shown with `textContent`, so it can never run as code. Lesson text is escaped before it's drawn.
- Video sites are only contacted when a student presses Play (YouTube in privacy-enhanced mode).

## Note on updates

Word, Excel and PowerPoint cache add-ins. After you change `taskpane.html`, close and reopen the document.
If the old version still shows, clear the browser cache.

## License

MIT. See [LICENSE](LICENSE).
