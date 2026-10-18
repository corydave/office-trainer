# Office Trainer (proof of concept)

A free practice-and-check sidebar for Word. Students do real tasks in real Word,
and the checker grades the **result** (is the title bold?) rather than the clicks
used to get there (ribbon, Ctrl+B, or a style all count).

## What's in this folder

| File | What it is |
| --- | --- |
| `taskpane.html` | The whole app: sidebar, lesson, checker. One file. |
| `manifest.xml` | Tells Word the add-in exists and adds a button to the Home tab. |
| `word/` | Everything for the Word course: lessons in Trainer Markdown, starter documents, the course map, the editor and specs. See `word/README.md`. |
| `word/starters/poc/` | The proof-of-concept practice documents (`word-1.docx`, `word-2.docx`...) that `taskpane.html` loads today. |
| `excel/`, `powerpoint/` | Placeholders for the other two courses (same layout as `word/`). |
| `.htaccess` | HostGator/Apache only: makes updates show up right away while testing. |
| `commands.html` | Empty page the manifest format requires. Leave it alone. |
| `index.html` | Forwards the site's home page to `taskpane.html`. |
| `assets/` | Icons for the ribbon button. |
| `word/fixtures/` | Test files for the proof-of-concept checks (unsolved, solved, partly solved...). Not needed for students. |

## 1. Put it on GitHub Pages

1. The repository is `corydave/office-trainer` (it must be **public** for free GitHub Pages).
2. Upload everything in this folder (you can skip `word/fixtures/`, `word/_build/` and `word/tools/`).
3. Go to **Settings > Pages**. Under "Build and deployment," choose **Deploy from a branch**, then **main** and **/ (root)**. Save.
4. After a minute or two, open https://corydave.github.io/office-trainer/ - you should see the trainer in upload mode.

If you ever host it somewhere else, find-and-replace `https://corydave.github.io/office-trainer/` in `manifest.xml`.

## 2. Try it in Word on the web (works on a Chromebook)

1. Open a **blank** document in Word on the web (office.com).
2. Go to **Home > Add-ins**. Look for **More Add-ins** or **Advanced**, then **My Add-ins > Upload My Add-in**. (Microsoft moves this menu around; it's always under Add-ins.)
3. Choose `manifest.xml` from your computer.
4. An **Office Trainer** button appears on the Home tab. Click it to open the sidebar.
5. Click **Load the practice document**, do the tasks, and click **Check my work**.

Keyboard users: **F6** moves between the document, the ribbon, and the sidebar.

## 3. Give it to everyone (later)

A Microsoft 365 admin can deploy it to students and faculty so nobody has to upload
anything: **Microsoft 365 admin center > Settings > Integrated apps > Upload custom apps**,
then choose `manifest.xml` and pick who gets it.

## 4. Students without the sidebar (phones, etc.)

Send them to https://corydave.github.io/office-trainer/. They download the practice file,
do the work in any version of Word (even the phone app), save, and upload the file to check it.
Same checker, same results.

## Updating it (two kinds of updates)

**Website update** - anything in `taskpane.html`, the lessons, or `word/starters/`:

1. Bump `APP_VERSION` near the top of PART 1 in `taskpane.html`.
2. Upload the changed files to the server.
3. In Word, close the Office Trainer pane and open it again. Check the version number at the
   bottom of the pane. If it still shows the old number, reload the Word tab.

No manifest changes and nothing for students to reinstall.

**Add-in update** - only when `manifest.xml` itself changes (button text, icons, web addresses):

1. Bump `<Version>` in `manifest.xml`.
2. Upload the new manifest again (Home > Add-ins > My Add-ins > Upload My Add-in), or update it
   in the Microsoft 365 admin center once it's deployed there.

## Lessons and links

Each lesson can be opened directly, which is handy for Brightspace links:
`taskpane.html?lesson=word-1`, `?lesson=word-2`, `?lesson=word-3`.
With no `?lesson=`, the lesson menu opens.

## How to add a task

1. In `taskpane.html`, find **PART 1: THE LESSONS** and copy one of the task blocks.
2. Change its `id`, `title`, `instructions`, and `steps`.
3. Point `check` at a check function. Copy one from **PART 4: THE CHECKS** as a starting point.
4. To add a video, paste the YouTube video ID (the part after `v=`) into `video.youtubeId`.

## Privacy and security

- The document is read **inside the student's browser**. No server, no database, no accounts.
- The page only loads code from Microsoft (office.js) and cdnjs (JSZip, locked with an integrity hash).
- Text from documents is only ever shown with `textContent`, so it can never run as code.
- YouTube is only contacted when a student clicks "Watch video" (privacy-enhanced mode).

## Note on updates

Word caches add-ins. After you change `taskpane.html`, close and reopen the document.
If the old version still shows, clear the browser cache.

## License

MIT. See [LICENSE](LICENSE).
