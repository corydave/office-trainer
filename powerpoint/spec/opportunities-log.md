# Opportunities log (PowerPoint)

Ideas, open questions and fixes found while building and testing the PowerPoint course (October 2026, add-in 0.7, PowerPoint checker 0.1.0). Each item says what it is, why it matters, and a suggested next step. Priority: **High** (affects learners now), **Medium**, **Low**.

## 1. Fixed while building

Found by testing:

| What was wrong | Fix |
|---|---|
| SmartArt text was read twice ("PlanPlan"): the data model wraps `a:t` inside `dgm:t` | Only the innermost text elements are read |
| Zoom targets were never matched to slides | Zoom objects are matched by slide id once every slide is read |
| python-pptx writes the picture's file name as alt text, so every picture looked described | A file name in the alt text box counts as no alt text, and the builders clear it |
| `list.is` with no text picked the slide title instead of the list | It skips titles and picks the longest list (and, across slides, one of the asked-for kind) |
| New shapes from python-pptx have a strong style, so "change the shape style" passed on a starter | The builders give new shapes PowerPoint's default style |
| "Bring to front" failed as soon as the learner drew ink afterward | In front now means nothing later on the slide overlaps it |
| `chart.is title=` meant both "the slide titled" and "the chart titled" | The chart's own title is `chartTitle=` |
| Push saves its default direction ("From Bottom") as `dir="u"`, so "change the effect option" would pass on a fresh Push | Known defaults are compared, not just "has an option" |
| Format Background > Apply to All puts the picture on the master, so the other slides still read as "no background" | `slide.is background=none` fails when the master has a picture background |
| The course certificate number started OT-WORD in every course (Excel too) | It now names the course: OT-WORD, OT-EXCEL, OT-PPT |
| A long web address in a quiz answer widened the panel past 320 px | Quiz options wrap; fieldsets can shrink |
| The course maps (all three) said every exam skill is "in the capstone" | They now say how many the capstone samples (Word 22, Excel 21, PowerPoint 22) |

From three independent content reviews (every fix is in the lesson files):

- **Compare is gone from Microsoft 365 PowerPoint for Windows** (version 2502, 2025) and was never in PowerPoint for Mac or the web. Lesson 5.2 teaches both Compare (PowerPoint 2024 and older) and a side-by-side method; the checks are the same.
- PowerPoint for the web has no Outline view, Notes Page view, WordArt styles or text columns: those steps are desktop, and D1.02 now works in either app.
- PowerPoint for the web **can** add footers and slide numbers ([[path: Insert > Footer > Footer]]): D1.10 and the SC1 and P1 steps are now web.
- Reuse Slides now opens files from your computer only (**Open a PowerPoint File**), not OneDrive.
- "Hide it again to show it" is **Unhide Slide**.
- Apply to All can't be undone with Reset Background: the hint says Ctrl+Z.
- D2.02's second background is a **solid** fill (the web app's gradient support is unclear).
- Templates appear under **File > New > Personal** only if saved in Custom Office Templates.
- PowerPoint for Mac has no Screen Recording, and **Cmd+Shift+Q logs you out of macOS**: the Mac note now points to QuickTime.
- Slide Zoom doesn't return to the zoom slide by default; Summary and Section Zoom do.
- Cycle layouts are under **More Layouts**, not the Layouts gallery's More arrow.
- Distribute spaces objects evenly between the outer two; it doesn't spread them across the slide.
- PowerPoint for the web's emphasis effects include Spin and Grow/Shrink, not Pulse.
- Captions work on Windows, Mac and the web (.vtt, and .srt in newer versions).
- Show Media Controls shows play controls; it isn't a recording option.
- PowerPoint for the web can't export video; D5.07's video and Word handout steps are desktop.
- @mentions need a work or school account; a private window on the same account doesn't show a second co-author; OneDrive saves versions every few minutes, so D5.02 says to wait.
- The add-in upload menu is now **Add-ins > More Settings > Upload My Add-in**; admins choose **Office Add-in** as the app type, and it can take a day to appear.
- Mac paths added: Cmd+Enter to post, File > Browse Version History, PowerPoint > Preferences > Save, Slide Show > Presenter View, Cmd+K.

## 2. Try these in real PowerPoint first (High)

The checker has been tested against files made by python-pptx and patched XML, not files saved by PowerPoint. Before students use it, open each starter in PowerPoint (Windows and the web), do the drill, and run Check my work:

1. **Every starter opens without a repair message**, especially the ones with modern comments (P5-L1, P5-L4, P5-SC5), animations (P4-L2, P4-SC4), sections and custom shows (P5-L3), the handout master (made in tests only) and the 3D model file (kayak.glb).
2. **Transition effect options:** what PowerPoint saves for each transition's default option. Only Push, Split and Random Bars have known defaults; if Vortex or Cube saves a direction by default, D4.01 step 3 would pass without a change. Fix by adding them to `TRANSITION_DEFAULTS`.
3. **3D transitions:** that Vortex, Ripple, Cube, Gallery and the other Exciting transitions are saved in the `p14`/`p15` namespaces (read as 3D), and Fade, Push and Wipe with a custom duration are still read as basic.
4. **3D model animations** (Turntable): D4.08 only checks that some animation targets the model. Look at the saved XML and, if it has its own preset class, check for it by name.
5. **Modern comments:** reply, resolve and delete in PowerPoint for the web, then check that `state=replied`, `resolved` and `absent` read correctly.
6. **Show settings:** that the Slide Show tab's Play Narrations and Use Timings boxes write `showNarration` and `useTimings` (D5.05 step 2).
7. **Inspect Document** removes the author and sets "remove personal information" (`doc.is personalInfo=removed`), and Mark as Final writes `_MarkAsFinal`.
8. **Screen Recording** saves an ordinary video shape (D4.07 checks for a video on the last slide).
9. **Show me where** with PowerPointApi 1.5 (selects the slide and shape) and without it (goes to the slide).
10. **PowerPoint for the web:** Crop > Aspect Ratio, the Shape tab name, Guides, animation Delay and Reorder, and transition Duration. The reviews couldn't confirm these; the drills are labeled web on the expectation that they work.

## 3. Content and course design

- **Medium: CLO 7 wording.** "Compare and Combine presentations" names a feature Microsoft retired. Suggested: "Collaborate on presentations: co-author, review, and combine changes from reviewers." Check which PowerPoint version the labs run.
- **Medium: videos.** Only three video slots exist (Week 1 tour, Week 1 sections, Week 2 master). Add one per lesson, like the Excel course, once recording starts. Every video needs checked captions before the lesson is marked ready.
- **Medium: example screenshots.** Seven screenshot steps need an example image (QAT, notes page print preview, template, co-authoring, Presenter View, video export, Word handouts).
- **Low: Facet theme.** Microsoft refreshed the theme gallery in 2026; the drill accepts any theme other than Office Theme, but a screenshot of the current gallery would help.
- **Low: Mac-only gaps.** Inspect Document and Create Handouts may be Windows-only. The lessons say "PowerPoint for Windows"; confirm on a Mac and add a Mac route if there is one.
- **Low: Northstar.** All 17 standards are covered, most in Week 1; if FLCC has a Northstar license, use it as a Day 1 diagnostic.

## 4. Checker and add-in ideas

- **Medium:** read handout "slides per page" if a future PowerPoint saves it, so D2.05 can check it.
- **Medium:** check captions on a video (a caption track part) once real files show how PowerPoint stores them.
- **Low:** a "readingOrder" fix-it hint that opens the Selection Pane.
- **Low:** recognize more animation effect names (only the common ones are named in messages).
- **Low:** the "Upload my file and check" button could say "Open my file and check", since nothing is uploaded. It's shared by all three courses, so change it everywhere at once.
