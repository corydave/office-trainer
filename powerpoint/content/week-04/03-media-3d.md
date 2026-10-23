---
type: lesson
id: PP4-L3
week: 4
order: 3
title: Video, Audio and 3D Models
objectives: [W4.3]
clos: [CLO4, CLO6]
minutes: 45
status: draft
---

::: callout why
Ten seconds of video can show what a paragraph can't, and a 3D model lets people turn an object around. But media has to start at the right moment, stop where it should, and work for people who can't hear or see it.
:::

## Insert and play video and audio

- **Video:** [[path: Insert > Video > This Device]] (or **Stock Videos**, or **Online Videos** for a link).
- **Audio:** [[path: Insert > Audio > Audio on My PC]] (desktop; on a Mac, **Audio from File**).

Select the clip, then use the **Playback** tab (desktop):

| Setting | What it does |
|---|---|
| **Trim Video** / **Trim Audio** | Cuts the start and end, without changing the file |
| **Fade In** / **Fade Out** | Eases the sound and picture in and out |
| **Start: Automatically** | Plays as soon as the slide appears. **When Clicked On** waits for a click on the clip; **In Click Sequence** plays on your next click |
| **Play Across Slides** | Keeps an audio clip playing when the slide changes |
| **Loop until Stopped** | Repeats it |
| **Play in Background** | One click for background music: start automatically, play across slides, loop, and hide the icon |

::: callout tip
Video needs **captions** for people who can't hear it: [[path: Playback > Insert Captions]] adds a caption file (.vtt, or .srt in newer versions). Give every clip alt text that says what it shows, and never let audio start on its own in a talk without telling people.
:::

::: drill D4.06
title: Add a tour video and music
badge: Media Mixer
starter: starters/week-04/P4-L3-media.pptx
files: starters/week-04/tour.mp4, starters/week-04/music.m4a
platform: desktop
mos: 4.5.1, 4.5.3
other: N7
minutes: 10

Download **tour.mp4** and **music.m4a** first.

- step: On **Tour video**, insert **tour.mp4**.
  check: media.is title="Tour video" kind=video
  hint: [[path: Insert > Video > This Device]].
- step: Trim it to play from about **0:01** to **0:06**.
  check: media.is title="Tour video" kind=video trimmed=true
  hint: [[path: Playback > Trim Video]], then drag the green (start) and red (end) markers.
- step: Set it to start **automatically** and **fade in** over 1 second.
  check: media.is title="Tour video" kind=video start=auto fadeIn=true
  hint: [[path: Playback > Start > Automatically]], and type 01.00 in **Fade In**.
- step: On the title slide, insert **music.m4a** and set it to **Play in Background**.
  check: media.is slide=1 kind=audio acrossSlides=true loop=true
  hint: [[path: Insert > Audio > Audio on My PC]], then [[path: Playback > Play in Background]].
:::

## Screen recordings

[[path: Insert > Screen Recording]] (Windows) records part of your screen, with your voice if you like, and drops the video on the slide. Choose **Select Area**, drag around what to record, then **Record**. Press [[key: Windows logo key+Shift+Q]] to stop. PowerPoint for Mac has no screen recorder: record with QuickTime Player (**File > New Screen Recording**), then insert the file with [[path: Insert > Video > Movie from File]].

::: callout privacy
Close email, chat and anything personal before you record. A screen recording keeps everything that was visible, including notifications that pop up.
:::

::: drill D4.07
title: Record the booking page
badge: Screen Recorder
starter: starters/week-04/P4-L3-media.pptx
platform: desktop
mos: 4.5.2
minutes: 6

- step: Open any web page in your browser (a page from your college's website is fine). Then, on the last slide, **How to book online**, insert a short screen recording (10 seconds or less) of you scrolling it.
  check: media.is slide=last kind=video
  hint: [[path: Insert > Screen Recording]], **Select Area**, drag around the browser, then **Record**. Stop with [[key: Windows logo key+Shift+Q]]. (Mac: record with QuickTime Player, then insert the video.)
:::

## 3D models

[[path: Insert > 3D Models > This Device]] (or **Stock 3D Models**) inserts a model you can turn in any direction: drag the rotation control in the middle, or pick a **3D Model Views** preset on the **3D Model** tab. 3D models also have their own animations: **Arrive**, **Turntable**, **Swing**, **Jump & Turn** and **Leave**.

::: drill D4.08
title: Show off the new kayak
badge: 3D Viewer
starter: starters/week-04/P4-L3-media.pptx
files: starters/week-04/kayak.glb
platform: desktop
mos: 4.4.1, 4.4.2, 5.2.2
minutes: 6

Download **kayak.glb** first.

- step: On **Our new kayak**, insert **kayak.glb** as a 3D model.
  check: object.is title="Our new kayak" kind=model3d
  hint: [[path: Insert > 3D Models > This Device]].
- step: Turn it to a new view.
  check: object.is title="Our new kayak" kind=model3d rotated=true
  hint: Pick a preset in [[path: 3D Model > 3D Model Views]], or drag the rotation control.
- step: Give it the **Turntable** animation.
  check: anim.is title="Our new kayak" kind=model3d
  hint: Select the model, then [[path: Animations > Turntable]] (in the 3D group of the gallery).
:::

::: quiz Q43.1
question: You want quiet music to play under a photo slideshow from the first slide to the last. Which setting does it in one click?
type: single
- [ ] Start: When Clicked On
- [ ] Trim Audio
- [x] Play in Background
- [ ] Fade Out
explain: Play in Background sets the clip to start automatically, play across slides, loop and hide its icon.
:::

## Check your understanding

- What does Trim Video change, and what does it leave alone?
- What two things make a video accessible?
- What should you close before a screen recording?
