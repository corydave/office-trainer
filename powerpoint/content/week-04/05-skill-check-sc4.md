---
type: skillcheck
id: PP4-SC4
week: 4
order: 5
title: "Skill Check 4: Motion, Media and Links"
objectives: [W4.1, W4.2, W4.3, W4.4]
clos: [CLO2, CLO4, CLO6]
minutes: 35
status: draft
---

One starter presentation, a set of tasks and a time limit, with no hints. Each task is worth one point. The timing, motion path, video and Zoom tasks need desktop PowerPoint.

::: skillcheck SC4
title: Kiosk tour
badge: Skill Check 4
starter: starters/week-04/P4-SC4-kiosk.pptx
files: starters/week-04/tour.mp4
platform: web
mos: 2.1.4, 3.2.1, 3.2.2, 4.5.1, 4.5.3, 5.1.1, 5.1.2, 5.2.1, 5.2.3, 5.2.4, 5.2.5
timeLimit: 35
attempts: 2
passing: 70

- task: Apply the **Push** transition to every slide.
  check: transition.is slide=all type=push
- task: Change the Push transition's effect option on every slide to any direction other than the default.
  check: transition.is slide=all optionChanged=true
- task: Set every slide to advance automatically after **8 seconds**.
  platform: desktop
  check: transition.is slide=all advanceAfter=8
- task: On **Our boats**, animate the bulleted list with **Fly In**, one paragraph at a time.
  check: anim.is title="Our boats" says="Kayaks" effect="Fly In" byParagraph=true
- task: Set the list's animation to start **After Previous**.
  check: anim.is title="Our boats" says="Kayaks" start=after
- task: Give the logo on **Our boats** the **Spin** emphasis effect.
  check: anim.is title="Our boats" name=Logo effect=Spin
- task: Move the logo's animation so it runs first.
  check: anim.is title="Our boats" name=Logo first=true
- task: On **The route**, give the kayak a motion path that moves it to the **right**.
  platform: desktop
  check: anim.is title="The route" name=Kayak class=path pathDir=right
- task: On **Tour**, insert **tour.mp4** and set it to start automatically.
  platform: desktop
  check: media.is title="Tour" kind=video start=auto
- task: Trim the video.
  platform: desktop
  check: media.is title="Tour" kind=video trimmed=true
- task: On **Book**, link the words **Book now** to `https://www.lakeshorepaddle.example/rentals`.
  check: link.present title="Book" url="lakeshorepaddle.example/rentals" says="Book now"
- task: Insert a **Summary Zoom** of the presentation's sections.
  platform: desktop
  check: zoom.is kind=summary
- task: On **Menu**, insert a **Slide Zoom** to the **Book** slide.
  platform: desktop
  check: zoom.is title="Menu" kind=slide
:::
