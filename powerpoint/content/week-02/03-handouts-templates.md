---
type: lesson
id: PP2-L3
week: 2
order: 3
title: Handout and Notes Masters, and Templates
objectives: [W2.3, W2.4]
clos: [CLO2, CLO3]
minutes: 35
status: draft
---

::: callout why
Printed handouts and notes pages have masters too, so your organization's name and the date appear on every page without typing them. And once your master is set up, saving it as a template means nobody has to set it up again.
:::

## The handout master and the notes master

Both are on the **View** tab (desktop):

- **Handout Master** sets how printed handouts look: **Slides Per Page**, **Handout Orientation**, and the **header**, **footer**, **date** and **page number** placeholders in the corners.
- **Notes Master** sets how notes pages look: where the slide image and the notes box sit, the notes text font, and the same four corner placeholders.

Click in a placeholder (such as **Header**) and type, or clear a check box in the **Placeholders** group to remove it. Then select **Close Master View**. [[path: Insert > Header & Footer > Notes and Handouts]] fills the same placeholders for you.

::: drill D2.05
title: Set up printed handouts and notes pages
badge: Handout Pro
starter: starters/week-02/P2-L3-handouts.pptx
platform: desktop
mos: 1.1.5, 1.1.6
minutes: 6

- step: In the **Handout Master**, set **Slides Per Page** to **3 Slides** and type **Lakeshore Paddle Company** in the header placeholder.
  check: master.is kind=handout phSays="hdr:Lakeshore Paddle Company"
  hint: [[path: View > Handout Master]], then **Slides Per Page**. Click in the **Header** box at the top left and type.
- step: In the **Notes Master**, type **Staff use only** in the footer placeholder.
  check: master.is kind=notes phSays="ftr:Staff use only"
  hint: [[path: View > Notes Master]], then click in the **Footer** box at the bottom left.
- step: Close the master view and open [[path: File > Print]] with **Notes Pages** to see the footer.
:::

## Templates

A **template** (.potx) is a presentation that opens as a new, untitled copy, so the original never gets overwritten. Save one with [[path: File > Save As]] and choose **PowerPoint Template (*.potx)** as the file type (desktop). PowerPoint suggests your **Custom Office Templates** folder, which makes it appear under **File > New > Personal**; you can save it anywhere, such as your course folder.

To use it, double-click the .potx file, or choose it in [[path: File > New > Personal]]. You get a new presentation with the template's master, layouts and theme.

::: drill D2.06
title: Save and use a template
badge: Template Maker
starter: starters/week-02/P2-L3-handouts.pptx
platform: desktop
minutes: 5

- step: Save the presentation as a template named **Lastname_Lakeshore.potx** in your course folder.
  check: file.name pattern="*_Lakeshore.potx"
  hint: [[path: File > Save As]], **Browse**, set **Save as type** to **PowerPoint Template (*.potx)**, then pick your course folder.
- step: Close it, then double-click **Lastname_Lakeshore.potx** in File Explorer or Finder. Take a screenshot of the new, untitled presentation it opens.
  screenshot: The new presentation (titled Presentation1 or similar) made from your template.
  example: TODO
  describe: What's the name in the title bar, and why isn't it Lastname_Lakeshore?
:::

::: quiz Q23.1
question: Why save a branded deck as a .potx template instead of a .pptx you copy each time?
type: single
- [ ] A .potx file is smaller
- [x] Opening a template makes a new copy, so the original can't be overwritten by mistake
- [ ] Only templates can have a slide master
- [ ] A .potx file can't be edited
explain: Every presentation has a slide master. A template's job is to start new presentations safely: it opens as an untitled copy.
:::

## Check your understanding

- Which master controls the page number on printed handouts?
- What do you get when you double-click a .potx file?
- Where does PowerPoint list your own templates?
