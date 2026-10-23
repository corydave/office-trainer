---
type: lesson
id: W4-L1
week: 4
order: 1
title: Templates and Quick Parts
objectives: [W4.1]
clos: [CLO2]
minutes: 35
status: draft
---

::: callout why
If you rebuild your resume header every time you apply for something, you'll eventually send one with an old phone number. Templates and Quick Parts let you set something up once and reuse it, exactly the same, every time.
:::

## Your scenario

This week you're putting together a complete job application packet: a resume in two versions, cover letters personalized for several employers, envelopes or labels, and a web version of your resume. You can use your own details, or the made-up job seeker **Alex Morgan** who appears in the practice files.

::: callout privacy
Your resume and letters hold personal details. You may always use a made-up persona instead. The recipient list you'll build in Lesson 3 holds other people's details too: keep it in your course folder, and delete it when the project is graded.
:::

## Templates

A **template** (.dotx) is a starting point. Opening it makes a new, untitled document; the template itself stays unchanged for next time.

- **Built-in and online templates:** [[path: File > New]], then search (for example "resume").
- **Your own:** set up a document, then [[path: File > Save As]] and choose **Word Template (*.dotx)** as the type. Desktop Word saves it in your Custom Office Templates folder, and it appears under [[path: File > New > Personal]].

Many templates use **content controls**: boxes with placeholder text such as `[Your name]`. Click one and type; the placeholder disappears.

::: video
id: w4-l1-templates
url: TODO
title: Fill in a template and save your own
captions: none
duration: 3:00
:::

::: drill D4.01
title: Fill in a resume template and save your own
badge: Template Maker
starter: starters/week-04/W4-L1-resume-template.docx
platform: desktop
other: E1.1, E4.1
minutes: 8

- step: Replace every placeholder in the template with your own details (or Alex Morgan's).
  check: controls.filled
  hint: Click each bracketed placeholder and type. The Experience section has three, and Education has two.
- step: Save the finished file as a template named **Lastname_Resume.dotx**.
  check: file.name pattern="*_Resume.dotx"
  hint: [[path: File > Save As > Browse]], then set **Save as type** to **Word Template (*.dotx)**. On a Mac, choose **Word Template** in **File Format**.
- step: Close it, then open a new document from it with [[path: File > New > Personal]] to see that it starts filled in.
:::

## Quick Parts

A **Quick Part** is a saved piece of content you can drop into any document: a signature block, a contact line, a standard paragraph.

1. Select the content (for example your name and contact line).
2. [[path: Insert > Quick Parts > Save Selection to Quick Part Gallery]], give it a name, and select **OK**.
3. To use it, click where it should go and choose it from [[path: Insert > Quick Parts]]. Or type its name and press [[key: F3]].

::: platform web
Templates you saved in desktop Word open in Word for the web, but Quick Parts and saving as .dotx need desktop Word.
:::

::: platform mac
On a Mac, your saved templates are under [[path: File > New from Template]], and reusable text is called **AutoText** ([[path: Insert > AutoText > New]]). It works the same way as a Quick Part.
:::

::: callout tip
On a campus lab computer, Quick Parts are saved in a file called **Building Blocks**. When you close Word it may ask whether to save changes to it: choose **Save**. Lab computers may still clear it when you sign out, so keep a copy of your contact line in your course folder too.
:::

::: drill D4.02
title: Save and reuse a Quick Part
badge: Quick Parts Pro
starter: starters/week-04/W4-L1-cover-letter.docx
platform: desktop
other: E3.1
minutes: 6

- step: In any document, type your name and contact line (email, phone, city), select them, and save them as a Quick Part named **Contact**.
  screenshot: The Quick Parts gallery open, showing your Contact entry.
  example: TODO
  describe: Say what your Quick Part contains.
- step: In the cover letter, replace **[Contact line]** with your Contact Quick Part.
  check: text.absent value="[Contact line]"
  hint: Select [Contact line], then choose your entry from [[path: Insert > Quick Parts]].
:::

## Check your understanding

- What happens to a template file when you open a new document from it?
- When is a Quick Part better than copying and pasting?
