---
type: lesson
id: XL4-L4
week: 4
order: 4
title: Themes, Cell Styles and Templates
objectives: [W4.4]
clos: [CLO1, CLO2]
minutes: 30
status: draft
---

::: callout why
When every workbook a business sends out uses the same fonts, colors and layout, people trust it and find things faster. Themes, cell styles and templates let you set that up once and reuse it, instead of formatting every new workbook by hand.
:::

## Themes

A **theme** is a set of colors, fonts and effects for the whole workbook. [[path: Page Layout > Themes]] (desktop) changes them all at once: every cell, chart and shape that uses theme colors updates together. **Colors**, **Fonts** and **Effects** beside it change just one part.

::: callout tip
Pick colors from the **Theme Colors** rows of any color menu, not **Standard Colors**. Theme colors follow the theme when you change it; standard colors stay as they are.
:::

## Cell styles

A **cell style** is a saved set of formatting (number format, font, fill, borders) with a name. [[path: Home > Cell Styles]] has built-in ones like **Title**, **Total**, **Input** and **Good**, and you can make your own:

1. Format one cell the way you want it.
2. [[path: Home > Cell Styles > New Cell Style]], type a name, and choose **OK**.
3. Select other cells and click your style in the gallery.

Change a style later (right-click it in the gallery > **Modify**) and every cell that uses it updates.

## Templates

A **template** (.xltx) is a starting workbook. Opening one from [[path: File > New > Personal]], or by double-clicking it, makes a new, untitled copy, so the original stays clean for next time. ([[path: File > Open]] opens the template itself, for when you want to change it.) Save one with [[path: File > Save As]] and choose **Excel Template (*.xltx)** as the type. Excel suggests your Custom Office Templates folder, so it shows up under [[path: File > New > Personal]].

::: video
id: e4-l4-themes
url: TODO
title: Themes, cell styles and saving a template
captions: none
duration: 3:30
:::

::: drill D4.07
title: Make the trip planner a reusable template
badge: Template Maker
starter: starters/week-04/E4-L4-planner.xlsx
platform: desktop
mos: 2.2.7
minutes: 8

- step: Apply any theme other than **Office**.
  check: theme.is changed=true
  hint: [[path: Page Layout > Themes]].
- step: Format B3 with a light fill and a border, then save that as a new cell style named **Entry**.
  check: cellstyle.is name=Entry custom=true
  hint: [[path: Home > Cell Styles > New Cell Style]]. (Excel already has a built-in style called Input, so pick a new name.)
- step: Apply the **Entry** style to the three input cells, B3:B5.
  check: cell.is sheet=Planner range=B3:B5 style=Entry
- step: Save the workbook as an Excel template named **Lastname_Planner.xltx**, using your own last name.
  check: file.name pattern="*_Planner.xltx"
  hint: [[path: File > Save As]], then choose **Excel Template (*.xltx)** under the file name.
:::

::: quiz Q44.1
question: You open a template from File > New > Personal, fill it in and press Save. What happens to the template?
type: single
- [ ] It's overwritten with your numbers
- [x] Nothing: Excel asks where to save the new workbook, and the template stays blank
- [ ] It turns into a regular workbook
explain: Opening a template makes a new, untitled copy. Saving saves the copy, so the template is ready for the next person.
:::

## Check your understanding

- Why use theme colors instead of standard colors?
- What happens to cells that use a style when you modify the style?
- What's the difference between saving a workbook and saving a template?
