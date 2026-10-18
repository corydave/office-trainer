---
type: config
id: WORD
title: Microsoft Word
college: Finger Lakes Community College
license: CC BY 4.0
headingFont: Comfortaa
bodyFont: Atkinson Hyperlegible
learnerCanChangeFonts: true
videoHosts: [www.youtube-nocookie.com, www.youtube.com, player.vimeo.com]
defaultPlatform: web
---

## Course settings

These settings apply to every lesson in the course. Change a value above and rebuild; lesson files never repeat these settings.

## Fonts

`headingFont` and `bodyFont` set the default fonts for the course. When `learnerCanChangeFonts` is `true`, learners can pick their own pair in the add-in's settings (useful for dyslexia, low vision, or a simple preference). Their choice stays in their own browser.

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

Comfortaa and Atkinson Hyperlegible load from Google Fonts. Every other choice uses fonts already on the learner's computer, so nothing else is downloaded. OpenDyslexic is used only if the learner has installed it.

## Video hosts

Videos load only from the hosts in `videoHosts`. YouTube links are rewritten to `www.youtube-nocookie.com` (privacy-enhanced mode), which does not set tracking cookies until the learner presses play. Add FLCC's own video platform here (for example its Panopto, Kaltura or YuJa domain) once it is confirmed.

## License

Lesson text is shared under `CC BY 4.0`, so other colleges can reuse and adapt it with credit. Change this before publishing if FLCC prefers another license.
