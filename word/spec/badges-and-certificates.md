# Badges and certificates (v0.1)

Learners earn a **badge** for every drill, skill check and project, and a **certificate** at the end of every lesson, plus one for the whole course. This file defines what they look like, the record behind them, and how far that record can be trusted.

## 1. What is earned

| Award | Earned when | Shown as |
|---|---|---|
| Drill badge | Every checked step passes (screenshot steps: confirmed in practice mode, marked complete by the instructor in a section) | Hexagon in the week's color with a check mark |
| Screenshot-only badge | A drill made only of screenshot steps (for example D1.02, the Quick Access Toolbar) | Same hexagon with a camera |
| Skill check badge | The skill check is submitted at or above its `passing` score | Hexagon with a star |
| Project badge | Practice: the learner completes the rubric self-check. Section: the instructor grades it | Hexagon with a document |
| Lesson certificate | All of the lesson's drill badges are earned (`certificateEach: lesson`) | Two-page certificate (front and back) |
| Course certificate | Every lesson certificate, skill check and project in the course (`courseCertificate: true`) | Two-page certificate, all badges on the back |

Badge art is drawn by code (`TrainerMD.badgeSVG`), so there are no image files to keep in sync. The name comes from the activity's `badge:` line, or its title.

## 2. The certificate

- **Front:** a traditional certificate: issuer, "Certificate of Practice" (practice mode) or "Certificate of Completion" (section), the learner's name, the lesson, the date, a seal in the week's color and the certificate number. Practice certificates carry a "Practice record" stamp. Section certificates carry the instructor's name.
- **Back:** the record of learning: every badge earned with its own number and how it was assessed, the course learning outcomes the lesson supports, and the MO-110 skills it practiced.
- Each side is one landscape Letter page, so printing two-sided makes a real certificate. `certificate.html` previews it and prints or saves it as PDF.

## 3. The record

Every certificate is drawn from a small JSON record. The add-in creates it when a learner earns the lesson's last badge. **View and print** opens `certificate.html` with the record in the part of the address after `#`, which browsers never send to the web server; the page then removes it from the address bar. **Save record (.json)** keeps a copy that `certificate.html` can open later.

```json
{
  "format": "office-trainer-certificate/1",
  "id": "OT-W1L2-2610-RM65-YHSK9",
  "issuedAt": "2026-10-18",
  "issuer": "Finger Lakes Community College, Computing Sciences",
  "mode": "practice",
  "learner": { "name": "Jordan Rivera" },
  "section": null,
  "achievement": { "kind": "lesson", "id": "w1-clean-up-the-text", "title": "Clean Up the Text", "week": 1 },
  "badges": [
    { "id": "OTB-D103-2610-YRH2-9M2D4", "activity": "D1.03", "name": "Space Cleaner", "kind": "drill",
      "week": 1, "assessedBy": ["auto"], "earnedAt": "2026-10-18" }
  ],
  "outcomes": ["CLO2"],
  "mos": ["1.1.1", "1.1.4", "2.1.1", "2.1.2"]
}
```

Only the learner's name is personal, and in practice mode the learner types it themselves. The record holds no email address, student ID or scores.

## 4. Numbers

Format: `PREFIX-YYMM-XXXX-XXXXC`

- `PREFIX` is `OT-<lesson>` for a certificate (for example `OT-W1L2`), `OT-WORD` for the course certificate, and `OTB-<activity>` for a badge (`OTB-D103`).
- `YYMM` is the month issued.
- `XXXXXXXX` is 8 random characters from a browser's secure random generator, in Crockford base32. It has no I, L, O or U, so a number can be read aloud without confusing 1 with I or 0 with O.
- `C` is a check character that catches most typing mistakes (`TrainerMD.idLooksValid`).

**What a number proves:** that it was typed correctly. **What it doesn't prove:** that the person earned it. In practice mode, anyone with a browser could make a certificate that looks real. That's why practice certificates say "Practice record" on both sides.

## 5. Should we keep a lookup registry?

A registry means a database of every certificate and badge, plus a public page where an employer types a number and sees "valid, issued to Jordan Rivera on October 18, 2026". It's possible, but it's a real commitment.

**What running a registry commits us to:**

- **Keeping it running for years.** A certificate is only as good as its lookup page, and employers check old ones.
- **Treating it as student records.** In a section, a certificate reflects course work, so it is an education record under FERPA. Publishing it needs the student's consent.
- **Handling corrections.** Revoking certificates issued in error, and fixing names.
- **Protecting it.** A public lookup page is something people will try to scrape or abuse.

**The three options:**

| Option | Who can trust it | What we run | When |
|---|---|---|---|
| **A. Practice records** (built now) | The learner, as their own record | Nothing | Now |
| **B. Registry with a lookup page** | Anyone who checks the page | A database and public page on the Phase 2 server, forever | Not recommended unless FLCC wants to own it |
| **C. Signed digital credentials (Open Badges 3.0)** | Anyone, by checking FLCC's digital signature | The Phase 2 server signs each badge when it verifies the work; no lookup database needed | With Phase 2 |

**Recommendation: A now, C later, and avoid B.**

- **Open Badges 3.0** is the 1EdTech standard released in May 2024. It is built on W3C Verifiable Credentials. A signed badge carries its own proof: anyone can verify it against FLCC's published key, and learners can keep it in a digital wallet or add it to LinkedIn.
- Signing must happen on a server, because a key kept in the browser could sign anything. It belongs in Phase 2, once that server re-checks submitted work.
- **Check first:** if FLCC already uses a digital badge platform (for example for the AI Microcredential Program), section badges could be issued there instead of building our own signing.

## 6. Mapping to Open Badges 3.0 later

| Record field | Open Badges 3.0 |
|---|---|
| `achievement` | `achievement` (an Achievement with name, description and criteria) |
| `badges[]` | One OpenBadgeCredential each |
| `outcomes`, `mos` | `achievement.alignment[]`, pointing to the course outcomes and the MO-110 skills |
| `issuer` | `issuer` (a Profile for FLCC) |
| `learner.name` | `credentialSubject` |
| `id` | Kept as the human-readable number. The credential's own id is a URL |

## 7. Credly (FLCC's current badge platform)

FLCC uses Credly today, for how long is unknown. Treat it as an extra place to publish milestones, not the home of the record.

- **Milestones only:** the course certificate, and optionally one badge per week. Drill badges stay in the trainer; dozens of small badges would bury a learner's Credly profile.
- **Sections only:** Credly needs an earner's email address, which a section has (Microsoft sign-in) and practice mode deliberately doesn't.
- **Before Phase 2:** the instructor issues the Credly badges at the end of term, using the trainer's export of who finished what.
- **Phase 2:** the server issues them through Credly's API. The API key stays on the server, never in the add-in.
- **If Credly goes away,** nothing is lost: this record format and the Open Badges 3.0 plan (section 5) don't depend on it. Credly has supported Open Badges 3.0 since October 2024.
