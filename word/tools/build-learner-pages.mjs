// build-learner-pages.mjs - builds the two learner-facing pages from the Markdown:
//   roadmap.html      the subway-map course roadmap
//   certificate.html  a certificate viewer (front + back), with a sample record
// Run from the word/ folder:  node tools/build-learner-pages.mjs
// For another course, name its folder:  node tools/build-learner-pages.mjs ../excel
// (the pages are written into that folder). Nothing is typed into these pages
// by hand: change the Markdown, then rebuild.
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { join, dirname, resolve } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const TM = require("./trainer-md.js");
const here = dirname(fileURLToPath(import.meta.url));
const root = process.argv[2] ? resolve(process.argv[2]) : join(here, "..");
const read = (p) => readFileSync(join(root, p), "utf8");

// ---------- 1. Course settings and the outline ----------
const config = TM.parse(read("course.config.md")).meta;
const outline = TM.parse(read("content/outline.md"));
const problems = TM.validate(outline, { fileExists: (p) => existsSync(join(root, p)) }).filter((i) => i.level === "error");
if (problems.length) { console.error("Fix content/outline.md first:", problems); process.exit(1); }

// ---------- 2. The course map: one station per outline item ----------
// (TrainerMD.courseMap is shared with the add-in, so both always agree.)
const docs = {};
for (const w of outline.blocks.filter((b) => b.name === "week")) {
  for (const it of w.steps) {
    const f = it.extra.file;
    if (f && existsSync(join(root, f))) docs[f] = TM.parse(read(f));
  }
}
const { weeks, outcomes, mosNames } = TM.courseMap(outline, config, docs);

// Which app and exam this course is about, for the page text ("Microsoft Excel", "MO-210").
const APP_NAMES = { word: "Word", excel: "Excel", powerpoint: "PowerPoint" };
const appName = APP_NAMES[String(config.app || "word").toLowerCase()] || "Word";
const examCode = ((outline.blocks.find((b) => b.name === "certification") || {}).arg || "").trim();

const DATA = {
  course: { title: `${config.title} Course Roadmap`, college: config.college, issuer: config.issuer, license: config.license,
    mode: config.mode || "practice", certificateEach: config.certificateEach || "lesson" },
  app: { name: appName, fullName: "Microsoft " + appName, exam: examCode },
  weeks, outcomes, mosNames, labels: TM.ASSESS_LABELS,
  builtFrom: "content/outline.md and the lesson files",
};

// ---------- 3. Write the pages ----------
const lib = readFileSync(join(here, "trainer-md.js"), "utf8");
const json = JSON.stringify(DATA).replace(/</g, "\\u003c"); // safe inside a <script> tag
for (const [tpl, out] of [["roadmap.src.html", "roadmap.html"], ["certificate.src.html", "certificate.html"]]) {
  const src = readFileSync(join(here, "src", tpl), "utf8");
  if (lib.toLowerCase().includes("</script")) throw new Error("trainer-md.js must not contain a closing script tag");
  // __APP__ in a page's <head> becomes "Word" or "Excel" (the rest of the page reads DATA.app).
  writeFileSync(join(root, out), src.replace("/*__TRAINER_MD__*/", lib).replace("__DATA_JSON__", json).replace(/__APP__/g, appName));
  console.log("Wrote", out);
}
const count = (f) => weeks.reduce((n, w) => n + w.stations.filter(f).length, 0);
console.log(`${weeks.length} weeks, ${count(() => true)} stations, ${count((s) => s.status !== "planned")} written, ` +
  `${weeks.reduce((n, w) => n + w.stations.reduce((m, s) => m + s.activities.length, 0), 0)} badges, ${count((s) => s.certificate)} certificates`);
