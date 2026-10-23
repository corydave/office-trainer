// validate.mjs - check every lesson file from the command line.
// Usage (from the word/ folder):  node tools/validate.mjs content/week-01
// Prints each problem with its file and line, then a summary. Exits with
// code 1 if any file has errors, so it can run in an automated check later.
import { readFileSync, readdirSync, existsSync, statSync } from "node:fs";
import { join, dirname, resolve } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const TM = require("./trainer-md.js");
const here = dirname(fileURLToPath(import.meta.url));

const targets = process.argv.slice(2);
if (!targets.length) { console.log("Usage: node tools/validate.mjs <folder or file> ...  (for Excel: node word/tools/validate.mjs excel/content ...)"); process.exit(2); }

// The course folder is the nearest folder above the first target that has a
// course.config.md (word/ or excel/). Its "app:" setting picks the checks.
function findCourseRoot(start) {
  let dir = resolve(start);
  if (existsSync(dir) && !statSync(dir).isDirectory()) dir = dirname(dir);
  for (let i = 0; i < 8; i++) {
    if (existsSync(join(dir, "course.config.md"))) return dir;
    const up = dirname(dir);
    if (up === dir) break;
    dir = up;
  }
  return join(here, "..");
}
const courseRoot = findCourseRoot(targets[0]);

// Read approved video hosts and the app from course.config.md, if it is there.
let videoHosts, app = "word";
const configPath = join(courseRoot, "course.config.md");
if (existsSync(configPath)) {
  const meta = TM.parse(readFileSync(configPath, "utf8")).meta;
  videoHosts = meta.videoHosts;
  app = meta.app || "word";
}

let files = [];
for (const t of targets) {
  if (statSync(t).isDirectory()) files.push(...readdirSync(t).filter(f => f.endsWith(".md")).sort().map(f => join(t, f)));
  else files.push(t);
}

// Index every video id and activity id in the course, so section files can be
// checked against them (a typo in a video id would otherwise go unnoticed).
const walk = (dir) => readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
  e.isDirectory() ? walk(join(dir, e.name)) : e.name.endsWith(".md") ? [join(dir, e.name)] : []);
const index = TM.courseIndex(walk(join(courseRoot, "content")).map((f) => TM.parse(readFileSync(f, "utf8"))));

let errors = 0, warnings = 0;
for (const f of files) {
  const doc = TM.parse(readFileSync(f, "utf8"));
  const issues = TM.validate(doc, { app, videoHosts, starterExists: p => existsSync(join(courseRoot, p)), fileExists: p => existsSync(join(courseRoot, p)),
    videoIds: index.videoIds, anchorIds: index.anchorIds, stopIds: index.stopIds, fileName: f.split("/").pop() });
  const s = TM.summarize(doc);
  console.log(`\n${f}  (${doc.meta.type}, ${s.drills} drills, ${s.checked}/${s.steps} steps checked, ${s.quizzes} quizzes, ${s.videos} videos)`);
  for (const i of issues) {
    console.log(`  ${i.level === "error" ? "ERROR  " : "warning"} line ${String(i.line).padStart(3)}  ${i.message}`);
    if (i.level === "error") errors++; else warnings++;
  }
  if (!issues.length) console.log("  no problems");
}
console.log(`\n${files.length} files, ${errors} errors, ${warnings} warnings`);
process.exit(errors ? 1 : 0);
