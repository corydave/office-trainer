// validate.mjs - check every lesson file from the command line.
// Usage (from the word/ folder):  node tools/validate.mjs content/week-01
// Prints each problem with its file and line, then a summary. Exits with
// code 1 if any file has errors, so it can run in an automated check later.
import { readFileSync, readdirSync, existsSync, statSync } from "node:fs";
import { join, dirname } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const TM = require("./trainer-md.js");
const here = dirname(fileURLToPath(import.meta.url));
const courseRoot = join(here, "..");

// Read approved video hosts from course.config.md, if it is there.
let videoHosts;
const configPath = join(courseRoot, "course.config.md");
if (existsSync(configPath)) videoHosts = TM.parse(readFileSync(configPath, "utf8")).meta.videoHosts;

const targets = process.argv.slice(2);
if (!targets.length) { console.log("Usage: node tools/validate.mjs <folder or file> ..."); process.exit(2); }

let files = [];
for (const t of targets) {
  if (statSync(t).isDirectory()) files.push(...readdirSync(t).filter(f => f.endsWith(".md")).sort().map(f => join(t, f)));
  else files.push(t);
}

let errors = 0, warnings = 0;
for (const f of files) {
  const doc = TM.parse(readFileSync(f, "utf8"));
  const issues = TM.validate(doc, { videoHosts, starterExists: p => existsSync(join(courseRoot, p)) });
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
