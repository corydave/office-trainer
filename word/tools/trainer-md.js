/* ===========================================================================
   trainer-md.js - the Trainer Markdown parser, validator and renderer
   ---------------------------------------------------------------------------
   One file, no dependencies. The editor, the Word add-in and the command-line
   validator all use this same code, so a lesson means the same thing everywhere.

   What it does, in order:
     1. parse(text)     -> turns a lesson file into a plain JavaScript object
     2. validate(doc)   -> lists errors and warnings (with line numbers)
     3. render(doc)     -> turns the object into safe HTML for a preview

   Security rules this file enforces:
     - Every piece of text is HTML-escaped before anything else happens.
       Only tags this file writes itself ever reach the page.
     - Raw HTML in a lesson file is reported as an error and shown as text.
     - Links must be https:// (or a relative .md file); videos must come from
       an approved host and render in a sandboxed frame.
     - A lesson can only affect grading by NAMING a check in CHECKS below.
       It can never supply code.
   =========================================================================== */
(function (root) {
  "use strict";

  const VERSION = "0.1.0";

  // -------------------------------------------------------------------------
  // Fonts: each choice expands to a fallback chain (see course.config.md).
  // Look-alike fonts (Arimo, Tinos, Carlito) cover Chromebooks and Linux.
  // -------------------------------------------------------------------------
  const FONT_STACKS = {
    "Comfortaa": '"Comfortaa", "Atkinson Hyperlegible", Arial, sans-serif',
    "Atkinson Hyperlegible": '"Atkinson Hyperlegible", Arial, Arimo, sans-serif',
    "Arial": 'Arial, Arimo, "Liberation Sans", Helvetica, sans-serif',
    "Helvetica": 'Helvetica, "Helvetica Neue", Arial, Arimo, sans-serif',
    "Verdana": 'Verdana, "DejaVu Sans", Arial, sans-serif',
    "Calibri": 'Calibri, Carlito, Arial, sans-serif',
    "Aptos": 'Aptos, Calibri, Carlito, Arial, sans-serif',
    "Times New Roman": '"Times New Roman", Tinos, "Liberation Serif", Times, serif',
    "Georgia": 'Georgia, Gelasio, "Times New Roman", Tinos, serif',
    "OpenDyslexic": 'OpenDyslexic, "Atkinson Hyperlegible", Arial, sans-serif',
    "System default": 'system-ui, -apple-system, "Segoe UI", Roboto, sans-serif',
  };
  const DEFAULT_FONTS = { heading: "Comfortaa", body: "Atkinson Hyperlegible" };
  function fontStack(name, fallbackName) {
    return FONT_STACKS[name] || FONT_STACKS[fallbackName] || FONT_STACKS["System default"];
  }

  // Hosts videos may load from unless course.config.md says otherwise.
  const DEFAULT_VIDEO_HOSTS = ["www.youtube-nocookie.com", "www.youtube.com", "player.vimeo.com"];

  // -------------------------------------------------------------------------
  // The checker library catalog. Each entry lists the parameters a check
  // accepts and their type. The real check code lives in the add-in; this
  // list is what the validator uses to reject typos and unknown checks.
  // -------------------------------------------------------------------------
  const SELECTORS = ["in", "para", "paras", "text", "style"];
  const CHECKS = {
    "file.name":            { params: { pattern: "string!" } },
    "text.present":         { params: { value: "string!", min: "int", max: "int", case: "enum:sensitive|insensitive" }, selector: true },
    "text.absent":          { params: { value: "string!", case: "enum:sensitive|insensitive" }, selector: true },
    "text.noDoubleSpaces":  { params: {}, selector: true },
    "para.noEmpty":         { params: { max: "int", after: "string", before: "string" } },
    "para.order":           { params: { first: "string!", then: "string!" } },
    "font.is":              { params: { name: "string", size: "number", bold: "bool", italic: "bool", underline: "bool", color: "color" }, selector: true, needsOneOf: ["name", "size", "bold", "italic", "underline", "color"] },
    "effect.present":       { params: { kind: "enum:shadow|glow|reflection|outline|any" }, selector: true },
    "format.same":          { params: { a: "string!", b: "string!", props: "string" } },
    "format.clean":         { params: {}, selector: true },
    "para.is":              { params: { align: "enum:left|center|right|justify", lineSpacing: "lineSpacing", spaceBefore: "length", spaceAfter: "length", firstLineIndent: "length", leftIndent: "length" }, selector: true, needsOneOf: ["align", "lineSpacing", "spaceBefore", "spaceAfter", "firstLineIndent", "leftIndent"] },
    "page.is":              { params: { margins: "length", top: "length", bottom: "length", left: "length", right: "length", size: "enum:letter|a4|legal", orientation: "enum:portrait|landscape" }, needsOneOf: ["margins", "top", "bottom", "left", "right", "size", "orientation"] },
    "props.is":             { params: { title: "string", author: "string", keywords: "string" }, needsOneOf: ["title", "author", "keywords"] },
    "lang.is":              { params: { value: "lang!" } },
    "self.attest":          { params: { label: "string!" }, notInSkillCheck: true },
  };

  const DIRECTIVES = ["video", "callout", "platform", "quiz", "drill", "skillcheck", "preflight", "rubric"];
  const CALLOUT_KINDS = ["tip", "note", "warning", "why", "ai", "privacy"];
  const PLATFORMS = ["web", "desktop", "windows", "mac"];
  const FILE_TYPES = ["lesson", "skillcheck", "project", "config"];

  // -------------------------------------------------------------------------
  // Small helpers
  // -------------------------------------------------------------------------

  /** Make text safe to put inside HTML (the most important function here). */
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  /** Convert "1in", "36pt", "2.54cm" to points. Returns null if it isn't a length. */
  function toPoints(v) {
    const m = /^(\d+(?:\.\d+)?)(in|pt|cm|mm)$/.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    return { in: n * 72, pt: n, cm: n * 72 / 2.54, mm: n * 72 / 25.4 }[m[2]];
  }

  /** "[a, b, c]" -> ["a","b","c"]; "a, b" -> ["a","b"]; anything else -> the string. */
  function parseValue(raw) {
    const v = raw.trim();
    if (/^\[.*\]$/.test(v)) return v.slice(1, -1).split(",").map(x => unquote(x.trim())).filter(Boolean);
    if (v === "true") return true;
    if (v === "false") return false;
    return unquote(v);
  }
  function unquote(v) { return /^".*"$/.test(v) ? v.slice(1, -1) : v; }
  function splitList(v) { return Array.isArray(v) ? v : String(v || "").split(",").map(x => x.trim()).filter(Boolean); }

  // -------------------------------------------------------------------------
  // 1. PARSE
  // -------------------------------------------------------------------------

  /**
   * Parse a whole Trainer Markdown file.
   * Returns { meta, blocks, issues } where issues are problems found while
   * reading (validate() adds more).
   */
  function parse(text) {
    const lines = String(text).replace(/\r\n?/g, "\n").split("\n");
    const issues = [];
    const meta = {};
    let i = 0;

    // --- front matter: between the first two "---" lines ---
    if (lines[0] && lines[0].trim() === "---") {
      i = 1;
      while (i < lines.length && lines[i].trim() !== "---") {
        const line = lines[i];
        if (line.trim()) {
          const m = /^([A-Za-z][\w-]*)\s*:\s*(.*)$/.exec(line);
          if (m) meta[m[1]] = parseValue(m[2]);
          else issues.push(issue("error", i + 1, `Front matter line is not "key: value": ${line.trim()}`));
        }
        i++;
      }
      if (i >= lines.length) issues.push(issue("error", 1, "Front matter is not closed with a --- line."));
      i++; // skip closing ---
    } else {
      issues.push(issue("error", 1, "The file must start with front matter between two --- lines."));
    }

    // --- body: Markdown chunks and ::: directives ---
    const blocks = [];
    let mdBuffer = [];
    let mdStart = i + 1;
    const flushMd = () => {
      if (mdBuffer.length) blocks.push({ kind: "md", lines: mdBuffer, line: mdStart });
      mdBuffer = [];
    };

    while (i < lines.length) {
      const line = lines[i];
      const open = /^:::\s*([a-z]+)\s*(.*)$/.exec(line.trim());
      if (open && line.trim() !== ":::") {
        flushMd();
        const start = i + 1;
        const inner = [];
        i++;
        while (i < lines.length && lines[i].trim() !== ":::") { inner.push(lines[i]); i++; }
        if (i >= lines.length) issues.push(issue("error", start, `The ::: ${open[1]} block that starts here is never closed with :::`));
        blocks.push(parseDirective(open[1], open[2].trim(), inner, start, issues));
        i++;
        mdStart = i + 1;
        continue;
      }
      if (!mdBuffer.length) mdStart = i + 1;
      mdBuffer.push(line);
      i++;
    }
    flushMd();
    return { meta, blocks, issues };
  }

  function issue(level, line, message) { return { level, line, message }; }

  /** Parse what is inside a ::: block. `start` is the line number of the ::: line. */
  function parseDirective(name, arg, inner, start, issues) {
    const block = { kind: "directive", name, arg, line: start, header: {}, headerLines: {}, intro: [], steps: [], body: inner, items: [] };
    if (!DIRECTIVES.includes(name)) {
      issues.push(issue("error", start, `Unknown directive "::: ${name}". Use one of: ${DIRECTIVES.join(", ")}.`));
      return block;
    }

    // Free-text directives: their inside is Markdown.
    if (name === "callout" || name === "platform") { block.md = inner; return block; }

    // Rubric: "- id | points | CLOs | description"
    if (name === "rubric") {
      inner.forEach((l, k) => {
        if (!l.trim()) return;
        const m = /^-\s*(.+)$/.exec(l.trim());
        const parts = m ? m[1].split("|").map(s => s.trim()) : [];
        if (parts.length !== 4) { issues.push(issue("error", start + k + 1, 'Rubric lines look like "- P1-a | 25 | CLO1 | Description".')); return; }
        block.items.push({ id: parts[0], points: Number(parts[1]), clos: splitList(parts[2]), text: parts[3], line: start + k + 1 });
      });
      return block;
    }

    // Everything else: header keys, then (for drills) intro Markdown and steps.
    let k = 0;
    // header: contiguous "key: value" lines at the top
    while (k < inner.length) {
      const l = inner[k];
      if (!l.trim()) { k++; if (Object.keys(block.header).length) break; continue; }
      const m = /^([A-Za-z][\w-]*)\s*:\s*(.*)$/.exec(l);
      if (!m || /^-\s/.test(l)) break;
      block.header[m[1]] = m[2].trim();
      block.headerLines[m[1]] = start + k + 1;
      k++;
    }

    if (name === "quiz") {
      for (; k < inner.length; k++) {
        const l = inner[k].trim();
        const lineNo = start + k + 1;
        if (!l) continue;
        const opt = /^-\s*\[( |x|X)\]\s*(.+)$/.exec(l);
        const kv = /^([A-Za-z]+)\s*:\s*(.*)$/.exec(l);
        if (opt) block.items.push({ correct: opt[1].toLowerCase() === "x", text: opt[2], line: lineNo });
        else if (kv) { block.header[kv[1]] = kv[2].trim(); block.headerLines[kv[1]] = lineNo; }
        else issues.push(issue("error", lineNo, 'Quiz lines are "key: value" or options like "- [x] Right answer".'));
      }
      return block;
    }

    if (name === "video") return block;

    // drill / skillcheck / preflight: intro Markdown, then "- step:" or "- task:" items
    const stepWord = name === "skillcheck" ? "task" : "step";
    let current = null;
    for (; k < inner.length; k++) {
      const raw = inner[k];
      const lineNo = start + k + 1;
      const stepStart = new RegExp("^-\\s*(step|task)\\s*:\\s*(.*)$").exec(raw.trim());
      if (stepStart && !/^\s/.test(raw)) {
        if (stepStart[1] !== stepWord) issues.push(issue("error", lineNo, `Use "- ${stepWord}:" inside a ${name} block.`));
        current = { text: stepStart[2], checks: [], hints: [], line: lineNo, extra: {} };
        block.steps.push(current);
        continue;
      }
      if (current && /^\s+\S/.test(raw)) {
        const m = /^\s+([A-Za-z]+)\s*:\s*(.*)$/.exec(raw);
        if (!m) { issues.push(issue("error", lineNo, 'Step details are indented "key: value" lines, such as "  check: text.absent value=\\"Acme\\"".')); continue; }
        const key = m[1], val = m[2].trim();
        if (key === "check") current.checks.push(Object.assign(parseCheck(val, lineNo, issues), { line: lineNo }));
        else if (key === "hint") current.hints.push(val);
        else if (["web", "desktop", "video", "points", "answer", "platform"].includes(key)) current.extra[key] = val;
        else issues.push(issue("error", lineNo, `Unknown step key "${key}". Use check, hint, web, desktop, video, points, answer or platform.`));
        continue;
      }
      if (current && raw.trim()) {
        issues.push(issue("error", lineNo, "Text after a step must be indented under it, or start a new \"- " + stepWord + ":\"."));
        continue;
      }
      if (!current) block.intro.push(raw);
    }
    return block;
  }

  /**
   * Parse the part after "check:".
   * Example: text.present value="Lakeshore Paddle Company" min=4 in=body
   */
  function parseCheck(src, lineNo, issues) {
    const out = { name: "", params: {}, raw: src };
    const nameMatch = /^([a-z]+\.[A-Za-z]+)(.*)$/.exec(src);
    if (!nameMatch) { issues.push(issue("error", lineNo, `"${src}" is not a check. It should start with a name like text.absent.`)); return out; }
    out.name = nameMatch[1];
    let rest = nameMatch[2];
    // Tokens: key=value or key="value with spaces" (\" for a quote inside)
    const re = /\s*([A-Za-z]+)=("((?:[^"\\]|\\.)*)"|[^\s"]+)/y;
    let pos = 0;
    while (pos < rest.length) {
      if (!rest.slice(pos).trim()) break;
      re.lastIndex = pos;
      const m = re.exec(rest);
      if (!m) { issues.push(issue("error", lineNo, `Could not read "${rest.slice(pos).trim()}". Parameters look like key=value or key="two words".`)); break; }
      out.params[m[1]] = m[3] !== undefined ? m[3].replace(/\\"/g, '"') : m[2];
      pos = re.lastIndex;
    }
    return out;
  }

  // -------------------------------------------------------------------------
  // 2. VALIDATE
  // -------------------------------------------------------------------------

  /**
   * Validate a parsed document.
   * ctx (all optional):
   *   videoHosts: approved host list
   *   starterExists(path) -> true | false | null (null = can't tell)
   */
  function validate(doc, ctx) {
    ctx = ctx || {};
    const issues = doc.issues.slice();
    const add = (level, line, message) => issues.push(issue(level, line, message));
    const meta = doc.meta;
    const hosts = ctx.videoHosts || DEFAULT_VIDEO_HOSTS;
    const ready = meta.status === "ready";

    // --- front matter ---
    const type = meta.type;
    if (!FILE_TYPES.includes(type)) add("error", 1, `type must be one of ${FILE_TYPES.join(", ")}.`);
    if (type !== "config") {
      ["id", "week", "order", "title", "objectives"].forEach(k => { if (meta[k] === undefined || meta[k] === "") add("error", 1, `Front matter is missing "${k}".`); });
      splitList(meta.objectives).forEach(o => { if (!/^W[1-5]\.[1-9]$/.test(o)) add("warning", 1, `Objective "${o}" is not in the course map (expected W1.1 to W5.6).`); });
      splitList(meta.clos).forEach(c => { if (!/^CLO[1-8]$/.test(c)) add("warning", 1, `"${c}" is not a course learning outcome (CLO1 to CLO8).`); });
      if (meta.status && !["draft", "ready"].includes(meta.status)) add("error", 1, 'status is "draft" or "ready".');
    }

    const ids = new Set();
    let drillCount = 0;
    for (const b of doc.blocks) {
      if (b.kind === "md") { validateMarkdown(b, add); continue; }
      if (!DIRECTIVES.includes(b.name)) continue;
      const h = b.header;
      const need = (key, why) => { if (!h[key]) add("error", b.line, `::: ${b.name} needs "${key}:"${why ? " (" + why + ")" : ""}.`); };

      if (["quiz", "drill", "skillcheck", "preflight", "rubric"].includes(b.name)) {
        if (!b.arg) add("error", b.line, `::: ${b.name} needs an id, for example "::: ${b.name} ${b.name === "quiz" ? "Q1.1" : "D1.01"}".`);
        else if (ids.has(b.arg)) add("error", b.line, `The id "${b.arg}" is used twice in this file.`);
        else ids.add(b.arg);
      }

      switch (b.name) {
        case "video": {
          need("url"); need("title", "it names the player for screen readers"); need("captions");
          const url = h.url || "";
          if (url === "TODO") add(ready ? "error" : "warning", b.headerLines.url || b.line, `Video "${h.title || ""}" has no URL yet.`);
          else if (url) {
            const host = hostOf(url);
            if (!host) add("error", b.headerLines.url, "Video url must be a full https:// address.");
            else if (!hosts.includes(host) && !(host === "youtu.be" && hosts.includes("www.youtube.com"))) add("error", b.headerLines.url, `Video host "${host}" is not on the approved list (${hosts.join(", ")}).`);
          }
          if (h.captions && !["verified", "auto", "none"].includes(h.captions)) add("error", b.headerLines.captions, 'captions is "verified", "auto" or "none".');
          if (h.captions && h.captions !== "verified") add(ready ? "error" : "warning", b.headerLines.captions || b.line, `Video "${h.title || ""}" needs human-verified captions before this file is ready (ADA Title II, WCAG 1.2.2).`);
          break;
        }
        case "callout":
          if (!CALLOUT_KINDS.includes(b.arg)) add("error", b.line, `Callout kind must be one of: ${CALLOUT_KINDS.join(", ")}.`);
          validateMarkdown({ lines: b.md, line: b.line + 1 }, add);
          break;
        case "platform":
          if (!PLATFORMS.includes(b.arg)) add("error", b.line, `Platform must be one of: ${PLATFORMS.join(", ")}.`);
          validateMarkdown({ lines: b.md, line: b.line + 1 }, add);
          break;
        case "quiz": {
          need("question");
          const qtype = h.type || "single";
          if (!["single", "multi", "number"].includes(qtype)) add("error", b.headerLines.type, 'Quiz type is "single", "multi" or "number".');
          if (qtype === "number") {
            if (h.answer === undefined || isNaN(Number(h.answer))) add("error", b.line, 'A number quiz needs "answer:" with a number.');
          } else {
            const right = b.items.filter(o => o.correct).length;
            if (b.items.length < 2) add("error", b.line, "A quiz needs at least two options.");
            if (qtype === "single" && right !== 1) add("error", b.line, "A single-answer quiz needs exactly one [x] option.");
            if (qtype === "multi" && right < 1) add("error", b.line, "A multi-answer quiz needs at least one [x] option.");
          }
          if (!h.explain) add("warning", b.line, 'Add "explain:" so learners learn from a wrong answer.');
          break;
        }
        case "drill": case "skillcheck": case "preflight": {
          if (b.name === "drill") drillCount++;
          need("title");
          if (b.name !== "preflight") need("starter", "checks are written against it");
          need("platform");
          if (h.platform && !["web", "desktop"].includes(h.platform)) add("error", b.headerLines.platform, 'platform is "web" (works in both) or "desktop" (needs desktop Word).');
          if (h.starter && h.starter !== "none") {
            if (!/^starters\/[\w\-./]+\.docx$/.test(h.starter) || h.starter.includes("..")) add("error", b.headerLines.starter, 'starter should be a path like "starters/week-01/W1-L2-letter.docx".');
            else if (ctx.starterExists) {
              const found = ctx.starterExists(h.starter);
              if (found === false) add(ready ? "error" : "warning", b.headerLines.starter, `Starter file not found: ${h.starter}`);
            }
          }
          if (b.name === "skillcheck") {
            ["timeLimit", "attempts", "passing"].forEach(key => { if (h[key] && isNaN(Number(h[key]))) add("error", b.headerLines[key], `${key} must be a number.`); });
            need("timeLimit");
          }
          if (!b.steps.length) add("error", b.line, `::: ${b.name} ${b.arg} has no steps.`);
          let checked = 0;
          b.steps.forEach(s => {
            if (!s.text) add("error", s.line, "This step has no text.");
            if (s.checks.length || s.extra.answer !== undefined) checked++;
            if (b.name === "skillcheck") {
              if (s.hints.length) add("error", s.line, "Skill check tasks can't have hints.");
              if (!s.checks.length && s.extra.answer === undefined) add("error", s.line, "Every skill check task needs a check or an answer.");
            }
            if (s.extra.platform && !["web", "desktop"].includes(s.extra.platform)) add("error", s.line, 'A step platform is "web" or "desktop".');
            if (s.extra.points && isNaN(Number(s.extra.points))) add("error", s.line, "points must be a number.");
            if (s.extra.video) checkUrl(s.extra.video, s.line, add, hosts, true);
            s.checks.forEach(c => validateCheck(c, b.name, add));
            [s.text].concat(s.hints, s.extra.web || [], s.extra.desktop || []).forEach(t => checkInline(t, s.line, add));
          });
          if (b.name === "drill" && !checked) add("warning", b.line, `Drill ${b.arg} has no checked steps, so the trainer can't give feedback.`);
          validateMarkdown({ lines: b.intro, line: b.line + 1 }, add);
          break;
        }
        case "rubric": {
          const total = b.items.reduce((sum, it) => sum + (isNaN(it.points) ? 0 : it.points), 0);
          b.items.forEach(it => { if (isNaN(it.points)) add("error", it.line, "Rubric points must be a number."); });
          if (total !== 100) add("error", b.line, `Rubric points total ${total}; they must total 100.`);
          break;
        }
      }
    }
    if (type === "lesson" && !drillCount) add("warning", 1, "This lesson has no drills.");
    return issues.sort((a, b) => a.line - b.line);
  }

  /** Check one parsed "check:" against the catalog. */
  function validateCheck(c, where, add) {
    const def = CHECKS[c.name];
    if (!c.name) return;
    if (!def) { add("error", c.line, `Unknown check "${c.name}". Known checks: ${Object.keys(CHECKS).join(", ")}.`); return; }
    if (def.notInSkillCheck && where === "skillcheck") add("error", c.line, `${c.name} can't be used in a skill check.`);
    const used = Object.keys(c.params);
    const sel = used.filter(k => SELECTORS.includes(k));
    if (sel.length && !def.selector) add("error", c.line, `${c.name} doesn't take a selector (${sel.join(", ")}).`);
    if (sel.length > 1) add("error", c.line, `Use only one selector; found ${sel.join(" and ")}.`);
    if (c.params.in && !["body", "header", "footer", "document"].includes(c.params.in)) add("error", c.line, 'in= is body, header, footer or document.');
    for (const [key, type] of Object.entries(def.params)) {
      if (type.endsWith("!") && c.params[key] === undefined) add("error", c.line, `${c.name} needs ${key}=...`);
    }
    if (def.needsOneOf && !def.needsOneOf.some(k => c.params[k] !== undefined)) add("error", c.line, `${c.name} needs at least one of: ${def.needsOneOf.join(", ")}.`);
    for (const key of used) {
      if (SELECTORS.includes(key)) continue;
      const type = def.params[key];
      if (!type) { add("error", c.line, `${c.name} has no parameter "${key}". It takes: ${Object.keys(def.params).concat(def.selector ? SELECTORS : []).join(", ") || "no parameters"}.`); continue; }
      const v = c.params[key];
      const base = type.replace("!", "");
      const bad = (msg) => add("error", c.line, `${c.name} ${key}=${v}: ${msg}`);
      if (base === "int" && !/^\d+$/.test(v)) bad("expected a whole number.");
      if (base === "number" && isNaN(Number(v))) bad("expected a number.");
      if (base === "bool" && !["true", "false"].includes(v)) bad("expected true or false.");
      if (base === "length" && toPoints(v) === null) bad("expected a length with units, like 1in, 36pt or 2.54cm.");
      if (base === "color" && !/^#[0-9A-Fa-f]{6}$/.test(v)) bad("expected a color like #0046AD.");
      if (base === "lang" && !/^[a-z]{2}-[A-Z]{2}$/.test(v)) bad("expected a language tag like en-US.");
      if (base === "lineSpacing" && !(["single", "1.15", "1.5", "double"].includes(v) || toPoints(v) !== null)) bad("expected single, 1.15, 1.5, double or an exact value like 14pt.");
      if (base.startsWith("enum:") && !base.slice(5).split("|").includes(v)) bad(`expected one of ${base.slice(5).split("|").join(", ")}.`);
    }
  }

  /** Markdown-level checks: H1, raw HTML, images, links, inline syntax. */
  function validateMarkdown(b, add) {
    let inCode = false;
    (b.lines || []).forEach((l, k) => {
      const lineNo = b.line + k;
      if (/^```/.test(l.trim())) { inCode = !inCode; return; }
      if (inCode) return;
      if (/^#\s/.test(l)) add("error", lineNo, "Don't use a # heading; the title is the page's only H1. Start at ##.");
      if (/^#{4,}\s/.test(l)) add("warning", lineNo, "Headings deeper than ### are hard to scan in the narrow add-in panel.");
      checkInline(l, lineNo, add);
    });
  }

  function checkInline(text, lineNo, add) {
    const t = String(text || "");
    const noCode = t.replace(/`[^`]*`/g, "");
    if (/<\/?[A-Za-z][^>]*>/.test(noCode)) add("error", lineNo, "Raw HTML isn't allowed. It will be shown as text.");
    let m;
    const img = /!\[([^\]]*)\]\(([^)\s]+)\)(\{decorative\})?/g;
    while ((m = img.exec(t))) {
      if (!m[1].trim() && !m[3]) add("error", lineNo, "This image needs alt text, or {decorative} if it carries no information.");
      if (/^https?:/i.test(m[2])) add("warning", lineNo, "External images load from another site; prefer a file in images/.");
      else if (m[2].includes("..") || /^[a-z]+:/i.test(m[2])) add("error", lineNo, "Image paths must be relative, like images/indent.png.");
    }
    const link = /(?<!!)\[([^\]]+)\]\(([^)\s]+)\)/g;
    while ((m = link.exec(t))) {
      const href = m[2];
      if (!/^https:\/\//i.test(href) && !/^[\w\-./]+\.md(#[\w-]+)?$/.test(href)) add("error", lineNo, `Link "${href}" must start with https:// or point to a lesson .md file.`);
    }
    const ext = /\[\[([a-z]+)\s*:([^\]]*)\]\]/g;
    while ((m = ext.exec(t))) {
      if (!["key", "path"].includes(m[1])) add("error", lineNo, `Unknown inline tag [[${m[1]}: ...]]. Use [[key: ...]] or [[path: ...]].`);
    }
  }

  function hostOf(url) { const m = /^https:\/\/([^/:?#]+)/i.exec(url); return m ? m[1].toLowerCase() : null; }
  function checkUrl(url, lineNo, add, hosts) {
    const host = hostOf(url.split("#")[0]);
    if (!host) add("error", lineNo, "Video links must be full https:// addresses.");
    else if (!hosts.includes(host) && host !== "youtu.be") add("error", lineNo, `Video host "${host}" is not approved.`);
  }

  // -------------------------------------------------------------------------
  // 3. RENDER (safe HTML string for previews)
  // -------------------------------------------------------------------------

  /** Inline Markdown on ONE line of text. Escapes first, then adds our own tags. */
  function inline(text) {
    // Pull out code spans first so nothing inside them is formatted.
    const codes = [];
    let s = String(text).replace(/`([^`]+)`/g, (_, c) => { codes.push(c); return "\u0000" + (codes.length - 1) + "\u0000"; });
    s = esc(s);
    // [[key: Ctrl+S / Cmd+S]] -> <kbd> groups
    s = s.replace(/\[\[key:\s*([^\]]+?)\s*\]\]/g, (_, keys) => keys.split(" / ").map((combo, n, all) =>
      `<span class="tm-keys">${all.length > 1 ? `<span class="tm-keyos">${n === 0 ? "Win" : "Mac"}</span>` : ""}${combo.split("+").map(k => `<kbd>${k}</kbd>`).join("+")}</span>`).join(" "));
    // [[path: Home &gt; Paragraph]] -> breadcrumb
    s = s.replace(/\[\[path:\s*([^\]]+?)\s*\]\]/g, (_, p) => `<span class="tm-path">${p.split(" &gt; ").map(x => `<span>${x}</span>`).join('<span class="tm-sep" aria-hidden="true">›</span>')}</span>`);
    // images, then links
    s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)\)(\{decorative\})?/g, (_, alt, src, deco) =>
      (/^https?:|^[a-z]+:|\.\./i.test(src.replace(/&amp;/g, "&")) && !/^https:/i.test(src)) ? "" :
      `<img src="${src}" alt="${deco ? "" : alt}"${deco ? ' role="presentation"' : ""} loading="lazy">`);
    s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (all, label, href) =>
      /^https:\/\//i.test(href) ? `<a href="${href}" target="_blank" rel="noopener noreferrer">${label}<span class="tm-vh"> (opens in a new tab)</span></a>`
        : /^[\w\-./]+\.md(#[\w-]+)?$/.test(href) ? `<a href="${href}" class="tm-lesson-link">${label}</a>` : label);
    s = s.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>").replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>");
    // Put code spans back (escaped).
    s = s.replace(/\u0000(\d+)\u0000/g, (_, n) => `<code>${esc(codes[Number(n)])}</code>`);
    return s;
  }

  /** Block Markdown: headings, paragraphs, lists (nested by indent), tables. */
  function markdown(lines) {
    const out = [];
    let i = 0;
    const L = lines || [];
    while (i < L.length) {
      const line = L[i];
      if (!line.trim()) { i++; continue; }
      let m;
      if ((m = /^(#{2,3})\s+(.*)$/.exec(line))) { out.push(`<h${m[1].length}>${inline(m[2])}</h${m[1].length}>`); i++; continue; }
      if (/^#\s/.test(line)) { out.push(`<p>${inline(line.replace(/^#\s/, ""))}</p>`); i++; continue; } // demoted: only one H1
      if (/^```/.test(line.trim())) {
        const code = []; i++;
        while (i < L.length && !/^```/.test(L[i].trim())) { code.push(L[i]); i++; }
        i++; out.push(`<pre><code>${esc(code.join("\n"))}</code></pre>`); continue;
      }
      if (/^\|.*\|\s*$/.test(line.trim()) && L[i + 1] && /^\|[\s:|-]+\|\s*$/.test(L[i + 1].trim())) {
        const cells = r => r.trim().replace(/^\||\|$/g, "").split("|").map(c => c.trim());
        const head = cells(line); i += 2;
        const rows = [];
        while (i < L.length && /^\|.*\|\s*$/.test(L[i].trim())) { rows.push(cells(L[i])); i++; }
        out.push(`<div class="tm-table"><table><thead><tr>${head.map(h => `<th scope="col">${inline(h)}</th>`).join("")}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${inline(c)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`);
        continue;
      }
      if (/^\s*([-*]|\d+\.)\s+/.test(line)) {
        const items = [];
        while (i < L.length && (/^\s*([-*]|\d+\.)\s+/.test(L[i]) || (/^\s{2,}\S/.test(L[i]) && items.length))) {
          const lm = /^(\s*)([-*]|\d+\.)\s+(.*)$/.exec(L[i]);
          if (lm) items.push({ indent: lm[1].length, ordered: /\d/.test(lm[2]), text: lm[3] });
          else items[items.length - 1].text += " " + L[i].trim();
          i++;
        }
        out.push(listHtml(items, 0, items.length ? items[0].indent : 0).html);
        continue;
      }
      const para = [];
      while (i < L.length && L[i].trim() && !/^(#{1,3}\s|```|\s*([-*]|\d+\.)\s+|\|)/.test(L[i])) { para.push(L[i].trim()); i++; }
      if (para.length) out.push(`<p>${inline(para.join(" "))}</p>`);
      else { out.push(`<p>${inline(line)}</p>`); i++; }
    }
    return out.join("\n");
  }
  function listHtml(items, start, indent) {
    const ordered = items[start] && items[start].ordered;
    let html = ordered ? "<ol>" : "<ul>";
    let i = start;
    while (i < items.length && items[i].indent >= indent) {
      if (items[i].indent > indent) {
        const sub = listHtml(items, i, items[i].indent);
        html = html.replace(/<\/li>$/, sub.html + "</li>");
        i = sub.next; continue;
      }
      html += `<li>${inline(items[i].text)}</li>`;
      i++;
    }
    return { html: html + (ordered ? "</ol>" : "</ul>"), next: i };
  }

  /** Turn a YouTube or Vimeo address into a privacy-friendly embed address. */
  function embedUrl(url) {
    let m;
    if ((m = /^https:\/\/(?:www\.)?youtube\.com\/watch\?v=([\w-]{6,})/.exec(url)) || (m = /^https:\/\/youtu\.be\/([\w-]{6,})/.exec(url))
      || (m = /^https:\/\/(?:www\.)?youtube(?:-nocookie)?\.com\/embed\/([\w-]{6,})/.exec(url))) return "https://www.youtube-nocookie.com/embed/" + m[1];
    if ((m = /^https:\/\/(?:www\.)?vimeo\.com\/(\d+)/.exec(url))) return "https://player.vimeo.com/video/" + m[1];
    return url;
  }

  /** Describe a check in plain words for the author view. */
  function describeCheck(c) {
    const p = Object.entries(c.params).map(([k, v]) => `${k}=${/\s/.test(v) ? `"${v}"` : v}`).join(" ");
    return `${c.name}${p ? " " + p : ""}`;
  }

  /**
   * Render a parsed document to an HTML string.
   * opts: platform ("all" | "web" | "desktop" | "windows" | "mac"),
   *       showChecks (true for authors), videoHosts
   */
  function render(doc, opts) {
    opts = Object.assign({ platform: "all", showChecks: true, videoHosts: DEFAULT_VIDEO_HOSTS }, opts || {});
    const meta = doc.meta;
    const shows = (p) => opts.platform === "all" || p === opts.platform ||
      (p === "desktop" && (opts.platform === "windows" || opts.platform === "mac"));
    const parts = [];
    parts.push(`<header class="tm-head"><p class="tm-eyebrow">${esc(`Week ${meta.week || "?"} · ${labelFor(meta.type)}${meta.minutes ? " · about " + meta.minutes + " min" : ""}`)}</p><h1>${esc(meta.title || "Untitled")}</h1></header>`);

    for (const b of doc.blocks) {
      if (b.kind === "md") { parts.push(markdown(b.lines)); continue; }
      const h = b.header;
      switch (b.name) {
        case "callout": {
          const label = { tip: "Tip", note: "Note", warning: "Watch out", why: "Why it matters", ai: "AI moment (optional)", privacy: "Privacy" }[b.arg] || "Note";
          parts.push(`<aside class="tm-callout tm-${esc(b.arg)}"><p class="tm-callout-label">${label}</p>${markdown(b.md)}</aside>`);
          break;
        }
        case "platform":
          if (shows(b.arg)) parts.push(`<section class="tm-platform"><p class="tm-callout-label">${esc(platformName(b.arg))}</p>${markdown(b.md)}</section>`);
          break;
        case "video": {
          const url = h.url || "";
          const host = hostOf(embedUrl(url));
          const title = esc(h.title || "Video");
          const capNote = h.captions === "verified" ? "Captions checked" : h.captions === "auto" ? "Automatic captions (not yet checked)" : "No captions yet";
          if (url === "TODO" || !host || !opts.videoHosts.concat(["www.youtube-nocookie.com"]).includes(host)) {
            parts.push(`<figure class="tm-video tm-video-empty"><div class="tm-video-box"><span>${url === "TODO" ? "Video coming soon" : "Video not shown: host not approved"}</span></div><figcaption><strong>${title}</strong>${h.duration ? " · " + esc(h.duration) : ""} · ${capNote}</figcaption></figure>`);
          } else {
            parts.push(`<figure class="tm-video"><div class="tm-video-box"><iframe src="${esc(embedUrl(url))}" title="${title}" loading="lazy" sandbox="allow-scripts allow-same-origin allow-presentation allow-popups" referrerpolicy="strict-origin-when-cross-origin" allow="encrypted-media; picture-in-picture; fullscreen" allowfullscreen></iframe></div><figcaption><strong>${title}</strong>${h.duration ? " · " + esc(h.duration) : ""} · ${capNote}${h.transcript && /^https:/.test(h.transcript) ? ` · <a href="${esc(h.transcript)}" target="_blank" rel="noopener noreferrer">Transcript</a>` : ""}</figcaption></figure>`);
          }
          break;
        }
        case "quiz": {
          const multi = h.type === "multi", num = h.type === "number";
          const name = "q-" + esc(b.arg);
          const body = num
            ? `<label class="tm-answer">Your answer <input type="number" inputmode="numeric" data-answer="${esc(h.answer || "")}"></label>`
            : `<fieldset><legend class="tm-vh">${inline(h.question || "")}</legend>${b.items.map((o, n) =>
              `<label class="tm-opt"><input type="${multi ? "checkbox" : "radio"}" name="${name}" value="${n}" data-correct="${o.correct}"> <span>${inline(o.text)}</span></label>`).join("")}</fieldset>`;
          parts.push(`<section class="tm-quiz" data-quiz="${esc(b.arg)}" data-type="${esc(h.type || "single")}"><p class="tm-callout-label">Quick check${multi ? " · select all that apply" : ""}</p><p class="tm-q">${inline(h.question || "")}</p>${body}<button type="button" class="tm-btn tm-quiz-check">Check answer</button><p class="tm-feedback" role="status" aria-live="polite"></p>${h.explain ? `<p class="tm-explain" hidden>${inline(h.explain)}</p>` : ""}</section>`);
          break;
        }
        case "drill": case "skillcheck": case "preflight": {
          const label = { drill: "Practice drill", skillcheck: "Skill check", preflight: "Preflight check" }[b.name];
          const facts = [];
          if (h.minutes) facts.push(`about ${esc(h.minutes)} min`);
          if (h.timeLimit) facts.push(`${esc(h.timeLimit)} min limit`);
          if (h.attempts) facts.push(`${esc(h.attempts)} attempts`);
          if (h.passing) facts.push(`pass at ${esc(h.passing)}%`);
          const badge = h.platform === "desktop" ? '<span class="tm-pill tm-desktop">Desktop Word</span>' : '<span class="tm-pill tm-web">Web or desktop</span>';
          const steps = b.steps.filter(s => !s.extra.platform || shows(s.extra.platform)).map(s => {
            const checked = s.checks.length || s.extra.answer !== undefined;
            const variants = [];
            if (s.extra.web && shows("web")) variants.push(`<p class="tm-variant"><span class="tm-pill tm-web">Web</span> ${inline(s.extra.web)}</p>`);
            if (s.extra.desktop && shows("desktop")) variants.push(`<p class="tm-variant"><span class="tm-pill tm-desktop">Desktop</span> ${inline(s.extra.desktop)}</p>`);
            const hint = s.hints.length && b.name !== "skillcheck" ? `<details class="tm-hint"><summary>Hint</summary>${s.hints.map(x => `<p>${inline(x)}</p>`).join("")}</details>` : "";
            const answer = s.extra.answer !== undefined ? `<label class="tm-answer">Your answer <input type="text" inputmode="numeric"></label>` : "";
            const checks = opts.showChecks && s.checks.length ? `<ul class="tm-checks" aria-label="Checks (author view)">${s.checks.map(c => `<li><code>${esc(describeCheck(c))}</code></li>`).join("")}${s.extra.answer !== undefined ? `<li><code>answer = ${esc(s.extra.answer)}</code></li>` : ""}</ul>` : (opts.showChecks && s.extra.answer !== undefined ? `<ul class="tm-checks"><li><code>answer = ${esc(s.extra.answer)}</code></li></ul>` : "");
            const only = s.extra.platform ? ` <span class="tm-pill tm-${s.extra.platform === "desktop" ? "desktop" : "web"}">${s.extra.platform === "desktop" ? "Desktop only" : "Web only"}</span>` : "";
            return `<li class="tm-step${checked ? " tm-checked" : ""}"><div class="tm-step-text">${inline(s.text)}${only}</div>${variants.join("")}${answer}${hint}${checks}<p class="tm-step-status">${checked ? '<span class="tm-status-dot" aria-hidden="true"></span> Checked by the trainer' : "Instruction only"}</p></li>`;
          }).join("");
          parts.push(`<section class="tm-drill tm-${b.name}" aria-labelledby="h-${esc(b.arg)}"><div class="tm-drill-head"><p class="tm-callout-label">${label} ${esc(b.arg)}</p><h2 id="h-${esc(b.arg)}">${inline(h.title || "")}</h2><p class="tm-facts">${badge}${facts.length ? " · " + facts.join(" · ") : ""}</p></div>${markdown(b.intro)}<ol class="tm-steps">${steps}</ol><div class="tm-actions"><button type="button" class="tm-btn tm-primary">${b.name === "skillcheck" ? "Start" : "Check my work"}</button>${opts.showChecks && h.starter && h.starter !== "none" ? `<span class="tm-starter">Starter: <code>${esc(h.starter)}</code></span>` : ""}</div></section>`);
          break;
        }
        case "rubric": {
          const total = b.items.reduce((s, it) => s + (it.points || 0), 0);
          parts.push(`<div class="tm-table"><table class="tm-rubric"><caption>Rubric ${esc(b.arg)}</caption><thead><tr><th scope="col">Criterion</th><th scope="col">Points</th><th scope="col">Outcomes</th></tr></thead><tbody>${b.items.map(it => `<tr><td><strong>${esc(it.id)}</strong> ${inline(it.text)}</td><td>${esc(it.points)}</td><td>${esc(it.clos.join(", "))}</td></tr>`).join("")}<tr><td><strong>Total</strong></td><td><strong>${total}</strong></td><td></td></tr></tbody></table></div>`);
          break;
        }
      }
    }
    return parts.join("\n");
  }
  function labelFor(t) { return { lesson: "Lesson", skillcheck: "Skill check", project: "Project", config: "Course settings" }[t] || "Lesson"; }
  function platformName(p) { return { web: "In Word for the web", desktop: "In desktop Word", windows: "On Windows", mac: "On a Mac" }[p] || p; }

  /** Quick counts for a status bar. */
  function summarize(doc) {
    const s = { drills: 0, steps: 0, checked: 0, quizzes: 0, videos: 0, videosTodo: 0 };
    doc.blocks.forEach(b => {
      if (b.name === "drill" || b.name === "skillcheck" || b.name === "preflight") {
        if (b.name === "drill") s.drills++;
        s.steps += b.steps.length;
        s.checked += b.steps.filter(x => x.checks.length || x.extra.answer !== undefined).length;
      }
      if (b.name === "quiz") s.quizzes++;
      if (b.name === "video") { s.videos++; if ((b.header.url || "") === "TODO") s.videosTodo++; }
    });
    return s;
  }

  const api = { VERSION, parse, validate, render, summarize, inline, markdown, esc, toPoints, fontStack, embedUrl,
    FONT_STACKS, DEFAULT_FONTS, DEFAULT_VIDEO_HOSTS, CHECKS, DIRECTIVES, CALLOUT_KINDS };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.TrainerMD = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
