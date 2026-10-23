/* =====================================================================
   OFFICE TRAINER - shared helpers for the site's tool pages
   ---------------------------------------------------------------------
   build-site-pages.py copies this into badges.html, instructor/sections.html
   and instructor/grader.html, before each page's own code. It makes
   window.Site with:
     el, add, fill      make elements safely (text is always plain text)
     badgeArt           a badge picture from TrainerMD.badgeSVG
     announce           tell screen reader users something
     download           save a file the page made (nothing is uploaded)
     loadCourse         read one course's outline and lessons from this site
     APPS               the three courses: folder, name, colors
   It also wires the light/dark button in the page header.

   SECURITY RULES these pages follow:
     - Text from files (a learner's badges, a student's document, a key)
       is only ever added as text (textContent), never as HTML.
     - The only HTML strings used come from TrainerMD (which escapes
       everything first) or are fixed in the code.
     - Each page's Content Security Policy only lets it fetch files from
       this site, so nothing a page opens can be sent anywhere.
   ===================================================================== */
(function () {
  "use strict";
  const TM = window.TrainerMD;

  /** The three courses. folder is where each course's files are, from the top of the site. */
  const APPS = {
    word: { id: "word", name: "Word", fullName: "Microsoft Word", folder: "word/", exam: "MO-110", fileWord: "document", ext: /\.(docx|dotx)$/i },
    excel: { id: "excel", name: "Excel", fullName: "Microsoft Excel", folder: "excel/", exam: "MO-210", fileWord: "workbook", ext: /\.(xlsx|xltx|xlsm)$/i },
    powerpoint: { id: "powerpoint", name: "PowerPoint", fullName: "Microsoft PowerPoint", folder: "powerpoint/", exam: "MO-310", fileWord: "presentation", ext: /\.(pptx|potx)$/i },
  };

  /**
   * Make an element. Children can be text (added as plain text, always safe),
   * other elements, or lists of either. "on<event>" attributes take a function.
   *   el("button", { class: "btn", type: "button", onclick: save }, "Save")
   */
  function el(tag, attrs, ...kids) {
    const node = document.createElement(tag);
    for (const [key, value] of Object.entries(attrs || {})) {
      if (value === null || value === undefined || value === false) continue;
      if (/^on[a-z]+$/.test(key) && typeof value === "function") node.addEventListener(key.slice(2), value);
      else node.setAttribute(key, value === true ? "" : String(value));
    }
    return add(node, ...kids);
  }
  /** Add children, skipping null, undefined and false. */
  function add(node, ...kids) {
    for (const kid of kids.flat(Infinity)) {
      if (kid === null || kid === undefined || kid === false) continue;
      node.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    }
    return node;
  }
  /** Like add(), but first empties the node. */
  function fill(node, ...kids) { node.replaceChildren(); return add(node, ...kids); }

  /** A badge picture. TrainerMD.badgeSVG escapes every value it draws, so its SVG text is safe. */
  function badgeArt(options) {
    const holder = document.createElement("span");
    holder.innerHTML = TM.badgeSVG(options);
    return holder.firstChild;
  }

  /** Lesson text with its simple formatting (bold, menu paths, keys). TrainerMD.inline escapes everything first. */
  function richText(tag, md, attrs) {
    const node = el(tag, attrs || {});
    node.innerHTML = TM.inline(String(md || ""));
    return node;
  }

  /** Tell screen reader users something (uses the page's #announce region). */
  function announce(text) {
    const region = document.getElementById("announce");
    if (!region) return;
    region.textContent = "";
    setTimeout(function () { region.textContent = text; }, 50);
  }

  /** Save a file the page made. It goes straight to the person's Downloads folder; nothing is uploaded. */
  function download(content, fileName, type) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: type || "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = el("a", { href: url, download: fileName, class: "sr-only" });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  /** Today's date as "2026-10-18". */
  function today() { return new Date().toISOString().slice(0, 10); }
  /** "October 18, 2026" from "2026-10-18" (or the text itself if it isn't a date). */
  function longDate(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso))) return String(iso || "");
    return new Date(iso + "T12:00:00").toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  }

  /**
   * Read one course from this site, the same way the add-in does:
   * course.config.md, content/outline.md, then every lesson file the outline
   * names. root is the way back to the top of the site ("" or "../").
   * Returns { app, config, outline, docs, map } where map is TrainerMD.courseMap.
   */
  const courseCache = {};
  function loadCourse(appId, root) {
    if (courseCache[appId]) return courseCache[appId];
    const app = APPS[appId];
    const base = (root || "") + app.folder;
    const getText = async function (path) {
      const r = await fetch(base + path, { cache: "no-cache" });
      if (!r.ok) throw new Error("Couldn't load " + app.folder + path + " (" + r.status + ").");
      return r.text();
    };
    courseCache[appId] = (async function () {
      const config = TM.parse(await getText("course.config.md")).meta;
      const outline = TM.parse(await getText("content/outline.md"));
      const files = [];
      outline.blocks.filter((b) => b.name === "week").forEach((w) => w.steps.forEach((it) => {
        const f = it.extra.file;
        if (f && /^content\/[A-Za-z0-9_\-/]+\.md$/.test(f) && !f.includes("..")) files.push(f);
      }));
      const docs = {};
      const results = await Promise.allSettled(files.map((f) => getText(f)));
      results.forEach((r, i) => { if (r.status === "fulfilled") docs[files[i]] = TM.parse(r.value); });
      const map = TM.courseMap(outline, config, docs);
      map.weeks.forEach((w) => w.stations.forEach((s) => { s.week = w.n; }));
      return { app, config, outline, docs, map, stations: map.weeks.flatMap((w) => w.stations) };
    })();
    courseCache[appId].catch(function () { delete courseCache[appId]; }); // a failed load can be tried again
    return courseCache[appId];
  }

  // ---------- The light/dark button (the same choice as the home page) ----------
  (function themeButton() {
    const KEY = "officeTrainer.siteTheme";
    const root = document.documentElement;
    const btn = document.getElementById("theme-btn");
    if (!btn) return;
    const systemDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)");
    const isDark = function () {
      const chosen = root.getAttribute("data-theme");
      return chosen ? chosen === "dark" : !!(systemDark && systemDark.matches);
    };
    // The button is a switch: it always says "Dark mode", and aria-pressed (and the check mark) says whether it's on.
    const label = function () { btn.textContent = (isDark() ? "✓ " : "") + "Dark mode"; btn.setAttribute("aria-pressed", String(isDark())); };
    try { const saved = localStorage.getItem(KEY); if (saved === "light" || saved === "dark") root.setAttribute("data-theme", saved); } catch (e) { /* storage blocked */ }
    btn.hidden = false;
    label();
    btn.addEventListener("click", function () {
      const next = isDark() ? "light" : "dark";
      root.setAttribute("data-theme", next);
      try { localStorage.setItem(KEY, next); } catch (e) { /* not important */ }
      label();
    });
    if (systemDark && systemDark.addEventListener) systemDark.addEventListener("change", label);
  })();

  window.Site = { APPS, el, add, fill, badgeArt, richText, announce, download, today, longDate, loadCourse };
})();
