/* =============================================================================
   ppt-checker.js - the PowerPoint checker library (v0.1)
   -----------------------------------------------------------------------------
   This is the ONLY code that decides whether a PowerPoint step passed. Lesson
   files name a check ("slide.is", "anim.is" ...) and give it values; they can't
   add logic. It works exactly like the Word and Excel checkers:

     1. readPptx(bytes) opens a saved .pptx file and builds a "model" of it:
        the slides, their shapes and text, transitions, animations, notes,
        sections, masters, layouts, comments and slide show settings.
     2. runStep(model, step) runs a step's checks against that model and
        returns { pass, messages, where }.

   Why read the .pptx file instead of asking PowerPoint through Office.js?
     - PowerPoint for the web, desktop PowerPoint and "upload a saved file" all
       give us the same .pptx bytes, so every learner is checked the same way.
     - The file holds things Office.js can't show (animation timing, transition
       options, media trimming, Zoom links, slide show settings).
     - The same code can be tested in any browser with known-good and
       known-bad test decks (see powerpoint/tools/check-tests.html).

   Principle: check the OUTCOME, not the clicks. A slide is hidden whether the
   learner right-clicked it or used Slide Show > Hide Slide.

   A few PowerPoint facts this file relies on:
     - Sizes and positions are in EMUs: 914400 EMUs = 1 inch, 12700 = 1 point.
     - Each slide is its own XML part (ppt/slides/slide3.xml). The order of the
       slides is the list in ppt/presentation.xml, not the file names.
     - Shapes are listed back to front: the last shape on a slide is in front.
     - A slide's look comes from its layout, which comes from its slide master.
     - Animations are saved as a "timing tree" on each slide (<p:timing>).
     - Newer features (Morph, many transitions, Zoom, 3D models, trimming
       video) are saved inside <mc:AlternateContent> blocks with a fallback.
       We read the first choice (the newer one).

   Needs JSZip (loaded by the page) to open the .pptx zip file.
   Safety: DOMParser in "application/xml" mode never runs scripts and never
   fetches outside files. Text from the presentation is returned as plain
   strings; the page must show it with textContent, never innerHTML. Nothing
   leaves the learner's device.
   ============================================================================= */
(function (root) {
  "use strict";

  const VERSION = "0.1.0";

  // XML namespaces: the long names behind tag prefixes like <p:sld> or <a:t>.
  const NS = {
    p: "http://schemas.openxmlformats.org/presentationml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    dc: "http://purl.org/dc/elements/1.1/",
    cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  };

  // Safety limits (a .pptx is a zip, and zips can "explode"). Decks with video are big.
  const MAX_FILE_BYTES = 150 * 1024 * 1024;       // 150 MB as uploaded
  const MAX_UNZIPPED_BYTES = 400 * 1024 * 1024;   // 400 MB once unzipped
  const MAX_SLIDES = 500;

  const EMU_PER_INCH = 914400;

  // Slide sizes PowerPoint offers, in EMUs (width x height).
  const SLIDE_SIZES = {
    widescreen: [12192000, 6858000],   // 16:9, 13.333 x 7.5 in (the default today)
    standard: [9144000, 6858000],      // 4:3, 10 x 7.5 in
  };

  // What PowerPoint calls its views, by the name saved in viewProps.xml.
  const VIEW_NAMES = {
    sldView: "Normal", outlineView: "Outline", sldSorterView: "Slide Sorter", notesView: "Notes Page",
    sldMasterView: "Slide Master", notesMasterView: "Notes Master", handoutView: "Handout Master", sldThumbnailView: "Normal",
  };
  const VIEW_PARAM = { normal: ["sldView", "sldThumbnailView", null], outline: ["outlineView"], sorter: ["sldSorterView"], notes: ["notesView"],
    master: ["sldMasterView"], notesMaster: ["notesMasterView"], handoutMaster: ["handoutView"] };

  // Animation effects we name in messages, by preset class and id (from PowerPoint's preset list).
  const EFFECTS = {
    entr: { 1: "Appear", 2: "Fly In", 10: "Fade", 16: "Split", 21: "Wheel", 22: "Wipe", 23: "Zoom", 26: "Bounce", 53: "Grow & Turn", 55: "Float In" },
    emph: { 6: "Grow/Shrink", 8: "Spin", 26: "Pulse", 32: "Teeter" },
    exit: { 1: "Disappear", 2: "Fly Out", 10: "Fade", 16: "Split", 22: "Wipe", 23: "Zoom" },
    path: {},
  };
  const CLASS_WORDS = { entr: "an entrance", emph: "an emphasis", exit: "an exit", path: "a motion path", mediacall: "a media" };
  const CLASS_PARAM = { entrance: "entr", emphasis: "emph", exit: "exit", path: "path", entr: "entr", emph: "emph" };
  const START_WORDS = { onClick: "On Click", with: "With Previous", after: "After Previous" };

  // Transitions. Older, simple ones are <p:fade/>, <p:push/> ...; newer and 3D ones
  // live in the p14/p15/p159 namespaces. These names are what PowerPoint shows.
  const TRANSITION_NAMES = {
    fade: "Fade", push: "Push", wipe: "Wipe", split: "Split", cut: "Cut", randomBar: "Random Bars", circle: "Shape",
    diamond: "Shape", plus: "Shape", wedge: "Clock", wheel: "Clock", zoom: "Zoom", cover: "Cover", pull: "Uncover",
    dissolve: "Dissolve", checker: "Checkerboard", blinds: "Blinds", comb: "Comb", newsflash: "Newsflash", strips: "Strips", random: "Random",
    reveal: "Reveal", flash: "Flash", vortex: "Vortex", switch: "Switch", flip: "Flip", ripple: "Ripple", honeycomb: "Honeycomb",
    prism: "Cube", doors: "Doors", window: "Window", ferris: "Ferris Wheel", gallery: "Gallery", conveyor: "Conveyor", pan: "Pan",
    glitter: "Glitter", warp: "Zoom", flythrough: "Fly Through", shred: "Shred", wheelReverse: "Clock", morph: "Morph",
  };
  // Effect options PowerPoint writes for a transition's default choice (Push "From Bottom" is saved as dir="u").
  // Any other attributes mean the learner picked a different effect option. To verify against real PowerPoint files.
  const TRANSITION_DEFAULTS = { push: { dir: "u" }, split: { orient: "vert" }, randomBar: { dir: "vert" } };
  function optionsChanged(t) {
    const want = TRANSITION_DEFAULTS[t.type] || {};
    const keys = Array.from(new Set(Object.keys(want).concat(Object.keys(t.options))));
    return keys.some((k) => String(t.options[k] === undefined ? "" : t.options[k]) !== String(want[k] === undefined ? "" : want[k]));
  }
  // Subtle transitions saved in the newer namespaces (not "3D").
  const SUBTLE_NEW = ["reveal", "flash", "morph", "wheelReverse"];

  // =====================================================================
  // PART 1: SMALL HELPERS
  // =====================================================================

  /** Turn XML text into a document we can search. Throws if the XML is damaged. */
  function parseXml(text) {
    const xml = new DOMParser().parseFromString(text, "application/xml");
    if (xml.getElementsByTagName("parsererror").length > 0) throw new Error("Part of this PowerPoint file is damaged and could not be read.");
    return xml;
  }
  /** Direct children of `el` with this local tag name (any namespace prefix). */
  function kids(el, name) { return el ? Array.from(el.children).filter((c) => c.localName === name) : []; }
  /** The first direct child with this local tag name, or null. */
  function kid(el, name) { return kids(el, name)[0] || null; }
  /** Every element anywhere inside `el` with this local tag name. */
  function all(el, name) {
    if (!el) return [];
    const out = [];
    const walk = (node) => { for (const c of node.children) { if (c.localName === name) out.push(c); walk(c); } };
    walk(el);
    return out;
  }
  /** The first element anywhere inside `el` with this local name, or null. */
  function find(el, name) { return all(el, name)[0] || null; }
  /** Read an attribute with no prefix, or null. */
  function at(el, name) { return el && el.hasAttribute(name) ? el.getAttribute(name) : null; }
  /** Read a prefixed attribute such as r:id, or null. */
  function atNS(el, ns, name) { return el && el.hasAttributeNS(ns, name) ? el.getAttributeNS(ns, name) : null; }
  /** On/off attributes: "1" or "true" is on. */
  function isOn(value) { return value === "1" || value === "true"; }
  /** All the text inside an element. */
  function textOf(el) {
    // Only innermost <t> elements: SmartArt wraps an <a:t> inside a <dgm:t>, and we must not count it twice.
    return el ? all(el, "t").filter((t) => t.children.length === 0).map((t) => t.textContent).join("") : "";
  }

  /** Make text easy to compare: one kind of space and quote mark, no case. */
  function loose(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/[‘’‚′]/g, "'").replace(/[“”„″]/g, '"')
      .replace(/[   \u000b]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  }
  /** Short version of a long text for messages. */
  function snippet(text, max) {
    const t = String(text).replace(/\s+/g, " ").trim();
    const limit = max || 40;
    return t.length > limit ? t.slice(0, limit - 1).trimEnd() + "…" : t;
  }
  /** "1 slide" / "3 slides" */
  function count(n, word, plural) { return n + " " + (n === 1 ? word : (plural || word + "s")); }
  /** A comma-separated list in a lesson file as an array. */
  function listOf(value) { return String(value || "").split(",").map((x) => x.trim()).filter(Boolean); }
  function has(params, key) { return Object.prototype.hasOwnProperty.call(params, key) && params[key] !== undefined && params[key] !== ""; }
  function truthy(v) { return v === true || /^(true|yes|on|1)$/i.test(String(v)); }
  function pass(message) { return { pass: true, message: message }; }
  function fail(message, where) { return { pass: false, message: message, where: where || null }; }

  /** Wildcard pattern ("*_P1_Welcome.pptx") to a regular expression. Not case sensitive. */
  function wildcard(pattern) {
    const escaped = String(pattern).replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
    return new RegExp("^" + escaped + "$", "i");
  }
  /** "2.5in", "180pt", "3" (inches) -> EMUs. */
  function toEmu(value) {
    const m = /^\s*(-?[\d.]+)\s*(in|pt|cm|")?\s*$/i.exec(String(value));
    if (!m) return null;
    const n = Number(m[1]);
    const unit = (m[2] || "in").toLowerCase();
    return Math.round(unit === "pt" ? n * 12700 : unit === "cm" ? n * 360000 : n * EMU_PER_INCH);
  }
  /** EMUs -> "2.5 in" for messages. */
  function inches(emu) { return (Math.round(emu / EMU_PER_INCH * 100) / 100) + " in"; }
  /** "5", "5s", "0.5s", "500ms" -> milliseconds. */
  function toMs(value) {
    const m = /^\s*([\d.]+)\s*(ms|s)?\s*$/i.exec(String(value));
    if (!m) return null;
    return Math.round((m[2] || "s").toLowerCase() === "ms" ? Number(m[1]) : Number(m[1]) * 1000);
  }
  function seconds(ms) { return (Math.round(ms / 10) / 100) + " seconds"; }

  /** Inside <mc:AlternateContent>, use the first <mc:Choice> (the newer content). */
  function unwrapAlternate(el) {
    if (!el || el.localName !== "AlternateContent") return el;
    const choice = kid(el, "Choice") || kid(el, "Fallback");
    return choice ? choice : el;
  }
  /** The children of a shape tree, with AlternateContent blocks replaced by what's inside. */
  function treeChildren(el) {
    const out = [];
    for (const c of Array.from(el.children)) {
      if (c.localName === "AlternateContent") { const ch = unwrapAlternate(c); for (const x of Array.from(ch.children)) out.push(x); }
      else out.push(c);
    }
    return out;
  }

  /** Width and height of a PNG, JPEG or GIF picture, from its first bytes. Null if unknown. */
  function imageSize(bytes) {
    if (!bytes || bytes.length < 24) return null;
    const b = bytes;
    if (b[0] === 0x89 && b[1] === 0x50) { // PNG: IHDR holds width and height
      return { w: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, h: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
    }
    if (b[0] === 0x47 && b[1] === 0x49) return { w: b[6] | b[7] << 8, h: b[8] | b[9] << 8 }; // GIF
    if (b[0] === 0xff && b[1] === 0xd8) { // JPEG: walk the markers to a "start of frame"
      let i = 2;
      while (i + 9 < b.length) {
        if (b[i] !== 0xff) { i++; continue; }
        const marker = b[i + 1];
        const len = b[i + 2] << 8 | b[i + 3];
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return { w: b[i + 7] << 8 | b[i + 8], h: b[i + 5] << 8 | b[i + 6] };
        i += 2 + len;
      }
    }
    return null;
  }

  // =====================================================================
  // PART 2: OPENING A .PPTX FILE AND BUILDING THE MODEL
  // =====================================================================

  /**
   * Open a .pptx file (bytes from Office.js or an uploaded file) and build the model.
   * options.fileName - the file's name, for the file.name check.
   */
  async function readPptx(bytes, options) {
    const opts = options || {};
    if (typeof root.JSZip === "undefined") throw new Error("The checker couldn't load. Check your internet connection and reopen this panel.");
    const size = bytes.byteLength !== undefined ? bytes.byteLength : bytes.length;
    if (size > MAX_FILE_BYTES) throw new Error("That file is over 150 MB, which is too big to check. If it has long videos, check a copy without them.");
    const head = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    // A password to open a file encrypts it into a different kind of file (it starts D0 CF 11 E0).
    if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
      const err = new Error("This presentation has a password to open it, so the trainer can't read it (that's what the password is for). Check your other steps on a copy without the password.");
      err.encrypted = true;
      throw err;
    }
    let zip;
    try { zip = await root.JSZip.loadAsync(bytes); }
    catch (e) { throw new Error("That doesn't look like a PowerPoint (.pptx) file."); }
    let unzipped = 0;
    zip.forEach(function (path, entry) { if (entry._data && entry._data.uncompressedSize) unzipped += entry._data.uncompressedSize; });
    if (unzipped > MAX_UNZIPPED_BYTES) throw new Error("That file is too big to check.");

    /** Read one XML part, or null if the file doesn't have it. */
    async function part(path) { const f = path && zip.file(path); return f ? parseXml(await f.async("string")) : null; }
    /** The first bytes of a part (to measure a picture). */
    async function partBytes(path, max) {
      const f = path && zip.file(path);
      if (!f) return null;
      const all8 = await f.async("uint8array");
      return all8.subarray(0, max || 65536);
    }
    /** Resolve a relationship target against the part that points to it. */
    function resolve(fromPart, target) {
      if (/^\//.test(target)) return target.slice(1);
      const base = fromPart.split("/").slice(0, -1);
      for (const bit of target.split("/")) { if (bit === "..") base.pop(); else if (bit !== ".") base.push(bit); }
      return base.join("/");
    }
    /** A part's relationships: id -> { type, target, external }. */
    async function rels(partPath) {
      const dir = partPath.split("/").slice(0, -1).join("/");
      const name = partPath.split("/").pop();
      const xml = await part((dir ? dir + "/" : "") + "_rels/" + name + ".rels");
      const out = {};
      if (!xml) return out;
      for (const r of all(xml.documentElement, "Relationship")) {
        const external = at(r, "TargetMode") === "External";
        out[at(r, "Id")] = { type: (at(r, "Type") || "").split("/").pop(), target: external ? at(r, "Target") : resolve(partPath, at(r, "Target") || ""), external: external };
      }
      return out;
    }
    const io = { part, partBytes, rels, zip };

    const presXml = await part("ppt/presentation.xml");
    if (!presXml) throw new Error("That doesn't look like a PowerPoint (.pptx) file. If it's a .ppt or .key file, save it as a PowerPoint Presentation (.pptx) first.");
    const pres = presXml.documentElement;
    const presRels = await rels("ppt/presentation.xml");

    const model = {
      fileName: opts.fileName || "",
      props: { title: "", author: "", keywords: "", lastModifiedBy: "", subject: "" },
      custom: {},                 // custom file properties (Mark as Final is one)
      size: { cx: 0, cy: 0 },
      slides: [],                 // in show order
      sections: [],               // { name, slides: [slide numbers] }
      customShows: [],            // { name, slides: [slide numbers] }
      masters: [],                // slide masters, each with its layouts
      notesMaster: null, handoutMaster: null,
      theme: { name: "", colorsName: "", fontsName: "", majorFont: "", minorFont: "" },
      show: { loop: false, kiosk: false, browse: false, narration: true, animation: true, useTimings: true, customShow: null, range: null },
      view: { lastView: null, guides: false, guideCount: 0 },
      comments: [],               // { slide, text, author, replies: [...], resolved }
      pres: { embedFonts: false, embeddedFonts: [], removePersonalInfo: false, modifyPassword: false },
    };

    // --- slide size ---
    const sz = kid(pres, "sldSz");
    if (sz) model.size = { cx: Number(at(sz, "cx")) || 0, cy: Number(at(sz, "cy")) || 0 };

    // --- presentation-wide settings ---
    model.pres.embedFonts = isOn(at(pres, "embedTrueTypeFonts"));
    model.pres.embeddedFonts = kids(kid(pres, "embeddedFontLst"), "embeddedFont").map((e) => at(kid(e, "font"), "typeface") || "");
    model.pres.removePersonalInfo = isOn(at(pres, "removePersonalInfoOnSave"));
    model.pres.modifyPassword = !!kid(pres, "modifyVerifier");

    // --- properties ---
    const core = await part("docProps/core.xml");
    if (core) {
      const prop = (ns, name) => { const e = core.getElementsByTagNameNS(ns, name)[0]; return e ? e.textContent : ""; };
      model.props.title = prop(NS.dc, "title");
      model.props.subject = prop(NS.dc, "subject");
      model.props.author = prop(NS.dc, "creator");
      model.props.keywords = prop(NS.cp, "keywords");
      model.props.lastModifiedBy = prop(NS.cp, "lastModifiedBy");
    }
    const custom = await part("docProps/custom.xml");
    if (custom) for (const p of kids(custom.documentElement, "property")) {
      const v = p.firstElementChild;
      model.custom[at(p, "name") || ""] = v ? v.textContent : "";
    }

    // --- theme (the slide master's) ---
    const themeRel = Object.values(presRels).find((r) => r.type === "theme");
    const theme = await part(themeRel ? themeRel.target : "ppt/theme/theme1.xml");
    if (theme) readTheme(model.theme, theme.documentElement);

    // --- slide show settings and the last view ---
    const propsRel = Object.values(presRels).find((r) => r.type === "presProps");
    const presProps = await part(propsRel ? propsRel.target : "ppt/presProps.xml");
    if (presProps) {
      const showPr = kid(presProps.documentElement, "showPr");
      if (showPr) {
        model.show.loop = isOn(at(showPr, "loop"));
        model.show.narration = at(showPr, "showNarration") === null ? true : isOn(at(showPr, "showNarration"));
        model.show.animation = at(showPr, "showAnimation") === null ? true : isOn(at(showPr, "showAnimation"));
        model.show.useTimings = at(showPr, "useTimings") === null ? true : isOn(at(showPr, "useTimings"));
        model.show.kiosk = !!kid(showPr, "kiosk");
        model.show.browse = !!kid(showPr, "browse");
        if (kid(showPr, "kiosk")) model.show.loop = true; // a kiosk always loops
        const cs = kid(showPr, "custShow");
        if (cs) model.show.customShow = at(cs, "id");
        const rg = kid(showPr, "sldRg");
        if (rg) model.show.range = { st: Number(at(rg, "st")), end: Number(at(rg, "end")) };
      }
    }
    const viewRel = Object.values(presRels).find((r) => r.type === "viewProps");
    const viewProps = await part(viewRel ? viewRel.target : "ppt/viewProps.xml");
    if (viewProps) {
      const vp = viewProps.documentElement;
      model.view.lastView = at(vp, "lastView");
      const csv = find(vp, "cSldViewPr");
      model.view.guides = isOn(at(csv, "showGuides"));
      model.view.guideCount = all(vp, "guide").length;
    }

    // --- masters and layouts ---
    const layoutsByPath = {};
    for (const mId of kids(kid(pres, "sldMasterIdLst"), "sldMasterId")) {
      const rel = presRels[atNS(mId, NS.r, "id")];
      if (!rel) continue;
      const master = await readMaster(io, rel.target, "slide");
      model.masters.push(master);
      for (const l of master.layouts) layoutsByPath[l.path] = l;
    }
    const nmRel = Object.values(presRels).find((r) => r.type === "notesMaster");
    if (nmRel) model.notesMaster = await readMaster(io, nmRel.target, "notes");
    const hmRel = Object.values(presRels).find((r) => r.type === "handoutMaster");
    if (hmRel) model.handoutMaster = await readMaster(io, hmRel.target, "handout");

    // --- comment authors (both the classic and the modern kind) ---
    const authors = {};
    const caRel = Object.values(presRels).find((r) => r.type === "commentAuthors");
    const ca = await part(caRel ? caRel.target : null);
    if (ca) for (const x of kids(ca.documentElement, "cmAuthor")) authors["c" + at(x, "id")] = at(x, "name") || "";
    const auRel = Object.values(presRels).find((r) => r.type === "authors");
    const au = await part(auRel ? auRel.target : null);
    if (au) for (const x of kids(au.documentElement, "author")) authors[at(x, "id")] = at(x, "name") || "";

    // --- slides, in show order ---
    const idToNumber = {}; // slide id (from presentation.xml) -> slide number
    const ridToNumber = {};
    const sldIds = kids(kid(pres, "sldIdLst"), "sldId");
    if (sldIds.length > MAX_SLIDES) throw new Error("That presentation has too many slides to check.");
    for (let i = 0; i < sldIds.length; i++) {
      const rid = atNS(sldIds[i], NS.r, "id");
      const rel = presRels[rid];
      idToNumber[at(sldIds[i], "id")] = i + 1;
      ridToNumber[rid] = i + 1;
      const slide = await readSlide(io, rel ? rel.target : "", i + 1, layoutsByPath, authors, model);
      slide.id = at(sldIds[i], "id");
      model.slides.push(slide);
    }
    // Slide jumps (hyperlinks to other slides) point at slide parts; turn them into numbers.
    const pathToNumber = {};
    model.slides.forEach((s) => { pathToNumber[s.path] = s.number; });
    for (const s of model.slides) {
      for (const o of s.shapes) for (const link of o.links) if (link.slidePath) link.slide = pathToNumber[link.slidePath] || null;
    }
    for (const s of model.slides) for (const z of s.shapes.filter((o) => o.zoom)) z.zoom.targets = z.zoom.targetIds.map((id) => idToNumber[id]).filter(Boolean);

    // --- sections (saved in an extension of presentation.xml) ---
    for (const sec of all(pres, "section")) {
      if (sec.namespaceURI && !/powerpoint\/2010/.test(sec.namespaceURI)) continue;
      model.sections.push({ name: at(sec, "name") || "", slides: all(sec, "sldId").map((x) => idToNumber[at(x, "id")]).filter(Boolean) });
    }
    model.slides.forEach((s) => { const sec = model.sections.find((x) => x.slides.includes(s.number)); s.section = sec ? sec.name : null; });

    // --- custom shows ---
    for (const cs of kids(kid(pres, "custShowLst"), "custShow")) {
      model.customShows.push({ name: at(cs, "name") || "", id: at(cs, "id"), slides: kids(kid(cs, "sldLst"), "sld").map((x) => ridToNumber[atNS(x, NS.r, "id")]).filter(Boolean) });
    }
    return model;
  }

  /** The theme's name, color set and fonts. */
  function readTheme(out, el) {
    out.name = at(el, "name") || "";
    const cs = find(el, "clrScheme"); out.colorsName = cs ? at(cs, "name") || "" : "";
    const fs = find(el, "fontScheme"); out.fontsName = fs ? at(fs, "name") || "" : "";
    const major = find(fs, "majorFont"), minor = find(fs, "minorFont");
    out.majorFont = major && kid(major, "latin") ? at(kid(major, "latin"), "typeface") || "" : "";
    out.minorFont = minor && kid(minor, "latin") ? at(kid(minor, "latin"), "typeface") || "" : "";
    // The theme's 12 colors as hex (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink), for contrast checks.
    out.colors = {};
    if (cs) for (const c of Array.from(cs.children)) {
      const v = kid(c, "srgbClr") ? at(kid(c, "srgbClr"), "val") : kid(c, "sysClr") ? at(kid(c, "sysClr"), "lastClr") : null;
      if (v) out.colors[c.localName] = v.toUpperCase();
    }
  }

  /** A background: none (follows the master), solid, gradient, picture or pattern. */
  function readBackground(cSld) {
    const bg = kid(cSld, "bg");
    if (!bg) return { kind: "none" };
    const pr = kid(bg, "bgPr");
    if (pr) {
      if (kid(pr, "blipFill")) return { kind: "picture" };
      if (kid(pr, "gradFill")) return { kind: "gradient" };
      if (kid(pr, "pattFill")) return { kind: "pattern" };
      if (kid(pr, "solidFill")) {
        const f = kid(pr, "solidFill");
        return { kind: "solid", color: kid(f, "srgbClr") ? at(kid(f, "srgbClr"), "val") : kid(f, "schemeClr") ? "scheme:" + at(kid(f, "schemeClr"), "val") : null };
      }
      if (kid(pr, "noFill")) return { kind: "none" };
    }
    if (kid(bg, "bgRef")) return { kind: "style", idx: at(kid(bg, "bgRef"), "idx") };
    return { kind: "none" };
  }

  /** A slide master (or notes / handout master) with its layouts. */
  async function readMaster(io, path, kind) {
    const xml = await io.part(path);
    const master = { kind: kind, path: path, background: { kind: "none" }, shapes: [], layouts: [], titleFont: "", bodyFont: "", hf: null, theme: null };
    if (!xml) return master;
    const el = xml.documentElement;
    const masterRels = await io.rels(path);
    const cSld = kid(el, "cSld");
    master.background = readBackground(cSld);
    master.shapes = await readShapes(io, find(cSld, "spTree"), masterRels, path);
    const ts = kid(el, "txStyles");
    const font = (style) => { const lv = find(style, "lvl1pPr"); const latin = lv ? find(lv, "latin") : null; return latin ? at(latin, "typeface") || "" : ""; };
    if (ts) { master.titleFont = font(kid(ts, "titleStyle")); master.bodyFont = font(kid(ts, "bodyStyle")); }
    // "+mj-lt" and "+mn-lt" mean "the theme's heading / body font", not a font of their own.
    if (/^\+m[jn]-/.test(master.titleFont)) master.titleFont = "";
    if (/^\+m[jn]-/.test(master.bodyFont)) master.bodyFont = "";
    const hf = kid(el, "hf");
    master.hf = hf ? { hdr: at(hf, "hdr") !== "0", ftr: at(hf, "ftr") !== "0", dt: at(hf, "dt") !== "0", sldNum: at(hf, "sldNum") !== "0" } : null;
    const themeRel = Object.values(masterRels).find((r) => r.type === "theme");
    if (themeRel) { const t = await io.part(themeRel.target); if (t) { master.theme = {}; readTheme(master.theme, t.documentElement); } }
    if (kind === "slide") {
      for (const lid of kids(kid(el, "sldLayoutIdLst"), "sldLayoutId")) {
        const rel = masterRels[atNS(lid, NS.r, "id")];
        if (!rel) continue;
        const lx = await io.part(rel.target);
        if (!lx) continue;
        const le = lx.documentElement;
        const lc = kid(le, "cSld");
        const shapes = await readShapes(io, find(lc, "spTree"), await io.rels(rel.target), rel.target);
        master.layouts.push({
          path: rel.target, name: at(lc, "name") || "", type: at(le, "type") || "cust",
          custom: isOn(at(le, "userDrawn")) || !at(le, "type") || at(le, "type") === "cust",
          shapes: shapes, placeholders: shapes.filter((s) => s.ph).map((s) => s.ph.type),
          background: readBackground(lc),
        });
      }
    }
    return master;
  }

  /** One slide: its layout, shapes, notes, background, transition, animations and comments. */
  async function readSlide(io, path, number, layoutsByPath, authors, model) {
    const slide = { number: number, path: path, hidden: false, layout: "", layoutPath: "", title: "", shapes: [], notes: "",
      background: { kind: "none" }, transition: null, anims: [], builds: {}, mediaNodes: {}, section: null };
    const xml = await io.part(path);
    if (!xml) return slide;
    const el = xml.documentElement;
    slide.hidden = at(el, "show") === "0";
    const slideRels = await io.rels(path);
    const layoutRel = Object.values(slideRels).find((r) => r.type === "slideLayout");
    if (layoutRel && layoutsByPath[layoutRel.target]) { slide.layout = layoutsByPath[layoutRel.target].name; slide.layoutPath = layoutRel.target; }
    const cSld = kid(el, "cSld");
    slide.background = readBackground(cSld);
    slide.shapes = await readShapes(io, find(cSld, "spTree"), slideRels, path);
    const titleShape = slide.shapes.find((s) => s.ph && (s.ph.type === "title" || s.ph.type === "ctrTitle"));
    slide.title = titleShape ? titleShape.text.trim() : "";
    // --- notes ---
    const notesRel = Object.values(slideRels).find((r) => r.type === "notesSlide");
    if (notesRel) {
      const nx = await io.part(notesRel.target);
      if (nx) {
        const body = all(nx.documentElement, "sp").find((sp) => { const ph = find(sp, "ph"); return ph && at(ph, "type") === "body"; });
        slide.notes = body ? paragraphsText(find(body, "txBody")) : "";
      }
    }
    // --- transition (may sit inside AlternateContent for newer kinds) ---
    let trEl = null;
    for (const c of Array.from(el.children)) {
      if (c.localName === "transition") trEl = c;
      if (c.localName === "AlternateContent") { const t = find(unwrapAlternate(c), "transition"); if (t) trEl = t; }
    }
    if (trEl) slide.transition = readTransition(trEl, slideRels);
    // --- animations ---
    const timing = kid(el, "timing");
    if (timing) readTiming(slide, timing);
    // Media settings saved in the timing tree: Play Across Slides, Loop until Stopped.
    for (const o of slide.shapes) {
      const m = slide.mediaNodes[o.id];
      if (m) { o.acrossSlides = m.acrossSlides; o.loop = m.loop; o.hideWhenStopped = m.hide; }
    }
    // --- comments: classic (one part per slide) and modern ---
    for (const r of Object.values(slideRels)) {
      if (r.type !== "comments") continue;
      const cx = await io.part(r.target);
      if (!cx) continue;
      const root2 = cx.documentElement;
      if (root2.localName === "cmLst" && kids(root2, "cm").some((cm) => kid(cm, "text"))) {
        for (const cm of kids(root2, "cm")) model.comments.push({ slide: number, author: authors["c" + at(cm, "authorId")] || "", text: kid(cm, "text") ? kid(cm, "text").textContent : "", replies: [], resolved: false, modern: false });
      } else {
        for (const cm of kids(root2, "cm")) {
          const replies = all(cm, "reply").map((rp) => ({ author: authors[at(rp, "authorId")] || "", text: textOf(kid(rp, "txBody")) }));
          model.comments.push({ slide: number, author: authors[at(cm, "authorId")] || "", text: textOf(kid(cm, "txBody")), replies: replies,
            resolved: at(cm, "status") === "resolved" || at(cm, "status") === "closed", modern: true });
        }
      }
    }
    return slide;
  }

  /** The text of a text body, one line per paragraph. */
  function paragraphsText(txBody) {
    return txBody ? kids(txBody, "p").map((p) => textOf(p)).join("\n") : "";
  }

  /** A transition: its kind, whether it's 3D, effect options, duration, advance and sound. */
  function readTransition(el, slideRels) {
    const tr = { type: "", name: "", newer: false, options: {}, durationMs: null, advanceClick: true, advanceAfterMs: null, sound: false };
    const effect = Array.from(el.children).find((c) => !["sndAc", "extLst"].includes(c.localName));
    if (effect) {
      tr.type = effect.localName;
      tr.name = TRANSITION_NAMES[tr.type] || (effect.localName === "prstTrans" ? (at(effect, "prst") || "prstTrans") : tr.type);
      tr.newer = effect.namespaceURI !== NS.p;
      for (const a of Array.from(effect.attributes)) tr.options[a.localName] = a.value;
    }
    // Duration: newer files save it in milliseconds (p14:dur); older ones as slow / med / fast.
    for (const a of Array.from(el.attributes)) if (a.localName === "dur") tr.durationMs = Number(a.value);
    if (tr.durationMs === null && at(el, "spd")) tr.durationMs = { slow: 1000, med: 750, fast: 500 }[at(el, "spd")] || null;
    tr.advanceClick = at(el, "advClick") !== "0";
    if (at(el, "advTm") !== null) tr.advanceAfterMs = Number(at(el, "advTm"));
    tr.sound = !!find(el, "stSnd");
    tr.threeD = tr.newer && !SUBTLE_NEW.includes(tr.type);
    return tr;
  }

  /**
   * Animations. The timing tree looks like this (simplified):
   *   tmRoot > mainSeq > click group > time group > EFFECT (presetClass, nodeType)
   * Each EFFECT has a target shape (spTgt spid), maybe a paragraph range, a
   * start (clickEffect / withEffect / afterEffect), a delay and child behaviors
   * with durations. Media "play" commands are effects of class "mediacall".
   */
  function readTiming(slide, timing) {
    for (const b of all(kid(timing, "bldLst"), "bldP")) slide.builds[at(b, "spid")] = at(b, "build") || "whole";
    // Each video or audio clip has a media node: how many slides it plays across, and whether it loops.
    for (const m of all(timing, "audio").concat(all(timing, "video"))) {
      const node = kid(m, "cMediaNode");
      if (!node) continue;
      const tgt = find(node, "spTgt");
      const cTn = kid(node, "cTn");
      if (tgt) slide.mediaNodes[at(tgt, "spid")] = { acrossSlides: Number(at(node, "numSld") || 1) > 1, loop: at(cTn, "repeatCount") === "indefinite", hide: at(cTn, "display") === "0" };
    }
    const seq = all(timing, "cTn").find((c) => at(c, "nodeType") === "mainSeq");
    const effects = all(timing, "cTn").filter((c) => at(c, "presetClass"));
    let order = 0;
    for (const cTn of effects) {
      const tgt = find(cTn, "spTgt");
      const delayCond = kids(kid(cTn, "stCondLst"), "cond")[0];
      const delay = delayCond && /^\d+$/.test(at(delayCond, "delay") || "") ? Number(at(delayCond, "delay")) : 0;
      const durs = all(cTn, "cTn").map((c) => at(c, "dur")).filter((d) => d && /^\d+$/.test(d)).map(Number);
      const node = at(cTn, "nodeType") || "";
      const pRg = tgt ? find(tgt, "pRg") : null;
      const motion = find(cTn, "animMotion");
      slide.anims.push({
        order: ++order, inMainSeq: !!(seq && seq.contains(cTn)),
        cls: at(cTn, "presetClass"), presetId: Number(at(cTn, "presetID") || 0), subtype: at(cTn, "presetSubtype"),
        start: node === "withEffect" ? "with" : node === "afterEffect" ? "after" : node === "clickEffect" ? "onClick" : node,
        delayMs: delay, durationMs: durs.length ? Math.max.apply(null, durs) : null,
        target: tgt ? at(tgt, "spid") : null, paragraph: pRg ? Number(at(pRg, "st")) : null,
        path: motion ? at(motion, "path") || "" : null,
        pathDir: motion ? pathDirection(at(motion, "path") || "") : null,
        repeat: at(cTn, "repeatCount"),
      });
    }
  }

  /**
   * Which way a motion path goes overall: "right", "left", "up" or "down",
   * from its first and last points ("M 0 0 L 0.25 0 E" goes right).
   * Path numbers are fractions of the slide size; y grows downward.
   */
  function pathDirection(path) {
    const nums = String(path).replace(/[A-Za-z]/g, " ").trim().split(/[\s,]+/).filter(Boolean).map(Number).filter((n) => !isNaN(n));
    if (nums.length < 4) return null;
    const dx = nums[nums.length - 2] - nums[0], dy = nums[nums.length - 1] - nums[1];
    if (Math.abs(dx) < 0.001 && Math.abs(dy) < 0.001) return null;
    return Math.abs(dx) >= Math.abs(dy) ? (dx > 0 ? "right" : "left") : (dy > 0 ? "down" : "up");
  }

  /** Read every shape in a shape tree (recursing into groups). */
  async function readShapes(io, spTree, relMap, path) {
    const out = [];
    if (!spTree) return out;
    let order = 0;
    async function walk(container, groupId) {
      for (const el of treeChildren(container)) {
        const name = el.localName;
        if (["nvGrpSpPr", "grpSpPr", "extLst"].includes(name)) continue;
        if (name === "grpSp") {
          const nv = find(el, "cNvPr");
          const g = baseShape(el, nv, "group", ++order);
          g.groupId = groupId;
          out.push(g);
          await walk(el, g.id);
          continue;
        }
        if (!["sp", "pic", "graphicFrame", "cxnSp", "contentPart"].includes(name)) continue;
        const obj = await readShape(io, el, relMap, ++order);
        if (!obj) continue;
        obj.groupId = groupId;
        out.push(obj);
      }
    }
    await walk(spTree, null);
    return out;
  }

  /** The parts every shape has: id, name, alt text, position and size. */
  function baseShape(el, nv, kind, order) {
    // Shapes and pictures keep their position in spPr, groups in grpSpPr, tables and charts in p:xfrm.
    const holder = kid(el, "spPr") || kid(el, "grpSpPr");
    const xfrm = (holder && kid(holder, "xfrm")) || kid(el, "xfrm");
    const off = xfrm ? kid(xfrm, "off") : null, ext = xfrm ? kid(xfrm, "ext") : null;
    return {
      kind: kind, order: order, id: at(nv, "id"), name: at(nv, "name") || "",
      // A file name ("photo.png") left in the alt text box isn't a description, so it counts as no alt text.
      alt: /^[\w\- ().]+\.(png|jpe?g|gif|bmp|svg|tiff?|mp4|mov|wmv|m4a|mp3|wav|glb)$/i.test((at(nv, "descr") || "").trim()) ? "" : (at(nv, "descr") || "").trim(), decorative: !!all(nv, "decorative").find((d) => isOn(at(d, "val"))),
      x: off ? Number(at(off, "x")) : null, y: off ? Number(at(off, "y")) : null,
      w: ext ? Number(at(ext, "cx")) : null, h: ext ? Number(at(ext, "cy")) : null,
      rot: xfrm && at(xfrm, "rot") ? Number(at(xfrm, "rot")) / 60000 : 0,
      ph: null, text: "", paragraphs: [], cols: 1, links: [], geom: "", styleChanged: false, effects: false, outline: false, fill: false,
    };
  }

  /** One shape, picture, table, chart, SmartArt, media clip, 3D model, Zoom or ink drawing. */
  async function readShape(io, el, relMap, order) {
    const name = el.localName;
    const nv = find(el, "cNvPr");
    const obj = baseShape(el, nv, "shape", order);
    // A placeholder (title, body, slide number...) comes from the layout.
    const ph = find(kid(el, "nvSpPr") || kid(el, "nvPicPr") || kid(el, "nvGraphicFramePr") || el, "ph");
    if (ph) obj.ph = { type: at(ph, "type") || "body", idx: at(ph, "idx") };
    // Links on the shape itself (action buttons, linked pictures) and in its text.
    for (const h of all(el, "hlinkClick").concat(all(el, "hlinkHover"))) {
      const action = at(h, "action") || "";
      const rel = relMap[atNS(h, NS.r, "id")];
      const link = { url: rel && rel.external ? rel.target : "", slidePath: rel && !rel.external && rel.type === "slide" ? rel.target : null, slide: null,
        action: action, onShape: !!(nv && nv.contains(h)), text: "" };
      // The text a link is on: the run holding it.
      const run = h.parentNode && h.parentNode.parentNode && h.parentNode.parentNode.localName === "r" ? h.parentNode.parentNode : null;
      if (run) link.text = textOf(run);
      // PowerPoint marks every video with a "media" action; that isn't a link the learner made.
      if (/ppaction:\/\/media/.test(action)) continue;
      if (!link.url && !link.slidePath && !/hlinkshowjump|hlinksldjump|hlinkfile|program/.test(action)) continue;
      obj.links.push(link);
    }
    const spPr = kid(el, "spPr");
    if (spPr) {
      const geom = kid(spPr, "prstGeom");
      obj.geom = geom ? at(geom, "prst") || "" : (kid(spPr, "custGeom") ? "custom" : "");
      obj.effects = !!(kid(spPr, "effectLst") && kid(spPr, "effectLst").children.length) || !!kid(spPr, "scene3d") || !!kid(spPr, "sp3d");
      obj.outline = !!(kid(spPr, "ln") && !kid(kid(spPr, "ln"), "noFill"));
      obj.fill = !!(kid(spPr, "solidFill") || kid(spPr, "gradFill") || kid(spPr, "blipFill") || kid(spPr, "pattFill"));
    }
    // A shape style (Format > Shape Styles) changes the theme references in <p:style>.
    const style = kid(el, "style");
    if (style) {
      const idx = (n) => { const e = kid(style, n); return e ? at(e, "idx") : null; };
      obj.styleRefs = { ln: idx("lnRef"), fill: idx("fillRef"), effect: idx("effectRef") };
      obj.styleChanged = !(obj.styleRefs.ln === "2" && obj.styleRefs.fill === "1" && obj.styleRefs.effect === "0");
    }

    if (name === "sp" || name === "cxnSp") {
      const isText = find(el, "cNvSpPr") && isOn(at(find(el, "cNvSpPr"), "txBox"));
      obj.kind = ph ? "placeholder" : isText ? "textbox" : (/^actionButton/.test(obj.geom) ? "actionButton" : "shape");
      if (name === "cxnSp") obj.kind = "line";
      readText(obj, find(el, "txBody"));
    } else if (name === "pic") {
      obj.kind = all(el, "svgBlip").length ? "icon" : "picture";
      const blip = find(el, "blip");
      const rel = blip ? relMap[atNS(blip, NS.r, "embed")] : null;
      if (rel && !rel.external) {
        obj.imagePath = rel.target;
        obj.image = imageSize(await io.partBytes(rel.target, 262144));
      }
      const src = find(kid(el, "blipFill"), "srcRect");
      obj.crop = { l: 0, t: 0, r: 0, b: 0 };
      if (src) for (const k of ["l", "t", "r", "b"]) obj.crop[k] = Number(at(src, k) || 0) / 100000;
      // Video and audio are pictures with a media file attached.
      const vf = find(el, "videoFile"), af = find(el, "audioFile");
      if (vf || af) {
        obj.kind = vf ? "video" : "audio";
        const media = find(el, "media");
        const trim = media ? find(media, "trim") : null;
        const fade = media ? find(media, "fade") : null;
        obj.media = { trimStart: trim ? Number(at(trim, "st") || 0) : 0, trimEnd: trim ? Number(at(trim, "end") || 0) : 0,
          fadeIn: fade ? Number(at(fade, "in") || 0) : 0, fadeOut: fade ? Number(at(fade, "out") || 0) : 0, linked: !!(vf && atNS(vf, NS.r, "link") && relMap[atNS(vf, NS.r, "link")] && relMap[atNS(vf, NS.r, "link")].external) };
      }
    } else if (name === "graphicFrame") {
      const data = find(el, "graphicData");
      const uri = at(data, "uri") || "";
      if (/table$/.test(uri)) {
        obj.kind = "table";
        const tbl = find(data, "tbl");
        const tp = kid(tbl, "tblPr");
        const styleIdEl = tp ? kid(tp, "tableStyleId") : null;
        obj.table = { rows: kids(tbl, "tr").length, cols: kids(kid(tbl, "tblGrid"), "gridCol").length,
          styleId: styleIdEl ? styleIdEl.textContent.trim().toUpperCase() : "",
          firstRow: tp ? isOn(at(tp, "firstRow")) : false, bandRow: tp ? isOn(at(tp, "bandRow")) : false,
          cells: kids(tbl, "tr").map((tr) => kids(tr, "tc").map((tc) => textOf(tc))) };
        obj.text = obj.table.cells.map((r) => r.join(" ")).join("\n");
      } else if (/chart$/.test(uri)) {
        obj.kind = "chart";
        const chartEl = find(data, "chart");
        const rel = chartEl ? relMap[atNS(chartEl, NS.r, "id")] : null;
        obj.chart = rel ? await readChart(io, rel.target) : null;
      } else if (/diagram$/.test(uri)) {
        obj.kind = "smartart";
        const ids = find(data, "relIds");
        const dmRel = ids ? relMap[atNS(ids, NS.r, "dm")] : null;
        const loRel = ids ? relMap[atNS(ids, NS.r, "lo")] : null;
        const csRel = ids ? relMap[atNS(ids, NS.r, "cs")] : null;
        const dm = dmRel ? await io.part(dmRel.target) : null;
        const lo = loRel ? await io.part(loRel.target) : null;
        const cs = csRel ? await io.part(csRel.target) : null;
        const nodes = dm ? all(dm.documentElement, "pt").filter((p) => !at(p, "type") || at(p, "type") === "node") : [];
        obj.smartart = { layout: lo ? (at(lo.documentElement, "uniqueId") || "").split("/").pop() : "", colors: cs ? (at(cs.documentElement, "uniqueId") || "").split("/").pop() : "",
          items: nodes.filter((p) => textOf(p).trim()).length, text: nodes.map((p) => textOf(p)).join("\n") };
        obj.text = obj.smartart.text;
      } else if (/model3d$/.test(uri)) {
        obj.kind = "model3d";
        const rot = find(data, "rot");
        obj.model3d = { rotated: !!(rot && ["ax", "ay", "az"].some((k) => Number(at(rot, k) || 0) !== 0)) };
      } else if (/zoom$/i.test(uri)) {
        obj.kind = "zoom";
        const kind = /summaryzoom/i.test(uri) ? "summary" : /sectionzoom/i.test(uri) ? "section" : "slide";
        const zmPr = find(data, "zmPr");
        const targets = [];
        for (const t of all(data, "sldZmObj").concat(all(data, "sectionZmObj"), all(data, "summaryZmObj"))) {
          const sid = at(t, "sldId") || at(t, "sectionId");
          targets.push(sid);
        }
        // targetIds are slide ids (or section ids); they become slide numbers once every slide is read.
        obj.zoom = { kind: kind, returnToParent: zmPr ? isOn(at(zmPr, "returnToParent")) : false, count: Math.max(1, all(data, "zmPr").length), targetIds: targets, targets: [] };
      } else {
        obj.kind = "other";
      }
    } else if (name === "contentPart") {
      obj.kind = "ink";
    }
    if (!obj.text && obj.paragraphs.length) obj.text = obj.paragraphs.map((p) => p.text).join("\n");
    return obj;
  }

  /** The paragraphs of a text body: text, list level, bullets and each run's formatting. */
  function readText(obj, txBody) {
    if (!txBody) return;
    const bodyPr = kid(txBody, "bodyPr");
    obj.cols = bodyPr && at(bodyPr, "numCol") ? Number(at(bodyPr, "numCol")) : 1;
    for (const p of kids(txBody, "p")) {
      const pPr = kid(p, "pPr");
      const para = { text: textOf(p), lvl: pPr && at(pPr, "lvl") ? Number(at(pPr, "lvl")) : 0, bullet: null, align: pPr ? at(pPr, "algn") : null, runs: [] };
      if (pPr) {
        if (kid(pPr, "buAutoNum")) para.bullet = "number";
        else if (kid(pPr, "buChar") || kid(pPr, "buBlip")) para.bullet = "bullet";
        else if (kid(pPr, "buNone")) para.bullet = "none";
      }
      for (const r of kids(p, "r").concat(kids(p, "fld"))) {
        const rPr = kid(r, "rPr");
        const latin = rPr ? kid(rPr, "latin") : null;
        const fill = rPr ? kid(rPr, "solidFill") : null;
        const clr = fill ? (kid(fill, "srgbClr") ? at(kid(fill, "srgbClr"), "val") : (kid(fill, "schemeClr") ? "scheme:" + at(kid(fill, "schemeClr"), "val") : null)) : null;
        para.runs.push({
          text: textOf(r), b: rPr ? isOn(at(rPr, "b")) : false, i: rPr ? isOn(at(rPr, "i")) : false, u: rPr ? !!(at(rPr, "u") && at(rPr, "u") !== "none") : false,
          sz: rPr && at(rPr, "sz") ? Number(at(rPr, "sz")) / 100 : null, font: latin ? at(latin, "typeface") || "" : "", color: clr,
          // A WordArt style adds an outline, a gradient or effects to the letters themselves.
          wordart: !!(rPr && (kid(rPr, "ln") || kid(rPr, "gradFill") || kid(rPr, "effectLst") || kid(rPr, "pattFill"))),
        });
      }
      obj.paragraphs.push(para);
    }
    obj.text = obj.paragraphs.map((p) => p.text).join("\n");
  }

  /** A chart part: type, series, title, data labels, legend. */
  async function readChart(io, path) {
    const chart = { path: path, types: [], type: "", grouping: "", series: [], title: null, hasTitle: false, dataLabels: false, legend: "none" };
    const xml = await io.part(path);
    if (!xml) return chart;
    const ch = kid(xml.documentElement, "chart");
    if (!ch) return chart;
    const titleEl = kid(ch, "title");
    if (titleEl) { chart.hasTitle = true; chart.title = textOf(titleEl); }
    const plot = kid(ch, "plotArea");
    const groups = plot ? Array.from(plot.children).filter((e) => /Chart$/.test(e.localName)) : [];
    for (const g of groups) {
      let t = g.localName.replace(/3D/, "").replace(/Chart$/, "");
      if (t === "bar") t = kid(g, "barDir") && at(kid(g, "barDir"), "val") === "bar" ? "bar" : "column";
      chart.types.push(t);
      if (!chart.grouping && kid(g, "grouping")) chart.grouping = at(kid(g, "grouping"), "val") || "";
      if (all(g, "dLbls").some((d) => kid(d, "showVal") && isOn(at(kid(d, "showVal"), "val")) || kid(d, "showPercent") && isOn(at(kid(d, "showPercent"), "val")))) chart.dataLabels = true;
      for (const ser of kids(g, "ser")) {
        const tx = kid(ser, "tx");
        chart.series.push({ name: tx ? (find(tx, "v") ? find(tx, "v").textContent : textOf(tx)) : "" });
      }
    }
    chart.type = chart.types.length > 1 && new Set(chart.types).size > 1 ? "combo" : (chart.types[0] || "");
    const legend = kid(ch, "legend");
    if (legend) chart.legend = { b: "bottom", t: "top", l: "left", r: "right", tr: "topRight" }[kid(legend, "legendPos") ? at(kid(legend, "legendPos"), "val") : "r"] || "right";
    return chart;
  }

  // =====================================================================
  // PART 3: LOOKING THINGS UP IN THE MODEL
  // =====================================================================

  /** "slide 3" or "the title slide" for messages. */
  function slideLabel(s) { return s ? "slide " + s.number + (s.title ? ' ("' + snippet(s.title, 30) + '")' : "") : "the slide"; }
  /** Where "Show me where" should go: a slide, and maybe a shape on it. */
  function whereSlide(s, shape) { return s ? { slide: s.number, slideId: s.id, shape: shape ? shape.name : null } : null; }

  /**
   * Which slides a check is about:
   *   slide=3 (a number), slide=last, title="Our prices" (a slide whose title contains it),
   *   or slides=all / no slide at all (every slide).
   */
  function slidesFor(model, params) {
    if (has(params, "slide")) {
      const v = String(params.slide).trim().toLowerCase();
      if (v === "all") return { list: model.slides };
      if (v === "last") return { list: model.slides.slice(-1) };
      if (v === "first") return { list: model.slides.slice(0, 1) };
      const n = Number(v);
      if (!n || !model.slides[n - 1]) return { error: "The presentation doesn't have a slide " + params.slide + " (it has " + count(model.slides.length, "slide") + ")." };
      return { list: [model.slides[n - 1]] };
    }
    if (has(params, "title")) {
      const list = model.slides.filter((s) => loose(s.title).includes(loose(params.title)));
      if (!list.length) return { error: 'I can\'t find a slide titled "' + params.title + '".', missingTitle: true };
      return { list: list };
    }
    return { list: model.slides };
  }

  /** The shapes on the chosen slides that match kind=, says=, name= and phType=. */
  function shapesFor(model, params, slides) {
    const out = [];
    for (const s of slides) {
      for (const o of s.shapes) {
        if (has(params, "kind") && params.kind !== "any") {
          const k = params.kind;
          const ok = k === o.kind || (k === "shape" && ["shape", "textbox", "actionButton"].includes(o.kind)) || (k === "text" && ["textbox", "placeholder", "shape"].includes(o.kind))
            || (k === "media" && ["video", "audio"].includes(o.kind)) || (k === "image" && ["picture", "icon"].includes(o.kind)) || (k === "graphic" && !["placeholder", "group", "other"].includes(o.kind));
          if (!ok) continue;
        } else if (o.kind === "other") continue;
        if (has(params, "says") && !loose(o.text).includes(loose(params.says))) continue;
        if (has(params, "name") && !loose(o.name).includes(loose(params.name))) continue;
        out.push({ slide: s, shape: o });
      }
    }
    return out;
  }

  const KIND_WORDS = { picture: "picture", icon: "icon", shape: "shape", textbox: "text box", placeholder: "placeholder", smartart: "SmartArt graphic",
    chart: "chart", table: "table", video: "video", audio: "audio clip", model3d: "3D model", zoom: "Zoom", ink: "ink drawing", group: "group",
    actionButton: "action button", line: "line", text: "text box or placeholder", media: "video or audio clip", image: "picture or icon", graphic: "graphic", any: "object" };
  function kindWord(k) { const w = KIND_WORDS[k] || "object"; return (/^[aeiou]/i.test(w) ? "an " : "a ") + w; }

  /** Slides whose text (or notes) has the value. */
  function slideHasText(s, value, inNotes) {
    if (inNotes) return loose(s.notes).includes(loose(value));
    return s.shapes.some((o) => loose(o.text).includes(loose(value)));
  }

  /** Every run of text in a shape (or only in paragraphs that contain `says`). */
  function runsOf(shape, says) {
    const paras = has({ says: says }, "says") ? shape.paragraphs.filter((p) => loose(p.text).includes(loose(says))) : shape.paragraphs;
    const use = paras.length ? paras : shape.paragraphs;
    return use.flatMap((p) => p.runs.filter((r) => r.text.trim()));
  }

  // =====================================================================
  // PART 4: THE CHECKS
  // ---------------------------------------------------------------------
  // Each check is CHECKS[name](model, params) -> pass(...) or fail(...).
  // Messages are for learners: plain words, the PowerPoint command to use.
  // =====================================================================
  /** A color as hex: "1F3864" stays, "scheme:tx1" becomes the theme's dark 1 color. Tints and shades are ignored. */
  function resolveColor(model, c) {
    if (!c) return null;
    if (!/^scheme:/.test(c)) return c.toUpperCase();
    const map = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };
    const name = c.slice(7);
    const colors = (model.theme && model.theme.colors) || {};
    return colors[map[name] || name] || null;
  }
  /** WCAG contrast ratio between two hex colors (1 to 21). */
  function contrast(a, b) {
    const lum = (hex) => {
      const v = [0, 2, 4].map((i) => parseInt(hex.substr(i, 2), 16) / 255).map((x) => x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4));
      return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
    };
    const [l1, l2] = [lum(a), lum(b)].sort((x, y) => y - x);
    return (l1 + 0.05) / (l2 + 0.05);
  }
  /** The solid color behind a slide's text: its own background, else the master's, else the theme's light 1. Null for pictures and gradients. */
  function backgroundColor(model, s) {
    const m = model.masters[0];
    const bg = s.background.kind !== "none" ? s.background : (m && m.background && m.background.kind !== "style" && m.background.kind !== "none" ? m.background : { kind: "solid", color: "scheme:bg1" });
    if (bg.kind !== "solid") return null;
    return resolveColor(model, bg.color || "scheme:bg1");
  }

  const CHECKS = {};

  // ---------- files and properties ----------
  CHECKS["file.name"] = function (model, params) {
    const raw = String(model.fileName || "").split("?")[0].split("#")[0];
    let name = raw.split(/[\\/]/).pop();
    try { name = decodeURIComponent(name); } catch (e) { /* keep it */ }
    if (!name) return fail("Save the file first. Then select Check my work again.");
    if (/^lastname[_ ]/i.test(name)) return fail('Almost: put your own last name in place of "Lastname".');
    if (wildcard(params.pattern).test(name)) return pass('The file is saved as "' + name + '".');
    return fail('The file is named "' + name + '". It should look like "' + String(params.pattern).replace(/^\*/, "Lastname") + '".');
  };
  CHECKS["props.is"] = function (model, params) {
    for (const key of ["title", "author", "keywords", "subject"]) {
      if (!has(params, key)) continue;
      const got = model.props[key] || "";
      if (!loose(got).includes(loose(params[key]))) {
        const word = { title: "Title", author: "Author", keywords: "Tags", subject: "Subject" }[key];
        return fail(got ? "The " + word + ' property is "' + snippet(got) + '". It should be "' + params[key] + '" (File > Info).' : "The " + word + ' property is empty. It should be "' + params[key] + '" (File > Info).');
      }
    }
    return pass("The properties are set.");
  };
  CHECKS["self.attest"] = function () { return pass("Thanks for confirming."); };

  // ---------- text anywhere ----------
  CHECKS["text.present"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const inNotes = params.in === "notes";
    const hits = x.list.filter((s) => slideHasText(s, params.value, inNotes));
    const min = has(params, "min") ? Number(params.min) : 1;
    // Count slides that have it (min=2 means "on at least two slides").
    if (hits.length < min) {
      const where = has(params, "slide") || has(params, "title") ? (x.list.length === 1 ? slideLabel(x.list[0]) : "those slides") : "the presentation";
      return fail(hits.length ? 'I found "' + params.value + '" on ' + count(hits.length, "slide") + "; it should be on at least " + min + "." : 'I don\'t see "' + params.value + '" ' + (inNotes ? "in the speaker notes of " : "on ") + where + " yet.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    }
    return pass('I found "' + params.value + '".');
  };
  CHECKS["text.absent"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hit = x.list.find((s) => slideHasText(s, params.value, params.in === "notes"));
    if (hit) return fail('"' + params.value + '" is still on ' + slideLabel(hit) + ".", whereSlide(hit));
    return pass('"' + params.value + '" is gone.');
  };

  // ---------- slides ----------
  CHECKS["slide.count"] = function (model, params) {
    let list = model.slides;
    if (has(params, "title")) list = list.filter((s) => loose(s.title).includes(loose(params.title)));
    if (has(params, "visible") && truthy(params.visible)) list = list.filter((s) => !s.hidden);
    const n = list.length;
    const what = has(params, "title") ? 'slides titled "' + params.title + '"' : "slides";
    if (has(params, "min") && n < Number(params.min)) return fail("I see " + n + " " + what + "; there should be at least " + params.min + ".");
    if (has(params, "max") && n > Number(params.max)) return fail("I see " + n + " " + what + "; there should be no more than " + params.max + ".");
    if (has(params, "is") && n !== Number(params.is)) return fail("I see " + n + " " + what + "; there should be " + params.is + ".");
    return pass("The slide count is right.");
  };
  CHECKS["slide.is"] = function (model, params) {
    if (has(params, "exists") && !truthy(params.exists)) {
      const x0 = slidesFor(model, params);
      if (x0.missingTitle) return pass("That slide is gone.");
      if (x0.error) return pass("That slide is gone.");
      return fail(slideLabel(x0.list[0]) + " is still there. Delete it.", whereSlide(x0.list[0]));
    }
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const s = x.list[0];
    const wh = whereSlide(s);
    if (has(params, "number") && s.number !== Number(params.number)) return fail(slideLabel(s).replace(/^s/, "S") + " is slide " + s.number + ". It should be slide " + params.number + ".", wh);
    if (has(params, "after")) {
      const other = model.slides.find((t) => loose(t.title).includes(loose(params.after)));
      if (!other) return fail('I can\'t find a slide titled "' + params.after + '".');
      if (s.number !== other.number + 1) return fail(slideLabel(s).replace(/^s/, "S") + ' should come right after "' + other.title + '".', wh);
    }
    if (has(params, "layout") && !loose(s.layout).includes(loose(params.layout))) return fail(slideLabel(s).replace(/^s/, "S") + ' uses the "' + s.layout + '" layout. It should use "' + params.layout + '" (Home > Layout).', wh);
    if (has(params, "hidden") && s.hidden !== truthy(params.hidden)) return fail(truthy(params.hidden) ? slideLabel(s).replace(/^s/, "S") + " isn't hidden yet. Right-click it > Hide Slide." : slideLabel(s).replace(/^s/, "S") + " is hidden.", wh);
    if (has(params, "section") && !loose(s.section).includes(loose(params.section))) return fail(slideLabel(s).replace(/^s/, "S") + (s.section ? ' is in the "' + s.section + '" section' : " isn't in a section") + '. It should be in "' + params.section + '".', wh);
    if (has(params, "background")) {
      const want = params.background;
      const got = s.background.kind;
      const ok = want === "any" ? got !== "none" : want === "none" ? got === "none" : got === want;
      if (!ok) return fail(got === "none" ? slideLabel(s).replace(/^s/, "S") + " still uses the theme background. Use Design > Format Background." : slideLabel(s).replace(/^s/, "S") + " has a " + got + " background; it should have a " + want + " background.", wh);
      // "Apply to All" puts a picture on the slide master, so every plain slide shows it too.
      const mbg = model.masters[0] && model.masters[0].background;
      if (want === "none" && mbg && mbg.kind === "picture") return fail("Every slide shows the picture, because it was applied to all slides (it's now on the slide master). Undo it, or remove it in View > Slide Master, then give only the one slide its picture.", wh);
    }
    if (has(params, "notes") && !loose(s.notes).includes(loose(params.notes))) return fail(s.notes.trim() ? "The speaker notes for " + slideLabel(s) + ' don\'t mention "' + params.notes + '" yet.' : slideLabel(s).replace(/^s/, "S") + " has no speaker notes yet. Type them under the slide (View > Notes).", wh);
    if (has(params, "hasNotes") && truthy(params.hasNotes) && !s.notes.trim()) return fail(slideLabel(s).replace(/^s/, "S") + " has no speaker notes yet.", wh);
    if (has(params, "titleIs") && loose(s.title) !== loose(params.titleIs)) return fail(slideLabel(s).replace(/^s/, "S") + ' has the title "' + snippet(s.title) + '". It should be "' + params.titleIs + '".', wh);
    return pass(slideLabel(s).replace(/^s/, "S") + " is set up.");
  };
  CHECKS["notes.count"] = function (model, params) {
    const n = model.slides.filter((s) => s.notes.trim()).length;
    if (has(params, "min") && n < Number(params.min)) return fail(count(n, "slide") + " " + (n === 1 ? "has" : "have") + " speaker notes; at least " + params.min + " should.");
    if (has(params, "max") && n > Number(params.max)) return fail(count(n, "slide") + " still " + (n === 1 ? "has" : "have") + " speaker notes.");
    return pass("The speaker notes are there.");
  };

  // ---------- sections ----------
  CHECKS["section.is"] = function (model, params) {
    const sec = model.sections.find((x) => loose(x.name) === loose(params.name)) || model.sections.find((x) => loose(x.name).includes(loose(params.name)));
    if (has(params, "exists") && !truthy(params.exists)) return sec ? fail('The section "' + sec.name + '" is still there.') : pass("That section is gone.");
    if (!sec) return fail(model.sections.length ? 'I don\'t see a section named "' + params.name + '". The sections are: ' + model.sections.map((x) => x.name).join(", ") + "." : "There are no sections yet. Right-click between slides > Add Section.");
    if (has(params, "index")) {
      const i = model.sections.indexOf(sec) + 1;
      if (i !== Number(params.index)) return fail('"' + sec.name + '" is section ' + i + ". It should be section " + params.index + " (drag it, or right-click it > Move Section Up).");
    }
    if (has(params, "slides") && sec.slides.length !== Number(params.slides)) return fail('The "' + sec.name + '" section has ' + count(sec.slides.length, "slide") + "; it should have " + params.slides + ".");
    if (has(params, "minSlides") && sec.slides.length < Number(params.minSlides)) return fail('The "' + sec.name + '" section has ' + count(sec.slides.length, "slide") + "; it should have at least " + params.minSlides + ".");
    if (has(params, "starts")) {
      const first = model.slides[sec.slides[0] - 1];
      if (!first || !loose(first.title).includes(loose(params.starts))) return fail('The "' + sec.name + '" section should start with the slide "' + params.starts + '".');
    }
    if (has(params, "contains")) {
      const titles = sec.slides.map((n) => model.slides[n - 1] ? model.slides[n - 1].title : "");
      const missing = listOf(params.contains).filter((t) => !titles.some((x) => loose(x).includes(loose(t))));
      if (missing.length) return fail('The "' + sec.name + '" section doesn\'t include the slide' + (missing.length > 1 ? "s " : " ") + missing.map((t) => '"' + t + '"').join(", ") + ".");
    }
    return pass('The "' + sec.name + '" section is set up.');
  };
  CHECKS["section.count"] = function (model, params) {
    const n = model.sections.length;
    if (has(params, "min") && n < Number(params.min)) return fail(n ? "There are " + count(n, "section") + "; there should be at least " + params.min + "." : "There are no sections yet. Right-click between slides > Add Section.");
    if (has(params, "max") && n > Number(params.max)) return fail("There are " + count(n, "section") + "; there should be no more than " + params.max + ".");
    return pass("The sections are there.");
  };

  // ---------- slide size and views ----------
  CHECKS["size.is"] = function (model, params) {
    const ratio = model.size.cy ? model.size.cx / model.size.cy : 0;
    const name = Math.abs(ratio - 16 / 9) < 0.01 ? "widescreen" : Math.abs(ratio - 4 / 3) < 0.01 ? "standard" : "custom";
    const label = { widescreen: "Widescreen (16:9)", standard: "Standard (4:3)", custom: "a custom size" };
    if (has(params, "ratio")) {
      const want = { "16:9": "widescreen", widescreen: "widescreen", "4:3": "standard", standard: "standard" }[String(params.ratio).toLowerCase()];
      if (name !== want) return fail("The slides are " + label[name] + ". They should be " + label[want] + " (Design > Slide Size).");
    }
    return pass("The slide size is " + label[name] + ".");
  };
  CHECKS["view.is"] = function (model, params) {
    if (has(params, "view")) {
      const want = VIEW_PARAM[params.view] || [params.view];
      if (!want.includes(model.view.lastView)) return fail("The presentation was saved in " + (VIEW_NAMES[model.view.lastView] || "Normal") + " view. Switch to " + ({ normal: "Normal", outline: "Outline", sorter: "Slide Sorter", notes: "Notes Page" }[params.view] || params.view) + " view (View tab, or the buttons at the bottom right), then save.");
    }
    if (has(params, "guides") && model.view.guides !== truthy(params.guides)) return fail(truthy(params.guides) ? "The guides aren't turned on yet (View > Guides)." : "The guides are still on.");
    return pass("The view is right.");
  };

  // ---------- text formatting, lists and columns ----------
  CHECKS["text.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const shapeParams = Object.assign({}, params);
    if (!has(params, "kind")) shapeParams.kind = "text";
    if (has(params, "phType")) delete shapeParams.kind;
    let hits = shapesFor(model, shapeParams, x.list);
    if (has(params, "phType")) hits = hits.filter((h) => h.shape.ph && (h.shape.ph.type === params.phType || (params.phType === "title" && h.shape.ph.type === "ctrTitle")));
    if (!hits.length) return fail(has(params, "says") ? 'I can\'t find the text "' + params.says + '"' + (x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "") + "." : "I can't find that text.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    const { slide: s, shape: o } = hits[0];
    const wh = whereSlide(s, o);
    const label = '"' + snippet(has(params, "says") ? params.says : o.text, 30) + '"';
    const runs = runsOf(o, params.says);
    const allRuns = (test) => runs.length > 0 && runs.every(test);
    if (has(params, "bold") && !allRuns((r) => r.b === truthy(params.bold))) return fail(label + (truthy(params.bold) ? " isn't bold yet." : " is still bold."), wh);
    if (has(params, "italic") && !allRuns((r) => r.i === truthy(params.italic))) return fail(label + (truthy(params.italic) ? " isn't italic yet." : " is still italic."), wh);
    if (has(params, "underline") && !allRuns((r) => r.u === truthy(params.underline))) return fail(label + (truthy(params.underline) ? " isn't underlined yet." : " is still underlined."), wh);
    if (has(params, "size")) {
      const want = Number(String(params.size).replace(/pt$/i, ""));
      const bad = runs.find((r) => r.sz === null || Math.abs(r.sz - want) > 0.5);
      if (bad || !runs.length) return fail(label + " is " + (bad && bad.sz ? bad.sz + " pt" : "the layout's size") + ". It should be " + want + " pt.", wh);
    }
    if (has(params, "minSize")) {
      const bad = runs.find((r) => r.sz !== null && r.sz < Number(params.minSize));
      if (bad) return fail(label + " is " + bad.sz + " pt, which is too small to read from the back of a room. Make it at least " + params.minSize + " pt.", wh);
    }
    if (has(params, "font") && !allRuns((r) => loose(r.font) === loose(params.font))) {
      const got = runs.find((r) => loose(r.font) !== loose(params.font));
      return fail(label + " uses " + (got && got.font ? got.font : "the theme font") + ". It should use " + params.font + ".", wh);
    }
    if (has(params, "notFont")) {
      const bad = model.slides.flatMap((sl) => sl.shapes.flatMap((sh) => sh.paragraphs.flatMap((p) => p.runs.filter((r) => loose(r.font) === loose(params.notFont)).map(() => ({ sl, sh })))));
      if (bad.length) return fail(params.notFont + " is still used on " + slideLabel(bad[0].sl) + ". Use Home > Replace > Replace Fonts to change it everywhere.", whereSlide(bad[0].sl, bad[0].sh));
    }
    if (has(params, "wordart")) {
      const styled = runs.some((r) => r.wordart);
      if (styled !== truthy(params.wordart)) return fail(truthy(params.wordart) ? label + " doesn't have a WordArt style yet (Shape Format > WordArt Styles)." : label + " still has a WordArt style.", wh);
    }
    if (has(params, "color") && !allRuns((r) => r.color && loose(r.color.replace(/^#/, "")) === loose(String(params.color).replace(/^#/, "")))) return fail(label + " isn't that color yet.", wh);
    if (has(params, "colorChanged") && truthy(params.colorChanged) && !runs.some((r) => r.color)) return fail(label + " still uses the theme's text color.", wh);
    if (has(params, "colorChanged") && !truthy(params.colorChanged) && runs.some((r) => r.color && !/^scheme:tx1$/.test(r.color))) return fail(label + " has a new color. It should keep the theme's text color.", wh);
    if (has(params, "minContrast")) {
      const back = backgroundColor(model, s);
      if (back) {
        const low = runs.map((r) => ({ r: r, fg: resolveColor(model, r.color || "scheme:tx1") })).find((x) => x.fg && contrast(x.fg, back) < Number(params.minContrast));
        if (low) return fail(label + " is hard to read: its color has a contrast of only " + (Math.round(contrast(low.fg, back) * 10) / 10) + " to 1 with the background. Pick a darker (or lighter) color so it's at least " + params.minContrast + " to 1.", wh);
      }
    }
    if (has(params, "columns") && o.cols !== Number(params.columns)) return fail(label + (o.cols > 1 ? " is in " + o.cols + " columns" : " is in one column") + ". Set it to " + params.columns + " columns (Home > Add or Remove Columns).", wh);
    if (has(params, "align")) {
      const paras = o.paragraphs.filter((p) => p.text.trim());
      const want = { left: "l", center: "ctr", right: "r", justify: "just" }[params.align];
      if (!paras.every((p) => (p.align || "l") === want)) return fail(label + " isn't " + params.align + "-aligned yet.", wh);
    }
    return pass(label + " is formatted.");
  };
  CHECKS["list.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hits = shapesFor(model, Object.assign({}, params, { kind: "text" }), x.list);
    if (!hits.length) return fail('I can\'t find the text "' + (params.says || "") + '"' + (x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "") + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    // With no "says", a list means the text that isn't a slide title, preferably the one with the most lines.
    const isTitle = (h) => h.shape.ph && /title/i.test(h.shape.ph.type || "") && !/sub/i.test(h.shape.ph.type || "");
    const candidates = has(params, "says") ? hits : (hits.filter((h) => !isTitle(h)).length ? hits.filter((h) => !isTitle(h)) : hits)
      .slice().sort((a, b) => b.shape.paragraphs.filter((p) => p.text.trim()).length - a.shape.paragraphs.filter((p) => p.text.trim()).length);
    // Looking across several slides for "a numbered list somewhere": try the lists of that kind first.
    if (!has(params, "says") && has(params, "kind")) {
      const inherits = (h) => h.shape.kind === "placeholder" && (!h.shape.ph || h.shape.ph.type === "body" || h.shape.ph.type === "obj");
      const isKind = (h) => { const ps = h.shape.paragraphs.filter((p) => p.text.trim()); return ps.length > 0 && ps.every((p) => (p.bullet || (inherits(h) ? "bullet" : "none")) === params.kind); };
      candidates.sort((a, b) => (isKind(b) ? 1 : 0) - (isKind(a) ? 1 : 0));
    }
    const { slide: s, shape: o } = candidates[0];
    const wh = whereSlide(s, o);
    let paras = o.paragraphs.filter((p) => p.text.trim());
    if (has(params, "items")) paras = paras.filter((p) => listOf(params.items).some((it) => loose(p.text).includes(loose(it))));
    if (!paras.length) return fail("I can't find those list items.", wh);
    // A body placeholder is bulleted by its layout unless a paragraph turns bullets off.
    const inherited = o.kind === "placeholder" && (!o.ph || o.ph.type === "body" || o.ph.type === "obj");
    const kindOf = (p) => p.bullet || (inherited ? "bullet" : "none");
    if (has(params, "kind")) {
      const bad = paras.find((p) => kindOf(p) !== params.kind);
      if (bad) return fail('"' + snippet(bad.text, 30) + '" ' + (kindOf(bad) === "none" ? "isn't in a list yet" : "is a " + kindOf(bad) + "ed item") + ". It should be " + (params.kind === "number" ? "numbered (Home > Numbering)." : "bulleted (Home > Bullets)."), wh);
    }
    if (has(params, "level")) {
      const bad = paras.find((p) => p.lvl + 1 !== Number(params.level));
      if (bad) return fail('"' + snippet(bad.text, 30) + '" is at list level ' + (bad.lvl + 1) + ". It should be at level " + params.level + " (Tab or Increase List Level).", wh);
    }
    if (has(params, "maxItems")) {
      // Every slide in the selection (for example both halves of a split slide) must stay short.
      for (const sl of x.list) {
        const lists = sl.shapes.filter((z) => z.paragraphs.length && !(z.ph && /title/i.test(z.ph.type || "") && !/sub/i.test(z.ph.type || "")));
        const most = lists.reduce((n, z) => Math.max(n, z.paragraphs.filter((p) => p.text.trim()).length), 0);
        if (most > Number(params.maxItems)) return fail(slideLabel(sl).replace(/^s/, "S") + " has " + most + " lines in one list. Keep it to " + params.maxItems + " or fewer: split the slide (duplicate it, then delete half the points from each copy).", whereSlide(sl));
      }
    }
    if (has(params, "levels")) {
      const levels = new Set(o.paragraphs.filter((p) => p.text.trim()).map((p) => p.lvl));
      if (levels.size < Number(params.levels)) return fail("The list has " + count(levels.size, "level") + ". Indent the details under each point (Tab) to make " + params.levels + " levels.", wh);
    }
    return pass("The list is set up.");
  };

  // ---------- headers, footers and slide numbers ----------
  CHECKS["footer.is"] = function (model, params) {
    const ph = (s, type) => s.shapes.find((o) => o.ph && o.ph.type === type);
    const titleSlides = model.slides.filter((s) => s.shapes.some((o) => o.ph && o.ph.type === "ctrTitle"));
    const others = model.slides.filter((s) => !titleSlides.includes(s));
    if (has(params, "slideNumbers") && truthy(params.slideNumbers)) {
      const missing = others.find((s) => !ph(s, "sldNum"));
      if (missing) return fail(slideLabel(missing).replace(/^s/, "S") + " has no slide number. Use Insert > Header & Footer > Slide number > Apply to All.", whereSlide(missing));
    }
    if (has(params, "footer")) {
      const missing = others.find((s) => !ph(s, "ftr") || !loose(ph(s, "ftr").text).includes(loose(params.footer)));
      if (missing) return fail(slideLabel(missing).replace(/^s/, "S") + ' doesn\'t have the footer "' + params.footer + '". Use Insert > Header & Footer > Footer > Apply to All.', whereSlide(missing));
    }
    if (has(params, "skipTitle") && truthy(params.skipTitle)) {
      const shown = titleSlides.find((s) => ph(s, "sldNum") || ph(s, "ftr"));
      if (shown) return fail("The title slide still shows the slide number or footer. Tick Don't show on title slide in Insert > Header & Footer.", whereSlide(shown));
    }
    if (has(params, "date") && truthy(params.date)) {
      const missing = others.find((s) => !ph(s, "dt"));
      if (missing) return fail(slideLabel(missing).replace(/^s/, "S") + " has no date. Use Insert > Header & Footer > Date and time.", whereSlide(missing));
    }
    return pass("The footers and slide numbers are set.");
  };

  // ---------- themes, masters and layouts ----------
  CHECKS["theme.is"] = function (model, params) {
    const t = model.theme;
    if (has(params, "name") && !loose(t.name).includes(loose(params.name))) return fail("The presentation uses the " + (t.name || "Office") + " theme. It should use " + params.name + " (Design > Themes).");
    if (has(params, "changed") && truthy(params.changed) && /^office( theme)?$/i.test(t.name || "Office Theme")) return fail("The presentation still uses the default Office Theme. Pick another on the Design tab.");
    if (has(params, "fonts") && !loose(t.fontsName).includes(loose(params.fonts)) && !loose(t.majorFont).includes(loose(params.fonts))) return fail("The theme fonts are " + (t.fontsName || t.majorFont || "the defaults") + ". They should be " + params.fonts + " (Design > Variants > Fonts).");
    if (has(params, "fontsChanged") && truthy(params.fontsChanged) && /^office$/i.test(t.fontsName || "Office") && /^(aptos display|calibri light)$/i.test(t.majorFont || "Calibri Light")) return fail("The theme fonts haven't changed yet. Use Design > Variants > More > Fonts.");
    if (has(params, "colors") && !loose(t.colorsName).includes(loose(params.colors))) return fail("The theme colors are " + (t.colorsName || "the defaults") + ". They should be " + params.colors + ".");
    if (has(params, "variant") && truthy(params.variant) && !t.colorsName) return fail("Pick one of the theme's variants on the Design tab.");
    return pass("The theme is set.");
  };
  CHECKS["master.is"] = function (model, params) {
    const kind = params.kind || "slide";
    const m = kind === "notes" ? model.notesMaster : kind === "handout" ? model.handoutMaster : model.masters[0];
    const what = { slide: "slide master", notes: "notes master", handout: "handout master" }[kind] || "master";
    const menu = { slide: "View > Slide Master", notes: "View > Notes Master", handout: "View > Handout Master" }[kind];
    if (!m) return fail("This presentation has no " + what + " yet. Open " + menu + " once to create it.");
    if (has(params, "background")) {
      const got = m.background.kind;
      const ok = params.background === "any" ? got !== "none" && got !== "style" || (got === "style" && m.background.idx !== "1001") : got === params.background;
      if (!ok) return fail("The " + what + "'s background hasn't changed yet. In " + menu + ", select the master (the top thumbnail), then Background Styles or Format Background.");
    }
    if (has(params, "titleFont") && !loose(m.titleFont).includes(loose(params.titleFont)) && !(m.theme && loose(m.theme.majorFont).includes(loose(params.titleFont)))) return fail("The " + what + "'s title font is " + (m.titleFont || "the theme font") + ". It should be " + params.titleFont + ".");
    if (has(params, "picture") && truthy(params.picture) && !m.shapes.some((o) => o.kind === "picture" || o.kind === "icon")) return fail("The " + what + " doesn't have a picture on it yet. Select the master (the top thumbnail) in " + menu + ", then Insert > Pictures.");
    if (has(params, "says")) {
      const want = listOf(params.says);
      const missing = want.filter((w) => !m.shapes.some((o) => loose(o.text).includes(loose(w))));
      if (missing.length) return fail('The ' + what + ' doesn\'t say "' + missing[0] + '" yet. Type it in ' + menu + ".");
    }
    if (has(params, "phSays")) {
      const [type, text] = String(params.phSays).split(":");
      const ph = m.shapes.find((o) => o.ph && o.ph.type === type);
      if (!ph || !loose(ph.text).includes(loose(text))) return fail("The " + ({ hdr: "header", ftr: "footer", dt: "date" }[type] || type) + ' on the ' + what + ' doesn\'t say "' + text + '" yet (' + menu + ").");
    }
    if (has(params, "lacks")) {
      const type = params.lacks;
      if (m.shapes.some((o) => o.ph && o.ph.type === type)) return fail("The " + what + " still has its " + ({ hdr: "header", ftr: "footer", dt: "date", sldNum: "page number" }[type] || type) + " placeholder. Clear it in " + menu + " (Placeholders group).");
    }
    return pass("The " + what + " is set up.");
  };
  CHECKS["layout.is"] = function (model, params) {
    const layouts = model.masters.flatMap((m) => m.layouts);
    const l = layouts.find((x) => loose(x.name) === loose(params.name)) || layouts.find((x) => loose(x.name).includes(loose(params.name)));
    if (has(params, "exists") && !truthy(params.exists)) return l ? fail('The "' + l.name + '" layout is still there.') : pass("That layout is gone.");
    if (!l) return fail('There\'s no layout named "' + params.name + '" yet. In View > Slide Master, use Insert Layout, then Rename.');
    if (has(params, "custom") && truthy(params.custom) && !l.custom) return fail('"' + l.name + '" is one of the theme\'s built-in layouts. Make your own with Insert Layout.');
    if (has(params, "placeholder")) {
      const want = listOf(params.placeholder);
      const names = { pic: "Picture", body: "Text", obj: "Content", title: "Title", chart: "Chart", tbl: "Table", media: "Media" };
      const missing = want.filter((t) => !l.placeholders.includes(t) && !(t === "body" && l.placeholders.includes("obj")));
      if (missing.length) return fail('The "' + l.name + '" layout has no ' + names[missing[0]] + " placeholder yet. Use Insert Placeholder > " + names[missing[0]] + ".");
    }
    if (has(params, "lacks") && l.placeholders.includes(params.lacks)) return fail('The "' + l.name + '" layout still has its ' + ({ ftr: "footer", dt: "date", sldNum: "slide number", title: "title" }[params.lacks] || params.lacks) + " placeholder.");
    if (has(params, "used") && truthy(params.used) && !model.slides.some((s) => s.layoutPath === l.path)) return fail('No slide uses the "' + l.name + '" layout yet. Pick it in Home > Layout.');
    return pass('The "' + l.name + '" layout is set up.');
  };

  // ---------- graphics ----------
  CHECKS["object.count"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const n = shapesFor(model, params, x.list).length;
    const word = KIND_WORDS[params.kind] || "object";
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (has(params, "min") && n < Number(params.min)) return fail(n === 0 ? "I don't see " + kindWord(params.kind) + where + " yet." : "I see " + count(n, word) + where + "; there should be at least " + params.min + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    if (has(params, "max") && n > Number(params.max)) return fail("I see " + count(n, word) + where + "; there should be no more than " + params.max + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    return pass("The " + word + "s are there.");
  };
  CHECKS["object.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hits = shapesFor(model, params, x.list);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (!hits.length) return fail("I don't see " + (has(params, "says") ? kindWord(params.kind || "any") + ' that says "' + params.says + '"' : kindWord(params.kind || "any")) + where + " yet.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    const { slide: s, shape: o } = has(params, "index") ? (hits[Number(params.index) - 1] || hits[0]) : hits[0];
    const wh = whereSlide(s, o);
    const label = "The " + (KIND_WORDS[o.kind] || "object") + (o.name ? ' "' + o.name + '"' : "");
    if (has(params, "alt") && truthy(params.alt) && !o.alt && !o.decorative) return fail(label + " has no alt text yet. Right-click it > View Alt Text.", wh);
    if (has(params, "decorative") && o.decorative !== truthy(params.decorative)) return fail(truthy(params.decorative) ? label + " isn't marked as decorative yet." : label + " is marked decorative.", wh);
    if (has(params, "width")) { const w = toEmu(params.width); if (Math.abs(o.w - w) > EMU_PER_INCH * 0.05) return fail(label + " is " + inches(o.w) + " wide. Make it " + inches(w) + " (Shape Format > Size).", wh); }
    if (has(params, "height")) { const hh = toEmu(params.height); if (Math.abs(o.h - hh) > EMU_PER_INCH * 0.05) return fail(label + " is " + inches(o.h) + " tall. Make it " + inches(hh) + " (Shape Format > Size).", wh); }
    if (has(params, "minWidth") && o.w < toEmu(params.minWidth)) return fail(label + " is only " + inches(o.w) + " wide. Make it at least " + inches(toEmu(params.minWidth)) + ".", wh);
    if (has(params, "cropped") && truthy(params.cropped) && !(o.crop && (o.crop.l || o.crop.t || o.crop.r || o.crop.b))) return fail(label + " isn't cropped yet (Picture Format > Crop).", wh);
    if (has(params, "square") && truthy(params.square) && !(o.w && o.h && Math.abs(o.w / o.h - 1) < 0.03)) return fail(label + " is " + inches(o.w) + " by " + inches(o.h) + ". Crop it to a square (Crop > Aspect Ratio > 1:1).", wh);
    if (has(params, "stretched") && !truthy(params.stretched) && o.image && o.w && o.h) {
      const c = o.crop || { l: 0, t: 0, r: 0, b: 0 };
      const shown = (o.image.w * (1 - c.l - c.r)) / (o.image.h * (1 - c.t - c.b));
      if (Math.abs(o.w / o.h / shown - 1) > 0.03) return fail(label + " is stretched out of shape. Drag a corner handle (not a side), or tick Lock aspect ratio in Format Picture > Size.", wh);
    }
    if (has(params, "styled") && truthy(params.styled) && !(o.effects || o.outline || (o.geom && o.geom !== "rect") || o.styleChanged)) return fail(label + " doesn't have a style yet (" + (o.kind === "picture" ? "Picture Format > Picture Styles" : "Shape Format > Shape Styles") + ").", wh);
    if (has(params, "effects") && truthy(params.effects) && !o.effects) return fail(label + " doesn't have an effect yet (" + (o.kind === "picture" ? "Picture Effects" : "Shape Effects") + ").", wh);
    if (has(params, "styleChanged") && truthy(params.styleChanged) && !o.styleChanged && !o.fill) return fail(label + " still has the default shape style. Pick one in Shape Format > Shape Styles.", wh);
    if (has(params, "geom") && !loose(o.geom).includes(loose(params.geom))) return fail(label + " is a " + (o.geom || "custom") + " shape; it should be a " + params.geom + " shape.", wh);
    if (has(params, "front") && truthy(params.front)) {
      // In front = nothing that comes later on the slide (higher in the stack) covers any part of it.
      const overlaps = (a, b) => a.x !== null && b.x !== null && a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
      const later = s.shapes.slice(s.shapes.indexOf(o) + 1).filter((z) => z.groupId === o.groupId && !(z.ph && ["sldNum", "ftr", "dt"].includes(z.ph.type)));
      if (later.some((z) => overlaps(o, z))) return fail(label + " is behind another object. Right-click it > Bring to Front.", wh);
    }
    if (has(params, "grouped") && truthy(params.grouped) && !o.groupId) return fail(label + " isn't in a group yet. Select the pieces, then Shape Format > Group.", wh);
    if (has(params, "rotated") && truthy(params.rotated) && !(o.model3d && o.model3d.rotated) && !o.rot) return fail(label + " hasn't been turned yet. Drag its rotation control, or pick one of the 3D Model Views.", wh);
    if (has(params, "items") && o.smartart && o.smartart.items < Number(params.items)) return fail("The SmartArt graphic has " + count(o.smartart.items, "item") + "; it should have " + params.items + ". Use SmartArt Design > Add Shape.", wh);
    if (has(params, "layout") && o.smartart && !loose(o.smartart.layout).includes(loose(params.layout))) return fail("The SmartArt graphic uses the " + o.smartart.layout + " layout. Change it in SmartArt Design > Layouts.", wh);
    if (has(params, "notLayout") && o.smartart && loose(o.smartart.layout) === loose(params.notLayout)) return fail("The SmartArt graphic still uses its first layout. Pick another in SmartArt Design > Layouts.", wh);
    if (has(params, "colorsChanged") && truthy(params.colorsChanged) && o.smartart && /^accent1_2$/.test(o.smartart.colors)) return fail("The SmartArt graphic still has the default colors. Use SmartArt Design > Change Colors.", wh);
    return pass(label + " is set up.");
  };
  CHECKS["align.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hits = shapesFor(model, params, x.list).map((h) => h.shape);
    const s = x.list[0];
    if (hits.length < 2) return fail("I need at least two " + (KIND_WORDS[params.kind] || "object") + "s" + " on " + slideLabel(s) + " to check.", whereSlide(s));
    const tol = EMU_PER_INCH * 0.03;
    const edge = { top: (o) => o.y, bottom: (o) => o.y + o.h, middle: (o) => o.y + o.h / 2, left: (o) => o.x, right: (o) => o.x + o.w, center: (o) => o.x + o.w / 2 };
    // Shapes inside a group keep group coordinates; that's fine for comparing them with each other.
    if (has(params, "edge")) {
      const f = edge[params.edge];
      const vals = hits.map(f);
      if (Math.max.apply(null, vals) - Math.min.apply(null, vals) > tol) return fail("The " + (KIND_WORDS[params.kind] || "object") + "s aren't " + params.edge + "-aligned yet. Select them all, then Shape Format > Align > Align " + params.edge.charAt(0).toUpperCase() + params.edge.slice(1) + ".", whereSlide(s));
    }
    if (has(params, "distributed")) {
      const horiz = params.distributed !== "vertical";
      const sorted = hits.slice().sort((a, b) => horiz ? a.x - b.x : a.y - b.y);
      const gaps = [];
      for (let i = 1; i < sorted.length; i++) gaps.push(horiz ? sorted[i].x - (sorted[i - 1].x + sorted[i - 1].w) : sorted[i].y - (sorted[i - 1].y + sorted[i - 1].h));
      if (gaps.length > 1 && Math.max.apply(null, gaps) - Math.min.apply(null, gaps) > tol) return fail("The gaps between them aren't even yet. Use Align > Distribute " + (horiz ? "Horizontally." : "Vertically."), whereSlide(s));
    }
    return pass("They line up.");
  };

  // ---------- tables and charts ----------
  CHECKS["table.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hits = shapesFor(model, Object.assign({}, params, { kind: "table" }), x.list);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (!hits.length) return fail("I don't see a table" + (has(params, "says") ? ' with "' + params.says + '" in it' : "") + where + " yet. Use Insert > Table.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    const { slide: s, shape: o } = hits[0];
    const t = o.table, wh = whereSlide(s, o);
    if (has(params, "rows") && t.rows !== Number(params.rows)) return fail("The table has " + count(t.rows, "row") + "; it should have " + params.rows + ".", wh);
    if (has(params, "cols") && t.cols !== Number(params.cols)) return fail("The table has " + count(t.cols, "column") + "; it should have " + params.cols + ".", wh);
    if (has(params, "styleChanged") && truthy(params.styleChanged) && (!t.styleId || t.styleId === "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")) return fail("The table still has the default style. Pick another in Table Design > Table Styles.", wh);
    if (has(params, "headerRow") && t.firstRow !== truthy(params.headerRow)) return fail(truthy(params.headerRow) ? "Turn on Header Row in Table Design, so screen readers know the first row holds headings." : "Header Row is still on.", wh);
    if (has(params, "lacks") && loose(o.text).includes(loose(params.lacks))) return fail('The table still has "' + params.lacks + '" in it.', wh);
    if (has(params, "has") && !loose(o.text).includes(loose(params.has))) return fail('The table doesn\'t have "' + params.has + '" in it yet.', wh);
    return pass("The table is set up.");
  };
  CHECKS["chart.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const hits = shapesFor(model, Object.assign({}, params, { kind: "chart" }), x.list);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (!hits.length || !hits[0].shape.chart) return fail("I don't see a chart" + where + " yet. Use Insert > Chart.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    const { slide: s, shape: o } = hits[0];
    const c = o.chart, wh = whereSlide(s, o);
    if (has(params, "type") && c.type !== params.type) return fail("The chart is a " + (c.type || "chart") + " chart. It should be a " + params.type + " chart (Chart Design > Change Chart Type).", wh);
    if (has(params, "notType") && c.type === params.notType) return fail("The chart is still a " + c.type + " chart. Change its type (Chart Design > Change Chart Type).", wh);
    if (has(params, "grouping") && c.grouping !== params.grouping) return fail("The chart's bars are " + (c.grouping || "standard") + ". They should be " + params.grouping + ".", wh);
    if (has(params, "minSeries") && c.series.length < Number(params.minSeries)) return fail("The chart shows " + count(c.series.length, "series", "series") + "; it should show at least " + params.minSeries + ".", wh);
    if (has(params, "chartTitle") && !(c.hasTitle && loose(c.title).includes(loose(params.chartTitle)))) return fail(c.hasTitle && c.title ? 'The chart title is "' + snippet(c.title) + '". It should be "' + params.chartTitle + '".' : "The chart has no title yet. Use Chart Elements > Chart Title (or Chart Design > Add Chart Element > Chart Title).", wh);
    if (has(params, "dataLabels") && c.dataLabels !== truthy(params.dataLabels)) return fail(truthy(params.dataLabels) ? "The chart has no data labels yet (Chart Elements > Data Labels)." : "The chart still has data labels.", wh);
    if (has(params, "alt") && truthy(params.alt) && !o.alt && !o.decorative) return fail("The chart has no alt text yet. Right-click its border > View Alt Text.", wh);
    return pass("The chart is set up.");
  };

  // ---------- media ----------
  CHECKS["media.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const kind = params.kind || "media";
    const hits = shapesFor(model, Object.assign({}, params, { kind: kind }), x.list);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (!hits.length) return fail("I don't see " + kindWord(kind) + where + " yet. Use Insert > " + (kind === "audio" ? "Audio" : "Video") + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    const { slide: s, shape: o } = hits[0];
    const m = o.media, wh = whereSlide(s, o);
    const label = o.kind === "audio" ? "The audio clip" : "The video";
    // "Start automatically" adds a play command that runs with or after the slide, not on a click.
    const plays = s.anims.filter((a) => a.cls === "mediacall" && a.target === o.id);
    if (has(params, "trimmed") && truthy(params.trimmed) && !(m.trimStart || m.trimEnd)) return fail(label + " isn't trimmed yet (Playback > Trim " + (o.kind === "audio" ? "Audio" : "Video") + ").", wh);
    if (has(params, "fadeIn") && truthy(params.fadeIn) && !m.fadeIn) return fail(label + " doesn't fade in yet. Set Fade In on the Playback tab.", wh);
    if (has(params, "start")) {
      const auto = plays.some((a) => a.start === "after" || a.start === "with");
      if (params.start === "auto" && !auto) return fail(label + " starts when clicked. Set Start to Automatically on the Playback tab.", wh);
      if (params.start === "click" && auto) return fail(label + " starts automatically. Set Start to When Clicked On.", wh);
    }
    if (has(params, "loop") && truthy(params.loop) && !s.anims.some((a) => a.target === o.id && a.repeat === "indefinite") && !o.loop) return fail(label + " doesn't loop yet. Tick Loop until Stopped on the Playback tab.", wh);
    if (has(params, "acrossSlides") && truthy(params.acrossSlides) && !o.acrossSlides) return fail(label + " stops when the slide changes. Tick Play Across Slides (or Play in Background) on the Playback tab.", wh);
    if (has(params, "alt") && truthy(params.alt) && !o.alt) return fail(label + " has no alt text yet. Describe what it shows (right-click > View Alt Text).", wh);
    return pass(label + " is set up.");
  };

  // ---------- transitions ----------
  CHECKS["transition.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    for (const s of x.list) {
      const t = s.transition;
      const wh = whereSlide(s);
      const L = slideLabel(s).replace(/^s/, "S");
      if (has(params, "exists") && !truthy(params.exists)) { if (t && t.type && t.type !== "cut") return fail(L + " still has a transition.", wh); continue; }
      if (!t || !t.type) {
        if (has(params, "advanceAfter") && t && t.advanceAfterMs !== null) { /* timing only, no effect: fall through */ }
        else return fail(L + " has no transition yet. Pick one on the Transitions tab.", wh);
      }
      if (has(params, "type") && !(loose(t.type) === loose(params.type) || loose(t.name) === loose(params.type))) return fail(L + " has the " + t.name + " transition. It should have " + params.type + ".", wh);
      if (has(params, "kind")) {
        if (params.kind === "3d" && !t.threeD) return fail(L + " has " + t.name + ", a basic transition. Pick one from the Exciting or Dynamic Content group, such as Cube, Vortex or Gallery.", wh);
        if (params.kind === "basic" && t.threeD) return fail(L + " has a 3D transition (" + t.name + "). Pick a subtle one, such as Fade or Push.", wh);
      }
      if (has(params, "optionChanged") && truthy(params.optionChanged) && !optionsChanged(t)) return fail("The " + t.name + " transition on " + slideLabel(s) + " still has its default effect option. Pick another in Transitions > Effect Options.", wh);
      if (has(params, "duration")) {
        const want = toMs(params.duration);
        if (t.durationMs === null || Math.abs(t.durationMs - want) > 50) return fail("The transition on " + slideLabel(s) + " lasts " + (t.durationMs === null ? "the default time" : seconds(t.durationMs)) + ". Set Duration to " + seconds(want) + ".", wh);
      }
      if (has(params, "advanceAfter")) {
        const want = toMs(params.advanceAfter);
        if (t.advanceAfterMs === null || Math.abs(t.advanceAfterMs - want) > 100) return fail(L + (t.advanceAfterMs === null ? " doesn't advance on its own yet" : " advances after " + seconds(t.advanceAfterMs)) + ". Tick After and set " + seconds(want) + " (Transitions > Timing).", wh);
      }
      if (has(params, "onClick") && t.advanceClick !== truthy(params.onClick)) return fail(truthy(params.onClick) ? L + " won't advance on a click. Tick On Mouse Click." : L + " still advances on a click. Clear On Mouse Click.", wh);
      if (has(params, "sound") && t.sound !== truthy(params.sound)) return fail(truthy(params.sound) ? L + " has no transition sound yet (Transitions > Sound)." : L + " still has a transition sound.", wh);
    }
    return pass(x.list.length > 1 ? "The transitions are set on every slide." : "The transition is set.");
  };

  // ---------- animations ----------
  /** The animations on a slide that match target (says= / kind= / name=), class and so on. */
  function animsFor(model, s, params) {
    let list = s.anims.filter((a) => a.cls !== "mediacall");
    if (has(params, "says") || has(params, "kind") || has(params, "name")) {
      const targets = shapesFor(model, params, [s]).map((h) => h.shape.id);
      list = list.filter((a) => targets.includes(a.target));
    }
    if (has(params, "class")) list = list.filter((a) => a.cls === (CLASS_PARAM[params.class] || params.class));
    return list;
  }
  CHECKS["anim.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const s = x.list[0];
    const wh = whereSlide(s);
    const list = animsFor(model, s, params);
    const what = (has(params, "class") ? CLASS_WORDS[CLASS_PARAM[params.class] || params.class] + " animation" : "an animation") + (has(params, "says") ? ' on "' + snippet(params.says, 25) + '"' : has(params, "kind") ? " on the " + (KIND_WORDS[params.kind] || "object") : "");
    if (!list.length) {
      const other = animsFor(model, s, Object.assign({}, params, { class: undefined }));
      if (has(params, "class") && other.length) return fail("There's " + CLASS_WORDS[other[0].cls] + " animation there, but this step needs " + what + ". Use Animations > Add Animation to add another.", wh);
      return fail("I don't see " + what + " on " + slideLabel(s) + " yet (Animations tab).", wh);
    }
    const a = list[0];
    if (has(params, "effect")) {
      const name = (EFFECTS[a.cls] || {})[a.presetId] || "";
      if (!list.some((z) => loose((EFFECTS[z.cls] || {})[z.presetId] || "") === loose(params.effect))) return fail("The animation is " + (name || "another effect") + ". It should be " + params.effect + ".", wh);
    }
    if (has(params, "byParagraph") && truthy(params.byParagraph)) {
      const target = a.target;
      const paragraphs = new Set(s.anims.filter((z) => z.target === target && z.paragraph !== null).map((z) => z.paragraph));
      if (paragraphs.size < 2 && s.builds[target] !== "p") return fail("The whole list appears at once. Set Effect Options > Sequence > By Paragraph.", wh);
    }
    if (has(params, "start") && !list.some((z) => z.start === params.start)) return fail("The animation starts " + (START_WORDS[a.start] || a.start) + ". Set Start to " + START_WORDS[params.start] + " (Animations > Timing).", wh);
    if (has(params, "minDelay") && !list.some((z) => z.delayMs >= toMs(params.minDelay))) return fail("The animation has " + (a.delayMs ? "a " + seconds(a.delayMs) + " delay" : "no delay") + ". Set Delay to at least " + seconds(toMs(params.minDelay)) + ".", wh);
    if (has(params, "duration")) {
      const want = toMs(params.duration);
      if (!list.some((z) => z.durationMs !== null && Math.abs(z.durationMs - want) <= 50)) return fail("The animation lasts " + (a.durationMs ? seconds(a.durationMs) : "the default time") + ". Set Duration to " + seconds(want) + ".", wh);
    }
    if (has(params, "first") && truthy(params.first)) {
      const firstAnim = s.anims.filter((z) => z.cls !== "mediacall")[0];
      if (!firstAnim || !list.includes(firstAnim)) return fail("That animation isn't first yet. In the Animation Pane, drag it to the top (or use Move Earlier).", wh);
    }
    if (has(params, "pathDir")) {
      const p = list.find((z) => z.path !== null);
      if (!p) return fail("The animation isn't a motion path yet. Use Animations > Motion Paths.", wh);
      if (p.pathDir !== params.pathDir) return fail("The motion path moves " + (p.pathDir || "nowhere") + ". It should move " + params.pathDir + " (Effect Options > Direction, or Reverse Path Direction).", wh);
    }
    return pass("The animation is set up.");
  };
  CHECKS["anim.count"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const n = x.list.reduce((sum, s) => sum + animsFor(model, s, params).length, 0);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (has(params, "min") && n < Number(params.min)) return fail(n ? "I see " + count(n, "animation") + where + "; there should be at least " + params.min + "." : "I don't see any animations" + where + " yet.", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    if (has(params, "max") && n > Number(params.max)) return fail("I see " + count(n, "animation") + where + "; there should be no more than " + params.max + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    return pass("The animations are there.");
  };

  // ---------- links and Zoom ----------
  CHECKS["link.present"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const links = [];
    for (const s of x.list) for (const o of s.shapes) for (const l of o.links) links.push({ s, o, l });
    let hits = links;
    if (has(params, "url")) hits = hits.filter((h) => loose(h.l.url).includes(loose(params.url)));
    if (has(params, "toSlide")) {
      const target = model.slides.find((s) => loose(s.title).includes(loose(params.toSlide)));
      hits = hits.filter((h) => (target && h.l.slide === target.number) || (/hlinkshowjump/.test(h.l.action) && /first/i.test(params.toSlide) && /firstslide/.test(h.l.action)));
    }
    if (has(params, "says")) hits = hits.filter((h) => loose(h.l.text || h.o.text).includes(loose(params.says)));
    if (has(params, "button") && truthy(params.button)) hits = hits.filter((h) => h.o.kind === "actionButton" || h.l.onShape);
    const where = x.list.length === 1 ? " on " + slideLabel(x.list[0]) : "";
    if (!hits.length) {
      const what = has(params, "button") ? "an action button" : "a link";
      return fail("I don't see " + what + (has(params, "url") ? " to " + params.url : has(params, "toSlide") ? ' to the "' + params.toSlide + '" slide' : "") + (has(params, "says") ? ' on "' + params.says + '"' : "") + where + " yet. " + (has(params, "button") ? "Use Insert > Shapes > Action Buttons." : "Select the text, then Insert > Link."), x.list.length === 1 ? whereSlide(x.list[0]) : null);
    }
    return pass("The link is there.");
  };
  CHECKS["zoom.is"] = function (model, params) {
    const x = slidesFor(model, params);
    if (x.error) return fail(x.error);
    const zooms = shapesFor(model, { kind: "zoom" }, x.list).filter((h) => !has(params, "kind") || h.shape.zoom.kind === params.kind);
    const word = { summary: "Summary Zoom", section: "Section Zoom", slide: "Slide Zoom" }[params.kind] || "Zoom";
    if (!zooms.length) return fail("I don't see a " + word + " yet. Use Insert > Zoom > " + word + ".", x.list.length === 1 ? whereSlide(x.list[0]) : null);
    if (has(params, "returnToZoom") && truthy(params.returnToZoom) && !zooms.some((h) => h.shape.zoom.returnToParent)) return fail("Tick Return to Zoom (Zoom tab), so the show comes back after each part.", whereSlide(zooms[0].slide, zooms[0].shape));
    return pass("The " + word + " is there.");
  };

  // ---------- slide show settings ----------
  CHECKS["show.is"] = function (model, params) {
    const sh = model.show;
    if (has(params, "loop") && sh.loop !== truthy(params.loop)) return fail(truthy(params.loop) ? "The show doesn't loop yet. Tick Loop continuously until 'Esc' in Slide Show > Set Up Slide Show." : "The show still loops.");
    if (has(params, "kiosk") && sh.kiosk !== truthy(params.kiosk)) return fail(truthy(params.kiosk) ? "Choose Browsed at a kiosk (full screen) in Slide Show > Set Up Slide Show." : "The show is still set to run at a kiosk.");
    if (has(params, "narration") && sh.narration !== truthy(params.narration)) return fail(truthy(params.narration) ? "Narration is turned off in Set Up Slide Show." : "Tick Show without narration in Set Up Slide Show.");
    if (has(params, "useTimings") && sh.useTimings !== truthy(params.useTimings)) return fail(truthy(params.useTimings) ? "The show is set to advance manually. In Set Up Slide Show, choose Using timings, if present." : "The show still uses timings.");
    if (has(params, "customShow")) {
      const cs = model.customShows.find((c) => loose(c.name) === loose(params.customShow)) || model.customShows.find((c) => loose(c.name).includes(loose(params.customShow)));
      if (!cs) return fail(model.customShows.length ? 'I don\'t see a custom show named "' + params.customShow + '". The custom shows are: ' + model.customShows.map((c) => c.name).join(", ") + "." : "There are no custom shows yet. Use Slide Show > Custom Slide Show > Custom Shows > New.");
      if (has(params, "slides") && cs.slides.length !== Number(params.slides)) return fail('The "' + cs.name + '" custom show has ' + count(cs.slides.length, "slide") + "; it should have " + params.slides + ".");
      if (has(params, "includes")) {
        const titles = cs.slides.map((n) => model.slides[n - 1] ? model.slides[n - 1].title : "");
        const missing = listOf(params.includes).filter((t) => !titles.some((x) => loose(x).includes(loose(t))));
        if (missing.length) return fail('The "' + cs.name + '" custom show doesn\'t include "' + missing[0] + '".');
      }
      if (has(params, "excludes")) {
        const titles = cs.slides.map((n) => model.slides[n - 1] ? model.slides[n - 1].title : "");
        const extra = listOf(params.excludes).filter((t) => titles.some((x) => loose(x).includes(loose(t))));
        if (extra.length) return fail('The "' + cs.name + '" custom show still includes "' + extra[0] + '".');
      }
    }
    if (has(params, "timings") && truthy(params.timings)) {
      const missing = model.slides.filter((s) => !s.hidden && !(s.transition && s.transition.advanceAfterMs !== null && s.transition.advanceAfterMs > 0));
      if (missing.length) return fail(slideLabel(missing[0]).replace(/^s/, "S") + " has no timing yet. Use Slide Show > Rehearse Timings, and keep the timings when you finish.", whereSlide(missing[0]));
    }
    return pass("The slide show is set up.");
  };

  // ---------- comments ----------
  function commentsFor(model, params) {
    const x = slidesFor(model, params);
    if (x.error) return { error: x.error };
    const numbers = x.list.map((s) => s.number);
    let list = model.comments.filter((c) => numbers.includes(c.slide));
    if (has(params, "says")) list = list.filter((c) => loose(c.text).includes(loose(params.says)) || c.replies.some((r) => loose(r.text).includes(loose(params.says))));
    if (has(params, "author")) list = list.filter((c) => loose(c.author).includes(loose(params.author)));
    return { list: list, slides: x.list };
  }
  CHECKS["comment.is"] = function (model, params) {
    const x = commentsFor(model, params);
    if (x.error) return fail(x.error);
    const desc = has(params, "says") ? 'the comment "' + params.says + '"' : "a comment" + (x.slides.length === 1 ? " on " + slideLabel(x.slides[0]) : "");
    const state = params.state || "present";
    if (state === "absent") return x.list.length ? fail(desc.charAt(0).toUpperCase() + desc.slice(1) + " is still there, on slide " + x.list[0].slide + ". Delete it.", { slide: x.list[0].slide }) : pass("That comment is gone.");
    if (!x.list.length) return fail("I don't see " + desc + " yet. Use Review > New Comment.", x.slides.length === 1 ? whereSlide(x.slides[0]) : null);
    const c = x.list[0];
    if (state === "replied" && !c.replies.length) return fail("Nobody has replied to " + desc + " yet.", { slide: c.slide });
    if (state === "resolved" && !c.resolved) return fail(desc.charAt(0).toUpperCase() + desc.slice(1) + " isn't resolved yet. Use the ... menu on the comment > Resolve thread.", { slide: c.slide });
    return pass(state === "replied" ? "There's a reply." : state === "resolved" ? "That thread is resolved." : "The comment is there.");
  };
  CHECKS["comment.count"] = function (model, params) {
    const x = commentsFor(model, params);
    if (x.error) return fail(x.error);
    const n = x.list.length;
    if (has(params, "min") && n < Number(params.min)) return fail("I see " + count(n, "comment") + "; there should be at least " + params.min + ".");
    if (has(params, "max") && n > Number(params.max)) return fail(Number(params.max) === 0 ? "There's still " + (n === 1 ? "a comment" : n + " comments") + " (on slide " + x.list[0].slide + ")." : "There are " + n + " comments; there should be no more than " + params.max + ".", { slide: x.list[0].slide });
    return pass(Number(params.max) === 0 ? "No comments are left." : "The comments are there.");
  };

  // ---------- protecting and preparing to share ----------
  CHECKS["doc.is"] = function (model, params) {
    if (has(params, "personalInfo")) {
      const removed = model.pres.removePersonalInfo || (!model.props.author.trim() && !model.props.lastModifiedBy.trim());
      if (params.personalInfo === "removed" && !removed) return fail("The presentation still carries personal information" + (model.props.author ? " (Author: " + snippet(model.props.author) + ")" : "") + ". Run Inspect Document and remove Document Properties and Personal Information.");
    }
    if (has(params, "fontsEmbedded") && truthy(params.fontsEmbedded) && !model.pres.embedFonts) return fail("The fonts aren't embedded yet. Use File > Options > Save > Embed fonts in the file.");
    if (has(params, "final") && truthy(params.final) && String(model.custom._MarkAsFinal || "").toLowerCase() !== "true") return fail("The presentation isn't marked as final yet. Use File > Info > Protect Presentation > Mark as Final.");
    if (has(params, "modifyPassword") && truthy(params.modifyPassword) && !model.pres.modifyPassword) return fail("There's no password to modify yet. Use File > Save As > Tools > General Options > Password to modify.");
    if (has(params, "notes") && params.notes === "none") {
      const s = model.slides.find((sl) => sl.notes.trim());
      if (s) return fail(slideLabel(s).replace(/^s/, "S") + " still has speaker notes.", whereSlide(s));
    }
    if (has(params, "offSlide") && params.offSlide === "none") {
      for (const s of model.slides) for (const o of s.shapes) {
        if (o.x === null || o.groupId) continue;
        if (o.x + o.w <= 0 || o.y + o.h <= 0 || o.x >= model.size.cx || o.y >= model.size.cy) return fail("There's still something off the edge of " + slideLabel(s) + ' ("' + (o.name || o.kind) + '"). Delete it, or use Inspect Document > Off-Slide Content.', whereSlide(s, o));
      }
    }
    return pass("The presentation is ready to share.");
  };
  const RULE_WORDS = { alt: "alt text", titles: "slide titles", uniqueTitles: "unique slide titles", tableHeaders: "table header rows", links: "link text", title: "the Title property", readingOrder: "reading order" };
  CHECKS["access.is"] = function (model, params) {
    const rules = has(params, "rules") ? listOf(params.rules) : ["alt", "titles", "uniqueTitles", "tableHeaders", "links"];
    const problems = [];
    let where = null;
    const needsAlt = ["picture", "icon", "chart", "smartart", "video", "model3d", "group", "shape", "audio"];
    if (rules.includes("alt")) {
      for (const s of model.slides) {
        const missing = s.shapes.filter((o) => needsAlt.includes(o.kind) && !o.alt && !o.decorative && !(o.kind === "shape" && o.text.trim()) && !(o.groupId && s.shapes.find((g) => g.id === o.groupId && (g.alt || g.decorative))));
        const realMissing = missing.filter((o) => !(o.kind === "shape" && !o.fill && !o.outline && !o.styleRefs));
        if (realMissing.length) { problems.push(count(realMissing.length, "graphic") + " on " + slideLabel(s) + " " + (realMissing.length === 1 ? "has" : "have") + " no alt text"); where = where || whereSlide(s, realMissing[0]); break; }
      }
    }
    if (rules.includes("titles")) {
      const s = model.slides.find((sl) => !sl.title.trim());
      if (s) { problems.push("slide " + s.number + " has no title (add one, even if you hide it off the slide)"); where = where || whereSlide(s); }
    }
    if (rules.includes("uniqueTitles")) {
      const seen = {};
      for (const s of model.slides) {
        const k = loose(s.title);
        if (!k) continue;
        if (seen[k]) { problems.push('slides ' + seen[k] + " and " + s.number + ' have the same title ("' + snippet(s.title, 30) + '")'); where = where || whereSlide(s); break; }
        seen[k] = s.number;
      }
    }
    if (rules.includes("tableHeaders")) {
      for (const s of model.slides) { const t = s.shapes.find((o) => o.kind === "table" && !o.table.firstRow); if (t) { problems.push("the table on " + slideLabel(s) + " has no header row"); where = where || whereSlide(s, t); break; } }
    }
    if (rules.includes("links")) {
      for (const s of model.slides) for (const o of s.shapes) for (const l of o.links) {
        if (l.url && l.text && /^(https?:\/\/|www\.)/i.test(l.text.trim())) { problems.push('a link on ' + slideLabel(s) + ' shows a web address ("' + snippet(l.text, 30) + '") instead of saying where it goes'); where = where || whereSlide(s, o); }
      }
    }
    if (rules.includes("title") && !model.props.title.trim()) problems.push("the presentation has no Title property (File > Info)");
    if (rules.includes("readingOrder")) {
      for (const s of model.slides) {
        const content = s.shapes.filter((o) => !o.groupId && !(o.ph && ["sldNum", "ftr", "dt"].includes(o.ph.type)));
        const ti = content.findIndex((o) => o.ph && (o.ph.type === "title" || o.ph.type === "ctrTitle"));
        if (ti > 0) { problems.push("on " + slideLabel(s) + ", the title isn't read first (fix it in the Selection Pane: the reading order is bottom to top)"); where = where || whereSlide(s); break; }
      }
    }
    if (problems.length) return fail("Accessibility: " + problems.slice(0, 2).join("; ") + ".", where);
    return pass("No accessibility problems found (" + rules.map((r) => RULE_WORDS[r] || r).join(", ") + ").");
  };

  // =====================================================================
  // PART 5: RUNNING THE CHECKS
  // =====================================================================

  /** Run one check. Never throws: a broken check fails with a message. */
  function runCheck(model, check, context) {
    const fn = CHECKS[check.name];
    if (!fn) return fail('This step uses a check the trainer doesn\'t know ("' + check.name + '"). Please tell your instructor.');
    try { return fn(model, check.params || {}, context || {}); }
    catch (e) {
      if (root.console) root.console.error("Check failed to run:", check, e);
      return fail("The trainer couldn't run this check. Please tell your instructor (" + check.name + ").");
    }
  }
  /** Run every check in a step. A step passes when all its checks pass. */
  function runStep(model, step, context) {
    const results = step.checks.map((c) => runCheck(model, c, context));
    const failed = results.filter((r) => !r.pass);
    if (!failed.length) return { pass: true, messages: Array.from(new Set(results.map((r) => r.message))), results: results };
    // Two checks can fail for the same reason; say it once.
    return { pass: false, messages: Array.from(new Set(failed.map((r) => r.message))), where: failed[0].where, results: results };
  }

  const api = { VERSION, readPptx, read: readPptx, runCheck, runStep, CHECKS: Object.keys(CHECKS), _internal: { imageSize, toEmu, toMs, readTransition } };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.PptChecker = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
