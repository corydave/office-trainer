/* =============================================================================
   word-checker.js - the Word checker library (v0.1)
   -----------------------------------------------------------------------------
   This is the ONLY code that decides whether a step passed. Lesson files name a
   check ("para.is", "text.absent" ...) and give it values; they can't add logic.

   How it works, in two steps:
     1. readDocx(bytes) opens a .docx file and builds a "model" of it: the
        paragraphs, their text, and where to look up their formatting.
     2. runCheck(model, check) runs one check from a lesson file against that
        model and returns { pass, message, where }.

   Why read the .docx file instead of asking Word through Office.js?
     - Word for the web, desktop Word and "upload a saved file" all give us the
       same .docx bytes, so every learner is checked by exactly the same code.
     - Office.js can't see some things on every platform (page setup on the
       web, text effects, proofing language). The .docx file holds all of them.
     - The same code can be tested in any browser with known-good and
       known-bad test documents (see check-tests.html).

   Principle: check the OUTCOME, not the clicks. A heading is bold whether the
   learner clicked B, pressed Ctrl+B or used a bold style.

   Needs JSZip (loaded by the page) to open the .docx zip file.
   Safety: DOMParser in "application/xml" mode never runs scripts and never
   fetches outside files. Text from the document is returned as plain strings;
   the page must show it with textContent, never innerHTML.
   ============================================================================= */
(function (root) {
  "use strict";

  const VERSION = "0.2.0";

  // XML namespaces. Word tags look like <w:p>; the "w" stands for this address.
  const NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
    dc: "http://purl.org/dc/elements/1.1/",
    cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  };

  // Safety limits for files we open (a .docx is a zip, and zips can "explode").
  const MAX_FILE_BYTES = 25 * 1024 * 1024;       // 25 MB as uploaded
  const MAX_UNZIPPED_BYTES = 80 * 1024 * 1024;   // 80 MB once unzipped

  // Word measures most lengths in "twips": twentieths of a point. 1 in = 1440 twips.
  const TWIPS_PER_PT = 20;

  // Paper sizes in twips (width x height, portrait).
  const PAPER = {
    letter: { w: 12240, h: 15840, name: "Letter", label: "Letter (8.5 x 11 in)" },
    legal: { w: 12240, h: 20160, name: "Legal", label: "Legal (8.5 x 14 in)" },
    a4: { w: 11906, h: 16838, name: "A4", label: "A4 (8.27 x 11.69 in)" },
  };

  // Everyday names for language tags, for friendlier messages.
  const LANGUAGE_NAMES = {
    "en-US": "English (United States)", "en-GB": "English (United Kingdom)", "en-CA": "English (Canada)",
    "en-AU": "English (Australia)", "es-ES": "Spanish (Spain)", "es-MX": "Spanish (Mexico)",
    "es-US": "Spanish (United States)", "fr-FR": "French (France)", "fr-CA": "French (Canada)",
    "de-DE": "German (Germany)", "pt-BR": "Portuguese (Brazil)",
  };

  // =====================================================================
  // PART 1: SMALL HELPERS
  // =====================================================================

  /** The first direct child of `element` with this Word tag name (for example "rPr"). */
  function child(element, tagName, ns) {
    if (!element) return null;
    const wanted = ns || NS.w;
    for (const c of element.children) {
      if (c.namespaceURI === wanted && c.localName === tagName) return c;
    }
    return null;
  }

  /** Every direct child with this tag name. */
  function children(element, tagName, ns) {
    if (!element) return [];
    const wanted = ns || NS.w;
    return Array.from(element.children).filter((c) => c.namespaceURI === wanted && c.localName === tagName);
  }

  /** Read a Word attribute such as w:val. Returns null when it isn't there. */
  function attr(element, name, ns) {
    if (!element) return null;
    const value = element.getAttributeNS(ns || NS.w, name);
    return value === "" && !element.hasAttributeNS(ns || NS.w, name) ? null : value;
  }

  /**
   * Word saves on/off settings like bold as <w:b/> (on), <w:b w:val="0"/> (off),
   * or leaves them out (not decided here). Returns true, false or null.
   */
  function onOff(element) {
    if (!element) return null;
    const val = attr(element, "val");
    if (val === null || val === "") return true;
    return !["0", "false", "off", "none"].includes(val.toLowerCase());
  }

  /** Turn XML text into a document we can search. Throws if the XML is damaged. */
  function parseXml(text) {
    const xml = new DOMParser().parseFromString(text, "application/xml");
    if (xml.getElementsByTagName("parsererror").length > 0) {
      throw new Error("Part of this Word file is damaged and could not be read.");
    }
    return xml;
  }

  /**
   * Make text easy to compare: one kind of space, one kind of quote mark.
   * Word turns ' into ’ and " into “ ” as you type, so we treat them as the same.
   */
  function plain(text) {
    return String(text)
      .replace(/[‘’‛′]/g, "'")
      .replace(/[“”‟″]/g, '"')
      .replace(/ /g, " ");
  }

  /** Lowercase, single spaces, trimmed: for finding paragraphs by their wording. */
  function loose(text) {
    return plain(text).replace(/\s+/g, " ").trim().toLowerCase();
  }

  /** Shorten long document text for a message: "Thank you for helping at..." */
  function snippet(text, max) {
    const t = String(text).replace(/\s+/g, " ").trim();
    const limit = max || 40;
    return t.length > limit ? t.slice(0, limit - 1).trimEnd() + "…" : t;
  }

  /** "1 paragraph" / "3 paragraphs" */
  function count(n, word, plural) {
    return n + " " + (n === 1 ? word : (plural || word + "s"));
  }

  /** Points to a friendly string: 72 -> "1 in", 6 -> "6 pt". */
  function showLength(points, unit) {
    const rounded = (n) => String(Math.round(n * 100) / 100);
    if (unit === "in" || (unit === undefined && Math.abs(points) >= 18)) return rounded(points / 72) + " in";
    if (unit === "cm") return rounded(points / 72 * 2.54) + " cm";
    return rounded(points) + " pt";
  }

  /** "1in" / "36pt" / "2.54cm" -> points (72 per inch). null if it isn't a length. */
  function toPoints(value) {
    const m = /^(-?\d+(?:\.\d+)?)\s*(in|pt|cm)$/.exec(String(value).trim());
    if (!m) return null;
    const n = Number(m[1]);
    return m[2] === "in" ? n * 72 : m[2] === "cm" ? n / 2.54 * 72 : n;
  }

  /** The unit a lesson used, so messages answer in the same unit ("0.75 in" for "1in"). */
  function unitOf(value) {
    const m = /(in|pt|cm)\s*$/.exec(String(value));
    return m ? m[1] : "pt";
  }

  /** A pass or a fail. `where` (optional) helps the add-in show the learner the spot. */
  function pass(message) { return { pass: true, message: message }; }
  function fail(message, where) { return { pass: false, message: message, where: where || null }; }

  /** Wildcard pattern ("*_W1_Letter.docx") to a regular expression. Not case sensitive. */
  function wildcard(pattern) {
    const escaped = String(pattern).replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
    return new RegExp("^" + escaped + "$", "i");
  }

  // =====================================================================
  // PART 2: OPENING A .DOCX FILE AND BUILDING THE MODEL
  // ---------------------------------------------------------------------
  // A .docx is a zip folder. The parts we read:
  //   word/document.xml   the text, plus formatting put directly on it
  //   word/styles.xml     styles (Normal, Heading 1 ...) that ALSO format text
  //   word/numbering.xml  bullets and numbering
  //   word/theme/theme1.xml  the theme fonts ("Aptos (Body)")
  //   word/settings.xml   the document's default editing language
  //   word/headerN.xml, word/footerN.xml   headers and footers
  //   docProps/core.xml   Title, Author (creator) and Tags (keywords)
  // =====================================================================

  /**
   * Open a .docx and build the model every check uses.
   *   bytes    a Uint8Array or ArrayBuffer with the file
   *   options  { fileName }  the saved file's name, for file.name
   */
  async function readDocx(bytes, options) {
    const opts = options || {};
    if (typeof root.JSZip === "undefined") throw new Error("The checker couldn't load. Check your internet connection and reopen this panel.");
    const size = bytes.byteLength !== undefined ? bytes.byteLength : bytes.length;
    if (size > MAX_FILE_BYTES) throw new Error("That file is over 25 MB, which is too big to be a practice document.");

    let zip;
    try {
      zip = await root.JSZip.loadAsync(bytes);
    } catch (e) {
      throw new Error("That doesn't look like a Word (.docx) file.");
    }

    // Refuse "zip bombs": files that unzip to something enormous.
    let unzipped = 0;
    zip.forEach(function (path, entry) {
      if (entry._data && entry._data.uncompressedSize) unzipped += entry._data.uncompressedSize;
    });
    if (unzipped > MAX_UNZIPPED_BYTES) throw new Error("That file is too big to check.");

    // Read one XML part, or null if the file doesn't have it.
    async function part(path) {
      const f = zip.file(path);
      return f ? parseXml(await f.async("string")) : null;
    }

    const documentXml = await part("word/document.xml");
    if (!documentXml) throw new Error("That doesn't look like a Word (.docx) file.");

    const model = {
      fileName: opts.fileName || "",
      paragraphs: [],        // body paragraphs, in order (including ones in tables)
      headers: [],           // paragraphs in headers
      footers: [],           // paragraphs in footers
      styles: {},            // styleId -> { id, name, type, basedOn, pPr, rPr }
      styleIdsByName: {},    // "heading 1" -> "Heading1"
      defaultParagraphStyle: null,
      defaultRunProps: null, // document-wide default text formatting
      defaultParaProps: null,
      themeFonts: { major: null, minor: null },
      theme: { colorsName: "", colors: {} },   // the theme's color set ("Office") and its colors
      defaultLanguage: null,
      settings: null,        // word/settings.xml (tracking, protection, mail merge, compatibility)
      sections: [],          // page setup for each section
      props: { title: "", author: "", keywords: "", lastModifiedBy: "" },
      nums: {}, abstractNums: {}, numOverrideEls: {},
      fields: [],            // every field: { code, part } (PAGE, TOC, MERGEFIELD Position ...)
      links: [],             // hyperlinks: { text, url, anchor }
      bookmarks: [],         // { name, paragraph }
      objects: [],           // pictures, shapes, text boxes, SmartArt, charts, 3D models
      tables: [],            // top-level tables
      notes: { footnote: [], endnote: [] },   // the text of each footnote / endnote
      noteRefs: { footnote: 0, endnote: 0 },  // how many note marks are in the text
      comments: [],          // { id, author, text, anchor, parent, done }
      revisions: [],         // tracked changes: { kind, author, text }
      sources: { count: 0, style: "", version: "" },   // the citation manager's source list
      background: { color: null, watermarks: [] },
      breaks: { page: 0, column: 0 },
      hiddenRuns: 0,         // runs formatted as hidden text
      placeholders: 0,       // content controls still showing their placeholder text
      headerXml: [],         // each header part, for watermarks
    };

    readStyles(model, await part("word/styles.xml"));
    readTheme(model, await part("word/theme/theme1.xml"));
    model.settings = await part("word/settings.xml");
    readSettings(model, model.settings);
    readNumbering(model, await part("word/numbering.xml"));
    readCoreProps(model, await part("docProps/core.xml"));

    // --- Body paragraphs ---
    const body = documentXml.getElementsByTagNameNS(NS.w, "body")[0];
    if (!body) throw new Error("That doesn't look like a Word (.docx) file.");
    collectParagraphs(body, "body", model.paragraphs);

    // --- Sections (page setup) ---
    // A section ends at a paragraph that holds a w:sectPr; the last one is the body's own.
    // Tracked changes keep an old copy inside w:sectPrChange; we skip those.
    const sectPrs = Array.from(body.getElementsByTagNameNS(NS.w, "sectPr"))
      .filter((e) => e.parentNode.localName === "body" || e.parentNode.localName === "pPr");
    for (const sectPr of sectPrs) model.sections.push(readSection(sectPr));
    if (!model.sections.length) model.sections.push(readSection(null));
    let sectionIndex = 0;
    for (const p of model.paragraphs) {
      p.section = Math.min(sectionIndex, model.sections.length - 1);
      if (child(child(p.element, "pPr"), "sectPr")) sectionIndex++;
    }

    // --- Relationships: where the document's other parts are ---
    const rels = await part("word/_rels/document.xml.rels");
    const targets = {};
    if (rels) {
      for (const r of rels.getElementsByTagNameNS(NS.rel, "Relationship")) {
        targets[r.getAttribute("Id")] = { target: r.getAttribute("Target"), external: r.getAttribute("TargetMode") === "External" };
      }
    }
    const partPath = (id) => targets[id] && !targets[id].external
      ? "word/" + targets[id].target.replace(/^\/?word\//, "").replace(/^\.\//, "") : null;

    // --- Headers and footers: follow each section's links to its header/footer files ---
    const seen = {};
    for (const sectPr of sectPrs) {
      for (const kind of ["header", "footer"]) {
        for (const ref of children(sectPr, kind + "Reference")) {
          // "default" is the normal header/footer; "first" and "even" are special pages.
          const path = partPath(attr(ref, "id", NS.r));
          if (!path || seen[path]) continue;
          seen[path] = true;
          const xml = await part(path);
          if (!xml) continue;
          collectParagraphs(xml.documentElement, kind, kind === "header" ? model.headers : model.footers);
          collectFields(xml.documentElement, kind, model.fields);
          if (kind === "header") model.headerXml.push(xml);
        }
      }
    }

    // --- Everything else the newer checks look at ---
    collectFields(body, "body", model.fields);
    collectLinks(body, targets, model);
    collectBookmarks(body, model);
    collectLists(model);
    collectTables(body, model);
    await collectObjects(body, model, partPath, part);
    collectBodyFacts(documentXml, body, model);
    collectRevisions(body, model);
    await collectNotes(model, part);
    await collectComments(body, model, part);
    await collectSources(zip, model);
    collectWatermarks(model);
    return model;
  }

  /** Styles: who is based on whom, and what each one sets. */
  function readStyles(model, xml) {
    if (!xml) return;
    for (const s of xml.getElementsByTagNameNS(NS.w, "style")) {
      const id = attr(s, "styleId");
      const style = {
        id: id,
        name: attr(child(s, "name"), "val") || id,
        type: attr(s, "type"),
        basedOn: attr(child(s, "basedOn"), "val"),
        pPr: child(s, "pPr"),
        rPr: child(s, "rPr"),
      };
      model.styles[id] = style;
      model.styleIdsByName[String(style.name).toLowerCase()] = id;
      // The paragraph style marked w:default="1" is used when a paragraph names none ("Normal").
      if (style.type === "paragraph" && ["1", "true", "on"].includes(String(attr(s, "default")).toLowerCase())) {
        model.defaultParagraphStyle = id;
      }
    }
    const docDefaults = xml.getElementsByTagNameNS(NS.w, "docDefaults")[0];
    if (docDefaults) {
      model.defaultRunProps = child(child(docDefaults, "rPrDefault"), "rPr");
      model.defaultParaProps = child(child(docDefaults, "pPrDefault"), "pPr");
    }
  }

  /** Theme fonts: what "Aptos (Body)" and "Aptos Display (Headings)" really are. */
  function readTheme(model, xml) {
    if (!xml) return;
    // <a:majorFont><a:latin typeface="Aptos Display"/> ... (theme attributes have no namespace)
    const typeface = (tag) => {
      const group = xml.getElementsByTagNameNS(NS.a, tag)[0];
      const latin = child(group, "latin", NS.a);
      return latin ? latin.getAttribute("typeface") : null;
    };
    model.themeFonts.major = typeface("majorFont");
    model.themeFonts.minor = typeface("minorFont");
    // The theme's color set: its name ("Office", or one the learner made) and each color as #RRGGBB.
    const scheme = xml.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    if (scheme) {
      model.theme.colorsName = scheme.getAttribute("name") || "";
      for (const c of scheme.children) {
        const color = c.firstElementChild;
        const hex = color && (color.getAttribute("val") || color.getAttribute("lastClr"));
        if (hex) model.theme.colors[c.localName] = "#" + hex.toUpperCase();
      }
    }
  }

  /** settings.xml: the document's own default language (used when nothing else says). */
  function readSettings(model, xml) {
    if (!xml) return;
    const lang = xml.getElementsByTagNameNS(NS.w, "themeFontLang")[0];
    model.defaultLanguage = attr(lang, "val");
  }

  /** Bullets and numbering definitions. */
  function readNumbering(model, xml) {
    if (!xml) return;
    for (const a of xml.getElementsByTagNameNS(NS.w, "abstractNum")) {
      const levels = {};
      for (const lvl of a.getElementsByTagNameNS(NS.w, "lvl")) levels[attr(lvl, "ilvl")] = lvl;
      model.abstractNums[attr(a, "abstractNumId")] = { levels: levels };
    }
    for (const n of xml.getElementsByTagNameNS(NS.w, "num")) {
      const overrides = {};
      for (const o of n.getElementsByTagNameNS(NS.w, "lvlOverride")) {
        const lvl = child(o, "lvl");
        if (lvl) overrides[attr(o, "ilvl")] = lvl;
      }
      model.nums[attr(n, "numId")] = { abstractId: attr(child(n, "abstractNumId"), "val"), overrides: overrides };
      model.numOverrideEls[attr(n, "numId")] = n.getElementsByTagNameNS(NS.w, "lvlOverride");
    }
  }

  /** File > Info properties: Title, Author and Tags. */
  function readCoreProps(model, xml) {
    if (!xml) return;
    const text = (ns, tag) => {
      const e = xml.getElementsByTagNameNS(ns, tag)[0];
      return e ? e.textContent.trim() : "";
    };
    model.props.title = text(NS.dc, "title");
    model.props.author = text(NS.dc, "creator");
    model.props.keywords = text(NS.cp, "keywords");
    model.props.lastModifiedBy = text(NS.cp, "lastModifiedBy");
  }

  /** One section's page size, orientation and margins (in twips). */
  function readSection(sectPr) {
    const pgSz = child(sectPr, "pgSz");
    const pgMar = child(sectPr, "pgMar");
    const num = (e, name, fallback) => {
      const v = Number(attr(e, name));
      return Number.isFinite(v) && attr(e, name) !== null ? v : fallback;
    };
    const cols = child(sectPr, "cols");
    const borders = child(sectPr, "pgBorders");
    return {
      type: attr(child(sectPr, "type"), "val") || "nextPage",   // how this section starts
      columns: Number(attr(cols, "num")) || 1,
      firstPageDifferent: onOff(child(sectPr, "titlePg")) === true,
      border: !!borders && Array.from(borders.children).some((b) => (attr(b, "val") || "none") !== "none" && attr(b, "val") !== "nil"),
      footnoteFormat: attr(child(child(sectPr, "footnotePr"), "numFmt"), "val"),
      endnoteFormat: attr(child(child(sectPr, "endnotePr"), "numFmt"), "val"),
      width: num(pgSz, "w", 12240),           // 12240 twips = 8.5 in
      height: num(pgSz, "h", 15840),          // 15840 twips = 11 in
      top: Math.abs(num(pgMar, "top", 1440)), // Word allows "negative" top margins; the size is what matters
      bottom: Math.abs(num(pgMar, "bottom", 1440)),
      left: num(pgMar, "left", num(pgMar, "start", 1440)),
      right: num(pgMar, "right", num(pgMar, "end", 1440)),
    };
  }

  /**
   * Find every paragraph inside `container` and add it to `list`.
   * Paragraphs inside text boxes are skipped (they'd be counted twice).
   */
  function collectParagraphs(container, partName, list) {
    for (const p of container.getElementsByTagNameNS(NS.w, "p")) {
      if (insideTag(p, container, ["txbxContent"])) continue;
      const runs = getRuns(p);
      list.push({
        element: p,
        part: partName,
        runs: runs,
        text: runs.map((r) => r.text).join(""),
        inTable: insideTag(p, container, ["tc"]),
        index: list.length,
      });
    }
  }

  /** True if `node` sits inside a Word element with one of these tag names (below `stopAt`). */
  function insideTag(node, stopAt, tagNames) {
    let n = node.parentNode;
    while (n && n !== stopAt) {
      if (n.namespaceURI === NS.w && tagNames.includes(n.localName)) return true;
      n = n.parentNode;
    }
    return false;
  }

  // Symbols inserted with a symbol font (<w:sym w:font="Symbol" w:char="F0D3"/>).
  const SYMBOL_CHARS = { D3: "©", D2: "®", D4: "™", E3: "©", E2: "®", E4: "™" };

  /**
   * A paragraph's text is split into "runs": stretches of text with the same
   * formatting. We skip deleted text (tracked changes) and anything that
   * belongs to a nested paragraph (text boxes).
   */
  function getRuns(paragraph) {
    const runs = [];
    for (const r of paragraph.getElementsByTagNameNS(NS.w, "r")) {
      if (!belongsTo(r, paragraph)) continue;
      let text = "";
      for (const c of r.children) {
        if (c.namespaceURI !== NS.w) continue;
        if (c.localName === "t") text += c.textContent;
        else if (c.localName === "tab") text += "\t";
        else if (c.localName === "br" || c.localName === "cr") text += attr(c, "type") === "page" ? "" : "\n";
        else if (c.localName === "noBreakHyphen") text += "-";
        else if (c.localName === "sym") {
          const code = String(attr(c, "char") || "").toUpperCase().slice(-2);
          text += SYMBOL_CHARS[code] || "";
        }
      }
      runs.push({ element: r, text: text });
    }
    return runs;
  }

  /** True if this run really is part of this paragraph (not deleted, not in a nested paragraph). */
  function belongsTo(run, paragraph) {
    let n = run.parentNode;
    while (n && n !== paragraph) {
      if (n.namespaceURI === NS.w && ["del", "moveFrom", "p"].includes(n.localName)) return false;
      n = n.parentNode;
    }
    return n === paragraph;
  }

  /**
   * Is this paragraph "empty" - just a paragraph mark used to make space?
   * Spaces and tabs don't count as content. Pictures, page breaks and section
   * breaks do: those paragraphs are doing a job.
   */
  function isEmptyParagraph(p) {
    if (p.text.replace(/[\s ]/g, "") !== "") return false;
    const e = p.element;
    if (e.getElementsByTagNameNS(NS.w, "drawing").length || e.getElementsByTagNameNS(NS.w, "pict").length ||
        e.getElementsByTagNameNS(NS.w, "object").length) return false;
    for (const br of e.getElementsByTagNameNS(NS.w, "br")) if (attr(br, "type") === "page") return false;
    const pPr = child(e, "pPr");
    if (child(pPr, "sectPr") || onOff(child(pPr, "pageBreakBefore"))) return false;
    return true;
  }

  // =====================================================================
  // PART 2b: THE REST OF THE DOCUMENT
  // ---------------------------------------------------------------------
  // Fields, links, bookmarks, lists, tables, pictures and shapes, notes,
  // comments, tracked changes and citations. Each collect...() function
  // reads one kind of thing into the model, so the checks in PART 5 only
  // have to look things up.
  // =====================================================================

  // More namespaces used below.
  const NS2 = {
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    dgm: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    a14: "http://schemas.microsoft.com/office/drawing/2010/main",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    v: "urn:schemas-microsoft-com:vml",
    b: "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
  };
  const GRAPHIC = {
    "http://schemas.openxmlformats.org/drawingml/2006/picture": "picture",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape": "shape",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup": "group",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram": "smartart",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "chart",
    "http://schemas.microsoft.com/office/drawing/2017/model3d": "model3d",
  };
  const DECORATIVE_EXT = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
  const SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

  /**
   * True if `node` is inside an mc:Fallback. Word saves newer things (shapes,
   * 3D models) twice: the real one, and an older "fallback" copy for old
   * versions of Word. We only count the real one.
   */
  function inFallback(node) {
    for (let n = node.parentNode; n; n = n.parentNode) {
      if (n.namespaceURI === NS2.mc && n.localName === "Fallback") return true;
    }
    return false;
  }

  /** Every element in a part, in document order, that isn't in a fallback copy. */
  function walk(container, visit) {
    const all = container.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
      if (!inFallback(all[i])) visit(all[i]);
    }
  }

  /**
   * Fields: Word's automatic pieces (PAGE, TOC, MERGEFIELD Position, SEQ
   * Figure, CITATION...). A field is saved either as <w:fldSimple w:instr="...">
   * or as a "complex field": a begin mark, instruction text, a separate mark,
   * the shown result, and an end mark. Complex fields can nest (IF holding a
   * MERGEFIELD), so we keep a stack.
   */
  function collectFields(container, partName, list) {
    const stack = [];
    walk(container, function (e) {
      if (e.namespaceURI !== NS.w) return;
      if (e.localName === "fldSimple") list.push({ code: tidyCode(attr(e, "instr")), part: partName });
      else if (e.localName === "fldChar") {
        const type = attr(e, "fldCharType");
        if (type === "begin") stack.push({ code: "", part: partName });
        else if (type === "end" && stack.length) { const f = stack.pop(); f.code = tidyCode(f.code); list.push(f); }
      } else if (e.localName === "instrText" && stack.length) stack[stack.length - 1].code += e.textContent;
    });
  }
  function tidyCode(code) { return String(code || "").replace(/\s+/g, " ").trim(); }

  /** Hyperlinks: their shown text, and where they go (a web address, or a bookmark). */
  function collectLinks(body, targets, model) {
    for (const h of body.getElementsByTagNameNS(NS.w, "hyperlink")) {
      if (inFallback(h)) continue;
      let text = "";
      for (const t of h.getElementsByTagNameNS(NS.w, "t")) text += t.textContent;
      const id = attr(h, "id", NS.r);
      model.links.push({ text: text, url: id && targets[id] ? targets[id].target : "", anchor: attr(h, "anchor") || "" });
    }
    // Links made as HYPERLINK fields: HYPERLINK "https://..." or HYPERLINK \l "Methods"
    model.fields.filter((f) => /^HYPERLINK\b/i.test(f.code)).forEach(function (f) {
      const anchor = /\\l\s+"([^"]+)"/.exec(f.code);
      const url = /^HYPERLINK\s+"([^"]+)"/i.exec(f.code);
      model.links.push({ text: "", url: url ? url[1] : "", anchor: anchor ? anchor[1] : "", fromField: true });
    });
  }

  /** Bookmarks the learner made (Word's own hidden ones start with "_"). */
  function collectBookmarks(body, model) {
    for (const b of body.getElementsByTagNameNS(NS.w, "bookmarkStart")) {
      const name = attr(b, "name") || "";
      if (!name || name.startsWith("_")) continue;
      // The paragraph it sits in, or (if it sits between paragraphs) the next one.
      let p = b.parentNode;
      while (p && !(p.namespaceURI === NS.w && p.localName === "p")) p = p.parentNode;
      if (!p) { p = b.nextElementSibling; while (p && p.localName !== "p") p = p.nextElementSibling; }
      model.bookmarks.push({ name: name, paragraph: model.paragraphs.find((x) => x.element === p) || null });
    }
  }

  /**
   * Lists: for every paragraph in a list, its kind, level, number format,
   * bullet symbol and the number Word shows. Word keeps one counter per
   * list design (abstractNum); "Restart at 1" or "Set numbering value" makes a
   * new list instance with a startOverride.
   */
  function collectLists(model) {
    const counters = {};   // abstractNum id -> [count at level 0, level 1, ...]
    const started = {};    // "numId:level" -> true once its startOverride has been applied
    for (const p of model.paragraphs) {
      let numPr = child(child(p.element, "pPr"), "numPr");
      if (!numPr) {
        for (const s of styleChain(model, paragraphStyleId(model, p))) { numPr = child(s.pPr, "numPr"); if (numPr) break; }
      }
      if (!numPr) continue;
      const numId = attr(child(numPr, "numId"), "val");
      const level = Number(attr(child(numPr, "ilvl"), "val") || 0);
      const num = model.nums[numId];
      if (!numId || numId === "0" || !num) continue;
      const abstract = model.abstractNums[num.abstractId] || { levels: {} };
      const lvl = num.overrides[level] || abstract.levels[level];
      if (!lvl) continue;
      const format = attr(child(lvl, "numFmt"), "val") || "decimal";
      if (format === "none") continue;
      const startAt = Number(attr(child(lvl, "start"), "val") || 1);
      const c = counters[num.abstractId] || (counters[num.abstractId] = []);
      // A restart or "set value": this list instance starts at its own number.
      const override = child(Array.from(model.numOverrideEls[numId] || []).find((o) => attr(o, "ilvl") === String(level)) || null, "startOverride");
      const key = numId + ":" + level;
      if (override && !started[key]) { c[level] = Number(attr(override, "val")) - 1; started[key] = true; }
      if (c[level] === undefined) c[level] = startAt - 1;
      c[level]++;
      c.length = level + 1; // deeper levels start over under a new item
      const lvlText = attr(child(lvl, "lvlText"), "val") || "";
      const font = attr(child(child(lvl, "rPr"), "rFonts"), "ascii") || "";
      p.list = { kind: format === "bullet" ? "bullet" : "number", level: level + 1, format: format,
        pattern: lvlText, symbol: format === "bullet" ? bulletSymbol(lvlText, font) : "", number: c[level], numId: numId };
    }
  }

  /** Bullet characters saved in symbol fonts, as the symbols people see. */
  const BULLET_FONTS = {
    symbol: { "": "•", "·": "•", "": "§" },
    wingdings: { "": "✓", "ü": "✓", "": "▪", "§": "▪", "": "➢", "Ø": "➢", "": "❖", "v": "❖", "": "❑", "q": "❑", "": "■", "n": "■", "": "➔", "à": "➔" },
  };
  function bulletSymbol(text, font) {
    const map = BULLET_FONTS[String(font).toLowerCase()];
    if (map && map[text] !== undefined) return map[text];
    if (text === "o" && /courier/i.test(font)) return "◦";
    return text;
  }

  /** Top-level tables (not tables inside tables or text boxes). */
  function collectTables(body, model) {
    for (const tbl of body.getElementsByTagNameNS(NS.w, "tbl")) {
      if (insideTag(tbl, body, ["tbl", "txbxContent"]) || inFallback(tbl)) continue;
      const tblPr = child(tbl, "tblPr");
      const grid = children(child(tbl, "tblGrid"), "gridCol").map((g) => Number(attr(g, "w")) || 0);
      const rows = children(tbl, "tr").map(function (tr) {
        const cells = children(tr, "tc").map(function (tc) {
          const tcPr = child(tc, "tcPr");
          let text = "";
          for (const p of children(tc, "p")) {
            if (text) text += "\n";
            for (const t of p.getElementsByTagNameNS(NS.w, "t")) text += t.textContent;
          }
          return { text: text, span: Number(attr(child(tcPr, "gridSpan"), "val")) || 1,
            vMerge: !!child(tcPr, "vMerge"), width: Number(attr(child(tcPr, "tcW"), "w")) || 0, widthType: attr(child(tcPr, "tcW"), "type") };
        });
        return { cells: cells, header: onOff(child(child(tr, "trPr"), "tblHeader")) === true };
      });
      const mar = child(tblPr, "tblCellMar");
      const marW = (side) => { const e = child(mar, side) || child(mar, side === "left" ? "start" : "end"); return e ? Number(attr(e, "w")) || 0 : null; };
      model.tables.push({
        element: tbl, rows: rows, grid: grid,
        cols: Math.max(grid.length, ...rows.map((r) => r.cells.reduce((n, c) => n + c.span, 0))),
        text: rows.map((r) => r.cells.map((c) => c.text).join("\t")).join("\n"),
        cellMarginLeft: marW("left"), cellMarginRight: marW("right"),
        cellSpacing: child(tblPr, "tblCellSpacing") ? Number(attr(child(tblPr, "tblCellSpacing"), "w")) || 0 : 0,
      });
    }
  }

  /**
   * Pictures, icons, shapes, text boxes, SmartArt, charts and 3D models.
   * Each is a <w:drawing> holding either wp:inline (in line with text) or
   * wp:anchor (floating, with a wrap style and a position).
   */
  async function collectObjects(body, model, partPath, part) {
    const drawings = Array.from(body.getElementsByTagNameNS(NS.w, "drawing")).filter((d) => !inFallback(d));
    for (const d of drawings) {
      const holder = child(d, "inline", NS2.wp) || child(d, "anchor", NS2.wp);
      if (!holder) continue;
      const docPr = child(holder, "docPr", NS2.wp);
      const graphicData = holder.getElementsByTagNameNS(NS.a, "graphicData")[0];
      const uri = graphicData ? graphicData.getAttribute("uri") : "";
      const o = {
        kind: GRAPHIC[uri] || "other",
        name: docPr ? docPr.getAttribute("name") || "" : "",
        alt: docPr ? (docPr.getAttribute("descr") || "").trim() : "",
        decorative: false,
        wrap: "inline", halign: null, valign: null,
        text: "", style: false, artistic: false, bgRemoved: false, fill: null, outline: null,
        layout: "", colorsChanged: false, styleChanged: false, items: 0, chartType: "",
      };
      if (docPr) {
        for (const ext of docPr.getElementsByTagNameNS(NS.a, "ext")) {
          if (ext.getAttribute("uri") === DECORATIVE_EXT) {
            const dec = ext.firstElementChild;
            o.decorative = !!dec && ["1", "true"].includes(String(dec.getAttribute("val")).toLowerCase());
          }
        }
      }
      // Position and wrapping (floating objects only)
      if (holder.localName === "anchor") {
        const wraps = { wrapSquare: "square", wrapTight: "tight", wrapThrough: "through", wrapTopAndBottom: "topAndBottom" };
        o.wrap = "inFront";
        for (const c of holder.children) {
          if (wraps[c.localName]) o.wrap = wraps[c.localName];
          if (c.localName === "wrapNone") o.wrap = holder.getAttribute("behindDoc") === "1" ? "behind" : "inFront";
        }
        const alignOf = (tag) => { const pos = child(holder, tag, NS2.wp); const al = child(pos, "align", NS2.wp); return al ? al.textContent.trim() : null; };
        o.halign = alignOf("positionH");
        o.valign = alignOf("positionV");
        if (o.valign === "center") o.valign = "middle";
      }
      if (o.kind === "picture") readPicture(o, graphicData);
      if (o.kind === "shape") readShape(o, graphicData);
      if (o.kind === "smartart") await readSmartArt(o, graphicData, partPath, part);
      if (o.kind === "chart") await readChart(o, graphicData, partPath, part);
      model.objects.push(o);
    }
  }

  function readPicture(o, graphicData) {
    const blip = graphicData.getElementsByTagNameNS(NS.a, "blip")[0];
    if (blip) {
      for (const ext of blip.getElementsByTagNameNS(NS.a, "ext")) if (ext.getAttribute("uri") === SVG_EXT) o.kind = "icon";
      o.artistic = Array.from(blip.getElementsByTagNameNS(NS2.a14, "imgEffect")).some((e) => Array.from(e.children).some((c) => /^artistic/.test(c.localName)));
      o.bgRemoved = blip.getElementsByTagNameNS(NS2.a14, "backgroundRemoval").length > 0;
    }
    // A picture style shows up as effects, a frame line, a 3-D look or a non-rectangle shape.
    const spPr = graphicData.getElementsByTagNameNS(NS2.pic, "spPr")[0];
    if (spPr) {
      const effects = child(spPr, "effectLst", NS.a);
      const line = child(spPr, "ln", NS.a);
      const geom = child(spPr, "prstGeom", NS.a);
      o.style = (!!effects && effects.children.length > 0) || (!!line && !child(line, "noFill", NS.a) && line.children.length > 0) ||
        !!child(spPr, "scene3d", NS.a) || !!child(spPr, "sp3d", NS.a) || (!!geom && geom.getAttribute("prst") !== "rect");
    }
  }

  function readShape(o, graphicData) {
    const wsp = graphicData.getElementsByTagNameNS(NS2.wps, "wsp")[0];
    if (!wsp) return;
    const cNv = child(wsp, "cNvSpPr", NS2.wps);
    if (cNv && cNv.getAttribute("txBox") === "1") o.kind = "textbox";
    const txbx = wsp.getElementsByTagNameNS(NS.w, "txbxContent")[0];
    if (txbx) o.text = Array.from(txbx.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join("");
    const spPr = child(wsp, "spPr", NS2.wps);
    const colorOf = (holder) => {
      if (!holder) return null;
      if (child(holder, "noFill", NS.a)) return "none";
      const solid = child(holder, "solidFill", NS.a);
      if (!solid) return child(holder, "gradFill", NS.a) || child(holder, "pattFill", NS.a) ? "changed" : null;
      const clr = solid.firstElementChild;
      if (!clr) return "changed";
      return clr.localName === "srgbClr" ? "#" + String(clr.getAttribute("val")).toUpperCase() : "theme:" + clr.getAttribute("val");
    };
    o.fill = colorOf(spPr);                      // null = the shape style's own fill
    o.outline = colorOf(child(spPr, "ln", NS.a));
  }

  async function readSmartArt(o, graphicData, partPath, part) {
    const relIds = graphicData.getElementsByTagNameNS(NS2.dgm, "relIds")[0];
    if (!relIds) return;
    const get = async (name) => { const path = partPath(attr(relIds, name, NS.r)); return path ? part(path) : null; };
    const data = await get("dm"), layout = await get("lo"), colors = await get("cs"), style = await get("qs");
    if (data) {
      // Each text "point" (pt) is one item, such as a step in a process.
      const points = Array.from(data.getElementsByTagNameNS(NS2.dgm, "pt")).filter((pt) => !pt.getAttribute("type") || pt.getAttribute("type") === "node");
      const texts = points.map((pt) => Array.from(pt.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join("").trim()).filter(Boolean);
      o.items = texts.length;
      o.text = texts.join(" | ");
    }
    const uid = (xml) => xml && xml.documentElement ? xml.documentElement.getAttribute("uniqueId") || "" : "";
    o.layout = uid(layout);
    // Word's starting choices; anything else means the learner changed them.
    const c = uid(colors), q = uid(style);
    o.colorsChanged = !!c && !/colors\/accent1_2$/.test(c);
    o.styleChanged = !!q && !/quickstyle\/simple1$/.test(q);
  }

  async function readChart(o, graphicData, partPath, part) {
    const ref = graphicData.getElementsByTagNameNS(NS2.c, "chart")[0];
    const path = ref ? partPath(attr(ref, "id", NS.r)) : null;
    const xml = path ? await part(path) : null;
    if (!xml) return;
    const plot = xml.getElementsByTagNameNS(NS2.c, "plotArea")[0];
    if (!plot) return;
    for (const c of plot.children) {
      if (c.localName === "barChart" || c.localName === "bar3DChart") {
        const dir = child(c, "barDir", NS2.c);
        o.chartType = dir && dir.getAttribute("val") === "bar" ? "bar" : "column";
      } else if (/^line/.test(c.localName)) o.chartType = "line";
      else if (/^(pie|doughnut)/.test(c.localName)) o.chartType = "pie";
      if (o.chartType) break;
    }
  }

  /** Page color, breaks, hidden text and unfilled content controls. */
  function collectBodyFacts(documentXml, body, model) {
    const bg = documentXml.getElementsByTagNameNS(NS.w, "background")[0];
    const color = attr(bg, "color");
    if (color && !["auto", "FFFFFF", "ffffff"].includes(color)) model.background.color = "#" + color.toUpperCase();
    for (const br of body.getElementsByTagNameNS(NS.w, "br")) {
      const type = attr(br, "type");
      if (type === "page") model.breaks.page++;
      if (type === "column") model.breaks.column++;
    }
    for (const p of model.paragraphs) {
      if (onOff(child(child(p.element, "pPr"), "pageBreakBefore"))) model.breaks.page++;
      for (const run of p.runs) if (run.text.trim() && runIsOn(model, run, p, "vanish")) model.hiddenRuns++;
    }
    for (const sdtPr of body.getElementsByTagNameNS(NS.w, "sdtPr")) {
      if (onOff(child(sdtPr, "showingPlcHdr"))) model.placeholders++;
    }
  }

  /** Tracked changes still waiting to be accepted or rejected. */
  function collectRevisions(body, model) {
    const kinds = { ins: "insert", del: "delete", moveFrom: "move", moveTo: "move", rPrChange: "format", pPrChange: "format",
      sectPrChange: "format", tblPrChange: "format", trPrChange: "format", tcPrChange: "format", numberingChange: "format" };
    walk(body, function (e) {
      if (e.namespaceURI !== NS.w || !kinds[e.localName]) return;
      // A tracked paragraph mark (w:ins inside w:rPr) belongs to the change next to it; skip it.
      if ((e.localName === "ins" || e.localName === "del") && e.parentNode.localName === "rPr") return;
      let text = "";
      for (const t of e.getElementsByTagNameNS(NS.w, e.localName === "del" || e.localName === "moveFrom" ? "delText" : "t")) text += t.textContent;
      model.revisions.push({ kind: kinds[e.localName], author: attr(e, "author") || "", text: text });
    });
  }

  /** Footnotes and endnotes (the separator lines Word keeps in the same file are not notes). */
  async function collectNotes(model, part) {
    for (const kind of ["footnote", "endnote"]) {
      const xml = await part("word/" + kind + "s.xml");
      if (xml) {
        for (const n of xml.getElementsByTagNameNS(NS.w, kind)) {
          const type = attr(n, "type");
          if (type && type !== "normal") continue;
          model.notes[kind].push(Array.from(n.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join("").trim());
        }
      }
      for (const p of model.paragraphs) model.noteRefs[kind] += p.element.getElementsByTagNameNS(NS.w, kind + "Reference").length;
    }
    // The number format: the section's setting wins over the document's.
    const docPr = (kind) => attr(child(child(model.settings && model.settings.documentElement, kind + "Pr"), "numFmt"), "val");
    model.noteFormats = {
      footnote: model.sections.map((s) => s.footnoteFormat).find(Boolean) || docPr("footnote") || "decimal",
      endnote: model.sections.map((s) => s.endnoteFormat).find(Boolean) || docPr("endnote") || "lowerRoman",
    };
  }

  /**
   * Comments: who wrote each one, what it says, the text it's attached to,
   * and (from commentsExtended.xml) whether it's a reply and whether it's resolved.
   */
  async function collectComments(body, model, part) {
    const xml = await part("word/comments.xml");
    if (!xml) return;
    const ext = await part("word/commentsExtended.xml");
    const extByPara = {};
    if (ext) for (const e of ext.getElementsByTagNameNS(NS2.w15, "commentEx")) {
      extByPara[e.getAttributeNS(NS2.w15, "paraId")] = { parent: e.getAttributeNS(NS2.w15, "paraIdParent") || "", done: e.getAttributeNS(NS2.w15, "done") === "1" };
    }
    // The document text each comment is attached to.
    const anchors = {}, open = new Set();
    walk(body, function (e) {
      if (e.namespaceURI !== NS.w) return;
      if (e.localName === "commentRangeStart") { open.add(attr(e, "id")); anchors[attr(e, "id")] = anchors[attr(e, "id")] || ""; }
      else if (e.localName === "commentRangeEnd") open.delete(attr(e, "id"));
      else if (e.localName === "t" && open.size) open.forEach((id) => { anchors[id] += e.textContent; });
    });
    const byPara = {};
    for (const c of xml.getElementsByTagNameNS(NS.w, "comment")) {
      const paras = c.getElementsByTagNameNS(NS.w, "p");
      const last = paras[paras.length - 1];
      const paraId = last ? last.getAttributeNS(NS.w14, "paraId") : "";
      const info = extByPara[paraId] || { parent: "", done: false };
      const entry = { id: attr(c, "id"), author: attr(c, "author") || "", text: Array.from(c.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join("").trim(),
        anchor: anchors[attr(c, "id")] || "", paraId: paraId, parentParaId: info.parent, done: info.done, replies: [] };
      byPara[paraId] = entry;
      model.comments.push(entry);
    }
    model.comments.forEach(function (c) {
      if (c.parentParaId && byPara[c.parentParaId]) { c.parent = byPara[c.parentParaId]; c.parent.replies.push(c); }
    });
  }

  /** The citation manager's sources live in a customXml part: count them and read the style (APA, MLA...). */
  async function collectSources(zip, model) {
    const files = zip.file(/^customXml\/item\d+\.xml$/);
    for (const f of files) {
      let xml;
      try { xml = parseXml(await f.async("string")); } catch (e) { continue; }
      const root = xml.documentElement;
      if (!root || root.namespaceURI !== NS2.b || root.localName !== "Sources") continue;
      // A file can carry more than one source list (an empty one from the template); keep the biggest.
      const n = root.getElementsByTagNameNS(NS2.b, "Source").length;
      if (n < model.sources.count) continue;
      model.sources.count = n;
      model.sources.style = root.getAttribute("StyleName") || root.getAttribute("SelectedStyle") || "";
      model.sources.version = root.getAttribute("Version") || "";
    }
  }

  /** Watermarks: Word draws them in the header as WordArt-style text ("DRAFT") or a faded picture. */
  function collectWatermarks(model) {
    for (const xml of model.headerXml) {
      for (const tp of xml.getElementsByTagNameNS(NS2.v, "textpath")) {
        const text = tp.getAttribute("string");
        if (text) model.background.watermarks.push(text);
      }
      for (const shape of xml.getElementsByTagNameNS(NS2.v, "shape")) {
        if (/WordPictureWatermark/i.test(shape.getAttribute("id") || "")) model.background.watermarks.push("(picture)");
      }
      // Newer Word may draw a text watermark as a shape whose name says so.
      for (const docPr of xml.getElementsByTagNameNS(NS2.wp, "docPr")) {
        if (/watermark/i.test(docPr.getAttribute("name") || "")) {
          const anchor = docPr.parentNode;
          const t = Array.from(anchor.getElementsByTagNameNS(NS.w, "t")).map((x) => x.textContent).join("");
          if (t) model.background.watermarks.push(t);
        }
      }
    }
  }

  // =====================================================================
  // PART 3: WORKING OUT FORMATTING
  // ---------------------------------------------------------------------
  // Word decides each setting by looking in this order; the first place
  // that has it wins:
  //   1. formatting put directly on the text (the B button, the font box)
  //   2. a character style on the text (like "Strong")
  //   3. the paragraph's style (like "Heading 1"), and the styles it's based on
  //   4. the document's defaults
  // =====================================================================

  /** A style and every style it's based on, closest first. */
  function styleChain(model, styleId) {
    const chain = [];
    let id = styleId;
    while (id && model.styles[id] && chain.length < 20) { // 20 = guard against loops
      chain.push(model.styles[id]);
      id = model.styles[id].basedOn;
    }
    return chain;
  }

  /** The paragraph style a paragraph uses (its id, like "Heading1"). */
  function paragraphStyleId(model, p) {
    return attr(child(child(p.element, "pPr"), "pStyle"), "val") || model.defaultParagraphStyle;
  }

  /** Every rPr (text formatting) element that applies to a run, in Word's order. */
  function runPropertyLevels(model, run, p) {
    const levels = [];
    const rPr = child(run.element, "rPr");
    levels.push(rPr);
    const charStyle = attr(child(rPr, "rStyle"), "val");
    styleChain(model, charStyle).forEach((s) => levels.push(s.rPr));
    styleChain(model, paragraphStyleId(model, p)).forEach((s) => levels.push(s.rPr));
    levels.push(model.defaultRunProps);
    return levels.filter(Boolean);
  }

  /** The first element with this tag in the run's formatting levels, or null. */
  function findRunSetting(model, run, p, tagName, ns) {
    for (const level of runPropertyLevels(model, run, p)) {
      const found = child(level, tagName, ns);
      if (found) return found;
    }
    return null;
  }

  /** The font name of a run, as Word would show it in the font box. */
  function runFontName(model, run, p) {
    for (const level of runPropertyLevels(model, run, p)) {
      const rFonts = child(level, "rFonts");
      if (!rFonts) continue;
      // A theme font ("Aptos (Body)") wins over a named font at the same level.
      const theme = attr(rFonts, "asciiTheme") || attr(rFonts, "hAnsiTheme");
      if (theme) {
        const font = /^major/.test(theme) ? model.themeFonts.major : model.themeFonts.minor;
        if (font) return font;
      }
      const name = attr(rFonts, "ascii") || attr(rFonts, "hAnsi");
      if (name) return name;
      // This level only set fonts for other alphabets: keep looking.
    }
    return model.themeFonts.minor || "Times New Roman"; // Word's fallback when nothing is set
  }

  /** Font size in points. Word stores HALF points, so 22 means 11 pt. */
  function runSize(model, run, p) {
    const sz = attr(findRunSetting(model, run, p, "sz"), "val");
    return sz ? Number(sz) / 2 : 10; // Word's built-in default is 10 pt
  }

  /** Is an on/off setting (b = bold, i = italic) on for this run? */
  function runIsOn(model, run, p, tagName) {
    return onOff(findRunSetting(model, run, p, tagName)) === true;
  }

  /** Underline: any underline style except "none". */
  function runIsUnderlined(model, run, p) {
    const u = findRunSetting(model, run, p, "u");
    if (!u) return false;
    const val = attr(u, "val");
    return !!val && val !== "none";
  }

  /** Text color as "#RRGGBB", or "auto" (Word's automatic black). */
  function runColor(model, run, p) {
    const val = attr(findRunSetting(model, run, p, "color"), "val");
    if (!val || val.toLowerCase() === "auto") return "auto";
    return "#" + val.toUpperCase();
  }

  /** Which text effects a run has: a sorted list such as ["glow", "shadow"]. */
  function runEffects(model, run, p) {
    const found = [];
    const levels = runPropertyLevels(model, run, p);
    const has = (tag, ns) => levels.some((level) => child(level, tag, ns));
    if (has("shadow", NS.w14) || levels.some((l) => onOff(child(l, "shadow")))) found.push("shadow");
    if (has("glow", NS.w14)) found.push("glow");
    if (has("reflection", NS.w14)) found.push("reflection");
    // An outline with "no fill" isn't a visible outline. (Word saves one with some presets.)
    const outline = levels.map((l) => child(l, "textOutline", NS.w14)).find(Boolean);
    if ((outline && !child(outline, "noFill", NS.w14)) || levels.some((l) => onOff(child(l, "outline")))) found.push("outline");
    return found.sort();
  }

  /** The proofing language of a run, such as "en-US", or null when nothing says. */
  function runLanguage(model, run, p) {
    const lang = attr(findRunSetting(model, run, p, "lang"), "val");
    return lang || model.defaultLanguage || null;
  }

  /** Runs that hold visible text (spaces alone don't count). */
  function textRuns(runs) {
    return runs.filter((r) => r.text.trim() !== "");
  }

  /**
   * The rPr children that do NOT change how text looks. A run whose only
   * formatting is these is "clean" (no direct formatting).
   */
  const HARMLESS_RUN_TAGS = ["rStyle", "lang", "noProof", "webHidden", "specVanish", "szCs", "bCs", "iCs", "rtl", "cs", "eastAsianLayout"];

  /** Direct formatting on a run that changes its look (empty list = clean). */
  function directFormatting(run) {
    const rPr = child(run.element, "rPr");
    if (!rPr) return [];
    const found = [];
    for (const c of rPr.children) {
      if (c.namespaceURI === NS.w && HARMLESS_RUN_TAGS.includes(c.localName)) continue;
      // rFonts that only names fonts for other alphabets doesn't change Latin text.
      if (c.namespaceURI === NS.w && c.localName === "rFonts" &&
          !c.hasAttributeNS(NS.w, "ascii") && !c.hasAttributeNS(NS.w, "hAnsi") &&
          !c.hasAttributeNS(NS.w, "asciiTheme") && !c.hasAttributeNS(NS.w, "hAnsiTheme")) continue;
      found.push(FORMAT_NAMES[c.localName] || c.localName);
    }
    return Array.from(new Set(found));
  }
  const FORMAT_NAMES = { rFonts: "font", sz: "size", b: "bold", i: "italic", u: "underline", color: "color",
    highlight: "highlight", shd: "shading", strike: "strikethrough", caps: "all caps", smallCaps: "small caps",
    vertAlign: "superscript or subscript", spacing: "character spacing", glow: "glow", shadow: "shadow",
    reflection: "reflection", textOutline: "outline", textFill: "text fill", kern: "kerning", w: "character scale" };

  // ----- Paragraph formatting -----

  /** Every pPr (paragraph formatting) element that applies to a paragraph, in Word's order. */
  function paragraphPropertyLevels(model, p) {
    const levels = [child(p.element, "pPr")];
    styleChain(model, paragraphStyleId(model, p)).forEach((s) => levels.push(s.pPr));
    levels.push(model.defaultParaProps);
    return levels.filter(Boolean);
  }

  /** One attribute of one paragraph setting, looked up level by level. Example: ("spacing", "after"). */
  function paragraphSetting(model, p, tagName, attrName) {
    for (const level of paragraphPropertyLevels(model, p)) {
      const value = attr(child(level, tagName), attrName);
      if (value !== null) return value;
    }
    return null;
  }

  /** Alignment as "left", "center", "right" or "justify". */
  function paragraphAlign(model, p) {
    const jc = paragraphSetting(model, p, "jc", "val") || "left";
    return { start: "left", left: "left", center: "center", right: "right", end: "right", both: "justify", distribute: "justify" }[jc] || jc;
  }

  /** Line spacing: { rule: "auto", lines: 1.5 } or { rule: "exact"|"atLeast", points: 14 }. */
  function paragraphLineSpacing(model, p) {
    const line = paragraphSetting(model, p, "spacing", "line");
    const rule = paragraphSetting(model, p, "spacing", "lineRule") || "auto";
    if (line === null) return { rule: "auto", lines: 1 };
    if (rule === "auto") return { rule: "auto", lines: Number(line) / 240 }; // 240 = one line
    return { rule: rule, points: Number(line) / TWIPS_PER_PT };
  }

  /** Space before or after a paragraph, in points. */
  function paragraphSpace(model, p, which) {
    // "Auto" spacing (text pasted from web pages) is about 14 pt.
    const auto = paragraphSetting(model, p, "spacing", which + "Autospacing");
    if (auto !== null && !["0", "false", "off"].includes(auto.toLowerCase())) return 14;
    const v = paragraphSetting(model, p, "spacing", which);
    return v === null ? 0 : Number(v) / TWIPS_PER_PT;
  }

  /** First-line indent in points (a hanging indent is negative). */
  function paragraphFirstLine(model, p) {
    for (const level of paragraphPropertyLevels(model, p)) {
      const ind = child(level, "ind");
      if (!ind) continue;
      if (attr(ind, "hanging") !== null) return -Number(attr(ind, "hanging")) / TWIPS_PER_PT;
      if (attr(ind, "firstLine") !== null) return Number(attr(ind, "firstLine")) / TWIPS_PER_PT;
    }
    return 0;
  }

  /** Left indent in points. */
  function paragraphLeftIndent(model, p) {
    const v = paragraphSetting(model, p, "ind", "left") || paragraphSetting(model, p, "ind", "start");
    return v === null ? 0 : Number(v) / TWIPS_PER_PT;
  }

  /** Line spacing in everyday words: "single", "1.5", "double", "exactly 14 pt". */
  function describeLineSpacing(ls) {
    if (ls.rule !== "auto") return (ls.rule === "exact" ? "exactly " : "at least ") + showLength(ls.points, "pt");
    const near = (n) => Math.abs(ls.lines - n) < 0.02;
    if (near(1)) return "single";
    if (near(2)) return "double";
    return String(Math.round(ls.lines * 100) / 100);
  }

  // =====================================================================
  // PART 4: SELECTORS - which part of the document a check looks at
  // ---------------------------------------------------------------------
  //   in=body (default) | in=header | in=footer | in=document
  //   para="text"   the first paragraph containing this text
  //   paras="text"  every paragraph containing this text
  //   text="text"   exactly this text (the first place it appears)
  //   style="Heading 1"  every paragraph using this style
  // Returns { targets: [{ p, runs }], label, missing } where `missing` is a
  // learner-friendly message when the text can't be found.
  // =====================================================================

  function partParagraphs(model, where) {
    if (where === "header") return model.headers;
    if (where === "footer") return model.footers;
    if (where === "document") return model.paragraphs.concat(model.headers, model.footers);
    return model.paragraphs;
  }

  /** The first body paragraph containing `text` (a paragraph that STARTS with it wins). */
  function findParagraph(model, text, list) {
    const wanted = loose(text);
    // Skip table-of-contents entries: they repeat the headings, and checks are about the real ones.
    const pool = (list || model.paragraphs).filter((p) => !isTocEntry(model, p));
    return pool.find((p) => loose(p.text).startsWith(wanted)) || pool.find((p) => loose(p.text).includes(wanted)) || null;
  }

  /** A line in a table of contents (style TOC 1, TOC 2...). */
  function isTocEntry(model, p) {
    const id = paragraphStyleId(model, p) || "";
    const style = model.styles[id];
    return /^toc ?\d$/i.test(id) || (!!style && /^toc \d$/i.test(style.name));
  }

  function notFound(text) {
    return 'I couldn\'t find "' + snippet(text) + '" in the document. If you\'ve already typed it, check the spelling; if you changed that wording, change it back or load a fresh copy of the practice file.';
  }

  function select(model, params) {
    if (params.para !== undefined) {
      const p = findParagraph(model, params.para, partParagraphs(model, "document"));
      return p ? { targets: [{ p: p, runs: p.runs }], label: 'the paragraph "' + snippet(p.text) + '"' }
        : { targets: [], missing: notFound(params.para) };
    }
    if (params.paras !== undefined) {
      const wanted = loose(params.paras);
      const ps = partParagraphs(model, "document").filter((p) => loose(p.text).includes(wanted));
      return ps.length ? { targets: ps.map((p) => ({ p: p, runs: p.runs })), label: 'the paragraphs with "' + snippet(params.paras) + '"' }
        : { targets: [], missing: notFound(params.paras) };
    }
    if (params.text !== undefined) {
      for (const p of partParagraphs(model, "document")) {
        const runs = runsCovering(p, params.text);
        if (runs) return { targets: [{ p: p, runs: runs }], label: '"' + snippet(params.text) + '"' };
      }
      return { targets: [], missing: notFound(params.text) };
    }
    if (params.style !== undefined) {
      const id = model.styleIdsByName[String(params.style).toLowerCase()] || params.style;
      const ps = model.paragraphs.filter((p) => paragraphStyleId(model, p) === id && p.text.trim() !== "");
      return ps.length ? { targets: ps.map((p) => ({ p: p, runs: p.runs })), label: 'the paragraphs in the "' + params.style + '" style' }
        : { targets: [], missing: 'No paragraphs use the "' + params.style + '" style yet.' };
    }
    const where = params.in || "body";
    const ps = partParagraphs(model, where).filter((p) => p.text.trim() !== "");
    const labels = { body: "the document", header: "the header", footer: "the footer", document: "the document" };
    return ps.length ? { targets: ps.map((p) => ({ p: p, runs: p.runs })), label: labels[where] }
      : { targets: [], missing: where === "footer" ? "The footer is empty." : where === "header" ? "The header is empty." : "The document is empty." };
  }

  /** The runs that hold a piece of text inside a paragraph (or null if it isn't there). */
  function runsCovering(p, text) {
    const start = plain(p.text).toLowerCase().indexOf(plain(text).toLowerCase());
    if (start === -1) return null;
    const end = start + text.length;
    const covering = [];
    let pos = 0;
    for (const run of p.runs) {
      const runStart = pos, runEnd = pos + run.text.length;
      if (runEnd > start && runStart < end) covering.push(run);
      pos = runEnd;
    }
    return covering;
  }

  // =====================================================================
  // PART 5: THE CHECKS
  // ---------------------------------------------------------------------
  // Each check takes (model, params, context) and returns
  //   { pass: true/false, message: "...", where: { search: "..." } }
  // Messages say WHAT is still wrong, in plain words, never "you clicked the
  // wrong thing". `where` lets the add-in's "Show me" button find the spot.
  // =====================================================================

  const CHECKS = {};

  // ---------- file.name ----------
  CHECKS["file.name"] = function (model, params) {
    // Office gives a web address or a path; we only want the last part.
    const raw = String(model.fileName || "").split("?")[0].split("#")[0];
    let name = raw.split(/[\\/]/).pop();
    try { name = decodeURIComponent(name); } catch (e) { /* keep it as it is */ }
    if (!name) return fail("Save the file first. Then select Check my work again.");
    if (/^lastname[_ ]/i.test(name)) return fail('Almost: put your own last name in place of "Lastname".');
    if (wildcard(params.pattern).test(name)) return pass('The file is saved as "' + name + '".');
    const example = String(params.pattern).replace(/^\*/, "Lastname");
    return fail('The file is named "' + name + '". It should look like "' + example + '".');
  };

  // ---------- text.present ----------
  CHECKS["text.present"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing && params.in) return fail(sel.missing);
    const n = countText(sel.targets, params.value, params.case);
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const place = params.in === "footer" ? " in the footer" : params.in === "header" ? " in the header" : "";
    if (n >= min && n <= max) {
      return pass(n === 1 ? '"' + params.value + '" is there' + place + "." : '"' + params.value + '" appears ' + n + " times" + place + ".");
    }
    if (n === 0) return fail('I don\'t see "' + params.value + '"' + place + " yet.");
    if (n < min) return fail('"' + params.value + '" appears ' + count(n, "time") + place + ". It should appear at least " + min + ".");
    return fail('"' + params.value + '" appears ' + count(n, "time") + place + ". It should appear no more than " + max + ".");
  };

  // ---------- text.absent ----------
  CHECKS["text.absent"] = function (model, params) {
    const sel = select(model, params);
    const n = countText(sel.targets, params.value, params.case);
    if (n === 0) return pass('"' + params.value + '" is gone.');
    return fail('"' + params.value + '" still appears ' + count(n, "time") + ".", { search: params.value, matchCase: params.case !== "insensitive" });
  };

  /** How many times `value` appears in the targets' text (never across two paragraphs). */
  function countText(targets, value, caseMode) {
    const insensitive = caseMode === "insensitive";
    const needle = insensitive ? plain(value).toLowerCase() : plain(value);
    if (!needle) return 0;
    let total = 0;
    for (const t of targets) {
      const text = t.runs.map((r) => r.text).join("");
      const hay = insensitive ? plain(text).toLowerCase() : plain(text);
      let at = hay.indexOf(needle);
      while (at !== -1) { total++; at = hay.indexOf(needle, at + needle.length); }
    }
    return total;
  }

  // ---------- text.noDoubleSpaces ----------
  CHECKS["text.noDoubleSpaces"] = function (model, params) {
    const sel = select(model, params);
    let n = 0, first = null;
    for (const t of sel.targets) {
      const text = t.runs.map((r) => r.text).join("").replace(/ /g, " ");
      const hits = text.match(/ {2,}/g);
      if (hits) { n += hits.length; if (!first) first = text; }
    }
    if (n === 0) return pass("There are no double spaces.");
    return fail("There " + (n === 1 ? "is still 1 double space" : "are still " + n + " double spaces") + ', starting in "' + snippet(first) + '".', { search: "  " });
  };

  // ---------- para.noEmpty ----------
  CHECKS["para.noEmpty"] = function (model, params) {
    // Only body paragraphs outside tables (an empty table cell isn't a spacing trick).
    let list = model.paragraphs.filter((p) => !p.inTable);
    // Word always keeps one last paragraph mark that can't be deleted, so a
    // single empty paragraph at the very end is allowed.
    if (list.length && isEmptyParagraph(list[list.length - 1])) list = list.slice(0, -1);
    // after="..." / before="..." limit the search to part of the document.
    if (params.after !== undefined) {
      const p = findParagraph(model, params.after, list);
      if (!p) return fail(notFound(params.after));
      list = list.slice(list.indexOf(p) + 1);
    }
    if (params.before !== undefined) {
      const p = findParagraph(model, params.before, list);
      if (!p) return fail(notFound(params.before));
      list = list.slice(0, list.indexOf(p));
    }
    const max = params.max !== undefined ? Number(params.max) : 0;
    // Count empty paragraphs that come more than `max` in a row.
    let tooMany = 0, inARow = 0, firstAfter = null, previous = null;
    for (const p of list) {
      if (isEmptyParagraph(p)) {
        inARow++;
        if (inARow > max) { tooMany++; if (firstAfter === null) firstAfter = previous; }
      } else {
        inARow = 0;
        previous = p;
      }
    }
    if (tooMany === 0) return pass(max === 0 ? "There are no empty paragraphs." : "No more than " + max + " empty paragraphs in a row.");
    const where = firstAfter ? ', the first one after "' + snippet(firstAfter.text, 30) + '"' : ", the first one at the top";
    return fail(count(tooMany, "empty paragraph") + " still " + (tooMany === 1 ? "makes" : "make") + " space" + where + ".", { emptyParagraph: true });
  };

  // ---------- para.order ----------
  CHECKS["para.order"] = function (model, params) {
    const all = model.paragraphs;
    const a = findParagraph(model, params.first, all);
    const b = findParagraph(model, params.then, all);
    if (!a) return fail(notFound(params.first));
    if (!b) return fail(notFound(params.then));
    if (all.indexOf(a) < all.indexOf(b)) return pass('"' + snippet(a.text, 30) + '" comes before "' + snippet(b.text, 30) + '".');
    return fail('"' + snippet(a.text, 30) + '" should come before "' + snippet(b.text, 30) + '". Right now it comes after.', { search: params.first.slice(0, 200) });
  };

  // ---------- font.is ----------
  CHECKS["font.is"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    const wanted = [];
    if (params.name !== undefined) wanted.push(["font", (r, p) => runFontName(model, r, p), (v) => v.toLowerCase() === String(params.name).toLowerCase(), params.name, (v) => v]);
    if (params.size !== undefined) wanted.push(["size", (r, p) => runSize(model, r, p), (v) => Math.abs(v - Number(params.size)) < 0.01, params.size + " pt", (v) => v + " pt"]);
    if (params.bold !== undefined) wanted.push(["bold", (r, p) => runIsOn(model, r, p, "b"), (v) => v === (params.bold === "true"), params.bold === "true" ? "bold" : "not bold", (v) => v ? "bold" : "not bold"]);
    if (params.italic !== undefined) wanted.push(["italic", (r, p) => runIsOn(model, r, p, "i"), (v) => v === (params.italic === "true"), params.italic === "true" ? "italic" : "not italic", (v) => v ? "italic" : "not italic"]);
    if (params.underline !== undefined) wanted.push(["underline", (r, p) => runIsUnderlined(model, r, p), (v) => v === (params.underline === "true"), params.underline === "true" ? "underlined" : "not underlined", (v) => v ? "underlined" : "not underlined"]);
    if (params.color !== undefined) {
      const c = String(params.color).toUpperCase();
      wanted.push(["color", (r, p) => runColor(model, r, p), (v) => v === c || (v === "auto" && c === "#000000"), c, (v) => v === "auto" ? "automatic (black)" : v]);
    }
    for (const t of sel.targets) {
      for (const run of textRuns(t.runs)) {
        for (const w of wanted) {
          const actual = w[1](run, t.p);
          if (!w[2](actual)) {
            const what = w[0] === "bold" || w[0] === "italic" || w[0] === "underline"
              ? '"' + snippet(run.text, 30) + '" is ' + w[4](actual) + ". It should be " + w[3] + "."
              : 'The ' + w[0] + ' of "' + snippet(run.text, 30) + '" is ' + w[4](actual) + ". It should be " + w[3] + ".";
            return fail(what, { search: run.text.trim().slice(0, 200) });
          }
        }
      }
    }
    return pass("The formatting of " + sel.label + " is right.");
  };

  // ---------- effect.present ----------
  CHECKS["effect.present"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    const kind = params.kind || "any";
    for (const t of sel.targets) {
      for (const run of textRuns(t.runs)) {
        const effects = runEffects(model, run, t.p);
        const ok = kind === "any" ? effects.length > 0 : effects.includes(kind);
        if (!ok) {
          const wanted = kind === "any" ? "a text effect (shadow, glow, reflection or outline)" : "a " + kind + " effect";
          return fail('"' + snippet(run.text, 30) + '" doesn\'t have ' + wanted + " yet.", { search: run.text.trim().slice(0, 200) });
        }
      }
    }
    return pass(sel.label.charAt(0).toUpperCase() + sel.label.slice(1) + " has a text effect.");
  };

  // ---------- format.same ----------
  const LOOK_PROPS = {
    "font.name": { label: "font", get: runFontName, show: (v) => v },
    "font.size": { label: "size", get: runSize, show: (v) => v + " pt" },
    "font.bold": { label: "bold", get: (m, r, p) => runIsOn(m, r, p, "b"), show: (v) => v ? "bold" : "not bold" },
    "font.italic": { label: "italic", get: (m, r, p) => runIsOn(m, r, p, "i"), show: (v) => v ? "italic" : "not italic" },
    "font.underline": { label: "underline", get: runIsUnderlined, show: (v) => v ? "underlined" : "not underlined" },
    "font.color": { label: "color", get: runColor, show: (v) => v === "auto" ? "automatic (black)" : v },
    "effect": { label: "text effects", get: (m, r, p) => runEffects(m, r, p).join(", ") || "none", show: (v) => v },
  };
  CHECKS["format.same"] = function (model, params) {
    const pa = findParagraph(model, params.a);
    const pb = findParagraph(model, params.b);
    if (!pa) return fail(notFound(params.a));
    if (!pb) return fail(notFound(params.b));
    const runsA = textRuns(pa.runs), runsB = textRuns(pb.runs);
    if (!runsA.length || !runsB.length) return fail("One of the two lines is empty.");
    const props = String(params.props || "font.name,font.size,font.bold,font.italic,font.color,effect").split(",").map((s) => s.trim()).filter((k) => LOOK_PROPS[k]);
    for (const key of props) {
      const prop = LOOK_PROPS[key];
      const target = prop.get(model, runsA[0], pa); // the look we're matching
      for (const run of runsB) {
        const actual = prop.get(model, run, pb);
        if (actual !== target) {
          return fail('The ' + prop.label + ' doesn\'t match yet: "' + snippet(pb.text, 30) + '" is ' + prop.show(actual) +
            ', but "' + snippet(pa.text, 30) + '" is ' + prop.show(target) + ".", { search: params.b.slice(0, 200) });
        }
      }
    }
    return pass('"' + snippet(pb.text, 30) + '" matches "' + snippet(pa.text, 30) + '".');
  };

  // ---------- format.clean ----------
  CHECKS["format.clean"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    for (const t of sel.targets) {
      for (const run of textRuns(t.runs)) {
        const extra = directFormatting(run);
        if (extra.length) {
          return fail('"' + snippet(t.p.text, 30) + '" still has its own formatting (' + extra.slice(0, 3).join(", ") + ').', { search: run.text.trim().slice(0, 200) });
        }
      }
    }
    return pass(sel.label.charAt(0).toUpperCase() + sel.label.slice(1) + " has no leftover formatting.");
  };

  // ---------- para.is ----------
  CHECKS["para.is"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    for (const t of sel.targets) {
      const p = t.p;
      const name = '"' + snippet(p.text, 30) + '"';
      const where = { search: p.text.trim().slice(0, 60) };
      if (params.align !== undefined) {
        const a = paragraphAlign(model, p);
        if (a !== params.align) return fail(name + " is " + ALIGN_WORDS[a] + ". It should be " + ALIGN_WORDS[params.align] + ".", where);
      }
      if (params.lineSpacing !== undefined) {
        const ls = paragraphLineSpacing(model, p);
        if (!lineSpacingMatches(ls, params.lineSpacing)) {
          return fail("The line spacing of " + name + " is " + describeLineSpacing(ls) + ". It should be " + params.lineSpacing.replace(/pt$/, " pt") + ".", where);
        }
      }
      if (params.dropCap !== undefined) {
        const frame = child(child(p.element, "pPr"), "framePr");
        const has = !!frame && ["drop", "margin"].includes(attr(frame, "dropCap"));
        if (has !== yes(params.dropCap)) return fail(has ? name + " has a drop cap. Remove it." : name + " doesn't start with a drop cap yet (Insert > Drop Cap).", where);
      }
      if (params.keepNext !== undefined) {
        const kn = paragraphPropertyLevels(model, p).map((l) => child(l, "keepNext")).find(Boolean);
        if ((onOff(kn) === true) !== yes(params.keepNext)) return fail("Keep with next is " + (onOff(kn) ? "on" : "off") + " for " + name + ". It should be " + (yes(params.keepNext) ? "on." : "off."), where);
      }
      const lengths = [
        ["spaceBefore", "Space Before", () => paragraphSpace(model, p, "before")],
        ["spaceAfter", "Space After", () => paragraphSpace(model, p, "after")],
        ["firstLineIndent", "first-line indent", () => paragraphFirstLine(model, p)],
        ["leftIndent", "left indent", () => paragraphLeftIndent(model, p)],
      ];
      for (const [key, label, get] of lengths) {
        if (params[key] === undefined) continue;
        const want = toPoints(params[key]);
        const actual = get();
        if (Math.abs(actual - want) > 0.5) { // within half a point counts
          const unit = unitOf(params[key]);
          return fail("The " + label + " of " + name + " is " + showLength(actual, unit) + ". It should be " + showLength(want, unit) + ".", where);
        }
      }
    }
    return pass(sel.targets.length === 1 ? "The paragraph settings are right." : "The paragraph settings are right in all " + sel.targets.length + " paragraphs.");
  };
  const ALIGN_WORDS = { left: "left-aligned", center: "centered", right: "right-aligned", justify: "justified" };

  function lineSpacingMatches(ls, wanted) {
    const multiples = { single: 1, "1.15": 1.15, "1.5": 1.5, double: 2 };
    if (multiples[wanted] !== undefined) return ls.rule === "auto" && Math.abs(ls.lines - multiples[wanted]) < 0.02;
    const pts = toPoints(wanted);
    return ls.rule !== "auto" && Math.abs(ls.points - pts) <= 0.5;
  }

  // ---------- page.is ----------
  CHECKS["page.is"] = function (model, params) {
    const many = model.sections.length > 1;
    for (let i = 0; i < model.sections.length; i++) {
      const s = model.sections[i];
      const which = many ? "In section " + (i + 1) + ", " : "";
      const problems = []; // collect everything wrong in this section, so the learner sees it all at once

      // Margins (Word stores twips; lessons use in, pt or cm). Within 0.01 in counts.
      const wrongSides = [];
      for (const side of ["top", "bottom", "left", "right"]) {
        const raw = params[side] !== undefined ? params[side] : params.margins;
        if (raw === undefined) continue;
        const want = toPoints(raw), actual = s[side] / TWIPS_PER_PT;
        if (Math.abs(actual - want) > 0.72) wrongSides.push({ side: side, actual: actual, want: want, unit: unitOf(raw) });
      }
      const allSame = wrongSides.length === 4 && wrongSides.every((w) => Math.abs(w.actual - wrongSides[0].actual) < 0.72);
      if (allSame) {
        const w = wrongSides[0];
        problems.push("all four margins are " + showLength(w.actual, w.unit) + ", not " + showLength(w.want, w.unit));
      } else {
        wrongSides.forEach((w) => problems.push("the " + w.side + " margin is " + showLength(w.actual, w.unit) + ", not " + showLength(w.want, w.unit)));
      }

      if (params.size !== undefined) {
        const paper = PAPER[params.size];
        const short = Math.min(s.width, s.height), long = Math.max(s.width, s.height);
        if (Math.abs(short - paper.w) > 20 || Math.abs(long - paper.h) > 20) {
          problems.push("the paper is " + paperName(s) + ", not " + paper.name);
        }
      }
      if (params.orientation !== undefined) {
        const orientation = s.width > s.height ? "landscape" : "portrait";
        if (orientation !== params.orientation) problems.push("the page is " + orientation + ", not " + params.orientation);
      }
      if (problems.length) {
        const list = problems.length === 1 ? problems[0]
          : problems.length === 2 ? problems.join(", and ")
          : problems.slice(0, -1).join("; ") + "; and " + problems[problems.length - 1];
        const sentence = which + list;
        return fail(sentence.charAt(0).toUpperCase() + sentence.slice(1) + ".");
      }
    }
    return pass("The page setup is right.");
  };

  /** Describe a section's paper: "A4 (8.27 x 11.69 in)" or "9 x 12 in". */
  function paperName(s) {
    const short = Math.min(s.width, s.height), long = Math.max(s.width, s.height);
    for (const key of Object.keys(PAPER)) {
      if (Math.abs(short - PAPER[key].w) <= 20 && Math.abs(long - PAPER[key].h) <= 20) return PAPER[key].label;
    }
    const inches = (t) => Math.round(t / 1440 * 100) / 100;
    return inches(s.width) + " x " + inches(s.height) + " in";
  }

  // ---------- props.is ----------
  CHECKS["props.is"] = function (model, params) {
    const names = { title: "Title", author: "Author", keywords: "Tags" };
    for (const key of ["title", "author", "keywords"]) {
      if (params[key] === undefined) continue;
      const actual = model.props[key] || "";
      if (params[key] === "*") {
        if (!actual.trim()) return fail("The " + names[key] + " property is empty. Fill it in on File > Info.");
      } else if (!loose(actual).includes(loose(params[key]))) {
        return fail(actual.trim()
          ? "The " + names[key] + ' property is "' + snippet(actual) + '". It should ' + (key === "keywords" ? "include" : "be") + ' "' + params[key] + '".'
          : "The " + names[key] + ' property is empty. It should ' + (key === "keywords" ? "include" : "be") + ' "' + params[key] + '".');
      }
    }
    return pass("The document properties are filled in.");
  };

  // ---------- lang.is ----------
  CHECKS["lang.is"] = function (model, params) {
    const wanted = String(params.value);
    const wrong = {};   // language -> how many paragraphs use it
    let firstWrong = null, unset = 0;
    for (const p of model.paragraphs) {
      for (const run of textRuns(p.runs)) {
        const lang = runLanguage(model, run, p);
        if (!lang) { unset++; continue; }
        if (lang.toLowerCase() !== wanted.toLowerCase()) {
          wrong[lang] = (wrong[lang] || 0) + 1;
          if (!firstWrong) firstWrong = p;
          break; // one wrong run is enough to count this paragraph
        }
      }
    }
    const langs = Object.keys(wrong);
    if (!langs.length && !unset) return pass("The proofing language is " + (LANGUAGE_NAMES[wanted] || wanted) + ".");
    if (!langs.length) return fail("Some text has no proofing language set. Select all, then set it to " + (LANGUAGE_NAMES[wanted] || wanted) + ".");
    const total = langs.reduce((s, l) => s + wrong[l], 0);
    return fail(count(total, "paragraph") + " " + (total === 1 ? "is" : "are") + " still marked " + (LANGUAGE_NAMES[langs[0]] || langs[0]) +
      ', starting with "' + snippet(firstWrong.text, 30) + '". It should be ' + (LANGUAGE_NAMES[wanted] || wanted) + ".", { search: firstWrong.text.trim().slice(0, 60) });
  };

  // ---------- self.attest ----------
  // The learner's own word: the page shows a checkbox. Nothing in the document to read.
  CHECKS["self.attest"] = function (model, params, context) {
    const ok = !!(context && context.attested);
    return ok ? pass("Confirmed.") : fail("Tick the box to confirm: " + params.label);
  };

  // =====================================================================
  // PART 5b: CHECKS FOR WEEKS 2 TO 5
  // Styles, navigation, lists, headers and sections, references, graphics,
  // tables, page design, templates, mail merge and collaboration.
  // =====================================================================

  /** "Heading 1" -> its style id, or null. Built-in style names are saved in lowercase ("heading 1"). */
  function styleIdFor(model, name) {
    const id = model.styleIdsByName[String(name).toLowerCase()];
    if (id) return id;
    return model.styles[name] ? name : (model.styles[String(name).replace(/\s+/g, "")] ? String(name).replace(/\s+/g, "") : null);
  }
  /** The name Word shows for a paragraph's style ("Heading 1", "Normal"). */
  function styleNameOf(model, p) {
    const id = paragraphStyleId(model, p);
    const s = model.styles[id];
    const name = s ? s.name : id || "Normal";
    return String(name).replace(/^heading (\d)$/i, "Heading $1").replace(/^normal$/i, "Normal").replace(/^title$/i, "Title").replace(/^caption$/i, "Caption");
  }
  const sameName = (a, b) => String(a).toLowerCase().replace(/\s+/g, "") === String(b).toLowerCase().replace(/\s+/g, "");
  const yes = (v) => v === true || v === "true";

  // ---------- style.is: these paragraphs use this style ----------
  CHECKS["style.is"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    for (const t of sel.targets) {
      const actual = styleNameOf(model, t.p);
      if (!sameName(actual, params.name)) {
        return fail('"' + snippet(t.p.text, 30) + '" uses the ' + actual + " style. It should use " + params.name + ".", { search: t.p.text.trim().slice(0, 60) });
      }
    }
    return pass(sel.targets.length === 1 ? '"' + snippet(sel.targets[0].p.text, 30) + '" uses ' + params.name + "." : "All " + sel.targets.length + " paragraphs use " + params.name + ".");
  };

  // ---------- style.def: how a style itself is set up (Modify Style) ----------
  CHECKS["style.def"] = function (model, params) {
    const id = styleIdFor(model, params.name);
    if (!id) return fail('This document has no style called "' + params.name + '" yet.');
    // A pretend paragraph and run in that style, so Word's lookup rules (PART 3) give the answer.
    const xmlDoc = document.implementation.createDocument(NS.w, "w:root", null);
    const pEl = xmlDoc.createElementNS(NS.w, "w:p");
    const pPr = xmlDoc.createElementNS(NS.w, "w:pPr");
    const pStyle = xmlDoc.createElementNS(NS.w, "w:pStyle");
    pStyle.setAttributeNS(NS.w, "w:val", id);
    pPr.append(pStyle); pEl.append(pPr);
    const rEl = xmlDoc.createElementNS(NS.w, "w:r");
    pEl.append(rEl);
    const p = { element: pEl, runs: [{ element: rEl, text: "x" }], text: "x" };
    const run = p.runs[0];
    const label = params.name + " style";
    if (params.font !== undefined && runFontName(model, run, p).toLowerCase() !== String(params.font).toLowerCase())
      return fail("The " + label + "'s font is " + runFontName(model, run, p) + ". It should be " + params.font + ".");
    if (params.size !== undefined && Math.abs(runSize(model, run, p) - Number(params.size)) > 0.01)
      return fail("The " + label + "'s size is " + runSize(model, run, p) + " pt. It should be " + params.size + " pt.");
    if (params.bold !== undefined && runIsOn(model, run, p, "b") !== yes(params.bold))
      return fail("The " + label + " is " + (runIsOn(model, run, p, "b") ? "bold" : "not bold") + ". It should " + (yes(params.bold) ? "be bold." : "not be bold."));
    if (params.italic !== undefined && runIsOn(model, run, p, "i") !== yes(params.italic))
      return fail("The " + label + " is " + (runIsOn(model, run, p, "i") ? "italic" : "not italic") + ". It should " + (yes(params.italic) ? "be italic." : "not be italic."));
    if (params.color !== undefined) {
      const c = runColor(model, run, p);
      if (!(c === String(params.color).toUpperCase() || (c === "auto" && params.color === "#000000"))) return fail("The " + label + "'s color is " + (c === "auto" ? "automatic (black)" : c) + ". It should be " + params.color + ".");
    }
    const lengths = [["spaceBefore", "Space Before", () => paragraphSpace(model, p, "before")], ["spaceAfter", "Space After", () => paragraphSpace(model, p, "after")]];
    for (const [key, name, get] of lengths) {
      if (params[key] === undefined) continue;
      const want = toPoints(params[key]), actual = get();
      if (Math.abs(actual - want) > 0.5) return fail("The " + label + "'s " + name + " is " + showLength(actual, "pt") + ". It should be " + showLength(want, "pt") + ".");
    }
    if (params.keepNext !== undefined) {
      const kn = paragraphPropertyLevels(model, p).map((l) => child(l, "keepNext")).find(Boolean);
      if ((onOff(kn) === true) !== yes(params.keepNext)) return fail("Keep with next is " + (onOff(kn) ? "on" : "off") + " in the " + label + ". It should be " + (yes(params.keepNext) ? "on." : "off."));
    }
    if (params.align !== undefined && paragraphAlign(model, p) !== params.align)
      return fail("The " + label + " is " + ALIGN_WORDS[paragraphAlign(model, p)] + ". It should be " + ALIGN_WORDS[params.align] + ".");
    if (params.lineSpacing !== undefined && !lineSpacingMatches(paragraphLineSpacing(model, p), params.lineSpacing))
      return fail("The " + label + "'s line spacing is " + describeLineSpacing(paragraphLineSpacing(model, p)) + ". It should be " + params.lineSpacing + ".");
    return pass("The " + label + " is set up correctly.");
  };

  // ---------- bookmark.is ----------
  CHECKS["bookmark.is"] = function (model, params) {
    const b = model.bookmarks.find((x) => x.name.toLowerCase() === String(params.name).toLowerCase());
    if (!b) {
      const others = model.bookmarks.map((x) => x.name);
      return fail('There\'s no bookmark named "' + params.name + '" yet.' + (others.length ? " (Bookmarks found: " + others.slice(0, 3).join(", ") + ".)" : ""));
    }
    if (params.at !== undefined && !(b.paragraph && loose(b.paragraph.text).includes(loose(params.at)))) {
      return fail('The "' + b.name + '" bookmark is ' + (b.paragraph ? 'at "' + snippet(b.paragraph.text, 30) + '"' : "not on a heading") + '. Put it on "' + params.at + '".');
    }
    return pass('The bookmark "' + b.name + '" is in place.');
  };

  // ---------- link.present ----------
  /** Link text that doesn't say where it goes (WCAG 2.4.4). */
  function vagueLinkText(text) {
    const t = loose(text);
    return !t || /^(https?:\/\/|www\.)/.test(t) || ["click here", "here", "link", "this link", "read more", "more", "click"].includes(t);
  }
  CHECKS["link.present"] = function (model, params) {
    let links = model.links;
    if (params.label !== undefined) links = links.filter((l) => loose(l.text).includes(loose(params.label)));
    if (params.anchor !== undefined) links = links.filter((l) => l.anchor.toLowerCase() === String(params.anchor).toLowerCase());
    if (params.url !== undefined) links = links.filter((l) => l.url.toLowerCase().includes(String(params.url).toLowerCase()));
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    if (links.length < min) {
      if (!model.links.length) return fail("There are no links in the document yet.");
      if (params.anchor !== undefined) return fail('No link goes to the "' + params.anchor + '" bookmark yet.' + (params.label ? ' Make "' + params.label + '" the link.' : ""));
      if (params.url !== undefined) return fail("No link goes to an address with " + params.url + " in it yet.");
      return fail('I don\'t see a link on "' + params.label + '" yet.');
    }
    if (yes(params.descriptive)) {
      const vague = links.find((l) => vagueLinkText(l.text));
      if (vague) return fail('A link\'s text is "' + snippet(vague.text || "(empty)", 40) + '". Link text should say where it goes, such as "my portfolio".', { search: vague.text.slice(0, 200) });
    }
    return pass(links.length === 1 ? "The link is in place." : links.length + " links are in place.");
  };

  // ---------- field.present ----------
  CHECKS["field.present"] = function (model, params) {
    const where = params.in || "document";
    const want = tidyCode(params.code).toUpperCase();
    const n = model.fields.filter((f) => (where === "document" || f.part === where) && (f.code.toUpperCase() === want || f.code.toUpperCase().startsWith(want + " ") || f.code.toUpperCase().startsWith(want + "\\"))).length;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const place = { header: " in the header", footer: " in the footer", body: " in the body" }[where] || "";
    const parts = tidyCode(params.code).split(" ");
    // "MERGEFIELD Position" -> "the «Position» merge field"; "SEQ Figure" -> "a caption number"
    const friendly = parts[0].toUpperCase() === "MERGEFIELD" && parts[1] ? "the «" + parts[1] + "» merge field" : FIELD_WORDS[want.split(" ")[0]] || "the " + params.code + " field";
    if (n >= min && n <= max) return pass(max === 0 ? "There are no " + friendly.replace(/^an? |^the /, "") + " fields left." : (n === 1 ? "There's " + friendly : "There are " + n + " of them: " + friendly) + place + ".");
    if (n === 0) return fail("I don't see " + friendly + place + " yet.");
    if (n < min) return fail("There " + (n === 1 ? "is " : "are ") + n + " of " + friendly + place + "; there should be at least " + min + ".");
    return fail("There are still " + n + " " + friendly.replace(/^an? |^the /, "") + " fields" + place + (max === 0 ? ". Finish the merge so they turn into real text." : "."));
  };
  const FIELD_WORDS = { PAGE: "a page number", NUMPAGES: "the total page count", TOC: "a table of contents", MERGEFIELD: "a merge field",
    ADDRESSBLOCK: "an Address Block", GREETINGLINE: "a Greeting Line", IF: "an If...Then...Else rule", NEXT: "a Next Record field",
    SEQ: "a caption number", CITATION: "a citation", BIBLIOGRAPHY: "a bibliography", HYPERLINK: "a link", DATE: "a date" };

  // ---------- section.is ----------
  CHECKS["section.is"] = function (model, params) {
    let list = model.sections.map((s, i) => ({ s: s, i: i }));
    if (params.para !== undefined) {
      const p = findParagraph(model, params.para);
      if (!p) return fail(notFound(params.para));
      list = [list[p.section]];
    } else if (params.index !== undefined) {
      const i = Number(params.index) - 1;
      if (!model.sections[i]) return fail("The document has " + count(model.sections.length, "section") + ", so there's no section " + params.index + " yet. Add a section break.");
      list = [list[i]];
    }
    for (const { s, i } of list) {
      const which = model.sections.length > 1 ? "Section " + (i + 1) + (params.para ? ' (with "' + snippet(params.para, 24) + '")' : "") : "The document";
      if (params.orientation !== undefined) {
        const o = s.width > s.height ? "landscape" : "portrait";
        if (o !== params.orientation) return fail(which + " is " + o + ". It should be " + params.orientation + (model.sections.length === 1 && params.para ? ". Add a section break first, so only part of the document changes." : "."));
      }
      if (params.columns !== undefined && s.columns !== Number(params.columns))
        return fail(which + " has " + count(s.columns, "column") + ". It should have " + params.columns + ".");
      if (params.breakType !== undefined && i > 0 && s.type !== params.breakType)
        return fail(which + " starts with a " + BREAK_WORDS[s.type] + " section break. It should be " + BREAK_WORDS[params.breakType] + ".");
      if (params.breakType !== undefined && i === 0 && model.sections.length === 1)
        return fail("There's no section break yet.");
      if (params.firstPageDifferent !== undefined && s.firstPageDifferent !== yes(params.firstPageDifferent))
        return fail(s.firstPageDifferent ? "Different First Page is on. Turn it off." : "Different First Page isn't on yet. Turn it on in the header's options.");
      if (params.border !== undefined && s.border !== yes(params.border))
        return fail(s.border ? "The page has a border. Remove it." : "The page doesn't have a border yet.");
    }
    return pass(list.length === 1 && model.sections.length > 1 ? "That section is set up correctly." : "The section setup is right.");
  };
  const BREAK_WORDS = { nextPage: "a Next Page", continuous: "a Continuous", evenPage: "an Even Page", oddPage: "an Odd Page", nextColumn: "a Next Column" };

  // ---------- break.count ----------
  CHECKS["break.count"] = function (model, params) {
    const n = params.kind === "section" ? model.sections.length - 1 : model.breaks[params.kind] || 0;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const word = params.kind + " break";
    if (n >= min && n <= max) return pass("There " + (n === 1 ? "is " : "are ") + count(n, word) + ".");
    if (n < min) return fail(n === 0 ? "There are no " + word + "s yet." : "There " + (n === 1 ? "is " : "are ") + count(n, word) + "; there should be at least " + min + ".");
    return fail("There are " + count(n, word) + "; there should be no more than " + max + ".");
  };

  // ---------- list.is ----------
  const FORMAT_WORDS = { decimal: "1, 2, 3", lowerLetter: "a, b, c", upperLetter: "A, B, C", lowerRoman: "i, ii, iii", upperRoman: "I, II, III", bullet: "bullets" };
  CHECKS["list.is"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    for (const t of sel.targets) {
      const L = t.p.list;
      const name = '"' + snippet(t.p.text, 30) + '"';
      const where = { search: t.p.text.trim().slice(0, 60) };
      if (!L) return fail(name + " isn't in a list yet.", where);
      if (params.kind !== undefined && L.kind !== params.kind) return fail(name + " is in a " + (L.kind === "bullet" ? "bulleted" : "numbered") + " list. It should be " + (params.kind === "bullet" ? "bulleted." : "numbered."), where);
      if (params.level !== undefined && L.level !== Number(params.level)) return fail(name + " is at list level " + L.level + ". It should be at level " + params.level + ".", where);
      if (params.format !== undefined && L.format !== params.format) return fail(name + " is numbered " + (FORMAT_WORDS[L.format] || L.format) + ". It should be " + (FORMAT_WORDS[params.format] || params.format) + ".", where);
      if (params.symbol !== undefined && L.symbol !== params.symbol) return fail("The bullet on " + name + " is " + (L.symbol || "a number") + ". It should be " + params.symbol + ".", where);
      if (params.pattern !== undefined && !loose(L.pattern.replace(/%\d/g, "#")).includes(loose(String(params.pattern).replace(/%\d/g, "#")))) {
        return fail("The number on " + name + ' looks like "' + L.pattern.replace(/%(\d)/g, "1") + '". It should look like "' + String(params.pattern).replace(/%(\d)/g, "1") + '".', where);
      }
      if (params.number !== undefined && L.number !== Number(params.number)) return fail(name + " is numbered " + L.number + ". It should be " + params.number + ".", where);
    }
    return pass(sel.targets.length === 1 ? "The list is set up correctly." : "All " + sel.targets.length + " list items are set up correctly.");
  };

  // ---------- notes.is (footnotes and endnotes) ----------
  CHECKS["notes.is"] = function (model, params) {
    const kind = params.kind;
    const n = Math.max(model.notes[kind].length, model.noteRefs[kind]);
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    if (n < min) return fail(n === 0 ? "There are no " + kind + "s yet." : "There " + (n === 1 ? "is " : "are ") + count(n, kind) + "; there should be at least " + min + ".");
    if (n > max) return fail("There are " + count(n, kind) + "; there should be no more than " + max + ".");
    if (params.format !== undefined && model.noteFormats[kind] !== params.format)
      return fail("The " + kind + "s are numbered " + (FORMAT_WORDS[model.noteFormats[kind]] || model.noteFormats[kind]) + ". They should be " + (FORMAT_WORDS[params.format] || params.format) + ".");
    return pass("The " + kind + "s are in place" + (params.format ? " and numbered " + FORMAT_WORDS[params.format] : "") + ".");
  };

  // ---------- toc.is ----------
  CHECKS["toc.is"] = function (model, params) {
    const toc = model.fields.find((f) => /^TOC\b/i.test(f.code) && f.part === "body");
    if (!toc) return fail("There's no automatic table of contents yet. Insert one from References > Table of Contents.");
    if (params.levels !== undefined) {
      const m = /\\o\s+"(\d)-(\d)"/.exec(toc.code);
      const levels = m ? Number(m[2]) : 9;
      if (levels !== Number(params.levels)) return fail("The table of contents shows " + levels + " heading levels. It should show " + params.levels + ".");
    }
    if (params.includes !== undefined) {
      const entries = model.paragraphs.filter((p) => /^toc ?\d$/i.test(styleNameOf(model, p)) || /^TOC\d$/i.test(paragraphStyleId(model, p) || ""));
      if (!entries.some((p) => loose(p.text).includes(loose(params.includes)))) return fail('The table of contents doesn\'t list "' + params.includes + '" yet. Update the table (Update Table > Update entire table).');
    }
    return pass("The table of contents is in place.");
  };

  // ---------- citation.is ----------
  CHECKS["citation.is"] = function (model, params) {
    const cites = model.fields.filter((f) => /^CITATION\b/i.test(f.code)).length;
    const bib = model.fields.some((f) => /^BIBLIOGRAPHY\b/i.test(f.code));
    const done = [];
    if (params.sources !== undefined && model.sources.count < Number(params.sources))
      return fail("The citation manager has " + count(model.sources.count, "source") + ". Add at least " + params.sources + " (References > Manage Sources).");
    if (params.citations !== undefined && cites < Number(params.citations))
      return fail("There " + (cites === 1 ? "is " : "are ") + count(cites, "in-text citation") + "; there should be at least " + params.citations + ". Use References > Insert Citation.");
    if (params.bibliography !== undefined && bib !== yes(params.bibliography))
      return fail(bib ? "Remove the bibliography." : "There's no reference list yet. Insert one from References > Bibliography.");
    if (params.citeStyle !== undefined && !String(model.sources.style).toUpperCase().includes(String(params.citeStyle).toUpperCase()))
      return fail("The citation style is " + (model.sources.style || "not set") + ". Choose " + params.citeStyle + " in References > Style.");
    if (params.sources !== undefined) done.push(count(model.sources.count, "source") + " in the citation manager");
    if (params.citations !== undefined) done.push(count(cites, "in-text citation"));
    if (params.bibliography !== undefined && bib) done.push("a reference list");
    if (params.citeStyle !== undefined) done.push("the " + params.citeStyle + " style");
    return pass("Done: " + done.join(", ") + ".");
  };

  // ---------- object.count / object.is (pictures, shapes, SmartArt...) ----------
  const KIND_WORDS = { any: "graphic", picture: "picture", icon: "icon", shape: "shape", textbox: "text box", smartart: "SmartArt graphic", chart: "chart", model3d: "3D model", group: "group" };
  function objectsOfKind(model, kind) {
    return model.objects.filter((o) => kind === "any" ? o.kind !== "group" : o.kind === kind);
  }
  CHECKS["object.count"] = function (model, params) {
    const n = objectsOfKind(model, params.kind).length;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const word = KIND_WORDS[params.kind];
    if (n >= min && n <= max) return pass("There " + (n === 1 ? "is " : "are ") + count(n, word) + ".");
    if (n < min) return fail(n === 0 ? "There's no " + word + " in the document yet." : "There " + (n === 1 ? "is " : "are ") + count(n, word) + "; there should be at least " + min + ".");
    return fail("There are " + count(n, word) + "; there should be no more than " + max + ".");
  };
  const WRAP_WORDS = { inline: "In Line with Text", square: "Square", tight: "Tight", through: "Through", topAndBottom: "Top and Bottom", behind: "Behind Text", inFront: "In Front of Text" };
  CHECKS["object.is"] = function (model, params) {
    let list = objectsOfKind(model, params.kind);
    const word = KIND_WORDS[params.kind];
    if (params.contains !== undefined) list = list.filter((o) => loose(o.text).includes(loose(params.contains)));
    if (params.name !== undefined) list = list.filter((o) => loose(o.name).includes(loose(params.name)));
    if (params.index !== undefined) list = list[Number(params.index) - 1] ? [list[Number(params.index) - 1]] : [];
    if (!list.length) return fail(params.contains !== undefined ? 'I don\'t see a ' + word + ' with "' + params.contains + '" in it yet.' :
      params.name !== undefined ? 'I can\'t find the ' + word + ' named "' + params.name + '". If you deleted it, load a fresh copy of the practice file.' : "There's no " + word + " in the document yet.");
    // any=true: passes if at least one of them is right (for "apply a style to the picture you inserted").
    if (yes(params.any)) {
      const tries = list.map((o) => objectProblem(o, params, "The " + word));
      if (tries.some((t) => !t)) return pass("The " + word + " is set up correctly.");
      return fail(list.length === 1 ? tries[0] : "None of the " + count(list.length, word) + " is set up yet. For example: " + tries[0].charAt(0).toLowerCase() + tries[0].slice(1));
    }
    for (let k = 0; k < list.length; k++) {
      const problem = objectProblem(list[k], params, list.length > 1 ? "The " + ordinal(k + 1) + " " + word : "The " + word);
      if (problem) return fail(problem);
    }
    return pass(list.length === 1 ? "The " + word + " is set up correctly." : "All " + list.length + " " + word + "s are set up correctly.");
  };
  /** What's wrong with one object, in words, or "" if nothing is. */
  function objectProblem(o, params, name) {
    const problem = (function () {
      if (params.alt !== undefined && yes(params.alt) && !o.alt && !o.decorative) return name + (o.name ? ' ("' + o.name + '")' : "") + " has no alt text yet.";
      if (params.decorative !== undefined && o.decorative !== yes(params.decorative)) return name + (o.decorative ? " is marked as decorative. It needs alt text instead." : " isn't marked as decorative yet.");
      if (params.wrap !== undefined && o.wrap !== params.wrap) return name + "'s text wrapping is " + WRAP_WORDS[o.wrap] + ". It should be " + WRAP_WORDS[params.wrap] + ".";
      if (params.halign !== undefined && o.halign !== params.halign) return name + " isn't lined up on the " + params.halign + (o.wrap === "inline" ? ". It's in line with text, so it can't be positioned yet: change its wrapping first." : " yet. Use Position, or Align > Align " + params.halign.charAt(0).toUpperCase() + params.halign.slice(1) + ".");
      if (params.valign !== undefined && o.valign !== params.valign) return name + " isn't at the " + params.valign + " of the page yet.";
      if (params.pictureStyle !== undefined && o.style !== yes(params.pictureStyle)) return name + " doesn't have a picture style yet. Pick one from Picture Format > Picture Styles.";
      if (params.artistic !== undefined && o.artistic !== yes(params.artistic)) return name + " doesn't have an artistic effect yet.";
      if (params.bgRemoved !== undefined && o.bgRemoved !== yes(params.bgRemoved)) return name + "'s background hasn't been removed yet.";
      if (params.fill !== undefined && !colorMatches(o.fill, params.fill)) return name + "'s fill " + (o.fill ? "is " + showColor(o.fill) : "is still the default") + ". " + (params.fill === "changed" ? "Change it." : "It should be " + params.fill + ".");
      if (params.outline !== undefined && !colorMatches(o.outline, params.outline)) return name + "'s outline " + (o.outline ? "is " + showColor(o.outline) : "is still the default") + ". " + (params.outline === "changed" ? "Change it." : "It should be " + params.outline + ".");
      if (params.says !== undefined && !loose(o.text).includes(loose(params.says))) return name + ' doesn\'t say "' + params.says + '" yet.';
      if (params.layout !== undefined && !o.layout.toLowerCase().includes(String(params.layout).toLowerCase())) return name + " doesn't use a " + params.layout + " layout.";
      if (params.items !== undefined && o.items < Number(params.items)) return name + " has " + count(o.items, "item") + " with text. It should have at least " + params.items + ".";
      if (params.colors !== undefined && o.colorsChanged !== yes(params.colors)) return name + "'s colors are still the default. Use Change Colors.";
      if (params.smartStyle !== undefined && o.styleChanged !== yes(params.smartStyle)) return name + "'s style is still the default. Pick another SmartArt style.";
      if (params.chartType !== undefined && o.chartType !== params.chartType) return name + " is a " + (o.chartType || "different kind of") + " chart. It should be a " + params.chartType + " chart.";
      return "";
    })();
    return problem;
  }
  function ordinal(n) { return ["first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth"][n - 1] || n + "th"; }
  function colorMatches(actual, wanted) {
    if (wanted === "changed") return !!actual;
    if (wanted === "none") return actual === "none";
    return !!actual && actual.toUpperCase() === String(wanted).toUpperCase();
  }
  function showColor(c) { return c === "none" ? "No Fill" : /^theme:/.test(c) ? "a theme color (" + c.slice(6) + ")" : c === "changed" ? "a gradient or pattern" : c; }

  // ---------- table.count / table.is ----------
  CHECKS["table.count"] = function (model, params) {
    // contains="text": count only the tables with that text in them.
    const n = model.tables.filter((t) => params.contains === undefined || loose(t.text).includes(loose(params.contains))).length;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const what = params.contains !== undefined ? ' with "' + snippet(params.contains, 24) + '" in ' + (n === 1 ? "it" : "them") : "";
    if (n >= min && n <= max) return pass(n === 0 ? "There are no tables" + (params.contains !== undefined ? ' with "' + snippet(params.contains, 24) + '"' : "") + " left." : "There " + (n === 1 ? "is " : "are ") + count(n, "table") + what + ".");
    if (n < min) return fail(n === 0 ? "There's no table" + (params.contains !== undefined ? ' with "' + snippet(params.contains, 24) + '" in it' : "") + " yet." : "There " + (n === 1 ? "is " : "are ") + count(n, "table") + what + "; there should be at least " + min + ".");
    return fail("There " + (n === 1 ? "is still " : "are ") + count(n, "table") + what + "; there should be " + (max === 0 ? "none. Convert it to text (Table Layout > Convert to Text)." : "no more than " + max + "."));
  };
  CHECKS["table.is"] = function (model, params) {
    // Which table: one with some text in it (contains=, and lacks= text it must not have),
    // the nth table (index=), or with any=true, any table that passes.
    let pool = model.tables.filter((t) => (params.contains === undefined || loose(t.text).includes(loose(params.contains))) &&
      (params.lacks === undefined || !loose(t.text).includes(loose(params.lacks))));
    if (params.index !== undefined) pool = pool[Number(params.index) - 1] ? [pool[Number(params.index) - 1]] : [];
    if (!pool.length) {
      if (params.contains !== undefined && params.lacks !== undefined) return fail('I don\'t see a table with "' + params.contains + '" but not "' + params.lacks + '" yet.');
      return fail(params.contains !== undefined ? 'I don\'t see a table with "' + params.contains + '" in it yet.' : "There's no table yet.");
    }
    if (yes(params.any)) {
      const results = pool.map((t) => checkOneTable(t, params));
      const ok = results.find((r) => r.pass);
      if (ok) return ok;
      if (params.rows !== undefined || params.cols !== undefined) {
        const size = [params.cols !== undefined ? count(Number(params.cols), "column") : "", params.rows !== undefined ? count(Number(params.rows), "row") : ""].filter(Boolean).join(" and ");
        return fail("I don't see a table with " + size + " yet.");
      }
      return results[0];
    }
    return checkOneTable(pool[0], params);
  };
  function checkOneTable(table, params) {
    const rows = table.rows.length;
    if (params.rows !== undefined && rows !== Number(params.rows)) return fail("The table has " + count(rows, "row") + ". It should have " + params.rows + ".");
    if (params.minRows !== undefined && rows < Number(params.minRows)) return fail("The table has " + count(rows, "row") + ". It should have at least " + params.minRows + ".");
    if (params.cols !== undefined && table.cols !== Number(params.cols)) return fail("The table has " + count(table.cols, "column") + ". It should have " + params.cols + ".");
    if (params.headerRepeat !== undefined && (table.rows[0] && table.rows[0].header) !== yes(params.headerRepeat))
      return fail("The top row isn't set to repeat as a header row yet (Table Layout > Repeat Header Rows).");
    if (params.merged !== undefined) {
      const merged = table.rows.some((r) => r.cells.some((c) => c.span > 1 || c.vMerge));
      if (merged !== yes(params.merged)) return fail(merged ? "Some cells are still merged." : "No cells are merged yet.");
    }
    if (params.sortedBy !== undefined) {
      const col = Number(params.sortedBy) - 1;
      // Sort the data rows: skip header rows, and title rows that span the whole table.
      const data = table.rows.filter((r) => !r.header && !(r.cells.length === 1 && r.cells[0].span > 1));
      let values = data.map((r) => cellAtColumn(r, col)).filter((v) => v !== null && v !== "");
      // Leave out leading rows that aren't the same kind of value as the rest (a title or column labels).
      const kinds = values.map(valueKind);
      const main = ["number", "date", "text"].map((k) => [k, kinds.filter((x) => x === k).length]).sort((a, b) => b[1] - a[1])[0][0];
      while (values.length && valueKind(values[0]) !== main) values = values.slice(1);
      const desc = params.order === "desc";
      for (let i = 1; i < values.length; i++) {
        const cmp = compareCells(values[i - 1], values[i]);
        if ((desc && cmp < 0) || (!desc && cmp > 0)) return fail("The table isn't sorted by column " + params.sortedBy + " yet: \"" + snippet(values[i - 1], 24) + '" comes before "' + snippet(values[i], 24) + '".');
      }
    }
    if (params.cellMargin !== undefined) {
      const want = toPoints(params.cellMargin) * TWIPS_PER_PT;
      const l = table.cellMarginLeft, r = table.cellMarginRight;
      if (l === null || r === null || Math.abs(l - want) > 15 || Math.abs(r - want) > 15)
        return fail("The left and right cell margins are " + (l === null ? "the default (0.08 in)" : showLength(l / TWIPS_PER_PT, unitOf(params.cellMargin))) + ". They should be " + showLength(want / TWIPS_PER_PT, unitOf(params.cellMargin)) + " (Table Layout > Cell Margins).");
    }
    if (params.cellSpacing !== undefined) {
      const want = toPoints(params.cellSpacing) * TWIPS_PER_PT;
      if (Math.abs(table.cellSpacing - want) > 15) return fail("The space between cells is " + showLength(table.cellSpacing / TWIPS_PER_PT, unitOf(params.cellSpacing)) + ". It should be " + params.cellSpacing.replace(/(in|pt|cm)$/, " $1") + ".");
    }
    if (params.width !== undefined) {
      const col = Number(params.col || 1) - 1;
      const want = toPoints(params.width) * TWIPS_PER_PT;
      const grid = table.grid[col];
      const cell = table.rows.map((r) => cellObjAt(r, col)).find((c) => c && c.span === 1 && c.widthType === "dxa");
      const actual = cell ? cell.width : grid;
      if (!actual || Math.abs(actual - want) > 72) return fail("Column " + (col + 1) + " is " + (actual ? showLength(actual / TWIPS_PER_PT, "in") : "an automatic width") + " wide. It should be " + showLength(want / TWIPS_PER_PT, "in") + ".");
    }
    return pass("The table is set up correctly.");
  }
  function cellObjAt(row, col) {
    let at = 0;
    for (const c of row.cells) { if (at === col) return c; at += c.span; if (at > col) return null; }
    return null;
  }
  function cellAtColumn(row, col) { const c = cellObjAt(row, col); return c ? c.text.trim() : null; }
  /** Is a cell a number, a date, or plain text? (Browsers will "parse" almost anything as a date, so we look for real date shapes.) */
  const MONTHS = "(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\\.?";
  const DATE_SHAPES = [new RegExp("^" + MONTHS + "\\s+\\d{1,2}(,?\\s+\\d{4})?$", "i"), new RegExp("^\\d{1,2}\\s+" + MONTHS + "(\\s+\\d{4})?$", "i"), /^\d{1,4}[\/-]\d{1,2}[\/-]\d{1,4}$/];
  function valueKind(v) {
    const t = String(v || "").trim();
    if (t && !isNaN(Number(t.replace(/[$,%]/g, "")))) return "number";
    if (DATE_SHAPES.some((re) => re.test(t)) && !isNaN(Date.parse(t))) return "date";
    return "text";
  }
  /** Compare two cells the way Sort does: as numbers, then dates, then text. */
  function compareCells(a, b) {
    const ka = valueKind(a), kb = valueKind(b);
    if (ka === "number" && kb === "number") return Number(a.replace(/[$,%]/g, "")) - Number(b.replace(/[$,%]/g, ""));
    if (ka === "date" && kb === "date") return Date.parse(a) - Date.parse(b);
    return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
  }

  // ---------- theme.is ----------
  CHECKS["theme.is"] = function (model, params) {
    if (params.colors !== undefined && !sameName(model.theme.colorsName, params.colors))
      return fail('The theme colors are "' + (model.theme.colorsName || "Office") + '". Save your own set named "' + params.colors + '" (Design > Colors > Customize Colors).');
    for (let i = 1; i <= 6; i++) {
      const key = "accent" + i;
      if (params[key] !== undefined && (model.theme.colors[key] || "").toUpperCase() !== String(params[key]).toUpperCase())
        return fail("Accent " + i + " is " + (model.theme.colors[key] || "not set") + ". It should be " + String(params[key]).toUpperCase() + ".");
    }
    return pass("The theme colors are right.");
  };

  // ---------- page.background ----------
  CHECKS["page.background"] = function (model, params) {
    if (params.watermark !== undefined) {
      const marks = model.background.watermarks;
      if (!marks.length) return fail("There's no watermark yet (Design > Watermark).");
      if (params.watermark !== "any" && !marks.some((m) => loose(m).includes(loose(params.watermark)))) return fail('The watermark says "' + marks[0] + '". It should say "' + params.watermark + '".');
    }
    if (params.color !== undefined) {
      const c = model.background.color;
      if (!c) return fail("The page color is still white (Design > Page Color).");
      if (params.color !== "any" && c !== String(params.color).toUpperCase()) return fail("The page color is " + c + ". It should be " + params.color + ".");
    }
    return pass("The page background is set up.");
  };

  // ---------- tabs.is ----------
  CHECKS["tabs.is"] = function (model, params) {
    const sel = select(model, params);
    if (sel.missing) return fail(sel.missing);
    for (const t of sel.targets) {
      const p = t.p, name = '"' + snippet(p.text, 30) + '"';
      const stops = [];
      paragraphPropertyLevels(model, p).forEach((level) => children(child(level, "tabs"), "tab").forEach((tab) => stops.push(tab)));
      const live = stops.filter((tab) => attr(tab, "val") !== "clear");
      const match = live.find((tab) => {
        const pos = Number(attr(tab, "pos")) / TWIPS_PER_PT;
        const align = { end: "right", start: "left" }[attr(tab, "val")] || attr(tab, "val");
        return (params.position === undefined || Math.abs(pos - toPoints(params.position)) <= 3.6) &&
          (params.align === undefined || align === params.align) &&
          (params.leader === undefined || (attr(tab, "leader") || "none") === params.leader);
      });
      if (!match) {
        const desc = [params.align ? params.align : "", "tab stop", params.position ? "at " + params.position.replace(/(in|pt|cm)$/, " $1") : "", params.leader && params.leader !== "none" ? "with a " + params.leader + " leader" : ""].filter(Boolean).join(" ");
        return fail(name + " doesn't have a " + desc + " yet." + (live.length ? "" : " (It has no tab stops.)"), { search: p.text.trim().slice(0, 60) });
      }
      if (!p.text.includes("\t")) return fail(name + " has the tab stop, but no tab yet. Press Tab before the date.", { search: p.text.trim().slice(0, 60) });
    }
    return pass("The tab stops are set up.");
  };

  // ---------- controls.filled (template placeholders) ----------
  CHECKS["controls.filled"] = function (model) {
    if (model.placeholders === 0) return pass("Every placeholder has been replaced.");
    return fail(count(model.placeholders, "placeholder") + " still " + (model.placeholders === 1 ? "shows" : "show") + " the template's sample text. Click each one and type your own.");
  };

  // ---------- merge.is (mail merge) ----------
  const MERGE_TYPES = { letters: "formLetters", labels: "mailingLabels", envelopes: "envelopes", directory: "catalog", email: "email" };
  const MERGE_WORDS = { formLetters: "letters", mailingLabels: "labels", envelopes: "envelopes", catalog: "a directory", email: "email messages" };
  CHECKS["merge.is"] = function (model, params) {
    const mm = model.settings ? model.settings.getElementsByTagNameNS(NS.w, "mailMerge")[0] : null;
    if (!mm) return fail("This isn't a mail merge document yet. Start one from Mailings > Start Mail Merge.");
    const type = attr(child(mm, "mainDocumentType"), "val") || "formLetters";
    if (params.type !== undefined && type !== MERGE_TYPES[params.type]) return fail("This merge is set up for " + (MERGE_WORDS[type] || type) + ". It should be for " + MERGE_WORDS[MERGE_TYPES[params.type]] + ".");
    if (params.source !== undefined && yes(params.source) && !child(mm, "dataSource") && !child(mm, "query") && !child(mm, "connectString"))
      return fail("No recipient list is connected yet. Use Mailings > Select Recipients > Use an Existing List.");
    return pass("The mail merge is set up for " + (MERGE_WORDS[type] || type) + ".");
  };

  // ---------- comment.count / comment.is ----------
  CHECKS["comment.count"] = function (model, params) {
    const n = model.comments.filter((c) => !c.parent).length;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1); // only a maximum given: none is fine
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    if (n >= min && n <= max) return pass(n === 0 ? "There are no comments left." : "There " + (n === 1 ? "is " : "are ") + count(n, "comment") + ".");
    if (n < min) return fail("There " + (n === 1 ? "is " : "are ") + count(n, "comment") + "; there should be at least " + min + ".");
    return fail(max === 0 ? "There " + (n === 1 ? "is still 1 comment" : "are still " + n + " comments") + " in the document." : "There are " + count(n, "comment") + "; there should be no more than " + max + ".");
  };
  CHECKS["comment.is"] = function (model, params) {
    const state = params.state || "present";
    const matches = (c) => (params.says === undefined || loose(c.text).includes(loose(params.says))) &&
      (params.anchor === undefined || loose(c.anchor).includes(loose(params.anchor))) &&
      (params.author === undefined || loose(c.author).includes(loose(params.author)));
    const pool = state === "present" || state === "absent" ? model.comments : model.comments.filter((c) => !c.parent);
    const found = pool.filter(matches);
    const what = params.says !== undefined ? 'a comment saying "' + snippet(params.says, 30) + '"' : params.anchor !== undefined ? 'a comment on "' + snippet(params.anchor, 30) + '"' : "a comment by " + params.author;
    if (state === "absent") return found.length ? fail("There's still " + what + ". Delete it.", { search: (found[0].anchor || "").slice(0, 60) }) : pass("That comment is gone.");
    if (!found.length) return fail("I don't see " + what + " yet.");
    if (state === "replied" && !found.some((c) => c.replies.length)) return fail(what.charAt(0).toUpperCase() + what.slice(1) + " doesn't have a reply yet.", { search: (found[0].anchor || "").slice(0, 60) });
    if (state === "resolved" && !found.some((c) => c.done)) return fail(what.charAt(0).toUpperCase() + what.slice(1) + " isn't resolved yet.", { search: (found[0].anchor || "").slice(0, 60) });
    if (state === "open" && !found.some((c) => !c.done)) return fail(what.charAt(0).toUpperCase() + what.slice(1) + " is resolved. It should still be open.");
    return pass({ present: "The comment is there.", replied: "The comment has a reply.", resolved: "The comment is resolved.", open: "The comment is still open." }[state]);
  };

  // ---------- revision.count (tracked changes) ----------
  CHECKS["revision.count"] = function (model, params) {
    const kind = params.kind || "any";
    const list = model.revisions.filter((r) => (kind === "any" || r.kind === kind) && (params.author === undefined || loose(r.author).includes(loose(params.author))));
    const n = list.length;
    const min = params.min !== undefined ? Number(params.min) : (params.max !== undefined ? 0 : 1);
    const max = params.max !== undefined ? Number(params.max) : Infinity;
    const word = { any: "tracked change", insert: "tracked insertion", delete: "tracked deletion", format: "tracked formatting change", move: "tracked move" }[kind];
    const by = params.author !== undefined ? " by " + params.author : "";
    if (n >= min && n <= max) return pass(n === 0 ? "There are no " + word + "s" + by + " left." : "There " + (n === 1 ? "is " : "are ") + count(n, word) + by + ".");
    if (n < min) return fail((n === 0 ? "There are no " + word + "s" + by + " yet" : "There " + (n === 1 ? "is " : "are ") + count(n, word) + by + "; there should be at least " + min) + (kind === "any" ? ". Is Track Changes on?" : "."));
    return fail("There " + (n === 1 ? "is still " : "are still ") + count(n, word) + by + (list[0].text ? ', starting with "' + snippet(list[0].text, 30) + '"' : "") + ". Accept or reject " + (n === 1 ? "it." : "them."),
      list[0].text ? { search: list[0].text.slice(0, 60) } : null);
  };

  // ---------- doc.is (tracking, protection, compatibility, personal info) ----------
  const PROTECT_WORDS = { readOnly: "read only (no changes)", comments: "comments only", trackedChanges: "tracked changes", forms: "filling in forms", none: "no restriction" };
  CHECKS["doc.is"] = function (model, params) {
    const st = model.settings ? model.settings.documentElement : null;
    if (params.tracking !== undefined) {
      const on = onOff(child(st, "trackRevisions")) === true;
      if (on !== yes(params.tracking)) return fail(on ? "Track Changes is still on." : "Track Changes isn't on yet (Review > Track Changes).");
    }
    if (params.protection !== undefined) {
      const prot = child(st, "documentProtection");
      // Restrictions only count once "Yes, Start Enforcing Protection" is selected (enforcement="1").
      const enforced = !!prot && ["1", "true", "on"].includes(String(attr(prot, "enforcement")).toLowerCase());
      const kind = enforced ? attr(prot, "edit") || "none" : "none";
      if (kind !== params.protection) return fail(kind === "none" ? "Editing isn't restricted yet (Review > Restrict Editing, then Yes, Start Enforcing Protection)." : "Editing is restricted to " + PROTECT_WORDS[kind] + ". It should be " + PROTECT_WORDS[params.protection] + ".");
    }
    if (params.password !== undefined) {
      const prot = child(st, "documentProtection");
      const has = !!prot && !!(attr(prot, "hashValue") || attr(prot, "hash") || attr(prot, "cryptProviderType"));
      if (has !== yes(params.password)) return fail(has ? "The protection has a password. Remove it." : "The protection doesn't have a password yet.");
    }
    if (params.compatMode !== undefined) {
      const setting = Array.from(st ? st.getElementsByTagNameNS(NS.w, "compatSetting") : []).find((c) => attr(c, "name") === "compatibilityMode");
      const mode = setting ? Number(attr(setting, "val")) : 12;
      if (mode < Number(params.compatMode)) return fail("The document is still in Compatibility Mode (made for an older version of Word). Use File > Info > Convert.");
    }
    if (params.personalInfo !== undefined) {
      const removed = onOff(child(st, "removePersonalInformation")) === true || (!model.props.author && !model.props.lastModifiedBy);
      if (params.personalInfo === "removed" && !removed) return fail('The file still names its author ("' + snippet(model.props.author || model.props.lastModifiedBy, 30) + '"). Remove document properties and personal information with Inspect Document.');
    }
    return pass("The document settings are right.");
  };

  // ---------- text.hidden ----------
  CHECKS["text.hidden"] = function (model, params) {
    const max = params.max !== undefined ? Number(params.max) : 0;
    const n = model.hiddenRuns;
    if (n <= max) return pass(n === 0 ? "There's no hidden text." : "Hidden text is within the limit.");
    return fail("There " + (n === 1 ? "is still 1 piece" : "are still " + n + " pieces") + " of hidden text. Show formatting marks to see it, or remove it with Inspect Document.");
  };

  // ---------- access.is (what the Accessibility Checker looks for) ----------
  CHECKS["access.is"] = function (model, params) {
    const rules = String(params.rules || "alt,tableHeaders,title,headings,links").split(",").map((r) => r.trim());
    const problems = [];
    if (rules.includes("alt")) {
      const missing = model.objects.filter((o) => o.kind !== "group" && !o.alt && !o.decorative);
      if (missing.length) problems.push(count(missing.length, "graphic") + " " + (missing.length === 1 ? "has" : "have") + " no alt text (and " + (missing.length === 1 ? "isn't" : "aren't") + " marked decorative)");
    }
    if (rules.includes("tableHeaders")) {
      const bad = model.tables.filter((t) => !(t.rows[0] && t.rows[0].header));
      if (bad.length) problems.push(count(bad.length, "table") + " " + (bad.length === 1 ? "has" : "have") + " no header row set (Repeat Header Rows)");
    }
    if (rules.includes("title") && !model.props.title.trim()) problems.push("the document has no Title property");
    if (rules.includes("headings")) {
      const levels = model.paragraphs.map((p) => { const m = /^heading (\d)$/i.exec(styleNameOf(model, p)); return m ? Number(m[1]) : 0; }).filter(Boolean);
      if (!levels.length) problems.push("there are no headings (use Heading 1, Heading 2...)");
      else {
        let prev = 0;
        for (const l of levels) { if (l > prev + 1) { problems.push("a heading skips a level (Heading " + prev + " to Heading " + l + ")"); break; } prev = l; }
      }
    }
    if (rules.includes("links")) {
      const vague = model.links.filter((l) => !l.fromField && vagueLinkText(l.text));
      if (vague.length) problems.push(count(vague.length, "link") + " " + (vague.length === 1 ? "has" : "have") + ' text that doesn\'t say where it goes ("' + snippet(vague[0].text || "(empty)", 24) + '")');
    }
    const RULE_WORDS = { alt: "alt text", tableHeaders: "table header rows", title: "the document title", headings: "heading levels", links: "link text" };
    if (!problems.length) return pass(params.rules ? "No problems with " + rules.map((r) => RULE_WORDS[r] || r).join(", ") + "." : "No accessibility problems found.");
    const first = problems[0].charAt(0).toUpperCase() + problems[0].slice(1);
    return fail(problems.length === 1 ? first + "." : first + "; " + problems.slice(1).join("; ") + ".");
  };

  // =====================================================================
  // PART 6: RUNNING CHECKS
  // =====================================================================

  /**
   * Run one check. `check` is what TrainerMD.parse makes from a "check:" line:
   * { name: "para.is", params: { para: "...", align: "left" } }.
   * A check that crashes never stops the others: it fails with a friendly note.
   */
  function runCheck(model, check, context) {
    const fn = CHECKS[check.name];
    if (!fn) return fail('This step uses a check the trainer doesn\'t know ("' + check.name + '"). Please tell your instructor.');
    try {
      return fn(model, check.params || {}, context || {});
    } catch (e) {
      if (root.console) root.console.error("Check failed to run:", check, e);
      return fail("Something went wrong checking this step. Please tell your instructor.");
    }
  }

  /**
   * Run every check on one step. A step passes when all its checks pass.
   * Returns { pass, messages: [...], where }.
   */
  function runStep(model, step, context) {
    const results = step.checks.map((c) => runCheck(model, c, context));
    const failed = results.filter((r) => !r.pass);
    if (!failed.length) {
      // Every check passed: show each one's (short) confirmation once.
      const messages = Array.from(new Set(results.map((r) => r.message)));
      return { pass: true, messages: messages, results: results };
    }
    return { pass: false, messages: failed.map((r) => r.message), where: failed[0].where, results: results };
  }

  const api = { VERSION, readDocx, runCheck, runStep, CHECKS: Object.keys(CHECKS), toPoints,
    // exposed for tests and for future tools:
    _internal: { isEmptyParagraph, runFontName, runSize, runEffects, directFormatting, paragraphLineSpacing, paragraphSpace, paragraphFirstLine, paragraphAlign, wildcard } };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.WordChecker = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
