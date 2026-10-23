/* =============================================================================
   excel-checker.js - the Excel checker library (v0.1)
   -----------------------------------------------------------------------------
   This is the ONLY code that decides whether an Excel step passed. Lesson files
   name a check ("cell.is", "formula.is" ...) and give it values; they can't add
   logic. It works exactly like the Word checker (word/tools/word-checker.js):

     1. readXlsx(bytes) opens a saved .xlsx file and builds a "model" of it:
        the sheets, their cells (values, formulas, formatting), tables, charts,
        comments, page setup and so on.
     2. runStep(model, step) runs a step's checks against that model and
        returns { pass, messages, where }.

   Why read the .xlsx file instead of asking Excel through Office.js?
     - Excel for the web, desktop Excel and "upload a saved file" all give us the
       same .xlsx bytes, so every learner is checked by exactly the same code.
     - The file holds things Office.js can't show on every platform (page setup,
       headers and footers, sparklines, scenarios, alt text on charts).
     - The same code can be tested in any browser with known-good and
       known-bad test workbooks (see excel/tools/check-tests.html).

   Principle: check the OUTCOME, not the clicks. A total is right whether the
   learner used AutoSum, typed =SUM(...) or used the status bar to check it -
   as long as the cell holds a formula that gives the right answer.

   A few Excel facts this file relies on:
     - Cells are stored row by row (<row><c r="B2" s="3" t="s"><v>5</v></c>).
       "s" is the cell's format (an index into the styles part), "t" is the type
       (s = shared string, str = formula text, b = true/false, e = error).
     - Excel saves the last calculated value of every formula in <v>, so we can
       check results without calculating anything ourselves.
     - When a formula is filled down, Excel saves it once ("shared formula") and
       the other cells just point at it. We rebuild each cell's own formula.
     - Newer functions are saved with a prefix: CONCAT is "_xlfn.CONCAT", SORT
       is "_xlfn._xlws.SORT". We remove the prefixes before comparing.

   Needs JSZip (loaded by the page) to open the .xlsx zip file.
   Safety: DOMParser in "application/xml" mode never runs scripts and never
   fetches outside files. Text from the workbook is returned as plain strings;
   the page must show it with textContent, never innerHTML. Nothing leaves the
   learner's device.
   ============================================================================= */
(function (root) {
  "use strict";

  const VERSION = "0.2.0";

  // XML namespaces: the long names behind tag prefixes like <c:chart> or <a:t>.
  const NS = {
    main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    rel: "http://schemas.openxmlformats.org/package/2006/relationships",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    xm: "http://schemas.microsoft.com/office/excel/2006/main",
    dc: "http://purl.org/dc/elements/1.1/",
    cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    tc: "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
  };

  // Safety limits for files we open (an .xlsx is a zip, and zips can "explode").
  const MAX_FILE_BYTES = 25 * 1024 * 1024;       // 25 MB as uploaded
  const MAX_UNZIPPED_BYTES = 120 * 1024 * 1024;  // 120 MB once unzipped
  const MAX_CELLS = 400000;                       // far more than any practice file

  // Paper sizes Excel saves as numbers.
  const PAPER = { 1: "letter", 5: "legal", 9: "a4" };
  const PAPER_LABEL = { letter: "Letter (8.5 x 11 in)", legal: "Legal (8.5 x 14 in)", a4: "A4" };

  // Excel's built-in number formats (the ones without their own code in the file).
  const BUILTIN_FORMATS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    5: "$#,##0_);($#,##0)", 6: "$#,##0_);[Red]($#,##0)", 7: "$#,##0.00_);($#,##0.00)", 8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yyyy h:mm",
    37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    41: "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", 42: "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
    43: "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", 44: "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
  };

  // Built-in cell style names, by Excel's builtinId (used when a file only saves the id).
  const BUILTIN_STYLES = {
    0: "Normal", 3: "Comma", 4: "Currency", 5: "Percent", 6: "Comma [0]", 7: "Currency [0]",
    8: "Hyperlink", 10: "Note", 11: "Warning Text", 15: "Title", 16: "Heading 1", 17: "Heading 2",
    18: "Heading 3", 19: "Heading 4", 20: "Input", 21: "Output", 22: "Calculation", 23: "Check Cell",
    24: "Linked Cell", 25: "Total", 26: "Good", 27: "Bad", 28: "Neutral", 53: "Explanatory Text",
  };

  // Each chart style's id when Excel first inserts that kind of chart (Excel 2013 and later).
  // "style=changed" compares with these. See the opportunities log: verify with real files.
  const DEFAULT_CHART_STYLE = { column: 201, bar: 216, line: 227, pie: 251, doughnut: 251, area: 276, scatter: 240 };

  // =====================================================================
  // PART 1: SMALL HELPERS
  // =====================================================================

  /** Turn XML text into a document we can search. Throws if the XML is damaged. */
  function parseXml(text) {
    const xml = new DOMParser().parseFromString(text, "application/xml");
    if (xml.getElementsByTagName("parsererror").length > 0) {
      throw new Error("Part of this Excel file is damaged and could not be read.");
    }
    return xml;
  }

  /** Direct children of `el` with this local tag name (any namespace prefix). */
  function kids(el, name) {
    if (!el) return [];
    return Array.from(el.children).filter((c) => c.localName === name);
  }
  /** The first direct child with this local tag name, or null. */
  function kid(el, name) {
    return kids(el, name)[0] || null;
  }
  /** Every element anywhere inside `el` with this local tag name. */
  function all(el, name) {
    if (!el) return [];
    const out = [];
    const walk = (node) => {
      for (const c of node.children) {
        if (c.localName === name) out.push(c);
        walk(c);
      }
    };
    walk(el);
    return out;
  }
  /** The first element anywhere inside `el` with this local name, or null. */
  function find(el, name) {
    return all(el, name)[0] || null;
  }
  /** Read an attribute with no prefix (Excel's own attributes), or null. */
  function at(el, name) {
    if (!el || !el.hasAttribute(name)) return null;
    return el.getAttribute(name);
  }
  /** Read a prefixed attribute such as r:id, or null. */
  function atNS(el, ns, name) {
    if (!el || !el.hasAttributeNS(ns, name)) return null;
    return el.getAttributeNS(ns, name);
  }
  /** Excel's on/off attributes: "1" or "true" is on. */
  function isOn(value) {
    return value === "1" || value === "true";
  }
  /** All the text inside an element (for example a shape's paragraphs). */
  function textOf(el) {
    if (!el) return "";
    return all(el, "t").map((t) => t.textContent).join("");
  }

  /**
   * Make text easy to compare: one kind of space, one kind of quote mark,
   * no case. Excel and AutoCorrect turn ' into ’ as you type.
   */
  function loose(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/[‘’‚′]/g, "'").replace(/[“”„″]/g, '"')
      .replace(/[   ]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
  }

  /** Short version of a long text for messages. */
  function snippet(text, max) {
    const t = String(text).replace(/\s+/g, " ").trim();
    const limit = max || 40;
    return t.length > limit ? t.slice(0, limit - 1).trimEnd() + "…" : t;
  }

  /** "1 cell" / "3 cells" */
  function count(n, word, plural) {
    return n + " " + (n === 1 ? word : (plural || word + "s"));
  }

  /** A comma-separated list in a lesson file ("SUM, AVERAGE") as an array. */
  function listOf(value) {
    return String(value || "").split(",").map((x) => x.trim()).filter(Boolean);
  }

  function pass(message) { return { pass: true, message: message }; }
  function fail(message, where) { return { pass: false, message: message, where: where || null }; }

  /** Wildcard pattern ("*_E1_Rentals.xlsx") to a regular expression. Not case sensitive. */
  function wildcard(pattern) {
    const escaped = String(pattern).replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".");
    return new RegExp("^" + escaped + "$", "i");
  }

  /** Is this a number (and not empty text)? */
  function isNumeric(v) {
    return v !== null && v !== undefined && v !== "" && !isNaN(Number(v));
  }

  /** A friendly number for messages: 1234.5 -> "1,234.5". */
  function showNumber(n) {
    if (!isNumeric(n)) return String(n);
    const x = Number(n);
    return (Math.round(x * 100) / 100).toLocaleString("en-US");
  }

  // ---------- cell addresses ----------

  /** "C" -> 3, "AA" -> 27 */
  function colNumber(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  /** 3 -> "C", 27 -> "AA" */
  function colLetters(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  /** "B12" -> { r: 12, c: 2 } (dollar signs are ignored). Null if it isn't an address. */
  function parseCell(ref) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
    return m ? { r: Number(m[2]), c: colNumber(m[1]) } : null;
  }
  /** { r: 12, c: 2 } -> "B12" */
  function cellName(r, c) { return colLetters(c) + r; }

  /**
   * "A1:D5", "B2", "B:D" (whole columns) or "2:4" (whole rows) ->
   * { r1, c1, r2, c2 }. Whole columns run to row 1048576, the last row.
   */
  function parseRange(ref) {
    const text = String(ref).trim().replace(/\$/g, "");
    let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(text);
    if (m) {
      const a = { r: Number(m[2]), c: colNumber(m[1]) };
      const b = m[3] ? { r: Number(m[4]), c: colNumber(m[3]) } : a;
      return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
    }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(text);
    if (m) return { r1: 1, c1: Math.min(colNumber(m[1]), colNumber(m[2])), r2: 1048576, c2: Math.max(colNumber(m[1]), colNumber(m[2])) };
    m = /^(\d+):(\d+)$/.exec(text);
    if (m) return { r1: Math.min(+m[1], +m[2]), c1: 1, r2: Math.max(+m[1], +m[2]), c2: 16384 };
    return null;
  }
  /** { r1, c1, r2, c2 } -> "A1:D5" (or "B2" for one cell). */
  function rangeName(g) {
    const a = cellName(g.r1, g.c1), b = cellName(g.r2, g.c2);
    return a === b ? a : a + ":" + b;
  }
  /** Do two ranges overlap? */
  function overlaps(g, h) {
    return g.r1 <= h.r2 && h.r1 <= g.r2 && g.c1 <= h.c2 && h.c1 <= g.c2;
  }
  /** Is range h completely inside range g? */
  function contains(g, h) {
    return h.r1 >= g.r1 && h.r2 <= g.r2 && h.c1 >= g.c1 && h.c2 <= g.c2;
  }
  /**
   * "June!B2", "'June 2027'!B2:B9" -> { sheet: "June", ref: "B2" }.
   * Without a sheet name, sheet is null.
   */
  function splitSheet(text) {
    const s = String(text).trim();
    const m = /^(?:'((?:[^']|'')+)'|([^!'\s][^!]*))!(.+)$/.exec(s);
    if (!m) return { sheet: null, ref: s };
    return { sheet: (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]), ref: m[3] };
  }
  /** A sqref ("A1:A5 C1:C5") as a list of ranges. */
  function parseSqref(sqref) {
    return String(sqref || "").split(/\s+/).map(parseRange).filter(Boolean);
  }

  // ---------- formulas ----------

  /** Remove the prefixes Excel saves on newer functions: _xlfn.CONCAT( -> CONCAT( */
  function cleanFormula(f) {
    return String(f || "").replace(/_xlfn\.|_xlws\.|_xlpm\./gi, "");
  }

  /**
   * Walk through a formula and call `onRef` for every cell, column or row
   * reference outside quoted text and outside [structured references]. onRef
   * gets the parts and returns the text to put in its place.
   * This is how shared formulas are rebuilt and how formulas are compared.
   */
  function mapRefs(formula, onRef) {
    const f = String(formula);
    let out = "";
    let i = 0;
    // A reference: optional $, 1-3 letters, optional $, digits; or a whole-column / whole-row range.
    const refRe = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\d(])|(\$?)(\d+):(\$?)(\d+)/y;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"') { // quoted text: copy it unchanged ("" is a quote inside the text)
        let j = i + 1;
        while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
        out += f.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === "'") { // a quoted sheet name: copy it unchanged
        let j = i + 1;
        while (j < f.length && !(f[j] === "'" && f[j + 1] !== "'")) j += f[j] === "'" ? 2 : 1;
        out += f.slice(i, j + 1); i = j + 1; continue;
      }
      if (ch === "[") { // structured reference like Rentals[@Hours]: copy unchanged
        let depth = 0, j = i;
        for (; j < f.length; j++) { if (f[j] === "[") depth++; if (f[j] === "]") { depth--; if (depth === 0) break; } }
        out += f.slice(i, j + 1); i = j + 1; continue;
      }
      const prev = i > 0 ? f[i - 1] : "";
      if (!/[A-Za-z0-9_.]/.test(prev)) {
        refRe.lastIndex = i;
        const m = refRe.exec(f);
        if (m) {
          const next = f[refRe.lastIndex] || "";
          // Not a reference if a letter, digit, _ or "(" follows (a function such as LOG10( or a name such as Q2Sales).
          if (!/[A-Za-z0-9_(]/.test(next)) {
            out += onRef(m);
            i = refRe.lastIndex;
            continue;
          }
        }
      }
      out += ch; i++;
    }
    return out;
  }

  /** Move a formula by dr rows and dc columns, the way filling it would. */
  function shiftFormula(formula, dr, dc) {
    return mapRefs(formula, (m) => {
      if (m[2] !== undefined) {
        const col = m[1] ? colNumber(m[2]) : colNumber(m[2]) + dc;
        const row = m[3] ? Number(m[4]) : Number(m[4]) + dr;
        return m[1] + colLetters(Math.max(1, col)) + m[3] + Math.max(1, row);
      }
      if (m[6] !== undefined) {
        const a = m[5] ? colNumber(m[6]) : colNumber(m[6]) + dc;
        const b = m[7] ? colNumber(m[8]) : colNumber(m[8]) + dc;
        return m[5] + colLetters(Math.max(1, a)) + ":" + m[7] + colLetters(Math.max(1, b));
      }
      const a = m[9] ? Number(m[10]) : Number(m[10]) + dr;
      const b = m[11] ? Number(m[12]) : Number(m[12]) + dr;
      return m[9] + Math.max(1, a) + ":" + m[11] + Math.max(1, b);
    });
  }

  /**
   * A formula written relative to its own cell ("R1C1" style): =B2*$B$1 in C2
   * becomes =RC[-1]*R1C2. Two cells hold "the same formula, filled" exactly
   * when these match.
   */
  function relativeFormula(formula, row, col) {
    return mapRefs(cleanFormula(formula).toUpperCase(), (m) => {
      if (m[2] !== undefined) {
        const c = colNumber(m[2]), r = Number(m[4]);
        return (m[3] ? "R" + r : "R[" + (r - row) + "]") + (m[1] ? "C" + c : "C[" + (c - col) + "]");
      }
      if (m[6] !== undefined) {
        const a = colNumber(m[6]), b = colNumber(m[8]);
        return (m[5] ? "C" + a : "C[" + (a - col) + "]") + ":" + (m[7] ? "C" + b : "C[" + (b - col) + "]");
      }
      const a = Number(m[10]), b = Number(m[12]);
      return (m[9] ? "R" + a : "R[" + (a - row) + "]") + ":" + (m[11] ? "R" + b : "R[" + (b - row) + "]");
    }).replace(/\s+/g, "");
  }

  /** The function names a formula uses: "=ROUND(SUM(A1:A3),2)" -> ["ROUND", "SUM"]. */
  function functionsIn(formula) {
    const names = [];
    const f = cleanFormula(formula).replace(/"(?:[^"]|"")*"/g, '""');
    const re = /([A-Za-z][A-Za-z0-9.]*)\s*\(/g;
    let m;
    while ((m = re.exec(f))) names.push(m[1].toUpperCase());
    return names;
  }

  // ---------- number formats ----------

  /** The number format code for a format id ("0.00%"), from the file or Excel's built-in list. */
  function formatCode(model, numFmtId) {
    const id = Number(numFmtId) || 0;
    return model.styles.numFmts[id] !== undefined ? model.styles.numFmts[id] : (BUILTIN_FORMATS[id] || "General");
  }

  /**
   * What kind of number format is this? general, number, currency, accounting,
   * percent, date, time, datetime, fraction, scientific or text.
   */
  function formatKind(code, id) {
    const n = Number(id);
    if ([41, 42, 43, 44].includes(n)) return "accounting";
    if ([5, 6, 7, 8].includes(n)) return "currency";
    const first = String(code).split(";")[0];
    // Remove quoted text, escaped characters and [color]/[$-409] parts before looking at the letters.
    const bare = first.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[(?!\$)[^\]]*\]/g, "");
    if (/^general$/i.test(bare.trim()) || bare.trim() === "") return "general";
    if (bare.trim() === "@") return "text";
    if (/\*/.test(bare) && /_\(/.test(String(code))) return "accounting";
    // [$-409] is a language tag, not a currency; [$€-x-euro2] is a currency.
    const noLocale = first.replace(/\[\$-[0-9A-Fa-f]+\]/g, "");
    if (/[$€£¥]|\[\$[^\]-]/.test(noLocale)) return /\*/.test(bare) ? "accounting" : "currency";
    if (/%/.test(bare)) return "percent";
    if (/E[+-]/i.test(bare)) return "scientific";
    if (/\?\/\?|#\s*\?+\/\?+/.test(bare)) return "fraction";
    const date = /[dy]|m(?![^\[]*\])/i.test(bare.replace(/\[\$[^\]]*\]/g, "")) && /[dy]/i.test(bare);
    const time = /[hs]/i.test(bare) || /AM\/PM/i.test(bare);
    if (date && time) return "datetime";
    if (date || /^m+(?:[-/]y+)?$/i.test(bare.trim())) return "date";
    if (time) return "time";
    return "number";
  }

  /** How many decimal places a number format shows ("#,##0.00" -> 2). */
  function formatDecimals(code) {
    const first = String(code).split(";")[0].replace(/"[^"]*"/g, "");
    const m = /\.([0#?]+)/.exec(first);
    return m ? m[1].length : 0;
  }

  /** Words for messages. */
  const FORMAT_WORDS = {
    general: "General", number: "Number", currency: "Currency", accounting: "Accounting", percent: "Percentage",
    date: "Date", time: "Time", datetime: "Date and time", fraction: "Fraction", scientific: "Scientific", text: "Text",
  };

  /** An Excel date serial number -> "2027-06-01". */
  function serialToISO(serial) {
    const ms = Math.round((Number(serial) - 25569) * 86400000);
    return new Date(ms).toISOString().slice(0, 10);
  }
  /** "2027-06-01" -> Excel date serial number. */
  function isoToSerial(iso) {
    const t = Date.parse(String(iso).trim() + "T00:00:00Z");
    return isNaN(t) ? null : t / 86400000 + 25569;
  }

  // ---------- colors ----------

  /** A color element (<color rgb="FF0070C0"/> or <color theme="4"/>) -> "#0070C0", "theme:4" or null. */
  function readColor(el) {
    if (!el) return null;
    const rgb = at(el, "rgb");
    if (rgb) return "#" + rgb.slice(-6).toUpperCase();
    if (at(el, "theme") !== null) return "theme:" + at(el, "theme") + (at(el, "tint") ? ":" + at(el, "tint") : "");
    if (at(el, "indexed") !== null) return at(el, "indexed") === "64" ? null : "indexed:" + at(el, "indexed");
    if (isOn(at(el, "auto"))) return null;
    return null;
  }

  // =====================================================================
  // PART 2: OPENING AN .XLSX FILE AND BUILDING THE MODEL
  // =====================================================================

  /**
   * Open an .xlsx file (bytes as an ArrayBuffer or Uint8Array) and build the
   * model the checks use. Throws an Error with a learner-friendly message.
   * options.fileName - the file's name, for the file.name check.
   */
  async function readXlsx(bytes, options) {
    const opts = options || {};
    if (typeof root.JSZip === "undefined") throw new Error("The checker couldn't load. Check your internet connection and reopen this panel.");
    const size = bytes.byteLength !== undefined ? bytes.byteLength : bytes.length;
    if (size > MAX_FILE_BYTES) throw new Error("That file is over 25 MB, which is too big to be a practice workbook.");

    let zip;
    try {
      zip = await root.JSZip.loadAsync(bytes);
    } catch (e) {
      throw new Error("That doesn't look like an Excel (.xlsx) file.");
    }
    // Refuse "zip bombs": files that unzip to something enormous.
    let unzipped = 0;
    zip.forEach(function (path, entry) {
      if (entry._data && entry._data.uncompressedSize) unzipped += entry._data.uncompressedSize;
    });
    if (unzipped > MAX_UNZIPPED_BYTES) throw new Error("That file is too big to check.");

    /** Read one XML part, or null if the file doesn't have it. */
    async function part(path) {
      const f = path && zip.file(path);
      return f ? parseXml(await f.async("string")) : null;
    }
    /** Resolve a relationship target ("../charts/chart1.xml") against the part that points to it. */
    function resolve(fromPart, target) {
      if (/^\//.test(target)) return target.slice(1);
      const base = fromPart.split("/").slice(0, -1);
      for (const bit of target.split("/")) {
        if (bit === "..") base.pop();
        else if (bit !== ".") base.push(bit);
      }
      return base.join("/");
    }
    /** A part's relationships: id -> { type, target (resolved path or outside address), external }. */
    async function rels(partPath) {
      const dir = partPath.split("/").slice(0, -1).join("/");
      const name = partPath.split("/").pop();
      const xml = await part((dir ? dir + "/" : "") + "_rels/" + name + ".rels");
      const out = {};
      if (!xml) return out;
      for (const r of all(xml.documentElement, "Relationship")) {
        const external = at(r, "TargetMode") === "External";
        out[at(r, "Id")] = {
          type: (at(r, "Type") || "").split("/").pop(),
          target: external ? at(r, "Target") : resolve(partPath, at(r, "Target") || ""),
          external: external,
        };
      }
      return out;
    }

    const workbookXml = await part("xl/workbook.xml");
    if (!workbookXml) throw new Error("That doesn't look like an Excel (.xlsx) file. If it's a .csv or .xls file, save it as an Excel Workbook (.xlsx) first.");
    const wbRels = await rels("xl/workbook.xml");

    const model = {
      fileName: opts.fileName || "",
      sheets: [],            // every sheet in tab order: worksheets and chart sheets
      names: [],             // defined names: { name, ref, sheetIndex (or null = workbook), hidden }
      styles: null,          // number formats, fonts, fills, borders, cell formats, cell styles
      sharedStrings: [],
      props: { title: "", author: "", keywords: "", lastModifiedBy: "" },
      workbook: { windows: 1, activeTab: 0, filterPrivacy: false, structureLocked: false, structurePassword: false, date1904: false },
      theme: { name: "", colorsName: "" },
      connections: 0,        // data connections (Get Data imports make these)
      externalLinks: [],     // other workbooks this one links to: { index, target }
      persons: {},           // threaded comment authors: id -> name
    };

    // --- shared strings (Excel stores each text once and points to it) ---
    const sstPath = Object.values(wbRels).find((r) => r.type === "sharedStrings");
    const sst = await part(sstPath ? sstPath.target : "xl/sharedStrings.xml");
    if (sst) {
      for (const si of kids(sst.documentElement, "si")) {
        // A plain string is <si><t>..</t></si>; formatted text has runs <si><r><t>..</t></r>...</si>.
        // Phonetic guides (<rPh>) are skipped.
        model.sharedStrings.push(kids(si, "t").concat(kids(si, "r").map((r) => kid(r, "t")).filter(Boolean)).map((t) => t.textContent).join(""));
      }
    }

    // --- styles ---
    const stylesPath = Object.values(wbRels).find((r) => r.type === "styles");
    model.styles = readStyles(await part(stylesPath ? stylesPath.target : "xl/styles.xml"));

    // --- theme ---
    const themePath = Object.values(wbRels).find((r) => r.type === "theme");
    const theme = await part(themePath ? themePath.target : "xl/theme/theme1.xml");
    if (theme) {
      model.theme.name = at(theme.documentElement, "name") || "";
      const scheme = find(theme.documentElement, "clrScheme");
      model.theme.colorsName = scheme ? at(scheme, "name") || "" : "";
    }

    // --- document properties ---
    const core = await part("docProps/core.xml");
    if (core) {
      const prop = (ns, name) => { const e = core.getElementsByTagNameNS(ns, name)[0]; return e ? e.textContent : ""; };
      model.props.title = prop(NS.dc, "title");
      model.props.author = prop(NS.dc, "creator");
      model.props.keywords = prop(NS.cp, "keywords");
      model.props.lastModifiedBy = prop(NS.cp, "lastModifiedBy");
    }

    // --- workbook settings ---
    const wb = workbookXml.documentElement;
    const wbPr = kid(wb, "workbookPr");
    model.workbook.filterPrivacy = isOn(at(wbPr, "filterPrivacy"));
    model.workbook.date1904 = isOn(at(wbPr, "date1904"));
    const views = kids(kid(wb, "bookViews"), "workbookView");
    model.workbook.windows = Math.max(1, views.length);
    model.workbook.activeTab = views.length ? Number(at(views[0], "activeTab") || 0) : 0;
    const prot = kid(wb, "workbookProtection");
    if (prot) {
      model.workbook.structureLocked = isOn(at(prot, "lockStructure"));
      model.workbook.structurePassword = !!(at(prot, "workbookPassword") || at(prot, "workbookHashValue"));
    }

    // --- defined names (named ranges, print areas, print titles) ---
    for (const dn of kids(kid(wb, "definedNames"), "definedName")) {
      model.names.push({
        name: at(dn, "name") || "",
        ref: dn.textContent || "",
        sheetIndex: at(dn, "localSheetId") !== null ? Number(at(dn, "localSheetId")) : null,
        hidden: isOn(at(dn, "hidden")),
      });
    }

    // --- data connections (Get Data / Power Query imports) ---
    const connPath = Object.values(wbRels).find((r) => r.type === "connections");
    const conn = await part(connPath ? connPath.target : null);
    if (conn) model.connections = kids(conn.documentElement, "connection").length;

    // --- links to other workbooks ---
    const extRels = Object.entries(wbRels).filter(([, r]) => r.type === "externalLink");
    const externalOrder = kids(kid(wb, "externalReferences"), "externalReference").map((e) => atNS(e, NS.r, "id"));
    for (let i = 0; i < externalOrder.length; i++) {
      const rel = wbRels[externalOrder[i]];
      if (!rel) continue;
      const linkRels = await rels(rel.target);
      const pathRel = Object.values(linkRels).find((r) => r.type === "externalLinkPath" || r.external);
      model.externalLinks.push({ index: i + 1, target: pathRel ? pathRel.target : "" });
    }
    if (!externalOrder.length && extRels.length) {
      extRels.forEach(([, r], i) => model.externalLinks.push({ index: i + 1, target: r.target }));
    }

    // --- people who wrote threaded comments ---
    const personsRel = Object.values(wbRels).find((r) => r.type === "person");
    const persons = await part(personsRel ? personsRel.target : null);
    if (persons) for (const p of kids(persons.documentElement, "person")) model.persons[at(p, "id")] = at(p, "displayName") || "";

    // --- each sheet, in tab order ---
    let cellTotal = 0;
    const sheetEls = kids(kid(wb, "sheets"), "sheet");
    for (let index = 0; index < sheetEls.length; index++) {
      const s = sheetEls[index];
      const rel = wbRels[atNS(s, NS.r, "id")];
      const sheet = {
        index: index,
        name: at(s, "name") || "",
        state: at(s, "state") || "visible",       // visible, hidden or veryHidden
        kind: rel && rel.type === "chartsheet" ? "chart" : "worksheet",
        path: rel ? rel.target : "",
        cells: new Map(),                        // "B2" -> cell
        rows: {}, cols: [], merges: [], tables: [], drawings: [], charts: [],
        cf: [], validations: [], sparklines: [], hyperlinks: [], notes: [], comments: [], scenarios: [], consolidate: null,
        view: { frozen: null, split: false, view: "normal", showFormulas: false, gridlines: true, zoom: 100, selected: false },
        page: { orientation: null, paper: null, fitToPage: false, fitWide: null, fitTall: null, scale: null, margins: {}, centerH: false, centerV: false, gridlines: false },
        headerFooter: { oddHeader: "", oddFooter: "", firstHeader: "", firstFooter: "", differentFirst: false },
        tabColor: null, protection: null, autoFilter: null, sortState: null,
      };
      model.sheets.push(sheet);
      if (!rel) continue;
      const xml = await part(rel.target);
      if (!xml) continue;
      const sheetRels = await rels(rel.target);
      if (sheet.kind === "chart") {
        await readChartSheet(sheet, xml.documentElement, sheetRels, part, rels, model);
      } else {
        cellTotal += readWorksheet(model, sheet, xml.documentElement, sheetRels);
        if (cellTotal > MAX_CELLS) throw new Error("That workbook has too many cells to check. Practice files are much smaller; check that you opened the right file.");
        await readSheetParts(model, sheet, sheetRels, part, rels);
      }
    }

    // Print areas and print titles are saved as hidden names tied to a sheet.
    for (const n of model.names) {
      if (n.sheetIndex === null || !model.sheets[n.sheetIndex]) continue;
      const sh = model.sheets[n.sheetIndex];
      if (/^_xlnm\.Print_Area$/i.test(n.name)) sh.page.printArea = n.ref;
      if (/^_xlnm\.Print_Titles$/i.test(n.name)) sh.page.printTitles = n.ref;
    }
    return model;
  }

  /** Read the styles part: number formats, fonts, fills, borders, cell formats and named cell styles. */
  function readStyles(xml) {
    const st = { numFmts: {}, fonts: [], fills: [], borders: [], xfs: [], styleXfs: [], cellStyles: [], dxfs: [] };
    if (!xml) { st.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, align: {}, locked: true }); st.fonts.push({}); return st; }
    const r = xml.documentElement;
    for (const nf of kids(kid(r, "numFmts"), "numFmt")) st.numFmts[Number(at(nf, "numFmtId"))] = at(nf, "formatCode") || "";
    for (const f of kids(kid(r, "fonts"), "font")) {
      const flag = (name) => { const e = kid(f, name); return e ? (at(e, "val") === null ? true : isOn(at(e, "val")) || at(e, "val") === "single" || at(e, "val") === "double") : false; };
      st.fonts.push({
        bold: flag("b"), italic: flag("i"), underline: !!kid(f, "u") && at(kid(f, "u"), "val") !== "none",
        size: kid(f, "sz") ? Number(at(kid(f, "sz"), "val")) : null,
        name: kid(f, "name") ? at(kid(f, "name"), "val") : null,
        color: readColor(kid(f, "color")),
      });
    }
    for (const f of kids(kid(r, "fills"), "fill")) {
      const p = kid(f, "patternFill");
      const g = kid(f, "gradientFill");
      st.fills.push({
        pattern: p ? at(p, "patternType") || "none" : (g ? "gradient" : "none"),
        color: p ? readColor(kid(p, "fgColor")) : null,
      });
    }
    for (const b of kids(kid(r, "borders"), "border")) {
      const side = (name) => { const e = kid(b, name); return e && at(e, "style") ? at(e, "style") : null; };
      st.borders.push({ left: side("left"), right: side("right"), top: side("top"), bottom: side("bottom") });
    }
    const readXf = (xf) => {
      const al = kid(xf, "alignment");
      const pr = kid(xf, "protection");
      return {
        numFmtId: Number(at(xf, "numFmtId") || 0), fontId: Number(at(xf, "fontId") || 0),
        fillId: Number(at(xf, "fillId") || 0), borderId: Number(at(xf, "borderId") || 0),
        xfId: Number(at(xf, "xfId") || 0),
        align: al ? {
          horizontal: at(al, "horizontal"), vertical: at(al, "vertical"), wrap: isOn(at(al, "wrapText")),
          indent: Number(at(al, "indent") || 0), rotation: Number(at(al, "textRotation") || 0),
        } : {},
        locked: pr ? at(pr, "locked") !== "0" : true,
      };
    };
    for (const xf of kids(kid(r, "cellXfs"), "xf")) st.xfs.push(readXf(xf));
    for (const xf of kids(kid(r, "cellStyleXfs"), "xf")) st.styleXfs.push(readXf(xf));
    for (const cs of kids(kid(r, "cellStyles"), "cellStyle")) {
      const builtin = at(cs, "builtinId");
      st.cellStyles.push({ name: at(cs, "name") || (builtin !== null ? BUILTIN_STYLES[builtin] : ""), xfId: Number(at(cs, "xfId") || 0), builtinId: builtin !== null ? Number(builtin) : null });
    }
    for (const d of kids(kid(r, "dxfs"), "dxf")) st.dxfs.push(d);
    if (!st.xfs.length) st.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, align: {}, locked: true });
    if (!st.fonts.length) st.fonts.push({});
    return st;
  }

  /** Read one worksheet: cells, rows, columns, views, page setup, merges and more. Returns how many cells it had. */
  function readWorksheet(model, sheet, ws, sheetRels) {
    // --- tab color ---
    const sheetPr = kid(ws, "sheetPr");
    sheet.tabColor = readColor(kid(sheetPr, "tabColor"));
    const fitPr = kid(sheetPr, "pageSetUpPr");
    sheet.page.fitToPage = isOn(at(fitPr, "fitToPage"));

    // --- views: frozen panes, Page Layout view, show formulas, gridlines ---
    const sv = kid(kid(ws, "sheetViews"), "sheetView");
    if (sv) {
      sheet.view.view = at(sv, "view") || "normal";
      sheet.view.showFormulas = isOn(at(sv, "showFormulas"));
      sheet.view.gridlines = at(sv, "showGridLines") !== "0" && at(sv, "showGridLines") !== "false";
      sheet.view.zoom = Number(at(sv, "zoomScale") || 100);
      sheet.view.selected = isOn(at(sv, "tabSelected"));
      const pane = kid(sv, "pane");
      if (pane) {
        const state = at(pane, "state") || "split";
        if (state === "frozen" || state === "frozenSplit") {
          const x = Number(at(pane, "xSplit") || 0), y = Number(at(pane, "ySplit") || 0);
          sheet.view.frozen = cellName(y + 1, x + 1);   // the top-left cell that stays scrollable
        } else {
          sheet.view.split = true;
        }
      }
    }

    // --- columns: width and hidden ---
    for (const col of kids(kid(ws, "cols"), "col")) {
      sheet.cols.push({
        min: Number(at(col, "min")), max: Number(at(col, "max")),
        width: at(col, "width") !== null ? Number(at(col, "width")) : null,
        custom: isOn(at(col, "customWidth")), hidden: isOn(at(col, "hidden")),
        style: at(col, "style") !== null ? Number(at(col, "style")) : null,
      });
    }
    const fmtPr = kid(ws, "sheetFormatPr");
    sheet.defaultRowHeight = fmtPr && at(fmtPr, "defaultRowHeight") ? Number(at(fmtPr, "defaultRowHeight")) : 15;
    sheet.defaultColWidth = fmtPr && at(fmtPr, "defaultColWidth") ? Number(at(fmtPr, "defaultColWidth")) : (fmtPr && at(fmtPr, "baseColWidth") ? Number(at(fmtPr, "baseColWidth")) + 0.71 : 8.43);

    // --- cells ---
    let n = 0;
    const shared = {};           // shared formula index -> { formula, r, c } of the cell that holds the text
    const pendingShared = [];    // cells that point at a shared formula we haven't seen yet
    for (const row of kids(kid(ws, "sheetData"), "row")) {
      const r = Number(at(row, "r"));
      sheet.rows[r] = {
        height: at(row, "ht") !== null ? Number(at(row, "ht")) : null,
        custom: isOn(at(row, "customHeight")), hidden: isOn(at(row, "hidden")),
        style: isOn(at(row, "customFormat")) && at(row, "s") !== null ? Number(at(row, "s")) : null,
      };
      let nextCol = 1;
      for (const c of kids(row, "c")) {
        const ref = at(c, "r");
        const pos = ref ? parseCell(ref) : { r: r, c: nextCol };
        nextCol = pos.c + 1;
        const cell = { r: pos.r, c: pos.c, ref: cellName(pos.r, pos.c), s: Number(at(c, "s") || 0), type: at(c, "t") || "n", value: null, formula: null, formulaKind: null };
        const v = kid(c, "v");
        const raw = v ? v.textContent : null;
        if (cell.type === "s" && raw !== null) cell.value = model.sharedStrings[Number(raw)] !== undefined ? model.sharedStrings[Number(raw)] : "";
        else if (cell.type === "inlineStr") cell.value = textOf(kid(c, "is"));
        else if (cell.type === "b") cell.value = raw === "1";
        else if (cell.type === "e") cell.value = raw;                       // an error such as #DIV/0!
        else if (cell.type === "str") cell.value = raw !== null ? raw : "";
        else if (raw !== null) cell.value = Number(raw);
        const f = kid(c, "f");
        if (f) {
          const kind = at(f, "t") || "normal";   // normal, shared, array or dataTable
          cell.formulaKind = kind;
          if (kind === "dataTable") {
            cell.dataTable = { ref: at(f, "ref"), r1: at(f, "r1"), r2: at(f, "r2"), twoD: isOn(at(f, "dt2D")), row: isOn(at(f, "dtr")) };
            cell.formula = "TABLE(" + (at(f, "r1") || "") + ")";
          } else if (kind === "shared") {
            const si = at(f, "si");
            if (f.textContent) { shared[si] = { formula: f.textContent, r: pos.r, c: pos.c }; cell.formula = f.textContent; }
            else pendingShared.push({ cell, si });
          } else {
            cell.formula = f.textContent;
            if (kind === "array") cell.spill = at(f, "ref");
          }
        }
        sheet.cells.set(cell.ref, cell);
        n++;
      }
    }
    // Rebuild each filled cell's own formula from the shared one.
    for (const p of pendingShared) {
      const src = shared[p.si];
      if (src) p.cell.formula = shiftFormula(src.formula, p.cell.r - src.r, p.cell.c - src.c);
    }

    // --- merged cells ---
    for (const m of kids(kid(ws, "mergeCells"), "mergeCell")) { const g = parseRange(at(m, "ref")); if (g) sheet.merges.push(g); }

    // --- protection ---
    const sp = kid(ws, "sheetProtection");
    if (sp && isOn(at(sp, "sheet"))) sheet.protection = { password: !!(at(sp, "password") || at(sp, "hashValue")) };

    // --- filter and sort settings ---
    const af = kid(ws, "autoFilter");
    if (af) sheet.autoFilter = readAutoFilter(af);
    const ss = kid(ws, "sortState");
    if (ss) sheet.sortState = { ref: at(ss, "ref") };

    // --- conditional formatting ---
    for (const cfEl of kids(ws, "conditionalFormatting")) {
      const ranges = parseSqref(at(cfEl, "sqref"));
      for (const rule of kids(cfEl, "cfRule")) sheet.cf.push(readCfRule(rule, ranges));
    }
    // Excel 2010 and later also save some rules (and every sparkline) in an "extension list".
    for (const ext of kids(kid(ws, "extLst"), "ext")) {
      for (const cfEl of all(ext, "conditionalFormatting")) {
        const sq = find(cfEl, "sqref");
        const ranges = parseSqref(sq ? sq.textContent : "");
        for (const rule of kids(cfEl, "cfRule")) {
          const parsed = readCfRule(rule, ranges);
          // Skip the extension half of a rule we already have (data bars are saved twice).
          if (!sheet.cf.some((x) => x.type === parsed.type && x.ranges.some((g) => ranges.some((h) => rangeName(g) === rangeName(h))))) sheet.cf.push(parsed);
        }
      }
      for (const group of all(ext, "sparklineGroup")) {
        const type = at(group, "type") || "line";
        for (const sl of all(group, "sparkline")) {
          const fEl = kid(sl, "f"), sq = kid(sl, "sqref");
          sheet.sparklines.push({
            type: type, source: fEl ? fEl.textContent : "", location: sq ? sq.textContent : "",
            high: isOn(at(group, "high")), low: isOn(at(group, "low")), markers: isOn(at(group, "markers")),
            first: isOn(at(group, "first")), last: isOn(at(group, "last")), negative: isOn(at(group, "negative")),
          });
        }
      }
    }

    // --- data validation ---
    for (const dv of kids(kid(ws, "dataValidations"), "dataValidation")) {
      const f1 = kid(dv, "formula1");
      sheet.validations.push({ type: at(dv, "type") || "any", ranges: parseSqref(at(dv, "sqref")), formula1: f1 ? f1.textContent : "" });
    }

    // --- hyperlinks ---
    for (const h of kids(kid(ws, "hyperlinks"), "hyperlink")) {
      const rel = sheetRels[atNS(h, NS.r, "id")];
      const g = parseRange(at(h, "ref") || "");
      const display = at(h, "display");
      const cell = g ? sheet.cells.get(cellName(g.r1, g.c1)) : null;
      sheet.hyperlinks.push({
        ref: at(h, "ref") || "", range: g,
        url: rel ? rel.target : "", location: at(h, "location") || "",
        label: cell && cell.value !== null ? String(cell.value) : (display || ""),
      });
    }

    // --- printing: options, margins, page setup, headers and footers ---
    const po = kid(ws, "printOptions");
    if (po) {
      sheet.page.gridlines = isOn(at(po, "gridLines"));
      sheet.page.centerH = isOn(at(po, "horizontalCentered"));
      sheet.page.centerV = isOn(at(po, "verticalCentered"));
    }
    const pm = kid(ws, "pageMargins");
    if (pm) for (const side of ["left", "right", "top", "bottom", "header", "footer"]) if (at(pm, side) !== null) sheet.page.margins[side] = Number(at(pm, side));
    const ps = kid(ws, "pageSetup");
    if (ps) {
      sheet.page.orientation = at(ps, "orientation");
      sheet.page.paper = at(ps, "paperSize") !== null ? (PAPER[at(ps, "paperSize")] || "other") : "letter";
      sheet.page.fitWide = at(ps, "fitToWidth") !== null ? Number(at(ps, "fitToWidth")) : 1;
      sheet.page.fitTall = at(ps, "fitToHeight") !== null ? Number(at(ps, "fitToHeight")) : 1;
      sheet.page.scale = at(ps, "scale") !== null ? Number(at(ps, "scale")) : 100;
    }
    const hf = kid(ws, "headerFooter");
    if (hf) {
      sheet.headerFooter.differentFirst = isOn(at(hf, "differentFirst"));
      for (const k of ["oddHeader", "oddFooter", "firstHeader", "firstFooter"]) sheet.headerFooter[k] = kid(hf, k) ? kid(hf, k).textContent : "";
    }

    // --- scenarios (What-If Analysis > Scenario Manager) ---
    for (const sc of kids(kid(ws, "scenarios"), "scenario")) {
      sheet.scenarios.push({ name: at(sc, "name") || "", inputs: kids(sc, "inputCells").map((i) => ({ ref: at(i, "r"), value: at(i, "val") })) });
    }

    // --- Data > Consolidate remembers its settings here, on the sheet that holds the result ---
    const dc = kid(ws, "dataConsolidate");
    if (dc) {
      sheet.consolidate = {
        fn: at(dc, "function") || "sum",                     // sum is the default, so Excel leaves it out
        links: isOn(at(dc, "link")),                         // "Create links to source data"
        topLabels: isOn(at(dc, "topLabels")), leftLabels: isOn(at(dc, "leftLabels")),
        refs: kids(kid(dc, "dataRefs"), "dataRef").map((d) => ({ ref: at(d, "ref") || at(d, "name") || "", sheet: at(d, "sheet") || "" })),
      };
    }

    // Remember which parts this sheet points at (tables, drawings, comments), for readSheetParts.
    sheet._tableIds = kids(kid(ws, "tableParts"), "tablePart").map((t) => atNS(t, NS.r, "id"));
    sheet._drawingId = atNS(kid(ws, "drawing"), NS.r, "id");
    return n;
  }

  /** An <autoFilter> element: its range and the criteria on each column. */
  function readAutoFilter(af) {
    const columns = kids(af, "filterColumn").map((fc) => {
      const filters = kid(fc, "filters");
      const custom = kid(fc, "customFilters");
      return {
        col: Number(at(fc, "colId") || 0),
        values: filters ? kids(filters, "filter").map((x) => at(x, "val") || "").concat(kids(filters, "dateGroupItem").map((d) => [at(d, "year"), at(d, "month"), at(d, "day")].filter(Boolean).join("-"))) : [],
        custom: custom ? kids(custom, "customFilter").map((x) => ({ operator: at(x, "operator") || "equal", val: at(x, "val") || "" })) : [],
        other: !!(kid(fc, "top10") || kid(fc, "dynamicFilter") || kid(fc, "colorFilter") || kid(fc, "iconFilter")),
      };
    });
    return { ref: at(af, "ref"), columns: columns };
  }

  /** One conditional formatting rule. */
  function readCfRule(rule, ranges) {
    const formulas = kids(rule, "formula").map((f) => f.textContent);
    return {
      type: at(rule, "type") || "", operator: at(rule, "operator") || "", text: at(rule, "text") || "",
      rank: at(rule, "rank"), bottom: isOn(at(rule, "bottom")), percent: isOn(at(rule, "percent")),
      aboveAverage: at(rule, "aboveAverage") !== "0", timePeriod: at(rule, "timePeriod") || "",
      formulas: formulas.length ? formulas : kids(rule, "f").map((f) => f.textContent), ranges: ranges,
    };
  }

  /** Read the parts a worksheet points at: tables, drawings (charts, pictures, shapes), notes and comments. */
  async function readSheetParts(model, sheet, sheetRels, part, rels) {
    // --- Excel tables ---
    for (const id of sheet._tableIds || []) {
      const rel = sheetRels[id];
      const xml = rel ? await part(rel.target) : null;
      if (!xml) continue;
      const t = xml.documentElement;
      const info = kid(t, "tableStyleInfo");
      const columns = kids(kid(t, "tableColumns"), "tableColumn").map((col) => ({
        name: at(col, "name") || "", totalsFunction: at(col, "totalsRowFunction") || "",
        totalsLabel: at(col, "totalsRowLabel") || "",
        calculated: kid(col, "calculatedColumnFormula") ? kid(col, "calculatedColumnFormula").textContent : "",
      }));
      sheet.tables.push({
        name: at(t, "displayName") || at(t, "name") || "", ref: at(t, "ref") || "", range: parseRange(at(t, "ref") || ""),
        headerRow: at(t, "headerRowCount") !== "0", totalRow: Number(at(t, "totalsRowCount") || 0) > 0,
        style: info ? at(info, "name") || "" : "",
        firstColumn: info ? isOn(at(info, "showFirstColumn")) : false, lastColumn: info ? isOn(at(info, "showLastColumn")) : false,
        bandedRows: info ? isOn(at(info, "showRowStripes")) : false, bandedColumns: info ? isOn(at(info, "showColumnStripes")) : false,
        columns: columns, queryTable: at(t, "tableType") === "queryTable",
        autoFilter: kid(t, "autoFilter") ? readAutoFilter(kid(t, "autoFilter")) : null,
        sortState: kid(t, "sortState") ? { ref: at(kid(t, "sortState"), "ref") } : null,
      });
    }

    // --- drawings: charts, pictures, icons, shapes, SmartArt ---
    const drawingRel = sheet._drawingId ? sheetRels[sheet._drawingId] : null;
    if (drawingRel) await readDrawing(model, sheet, drawingRel.target, part, rels);

    // --- notes (the classic yellow notes) and threaded comments ---
    const legacy = Object.values(sheetRels).find((r) => r.type === "comments");
    const threadedRel = Object.values(sheetRels).find((r) => r.type === "threadedComment");
    const threadedRefs = new Set();
    if (threadedRel) {
      const xml = await part(threadedRel.target);
      if (xml) {
        const list = kids(xml.documentElement, "threadedComment");
        const byId = {};
        for (const tcEl of list) {
          const item = {
            id: at(tcEl, "id"), ref: at(tcEl, "ref") || "", parentId: at(tcEl, "parentId"),
            done: isOn(at(tcEl, "done")), author: model.persons[at(tcEl, "personId")] || "",
            text: kid(tcEl, "text") ? kid(tcEl, "text").textContent : "",
          };
          byId[item.id] = item;
        }
        // Group replies under the comment they answer.
        for (const item of Object.values(byId)) {
          if (item.parentId && byId[item.parentId]) continue;
          const replies = Object.values(byId).filter((x) => x.parentId === item.id);
          // A thread is resolved when its first comment is marked done.
          sheet.comments.push({ ref: item.ref, author: item.author, text: item.text, replies: replies, resolved: item.done });
          threadedRefs.add(item.ref.toUpperCase());
        }
      }
    }
    if (legacy) {
      const xml = await part(legacy.target);
      if (xml) {
        const authors = kids(kid(xml.documentElement, "authors"), "author").map((a) => a.textContent);
        for (const cm of kids(kid(xml.documentElement, "commentList"), "comment")) {
          const ref = (at(cm, "ref") || "").toUpperCase();
          const text = textOf(kid(cm, "text"));
          // Excel also saves every threaded comment here as a fallback; those aren't notes.
          if (threadedRefs.has(ref) || /^\[Threaded comment\]/i.test(text)) continue;
          sheet.notes.push({ ref: ref, author: authors[Number(at(cm, "authorId") || 0)] || "", text: text });
        }
      }
    }
  }

  /** A chart sheet: its one chart. */
  async function readChartSheet(sheet, cs, sheetRels, part, rels, model) {
    sheet.tabColor = readColor(kid(kid(cs, "sheetPr"), "tabColor"));
    const drawingId = atNS(kid(cs, "drawing"), NS.r, "id");
    const drawingRel = drawingId ? sheetRels[drawingId] : null;
    if (drawingRel) await readDrawing(model, sheet, drawingRel.target, part, rels);
  }

  /** Read a drawing part: each anchored object becomes an entry in sheet.drawings, charts also in sheet.charts. */
  async function readDrawing(model, sheet, path, part, rels) {
    const xml = await part(path);
    if (!xml) return;
    const drawingRels = await rels(path);
    const anchors = Array.from(xml.documentElement.children).filter((a) => /Anchor$/.test(a.localName));
    for (const anchor of anchors) {
      const from = kid(anchor, "from");
      // The cell the object's top-left corner sits on (rows and columns count from 0 here).
      const at1 = from && kid(from, "row") && kid(from, "col") ? cellName(Number(kid(from, "row").textContent) + 1, Number(kid(from, "col").textContent) + 1) : null;
      for (const el of Array.from(anchor.children)) {
        await readDrawingObject(model, sheet, el, drawingRels, part, rels, at1);
      }
    }
  }

  /** One object in a drawing (recurses into groups). */
  async function readDrawingObject(model, sheet, el, drawingRels, part, rels, anchorCell) {
    const name = el.localName;
    if (name === "grpSp") {
      for (const c of Array.from(el.children)) await readDrawingObject(model, sheet, c, drawingRels, part, rels, anchorCell);
      return;
    }
    if (!["sp", "pic", "graphicFrame", "cxnSp"].includes(name)) return;
    const nv = find(el, "cNvPr");
    const obj = {
      kind: "shape", name: at(nv, "name") || "", alt: (at(nv, "descr") || "").trim(), title: at(nv, "title") || "",
      decorative: !!all(nv, "decorative").find((d) => isOn(at(d, "val"))), text: "", cell: anchorCell,
    };
    if (name === "pic") {
      // An icon is a picture saved as SVG (with a PNG copy for older versions).
      obj.kind = all(el, "svgBlip").length ? "icon" : "picture";
    } else if (name === "sp" || name === "cxnSp") {
      obj.kind = find(el, "cNvSpPr") && isOn(at(find(el, "cNvSpPr"), "txBox")) ? "textbox" : "shape";
      obj.text = textOf(find(el, "txBody"));
    } else if (name === "graphicFrame") {
      const data = find(el, "graphicData");
      const uri = at(data, "uri") || "";
      if (/chart$/.test(uri)) {
        obj.kind = "chart";
        const chartEl = find(data, "chart");
        const rel = chartEl ? drawingRels[atNS(chartEl, NS.r, "id")] : null;
        if (rel) {
          const chart = await readChart(rel.target, part, rels);
          chart.alt = obj.alt; chart.decorative = obj.decorative; chart.name = obj.name; chart.cell = anchorCell;
          chart.chartSheet = sheet.kind === "chart";
          sheet.charts.push(chart);
        }
      } else if (/diagram$/.test(uri)) {
        obj.kind = "smartart";
        const ids = find(data, "relIds");
        const dataRel = ids ? drawingRels[atNS(ids, NS.r, "dm")] : null;
        const dm = dataRel ? await part(dataRel.target) : null;
        obj.items = dm ? all(dm.documentElement, "pt").filter((p) => !at(p, "type") || at(p, "type") === "node").length : 0;
        obj.text = dm ? textOf(dm.documentElement) : "";
      } else {
        obj.kind = "other";
      }
    }
    sheet.drawings.push(obj);
  }

  /** Read a chart part: type, series, titles, legend, labels, style. */
  async function readChart(path, part, rels) {
    const chart = { path: path, types: [], type: "", barDir: "", grouping: "", series: [], title: null, hasTitle: false,
      axisTitles: [], dataLabels: false, legend: "none", dataTable: false, styleId: null, plotBy: "" };
    const xml = await part(path);
    if (!xml) return chart;
    const space = xml.documentElement;
    const ch = kid(space, "chart");
    if (!ch) return chart;
    // Title: a real title element, unless "auto title" is turned off.
    const titleEl = kid(ch, "title");
    const autoDeleted = kid(ch, "autoTitleDeleted") && isOn(at(kid(ch, "autoTitleDeleted"), "val"));
    if (titleEl) { chart.hasTitle = true; chart.title = textOf(titleEl); }
    const plot = kid(ch, "plotArea");
    const groups = plot ? Array.from(plot.children).filter((e) => /Chart$/.test(e.localName)) : [];
    for (const g of groups) {
      let t = g.localName.replace(/3D/, "").replace(/Chart$/, "");
      if (t === "bar") {
        const dir = kid(g, "barDir") ? at(kid(g, "barDir"), "val") : "col";
        t = dir === "bar" ? "bar" : "column";
        chart.barDir = dir;
      }
      chart.types.push(t);
      if (!chart.grouping && kid(g, "grouping")) chart.grouping = at(kid(g, "grouping"), "val") || "";
      if (all(g, "dLbls").some((d) => kid(d, "showVal") && isOn(at(kid(d, "showVal"), "val")) || kid(d, "showPercent") && isOn(at(kid(d, "showPercent"), "val")))) chart.dataLabels = true;
      for (const ser of kids(g, "ser")) {
        const tx = kid(ser, "tx");
        const nameRef = tx ? find(tx, "f") : null;
        const nameVal = tx ? (find(tx, "v") ? find(tx, "v").textContent : textOf(tx)) : "";
        const valRef = find(kid(ser, "val") || kid(ser, "yVal"), "f");
        const catRef = find(kid(ser, "cat") || kid(ser, "xVal"), "f");
        chart.series.push({ nameRef: nameRef ? nameRef.textContent : "", name: nameVal, values: valRef ? valRef.textContent : "", categories: catRef ? catRef.textContent : "" });
      }
    }
    chart.type = chart.types.length > 1 && new Set(chart.types).size > 1 ? "combo" : (chart.types[0] || "");
    if (!titleEl && !autoDeleted && chart.series.length === 1 && ["pie", "doughnut"].includes(chart.type)) {
      chart.hasTitle = true; chart.title = chart.series[0].name; // Excel shows the series name as the title
    }
    // Axis titles
    for (const ax of plot ? Array.from(plot.children).filter((e) => /Ax$/.test(e.localName)) : []) {
      const delEl = kid(ax, "delete");
      if (kid(ax, "title")) chart.axisTitles.push({ axis: ax.localName, text: textOf(kid(ax, "title")), deleted: delEl ? isOn(at(delEl, "val")) : false });
    }
    if (plot && kid(plot, "dTable")) chart.dataTable = true;
    const legend = kid(ch, "legend");
    if (legend) chart.legend = { b: "bottom", t: "top", l: "left", r: "right", tr: "topRight" }[kid(legend, "legendPos") ? at(kid(legend, "legendPos"), "val") : "r"] || "right";
    // Plotted by rows or columns? Look at the first series' values: one row wide = by rows.
    if (chart.series.length) {
      const g = parseRange(splitSheet(chart.series[0].values).ref);
      if (g) chart.plotBy = g.r1 === g.r2 && g.c1 !== g.c2 ? "rows" : (g.c1 === g.c2 && g.r1 !== g.r2 ? "columns" : "");
    }
    // Chart style: Excel 2013+ saves it in a separate "chart style" part with an id.
    const chartRels = await rels(path);
    const styleRel = Object.values(chartRels).find((r) => r.type === "chartStyle");
    if (styleRel) {
      const st = await part(styleRel.target);
      if (st && at(st.documentElement, "id") !== null) chart.styleId = Number(at(st.documentElement, "id"));
    }
    if (chart.styleId === null) {
      const styleEl = all(space, "style").find((e) => e.namespaceURI === NS.c || /chart/.test(e.namespaceURI || ""));
      if (styleEl && at(styleEl, "val") !== null) chart.styleId = Number(at(styleEl, "val"));
    }
    return chart;
  }

  // =====================================================================
  // PART 3: LOOKING THINGS UP IN THE MODEL
  // =====================================================================

  /** The worksheets (not chart sheets). */
  function worksheets(model) { return model.sheets.filter((s) => s.kind === "worksheet"); }

  /** Find a sheet by name (not case sensitive), or the active / first worksheet when no name is given. */
  function getSheet(model, name) {
    if (name === undefined || name === null || name === "") {
      const active = model.sheets[model.workbook.activeTab];
      return active && active.kind === "worksheet" ? active : worksheets(model)[0] || null;
    }
    const want = loose(name);
    return model.sheets.find((s) => loose(s.name) === want) || null;
  }

  /** A sheet name for messages: June or 'June 2027'. */
  function sheetLabel(sheet) { return sheet ? "the " + sheet.name + " sheet" : "the sheet"; }

  /**
   * Work out which sheet and cells a check is about, from its sheet=, cell=
   * and range= values. Returns { sheet, range, label, error }.
   */
  function target(model, params) {
    const ref = params.cell !== undefined ? params.cell : params.range;
    let sheetName = params.sheet;
    let rangeText = ref;
    if (ref !== undefined) {
      const sp = splitSheet(ref);
      if (sp.sheet) { sheetName = sp.sheet; rangeText = sp.ref; }
    }
    const sheet = getSheet(model, sheetName);
    if (!sheet) return { error: sheetName ? 'I can\'t find a sheet named "' + sheetName + '". Check the sheet tabs at the bottom.' : "This workbook has no worksheets." };
    if (sheet.kind !== "worksheet") return { error: sheet.name + " is a chart sheet, not a worksheet." };
    let range = null;
    if (rangeText !== undefined) {
      range = parseRange(rangeText);
      if (!range) {
        // Maybe it's a named range ("TaxRate").
        const named = lookupName(model, rangeText, sheet);
        if (named) return named;
        return { error: '"' + rangeText + '" isn\'t a cell address. Please tell your instructor.' };
      }
    }
    return { sheet: sheet, range: range, label: range ? rangeName(range) : "" };
  }

  /** A named range -> { sheet, range, label } (or null if there's no such name). */
  function lookupName(model, name, sheetHint) {
    const want = loose(name);
    const candidates = model.names.filter((n) => loose(n.name) === want);
    if (!candidates.length) return null;
    const n = candidates.find((x) => sheetHint && x.sheetIndex === sheetHint.index) || candidates.find((x) => x.sheetIndex === null) || candidates[0];
    const sp = splitSheet(n.ref.split(",")[0]);
    const sheet = sp.sheet ? getSheet(model, sp.sheet) : sheetHint;
    const range = parseRange(sp.ref);
    if (!sheet || !range) return null;
    return { sheet: sheet, range: range, label: name };
  }

  /** The cell at row r, column c of a sheet (an empty stand-in if it has nothing saved). */
  function cellAt(sheet, r, c) {
    return sheet.cells.get(cellName(r, c)) || { r: r, c: c, ref: cellName(r, c), s: null, type: "n", value: null, formula: null, empty: true };
  }

  /** Every cell in a range, row by row. Whole columns stop at the sheet's last used row. */
  function cellsIn(sheet, range) {
    let r2 = range.r2, c2 = range.c2;
    if (r2 > 100000 || c2 > 1000) {
      let maxR = 0, maxC = 0;
      for (const cell of sheet.cells.values()) { maxR = Math.max(maxR, cell.r); maxC = Math.max(maxC, cell.c); }
      r2 = Math.min(r2, Math.max(maxR, range.r1)); c2 = Math.min(c2, Math.max(maxC, range.c1));
    }
    const out = [];
    for (let r = range.r1; r <= r2; r++) for (let c = range.c1; c <= c2; c++) out.push(cellAt(sheet, r, c));
    return out;
  }

  /** The used area of a sheet (where cells have values). */
  function usedRange(sheet) {
    let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
    for (const cell of sheet.cells.values()) {
      if (cell.value === null && !cell.formula) continue;
      r1 = Math.min(r1, cell.r); c1 = Math.min(c1, cell.c); r2 = Math.max(r2, cell.r); c2 = Math.max(c2, cell.c);
    }
    return r2 ? { r1, c1, r2, c2 } : null;
  }

  /** The format (xf) a cell uses: its own, or its row's or column's when the cell itself has none saved. */
  function cellXf(model, sheet, cell) {
    let s = cell.s;
    if (s === null || s === undefined) {
      const row = sheet.rows[cell.r];
      if (row && row.style !== null) s = row.style;
      else {
        const col = sheet.cols.find((k) => cell.c >= k.min && cell.c <= k.max && k.style !== null);
        s = col ? col.style : 0;
      }
    }
    return model.styles.xfs[s] || model.styles.xfs[0];
  }

  /** The named cell style a cell uses ("Normal", "Title", "Heading 1", "Input" ...). */
  function cellStyleName(model, xf) {
    const st = model.styles.cellStyles.find((c) => c.xfId === xf.xfId);
    return st ? st.name : (xf.xfId === 0 ? "Normal" : "");
  }

  /** What a cell shows, as text, for searching and messages. */
  function displayText(cell) {
    if (cell.value === null || cell.value === undefined) return "";
    if (typeof cell.value === "boolean") return cell.value ? "TRUE" : "FALSE";
    return String(cell.value);
  }

  /** Is the cell blank (no value and no formula)? */
  function isBlank(cell) { return (cell.value === null || cell.value === "") && !cell.formula; }

  /** Is this merged area exactly the range? Returns the merge or null. */
  function mergeAt(sheet, range) {
    return sheet.merges.find((m) => rangeName(m) === rangeName(range)) || null;
  }

  /** Column width in Excel's units (characters) for column c. */
  function colWidth(sheet, c) {
    const col = sheet.cols.find((k) => c >= k.min && c <= k.max);
    return col && col.width !== null ? col.width : sheet.defaultColWidth;
  }

  /** The table (Excel table) a check means: by name, by a header it contains, or the first one. */
  function findTable(model, params) {
    const sheets = params.sheet ? [getSheet(model, params.sheet)].filter(Boolean) : worksheets(model);
    const tables = sheets.flatMap((s) => s.tables.map((t) => ({ table: t, sheet: s })));
    if (params.table || params.name) {
      const want = loose(params.table || params.name);
      return tables.find((x) => loose(x.table.name) === want) || null;
    }
    if (params.contains) {
      return tables.find((x) => x.table.columns.some((c) => loose(c.name).includes(loose(params.contains))) || tableText(x.sheet, x.table).includes(loose(params.contains))) || null;
    }
    return tables[0] || null;
  }

  /** All the text in a table, for "contains". */
  function tableText(sheet, table) {
    return table.range ? cellsIn(sheet, table.range).map((c) => loose(displayText(c))).join(" | ") : "";
  }

  /** The column number (in the sheet) of a header inside a range's first row, or by letter ("C"). */
  function headerColumn(sheet, range, header) {
    if (/^[A-Za-z]{1,3}$/.test(String(header).trim()) && !cellsIn(sheet, { r1: range.r1, c1: range.c1, r2: range.r1, c2: range.c2 }).some((c) => loose(displayText(c)) === loose(header))) {
      const c = colNumber(String(header).trim());
      if (c >= range.c1 && c <= range.c2) return c;
    }
    for (let c = range.c1; c <= range.c2; c++) {
      if (loose(displayText(cellAt(sheet, range.r1, c))) === loose(header)) return c;
    }
    for (let c = range.c1; c <= range.c2; c++) {
      if (loose(displayText(cellAt(sheet, range.r1, c))).includes(loose(header))) return c;
    }
    return null;
  }

  /** The data area a sort or filter check means: a table, or the sheet's used area (first row = headers). */
  function dataArea(model, params) {
    if (params.table) {
      const t = findTable(model, params);
      if (!t) return { error: 'I can\'t find a table named "' + params.table + '".' };
      const g = Object.assign({}, t.table.range);
      if (t.table.totalRow) g.r2 -= 1;
      return { sheet: t.sheet, range: g, table: t.table, label: t.table.name };
    }
    const tg = target(model, params);
    if (tg.error) return tg;
    const sheet = tg.sheet;
    if (tg.range) return { sheet: sheet, range: tg.range, label: tg.label };
    if (sheet.tables.length === 1) {
      const g = Object.assign({}, sheet.tables[0].range);
      if (sheet.tables[0].totalRow) g.r2 -= 1;
      return { sheet: sheet, range: g, table: sheet.tables[0], label: sheet.tables[0].name };
    }
    const used = usedRange(sheet);
    if (!used) return { error: sheetLabel(sheet) + " is empty." };
    return { sheet: sheet, range: used, label: rangeName(used) };
  }

  /** Compare two cell values the way Excel's sort does: numbers before text, text without case. */
  function compareValues(a, b) {
    const ea = a === null || a === "", eb = b === null || b === "";
    if (ea || eb) return ea === eb ? 0 : (ea ? 1 : -1);    // blanks always go last
    const na = typeof a === "number", nb = typeof b === "number";
    if (na && nb) return a - b;
    if (na !== nb) return na ? -1 : 1;
    return String(a).localeCompare(String(b), "en", { sensitivity: "base", numeric: true });
  }

  /** Is every cell of range `g` inside at least one of `ranges`? */
  function cellsAllIn(g, ranges) {
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (!ranges.some((h) => r >= h.r1 && r <= h.r2 && c >= h.c1 && c <= h.c2)) return false;
    }
    return true;
  }

  /** Every link a "Show me where" can select: { sheet, range }. */
  function whereRange(sheet, range) {
    return sheet && range ? { sheet: sheet.name, range: rangeName(range) } : (sheet ? { sheet: sheet.name } : null);
  }

  // =====================================================================
  // PART 4: THE CHECKS
  // ---------------------------------------------------------------------
  // Each check takes (model, params) and returns pass(message) or
  // fail(message, where). Messages are written for learners: they say what
  // is there now and what it should be. "where" lets "Show me where" select
  // the cells the message is about: { sheet, range }.
  // Lesson files pass values as text, so "true", "12" and "1in" arrive as
  // strings and are converted here.
  // =====================================================================
  const CHECKS = {};
  const truthy = (v) => v === true || v === "true";
  const has = (params, key) => params[key] !== undefined;

  // ---------- file.name ----------
  CHECKS["file.name"] = function (model, params) {
    const raw = String(model.fileName || "").split("?")[0].split("#")[0];
    let name = raw.split(/[\\/]/).pop();
    try { name = decodeURIComponent(name); } catch (e) { /* keep it as it is */ }
    if (!name) return fail("Save the file first. Then select Check my work again.");
    if (/^lastname[_ ]/i.test(name)) return fail('Almost: put your own last name in place of "Lastname".');
    if (wildcard(params.pattern).test(name)) return pass('The file is saved as "' + name + '".');
    const example = String(params.pattern).replace(/^\*/, "Lastname");
    return fail('The file is named "' + name + '". It should look like "' + example + '".');
  };

  // ---------- text.present / text.absent ----------
  /** The cells a text check searches: a range, one sheet, or every worksheet. */
  function searchCells(model, params) {
    if (has(params, "cell") || has(params, "range") || has(params, "sheet")) {
      const t = target(model, params);
      if (t.error) return t;
      return { cells: t.range ? cellsIn(t.sheet, t.range).map((c) => ({ cell: c, sheet: t.sheet })) : Array.from(t.sheet.cells.values()).map((c) => ({ cell: c, sheet: t.sheet })), label: t.range ? t.label + " on " + sheetLabel(t.sheet) : sheetLabel(t.sheet) };
    }
    return { cells: worksheets(model).flatMap((s) => Array.from(s.cells.values()).map((c) => ({ cell: c, sheet: s }))), label: "the workbook" };
  }
  function textMatches(cellText, value, caseMode) {
    if (caseMode === "sensitive") return String(cellText).includes(String(value));
    return loose(cellText).includes(loose(value));
  }
  CHECKS["text.present"] = function (model, params) {
    const sc = searchCells(model, params);
    if (sc.error) return fail(sc.error);
    const hits = sc.cells.filter((x) => textMatches(displayText(x.cell), params.value, params.case));
    const min = has(params, "min") ? Number(params.min) : 1;
    const max = has(params, "max") ? Number(params.max) : Infinity;
    if (hits.length < min) {
      return fail(hits.length === 0 ? 'I don\'t see "' + params.value + '" in ' + sc.label + "."
        : 'I see "' + params.value + '" in ' + count(hits.length, "cell") + " in " + sc.label + "; it should be in at least " + min + ".");
    }
    if (hits.length > max) return fail('"' + params.value + '" is in ' + count(hits.length, "cell") + " in " + sc.label + "; it should be in no more than " + max + ".", whereRange(hits[0].sheet, { r1: hits[0].cell.r, c1: hits[0].cell.c, r2: hits[0].cell.r, c2: hits[0].cell.c }));
    return pass(min > 1 ? '"' + params.value + '" is in ' + count(hits.length, "cell") + "." : 'I found "' + params.value + '".');
  };
  CHECKS["text.absent"] = function (model, params) {
    const sc = searchCells(model, params);
    if (sc.error) return fail(sc.error);
    const hits = sc.cells.filter((x) => textMatches(displayText(x.cell), params.value, params.case));
    if (!hits.length) return pass('"' + params.value + '" is gone from ' + sc.label + ".");
    const first = hits[0];
    return fail('"' + params.value + '" is still in ' + count(hits.length, "cell") + ", the first at " + first.cell.ref + " on " + sheetLabel(first.sheet) + ".",
      whereRange(first.sheet, { r1: first.cell.r, c1: first.cell.c, r2: first.cell.r, c2: first.cell.c }));
  };

  // ---------- props.is ----------
  CHECKS["props.is"] = function (model, params) {
    const names = { title: "Title", author: "Author", keywords: "Tags" };
    for (const key of ["title", "author", "keywords"]) {
      if (!has(params, key)) continue;
      const actual = model.props[key] || "";
      if (params[key] === "*") {
        if (!actual.trim()) return fail("The " + names[key] + " property is empty. Fill it in on File > Info.");
      } else if (!loose(actual).includes(loose(params[key]))) {
        return fail(actual.trim()
          ? "The " + names[key] + ' property is "' + snippet(actual) + '". It should ' + (key === "keywords" ? "include" : "be") + ' "' + params[key] + '".'
          : "The " + names[key] + ' property is empty. It should ' + (key === "keywords" ? "include" : "be") + ' "' + params[key] + '".');
      }
    }
    return pass("The workbook properties are filled in.");
  };

  // ---------- self.attest (the learner confirms something the file can't show) ----------
  CHECKS["self.attest"] = function () { return pass("Thanks for confirming."); };

  // ---------- sheet.is ----------
  CHECKS["sheet.is"] = function (model, params) {
    const sheet = getSheet(model, params.name);
    if (has(params, "exists") && !truthy(params.exists)) {
      return sheet ? fail('There\'s still a sheet named "' + sheet.name + '".') : pass('There\'s no sheet named "' + params.name + '".');
    }
    if (!sheet) {
      const names = model.sheets.map((s) => s.name).join(", ");
      return fail('I can\'t find a sheet named "' + params.name + '". The sheets are: ' + names + ".");
    }
    if (has(params, "index") && sheet.index + 1 !== Number(params.index)) {
      return fail(sheet.name + " is sheet " + (sheet.index + 1) + " from the left. It should be sheet " + params.index + ".");
    }
    if (has(params, "after")) {
      const before = getSheet(model, params.after);
      if (!before) return fail('I can\'t find a sheet named "' + params.after + '".');
      const prev = model.sheets[sheet.index - 1];
      if (!prev || prev !== before) return fail(sheet.name + " should come right after " + before.name + " in the sheet tabs.");
    }
    if (has(params, "tabColor")) {
      if (params.tabColor === "any" && !sheet.tabColor) return fail(sheet.name + " has no tab color yet.");
      if (params.tabColor === "none" && sheet.tabColor) return fail(sheet.name + " still has a tab color.");
      if (/^#/.test(params.tabColor) && String(sheet.tabColor).toUpperCase() !== params.tabColor.toUpperCase()) return fail(sheet.name + "'s tab color isn't " + params.tabColor + ".");
    }
    if (has(params, "hidden")) {
      const hidden = sheet.state !== "visible";
      if (hidden !== truthy(params.hidden)) return fail(truthy(params.hidden) ? sheet.name + " is still showing. Hide it." : sheet.name + " is hidden. Unhide it.");
    }
    if (has(params, "kind") && sheet.kind !== params.kind) return fail(sheet.name + " is a " + (sheet.kind === "chart" ? "chart sheet" : "worksheet") + ", not a " + (params.kind === "chart" ? "chart sheet" : "worksheet") + ".");
    return pass('The sheet "' + sheet.name + '" is set up.');
  };

  // ---------- sheet.count ----------
  CHECKS["sheet.count"] = function (model, params) {
    const list = truthy(params.visible) ? model.sheets.filter((s) => s.state === "visible") : model.sheets;
    const n = list.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    const what = truthy(params.visible) ? "visible sheet" : "sheet";
    if (n < min) return fail("The workbook has " + count(n, what) + "; it should have at least " + min + ".");
    if (n > max) return fail("The workbook has " + count(n, what) + "; it should have no more than " + max + ".");
    return pass("The workbook has " + count(n, what) + ".");
  };

  // ---------- cell.is: values and formatting of a cell or range ----------
  /** One cell against the cell.is parameters. Returns a problem sentence or null. */
  function cellProblem(model, sheet, cell, params) {
    const where = cell.ref;
    const xf = cellXf(model, sheet, cell);
    const font = model.styles.fonts[xf.fontId] || {};
    const baseFont = model.styles.fonts[0] || {};
    const fill = model.styles.fills[xf.fillId] || { pattern: "none" };
    const border = model.styles.borders[xf.borderId] || {};
    const al = xf.align || {};
    const text = displayText(cell);

    if (has(params, "blank")) {
      if (truthy(params.blank) && !isBlank(cell)) return where + ' isn\'t empty: it holds "' + snippet(text || cell.formula) + '".';
      if (!truthy(params.blank) && isBlank(cell)) return where + " is empty.";
    }
    if (has(params, "value") && loose(text) !== loose(params.value)) {
      return text === "" ? where + ' is empty. It should say "' + params.value + '".' : where + ' says "' + snippet(text) + '". It should say "' + params.value + '".';
    }
    if (has(params, "contains") && !loose(text).includes(loose(params.contains))) {
      return text === "" ? where + ' is empty. It should include "' + params.contains + '".' : where + ' says "' + snippet(text) + '". It should include "' + params.contains + '".';
    }
    if (has(params, "number")) {
      const want = Number(params.number);
      const tol = has(params, "tolerance") ? Number(params.tolerance) : Math.max(0.005, Math.abs(want) * 1e-6);
      if (typeof cell.value !== "number") return where + (cell.value === null ? " is empty." : ' holds "' + snippet(text) + '", not a number.') + " It should show " + showNumber(want) + ".";
      if (Math.abs(cell.value - want) > tol) return where + " shows " + showNumber(cell.value) + ". It should show " + showNumber(want) + ".";
    }
    if (has(params, "date")) {
      const want = isoToSerial(params.date);
      if (typeof cell.value !== "number" || Math.floor(cell.value) !== Math.floor(want)) return where + " should hold the date " + params.date + (typeof cell.value === "number" ? " (it holds " + serialToISO(cell.value) + ")." : ".");
    }
    if (has(params, "type")) {
      const kind = cell.value === null && !cell.formula ? "blank" : (cell.type === "e" ? "error" : (typeof cell.value === "number" ? "number" : (typeof cell.value === "boolean" ? "logical" : "text")));
      if (params.type === "formula" ? !cell.formula : (params.type === "date" ? (kind !== "number" || !/date/.test(formatKind(formatCode(model, xf.numFmtId), xf.numFmtId))) : kind !== params.type)) {
        const words = { number: "a number", text: "text", date: "a date", blank: "empty", formula: "a formula", logical: "TRUE or FALSE", error: "an error" };
        return where + " should be " + words[params.type] + "; it's " + (cell.formula && params.type !== "formula" ? "a formula that gives " : "") + (words[kind] || kind) + ".";
      }
    }
    // Number formats only matter for cells that hold something, so empty cells are skipped.
    if (has(params, "numFmt") && !isBlank(cell)) {
      const code = formatCode(model, xf.numFmtId);
      const kind = formatKind(code, xf.numFmtId);
      const want = params.numFmt;
      const ok = kind === want || (want === "date" && kind === "datetime") || (want === "currency" && kind === "accounting" && params.strict !== "true");
      if (!ok) return where + " uses the " + FORMAT_WORDS[kind] + " number format. It should use " + FORMAT_WORDS[want] + ".";
      if (has(params, "decimals") && formatDecimals(code) !== Number(params.decimals)) return where + " shows " + count(formatDecimals(code), "decimal place") + "; it should show " + params.decimals + ".";
    } else if (has(params, "decimals") && !isBlank(cell)) {
      const d = formatDecimals(formatCode(model, xf.numFmtId));
      if (d !== Number(params.decimals)) return where + " shows " + count(d, "decimal place") + "; it should show " + params.decimals + ".";
    }
    if (has(params, "bold") && !!font.bold !== truthy(params.bold)) return where + (truthy(params.bold) ? " isn't bold." : " is still bold.");
    if (has(params, "italic") && !!font.italic !== truthy(params.italic)) return where + (truthy(params.italic) ? " isn't italic." : " is still italic.");
    if (has(params, "underline") && !!font.underline !== truthy(params.underline)) return where + (truthy(params.underline) ? " isn't underlined." : " is still underlined.");
    if (has(params, "fontSize") && Number(font.size) !== Number(params.fontSize)) return where + " is " + (font.size || 11) + " pt. It should be " + params.fontSize + " pt.";
    if (has(params, "font") && loose(font.name) !== loose(params.font)) return where + " uses " + (font.name || "the default font") + ". It should use " + params.font + ".";
    if (has(params, "fontColor")) {
      if (params.fontColor === "changed") { if ((font.color || null) === (baseFont.color || null)) return where + "'s text color hasn't been changed."; }
      else if (String(font.color || "").toUpperCase() !== String(params.fontColor).toUpperCase()) return where + "'s text color isn't " + params.fontColor + ".";
    }
    if (has(params, "fill")) {
      const filled = fill.pattern && fill.pattern !== "none";
      if (params.fill === "any" && !filled) return where + " has no fill color.";
      if (params.fill === "none" && filled) return where + " still has a fill color.";
      if (/^#/.test(params.fill) && (!filled || String(fill.color).toUpperCase() !== params.fill.toUpperCase())) return where + "'s fill isn't " + params.fill + ".";
    }
    if (has(params, "align")) {
      const h = al.horizontal || "general";
      const okCenter = params.align === "center" && (h === "center" || h === "centerContinuous");
      if (h !== params.align && !okCenter) return where + " is aligned " + (h === "general" ? "by default (General)" : h) + ". It should be aligned " + params.align + ".";
    }
    if (has(params, "valign")) {
      const v = al.vertical || "bottom";
      if (v !== params.valign) return where + " is aligned to the " + v + " of the cell. It should be " + params.valign + ".";
    }
    if (has(params, "wrap") && !!al.wrap !== truthy(params.wrap)) return where + (truthy(params.wrap) ? " doesn't wrap its text yet." : " still wraps its text.");
    if (has(params, "indent")) {
      const want = params.indent === "any" ? null : Number(params.indent);
      if (want === null ? !(al.indent > 0) : (al.indent || 0) !== want) return where + " is indented " + (al.indent || 0) + (want === null ? "; it should be indented." : "; it should be indented " + want + ".");
    }
    if (has(params, "rotation")) {
      const rot = al.rotation || 0;
      if (params.rotation === "any" ? rot === 0 : rot !== Number(params.rotation)) return where + (rot === 0 ? " isn't angled or rotated." : " is rotated " + rot + " degrees.") + (params.rotation === "any" ? " Set an orientation." : " It should be " + params.rotation + ".");
    }
    if (has(params, "border")) {
      const sides = ["left", "right", "top", "bottom"].filter((s) => border[s]);
      const want = params.border;
      if (want === "none" && sides.length) return where + " still has a border.";
      if (want === "any" && !sides.length) return where + " has no border.";
      if (["left", "right", "top", "bottom"].includes(want) && !border[want]) return where + " has no " + want + " border.";
    }
    if (has(params, "style")) {
      const name = cellStyleName(model, xf);
      if (loose(name) !== loose(params.style)) return where + " uses the " + (name || "a custom") + " cell style. It should use " + params.style + ".";
    }
    if (has(params, "locked") && xf.locked !== truthy(params.locked)) return where + (truthy(params.locked) ? " is unlocked. Lock it." : " is still locked. Unlock it (Format Cells > Protection).");
    if (has(params, "clean") && truthy(params.clean)) {
      const fmtDirty = xf.numFmtId !== 0 || xf.fontId !== 0 || xf.fillId !== 0 || xf.borderId !== 0 || al.horizontal || al.vertical || al.wrap || al.indent || al.rotation;
      if (fmtDirty) return where + " still has formatting on it.";
    }
    if (has(params, "sameAs")) {
      const sp = splitSheet(params.sameAs);
      const other = getSheet(model, sp.sheet || sheet.name);
      const ref = parseCell(sp.ref);
      if (!other || !ref) return "the sameAs cell " + params.sameAs + " doesn't exist";
      const oxf = cellXf(model, other, cellAt(other, ref.r, ref.c));
      const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
      const fa = model.styles.fonts[xf.fontId], fb = model.styles.fonts[oxf.fontId];
      const la = model.styles.fills[xf.fillId], lb = model.styles.fills[oxf.fillId];
      const ba = model.styles.borders[xf.borderId], bb = model.styles.borders[oxf.borderId];
      if (!same(fa, fb) || !same(la, lb) || !same(ba, bb) || xf.numFmtId !== oxf.numFmtId || !same(xf.align || {}, oxf.align || {})) {
        return where + " isn't formatted the same as " + params.sameAs + " yet.";
      }
    }
    return null;
  }

  CHECKS["cell.is"] = function (model, params) {
    const t = target(model, params);
    if (t.error) return fail(t.error);
    if (!t.range) return fail("This step doesn't say which cells to check. Please tell your instructor.");
    if (has(params, "merged")) {
      const m = mergeAt(t.sheet, t.range);
      const anyMerge = t.sheet.merges.find((g) => overlaps(g, t.range));
      if (truthy(params.merged) && !m) return fail(anyMerge ? t.label + " is merged as " + rangeName(anyMerge) + ". It should be merged across exactly " + t.label + "." : t.label + " isn't merged yet.", whereRange(t.sheet, t.range));
      if (!truthy(params.merged) && anyMerge) return fail(rangeName(anyMerge) + " is still merged.", whereRange(t.sheet, anyMerge));
    }
    const cellParams = Object.assign({}, params);
    delete cellParams.merged;
    const keys = Object.keys(cellParams).filter((k) => !["sheet", "cell", "range", "tolerance", "strict"].includes(k));
    if (keys.length) {
      // Merged cells only hold their value in the top-left cell, so check that one for values.
      const cells = has(params, "merged") ? [cellAt(t.sheet, t.range.r1, t.range.c1)] : cellsIn(t.sheet, t.range);
      for (const cell of cells) {
        const problem = cellProblem(model, t.sheet, cell, cellParams);
        if (problem) return fail(problem.charAt(0).toUpperCase() + problem.slice(1) + (t.sheet.index > 0 || model.sheets.length > 1 ? " (" + t.sheet.name + " sheet)" : ""), whereRange(t.sheet, { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c }));
      }
    }
    return pass(t.label + " " + (t.range.r1 === t.range.r2 && t.range.c1 === t.range.c2 ? "is" : "are") + " right.");
  };

  // ---------- formula.is ----------
  CHECKS["formula.is"] = function (model, params) {
    const t = target(model, params);
    if (t.error) return fail(t.error);
    if (!t.range) return fail("This step doesn't say which cells to check. Please tell your instructor.");
    const all1 = cellsIn(t.sheet, t.range);
    // A dynamic array (SORT, UNIQUE, SEQUENCE) lives in its first cell and "spills" into the rest.
    const cells = truthy(params.spill) ? [all1[0]] : all1;
    const wantFormula = !has(params, "hasFormula") || truthy(params.hasFormula);
    const wh = (cell) => whereRange(t.sheet, { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c });
    for (const cell of cells) {
      if (!wantFormula) {
        if (cell.formula) return fail(cell.ref + " still has a formula (=" + snippet(cleanFormula(cell.formula)) + "). It should hold just the value.", wh(cell));
        continue;
      }
      if (!cell.formula) {
        return fail(cell.ref + (isBlank(cell) ? " is empty." : ' holds a typed value ("' + snippet(displayText(cell)) + '"), not a formula.') + " Start with = so Excel calculates it.", wh(cell));
      }
      const f = cleanFormula(cell.formula).toUpperCase().replace(/\s+/g, "");
      for (const fn of listOf(params.uses)) {
        if (!functionsIn(cell.formula).includes(fn.toUpperCase())) return fail(cell.ref + " has =" + snippet(cleanFormula(cell.formula), 50) + ". It should use the " + fn.toUpperCase() + " function.", wh(cell));
      }
      if (has(params, "contains")) {
        for (const piece of listOf(params.contains)) {
          if (!f.includes(piece.toUpperCase().replace(/\s+/g, ""))) return fail(cell.ref + " has =" + snippet(cleanFormula(cell.formula), 50) + ". It should include " + piece + ".", wh(cell));
        }
      }
      if (has(params, "lacks")) {
        for (const piece of listOf(params.lacks)) {
          if (f.includes(piece.toUpperCase().replace(/\s+/g, ""))) return fail(cell.ref + " has =" + snippet(cleanFormula(cell.formula), 50) + ". It shouldn't include " + piece + ".", wh(cell));
        }
      }
      if (has(params, "refs") && truthy(params.refs) && !/\$?[A-Z]{1,3}\$?\d+|\[[^\]]+\]/.test(f) && !model.names.some((n) => f.includes(n.name.toUpperCase()))) {
        return fail(cell.ref + " has =" + snippet(cleanFormula(cell.formula)) + ", which only uses typed numbers. Point it at the cells instead, so it updates when they change.", wh(cell));
      }
      if (truthy(params.structured) && !/\[/.test(cell.formula)) {
        return fail(cell.ref + " has =" + snippet(cleanFormula(cell.formula), 50) + ". Use the table's column names (a structured reference such as [@Hours]) instead of cell addresses.", wh(cell));
      }
      if (has(params, "external")) {
        const m = /\[(\d+)\]/.exec(cell.formula);
        const link = m ? model.externalLinks.find((l) => l.index === Number(m[1])) : null;
        const ok = (link && loose(decodeURIComponent(link.target)).includes(loose(params.external))) || loose(cell.formula).includes(loose(params.external));
        if (!ok) return fail(cell.ref + " doesn't link to the " + params.external + " workbook yet.", wh(cell));
      }
      if (has(params, "result")) {
        const want = params.result;
        // Text functions give text ("27"), so a number-looking answer may match as text too.
        const ok = (isNumeric(want) && typeof cell.value === "number" && Math.abs(cell.value - Number(want)) <= Math.max(0.005, Math.abs(Number(want)) * 1e-6))
          || (typeof cell.value === "string" && loose(cell.value) === loose(want));
        const shown = cell.value === null ? "nothing" : (typeof cell.value === "number" ? showNumber(cell.value) : '"' + snippet(displayText(cell)) + '"');
        if (!ok) return fail(cell.ref + " shows " + shown + ". It should show " + (isNumeric(want) ? showNumber(want) : '"' + want + '"') + ".", wh(cell));
      }
      // min / max: the answer must be at least (or at most) a number, for answers that change with the inputs.
      if ((has(params, "min") || has(params, "max")) && cell.type !== "e") {
        const v = typeof cell.value === "number" ? cell.value : NaN;
        const shownV = isNaN(v) ? "nothing" : showNumber(v);
        if (has(params, "min") && !(v >= Number(params.min))) return fail(cell.ref + " shows " + shownV + ". It should be at least " + showNumber(params.min) + (Number(params.min) === 0 && v < 0 ? ": a negative answer usually means a minus sign is missing." : "."), wh(cell));
        if (has(params, "max") && !(v <= Number(params.max))) return fail(cell.ref + " shows " + shownV + ". It should be no more than " + showNumber(params.max) + ".", wh(cell));
      }
      if (cell.type === "e" && !has(params, "result")) return fail(cell.ref + " shows the error " + cell.value + ". Fix the formula.", wh(cell));
    }
    if (truthy(params.filled) && cells.length > 1) {
      const first = relativeFormula(cells[0].formula, cells[0].r, cells[0].c);
      const odd = cells.find((c) => relativeFormula(c.formula, c.r, c.c) !== first);
      if (odd) return fail(odd.ref + " has =" + snippet(cleanFormula(odd.formula), 40) + ", which doesn't match the formula filled from " + cells[0].ref + ". Fill the formula again, or check the $ signs.", wh(odd));
    }
    if (truthy(params.spill)) {
      const anchor = cells[0];
      if (!(anchor.formulaKind === "array" || anchor.spill)) {
        // Older files save dynamic arrays like any other formula; accept it if the results are there.
        const filledCells = all1.filter((c) => c.value !== null).length;
        if (filledCells < Math.min(2, all1.length)) return fail(anchor.ref + "'s formula doesn't fill the cells below it yet.", wh(anchor));
      }
    }
    return pass(wantFormula ? (cells.length > 1 ? "The formulas in " + t.label + " are right." : t.label + "'s formula is right.") : t.label + " holds values, not formulas.");
  };

  // ---------- col.is / row.is ----------
  CHECKS["col.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const g = parseRange(String(params.col).includes(":") ? params.col : params.col + ":" + params.col);
    if (!g) return fail("This step names a column the trainer can't read. Please tell your instructor.");
    for (let c = g.c1; c <= g.c2; c++) {
      const letter = colLetters(c);
      const col = sheet.cols.find((k) => c >= k.min && c <= k.max);
      const w = colWidth(sheet, c);
      if (has(params, "hidden") && !!(col && col.hidden) !== truthy(params.hidden)) return fail("Column " + letter + (truthy(params.hidden) ? " is still showing." : " is hidden."), whereRange(sheet, { r1: 1, c1: c, r2: 1, c2: c }));
      if (has(params, "width") && Math.abs(w - Number(params.width)) > 0.75) return fail("Column " + letter + " is " + (Math.round(w * 100) / 100) + " wide. It should be " + params.width + ".", whereRange(sheet, { r1: 1, c1: c, r2: 1, c2: c }));
      if (has(params, "minWidth") && w < Number(params.minWidth) - 0.01) return fail("Column " + letter + " is only " + (Math.round(w * 100) / 100) + " wide. Make it at least " + params.minWidth + " (or use AutoFit).", whereRange(sheet, { r1: 1, c1: c, r2: 1, c2: c }));
      if (has(params, "maxWidth") && w > Number(params.maxWidth) + 0.01) return fail("Column " + letter + " is " + (Math.round(w * 100) / 100) + " wide. Make it no more than " + params.maxWidth + ".", whereRange(sheet, { r1: 1, c1: c, r2: 1, c2: c }));
    }
    return pass("The column widths are set.");
  };
  CHECKS["row.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const g = parseRange(String(params.row).includes(":") ? params.row : params.row + ":" + params.row);
    if (!g) return fail("This step names a row the trainer can't read. Please tell your instructor.");
    for (let r = g.r1; r <= g.r2; r++) {
      const row = sheet.rows[r] || {};
      const h = row.height !== null && row.height !== undefined ? row.height : sheet.defaultRowHeight;
      if (has(params, "hidden") && !!row.hidden !== truthy(params.hidden)) return fail("Row " + r + (truthy(params.hidden) ? " is still showing." : " is hidden."), whereRange(sheet, { r1: r, c1: 1, r2: r, c2: 1 }));
      if (has(params, "height") && Math.abs(h - Number(params.height)) > 1) return fail("Row " + r + " is " + h + " points tall. It should be " + params.height + ".", whereRange(sheet, { r1: r, c1: 1, r2: r, c2: 1 }));
      if (has(params, "minHeight") && h < Number(params.minHeight) - 0.01) return fail("Row " + r + " is " + h + " points tall. Make it at least " + params.minHeight + ".", whereRange(sheet, { r1: r, c1: 1, r2: r, c2: 1 }));
    }
    return pass("The row heights are set.");
  };

  // ---------- view.is: frozen panes, Page Layout view, show formulas ----------
  CHECKS["view.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const v = sheet.view;
    if (has(params, "frozen")) {
      const want = String(params.frozen).toUpperCase();
      if (want === "NONE") { if (v.frozen) return fail("Panes are still frozen on " + sheetLabel(sheet) + "."); }
      else if (v.frozen !== want) {
        const words = { A2: "the top row", B1: "the first column", B2: "the top row and the first column" };
        return fail(v.frozen ? "The frozen panes start at " + v.frozen + ". Freeze " + (words[want] || "at " + want) + " instead." : "Nothing is frozen on " + sheetLabel(sheet) + " yet. Freeze " + (words[want] || "at " + want) + ".");
      }
    }
    if (has(params, "view") && v.view !== params.view) {
      const words = { normal: "Normal", pageLayout: "Page Layout", pageBreakPreview: "Page Break Preview" };
      return fail(sheetLabel(sheet).replace(/^t/, "T") + " is in " + words[v.view] + " view. Switch it to " + words[params.view] + " view, then save.");
    }
    if (has(params, "showFormulas") && v.showFormulas !== truthy(params.showFormulas)) return fail(truthy(params.showFormulas) ? "Formulas aren't showing on " + sheetLabel(sheet) + ". Turn on Show Formulas." : "Show Formulas is still on for " + sheetLabel(sheet) + ". Turn it off.");
    if (has(params, "split") && v.split !== truthy(params.split)) return fail(truthy(params.split) ? sheetLabel(sheet) + " isn't split." : sheetLabel(sheet) + " is still split.");
    if (has(params, "gridlines") && v.gridlines !== truthy(params.gridlines)) return fail(truthy(params.gridlines) ? "Gridlines are hidden on " + sheetLabel(sheet) + "." : "Gridlines still show on " + sheetLabel(sheet) + ".");
    return pass("The view is set.");
  };

  // ---------- window.count (View > New Window) ----------
  CHECKS["window.count"] = function (model, params) {
    const n = model.workbook.windows;
    if (n < Number(params.min || 1)) return fail("The workbook has " + count(n, "window") + " open. Use View > New Window to open another, then save while both are open.");
    return pass("The workbook has " + count(n, "window") + ".");
  };

  // ---------- page.is: printing ----------
  function normalRef(text) {
    return String(text || "").replace(/\$/g, "").replace(/'/g, "").split(",").map((x) => splitSheet(x).ref).join(",").toUpperCase();
  }
  CHECKS["page.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const p = sheet.page;
    const problems = [];
    if (has(params, "orientation") && (p.orientation || "portrait") !== params.orientation) problems.push("the page is " + (p.orientation || "portrait") + ", not " + params.orientation);
    if (has(params, "size") && (p.paper || "letter") !== params.size) problems.push("the paper is " + (PAPER_LABEL[p.paper] || p.paper) + ", not " + PAPER_LABEL[params.size]);
    if (has(params, "fitWide")) {
      if (!p.fitToPage || p.fitWide !== Number(params.fitWide)) problems.push(p.fitToPage ? "it fits " + p.fitWide + " pages wide, not " + params.fitWide : "it isn't set to fit " + count(Number(params.fitWide), "page") + " wide (Scale to Fit)");
    }
    if (has(params, "fitTall")) {
      if (!p.fitToPage || p.fitTall !== Number(params.fitTall)) problems.push(Number(params.fitTall) === 0 ? "the height should be Automatic" : "it should fit " + count(Number(params.fitTall), "page") + " tall");
    }
    if (has(params, "scale") && (p.fitToPage || p.scale !== Number(params.scale))) problems.push("the scale is " + (p.fitToPage ? "set to fit pages" : p.scale + "%") + ", not " + params.scale + "%");
    const inches = (v) => { const m = /^([\d.]+)\s*(in|cm|pt)?$/.exec(String(v)); if (!m) return null; const n = Number(m[1]); return m[2] === "cm" ? n / 2.54 : m[2] === "pt" ? n / 72 : n; };
    for (const side of ["left", "right", "top", "bottom"]) {
      const want = has(params, side) ? inches(params[side]) : (has(params, "margins") ? inches(params.margins) : null);
      if (want === null) continue;
      const actual = p.margins[side] !== undefined ? p.margins[side] : ({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 })[side];
      if (Math.abs(actual - want) > 0.02) problems.push("the " + side + " margin is " + Math.round(actual * 100) / 100 + " in, not " + Math.round(want * 100) / 100 + " in");
    }
    if (has(params, "centerH") && p.centerH !== truthy(params.centerH)) problems.push(truthy(params.centerH) ? "it isn't centered across the page" : "it's still centered across the page");
    if (has(params, "gridlines") && p.gridlines !== truthy(params.gridlines)) problems.push(truthy(params.gridlines) ? "gridlines aren't set to print" : "gridlines still print");
    if (has(params, "printArea")) {
      if (!p.printArea) problems.push("there's no print area yet");
      else if (params.printArea !== "any" && normalRef(p.printArea) !== normalRef(params.printArea)) problems.push("the print area is " + normalRef(p.printArea) + ", not " + normalRef(params.printArea));
    }
    if (has(params, "printTitles")) {
      if (!p.printTitles) problems.push("no rows are set to repeat at the top (Print Titles)");
      else if (!normalRef(p.printTitles).includes(normalRef(params.printTitles))) problems.push("the rows that repeat are " + normalRef(p.printTitles) + ", not " + normalRef(params.printTitles));
    }
    if (problems.length) return fail("On " + sheetLabel(sheet) + ", " + problems.join("; ") + ".");
    return pass("The page setup is right.");
  };

  // ---------- header.is: headers and footers ----------
  /** Split a header/footer code ("&LPaddle&C&A&RPage &P") into its left, center and right parts. */
  function headerParts(code) {
    const parts = { left: "", center: "", right: "" };
    let current = "center";
    const s = String(code || "");
    for (let i = 0; i < s.length; i++) {
      if (s[i] === "&" && i + 1 < s.length) {
        const n = s[i + 1].toUpperCase();
        if (n === "L") { current = "left"; i++; continue; }
        if (n === "C") { current = "center"; i++; continue; }
        if (n === "R") { current = "right"; i++; continue; }
      }
      parts[current] += s[i];
    }
    return parts;
  }
  const HEADER_FIELDS = { page: "&P", pages: "&N", sheet: "&A", file: "&F", date: "&D", time: "&T", path: "&Z" };
  const HEADER_WORDS = { page: "the page number", pages: "the number of pages", sheet: "the sheet name", file: "the file name", date: "the date", time: "the time", path: "the file path" };
  CHECKS["header.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const which = params.part === "footer" ? "footer" : "header";
    const code = which === "footer" ? sheet.headerFooter.oddFooter : sheet.headerFooter.oddHeader;
    if (!code) return fail(sheetLabel(sheet).replace(/^t/, "T") + " has no " + which + " yet.");
    const parts = headerParts(code);
    const section = params.position ? parts[params.position] : parts.left + " " + parts.center + " " + parts.right;
    const sectionWord = params.position ? "the " + params.position + " part of the " + which : "the " + which;
    if (has(params, "field")) {
      const fieldCode = HEADER_FIELDS[params.field];
      if (!section.toUpperCase().includes(fieldCode)) return fail(sectionWord.charAt(0).toUpperCase() + sectionWord.slice(1) + " doesn't show " + HEADER_WORDS[params.field] + " yet. Use the Header & Footer buttons, not typed text.");
    }
    if (has(params, "says")) {
      // Remove font codes like &"Arial,Bold" and &12 before comparing the words.
      const words = section.replace(/&"[^"]*"/g, "").replace(/&\d+/g, "").replace(/&[A-Za-z]/g, " ").replace(/&&/g, "&");
      if (!loose(words).includes(loose(params.says))) return fail(sectionWord.charAt(0).toUpperCase() + sectionWord.slice(1) + ' doesn\'t say "' + params.says + '" yet.');
    }
    return pass("The " + which + " is set up.");
  };

  // ---------- name.is: named ranges ----------
  CHECKS["name.is"] = function (model, params) {
    const want = loose(params.name);
    const found = model.names.filter((n) => loose(n.name) === want && !/^_xlnm\./i.test(n.name));
    if (has(params, "exists") && !truthy(params.exists)) return found.length ? fail('The name "' + params.name + '" still exists. Delete it in the Name Manager.') : pass('There\'s no name "' + params.name + '".');
    if (!found.length) {
      const near = model.names.find((n) => loose(n.name).replace(/[_\s]/g, "") === want.replace(/[_\s]/g, ""));
      return fail(near ? 'There\'s a name "' + near.name + '", but it should be exactly "' + params.name + '".' : 'There\'s no named range called "' + params.name + '" yet. Select the cells, then type the name in the Name Box and press Enter.');
    }
    const n = found[0];
    if (has(params, "refersTo") && !normalRef(n.ref).includes(normalRef(params.refersTo))) return fail('"' + n.name + '" refers to ' + n.ref.replace(/\$/g, "") + ". It should refer to " + params.refersTo + ".");
    if (has(params, "scope")) {
      const scope = n.sheetIndex === null ? "workbook" : (model.sheets[n.sheetIndex] || {}).name;
      if (loose(scope) !== loose(params.scope)) return fail('"' + n.name + '" belongs to ' + (scope === "workbook" ? "the whole workbook" : "the " + scope + " sheet") + ". Its scope should be " + params.scope + ".");
    }
    return pass('The name "' + n.name + '" is set up.');
  };

  // ---------- link.present: hyperlinks ----------
  CHECKS["link.present"] = function (model, params) {
    const sheets = has(params, "sheet") || has(params, "cell") ? [target(model, params).sheet].filter(Boolean) : worksheets(model);
    if (!sheets.length) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    let links = sheets.flatMap((s) => s.hyperlinks.map((l) => Object.assign({ sheet: s }, l)));
    const t = has(params, "cell") ? target(model, params) : null;
    if (t && t.range) links = links.filter((l) => l.range && overlaps(l.range, t.range));
    // Links made with the HYPERLINK() function are formulas, not saved links; count them too.
    for (const s of sheets) {
      for (const cell of s.cells.values()) {
        if (cell.formula && /HYPERLINK\(/i.test(cell.formula) && (!t || !t.range || overlaps(t.range, { r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c }))) {
          const m = /HYPERLINK\(\s*"([^"]*)"/i.exec(cell.formula);
          links.push({ sheet: s, ref: cell.ref, url: m ? m[1] : "", location: m && m[1].startsWith("#") ? m[1].slice(1) : "", label: displayText(cell) });
        }
      }
    }
    if (has(params, "url")) links = links.filter((l) => loose(l.url).includes(loose(params.url)));
    if (has(params, "location")) links = links.filter((l) => loose(l.location).replace(/'/g, "").includes(loose(params.location).replace(/'/g, "")));
    if (has(params, "label")) links = links.filter((l) => loose(l.label).includes(loose(params.label)));
    const min = has(params, "min") ? Number(params.min) : (has(params, "max") ? 0 : 1);
    const max = has(params, "max") ? Number(params.max) : Infinity;
    const what = has(params, "location") ? "link to " + params.location : (has(params, "url") ? "link to " + params.url : "link");
    if (links.length < min) return fail(min === 1 ? "I don't see a " + what + (t && t.range ? " in " + t.label : "") + " yet." : "I see " + count(links.length, "link") + "; there should be at least " + min + ".");
    if (links.length > max) return fail(max === 0 ? "There's still a " + what + (links[0] ? " in " + links[0].ref : "") + ". Remove it (right-click > Remove Hyperlink)." : "There are " + count(links.length, "link") + "; there should be no more than " + max + ".", links[0] && links[0].range ? whereRange(links[0].sheet, links[0].range) : null);
    if (truthy(params.descriptive)) {
      const bad = links.find((l) => /^\s*(click here|here|link|https?:\/\/|www\.)/i.test(l.label) || loose(l.label) === loose(l.url));
      if (bad) return fail('The link in ' + bad.ref + ' says "' + snippet(bad.label) + '". Use words that say where it goes.', bad.range ? whereRange(bad.sheet, bad.range) : null);
    }
    return pass(max === 0 ? "That link is gone." : "The link is there.");
  };

  // ---------- table.is / table.count: Excel tables ----------
  const TOTAL_WORDS = { sum: "Sum", average: "Average", count: "Count", countNums: "Count Numbers", max: "Max", min: "Min", stdDev: "StdDev", var: "Var", custom: "a custom formula" };
  CHECKS["table.is"] = function (model, params) {
    const found = findTable(model, params);
    if (has(params, "exists") && !truthy(params.exists)) {
      return found ? fail('"' + found.table.name + '" is still an Excel table. Convert it to a range (Table Design > Convert to Range).', whereRange(found.sheet, found.table.range)) : pass("That table is now a normal range.");
    }
    if (!found) {
      const any = worksheets(model).some((s) => s.tables.length);
      if (params.name && any) {
        const names = worksheets(model).flatMap((s) => s.tables.map((t) => t.name)).join(", ");
        return fail('There\'s no table named "' + params.name + '". The tables are named: ' + names + ". Rename it in Table Design > Table Name.");
      }
      return fail("I don't see an Excel table" + (params.contains ? ' with "' + params.contains + '" in it' : "") + " yet. Click in the data, then Insert > Table (or Ctrl+T).");
    }
    const t = found.table, wh = whereRange(found.sheet, t.range);
    if (has(params, "style")) {
      const st = t.style || "";
      if (params.style === "any") { if (!st) return fail(t.name + " has no table style.", wh); }
      else if (["light", "medium", "dark"].includes(params.style)) { if (!new RegExp("TableStyle" + params.style, "i").test(st)) return fail(t.name + " uses " + (st.replace(/^TableStyle/, "") || "no style") + ". Pick a " + params.style.charAt(0).toUpperCase() + params.style.slice(1) + " table style.", wh); }
      else if (loose(st) !== loose(params.style)) return fail(t.name + " uses " + (st || "no style") + ". It should use " + params.style + ".", wh);
    }
    const opts = { bandedRows: "Banded Rows", bandedColumns: "Banded Columns", firstColumn: "First Column", lastColumn: "Last Column", headerRow: "Header Row", totalRow: "Total Row" };
    for (const [key, word] of Object.entries(opts)) {
      if (has(params, key) && !!t[key] !== truthy(params[key])) return fail(t.name + ": " + word + (truthy(params[key]) ? " isn't turned on." : " is still on.") + " Use the check boxes in Table Design > Table Style Options.", wh);
    }
    if (has(params, "total")) {
      const col = t.columns.find((c) => loose(c.name) === loose(params.total));
      if (!col) return fail(t.name + ' has no "' + params.total + '" column.', wh);
      if (!t.totalRow) return fail(t.name + " has no total row yet.", wh);
      const fnName = col.totalsFunction;
      if (has(params, "totalFunction") ? fnName !== params.totalFunction : !fnName) {
        return fail("The total row under " + col.name + (fnName ? " uses " + (TOTAL_WORDS[fnName] || fnName) : " is empty") + (has(params, "totalFunction") ? ". Choose " + TOTAL_WORDS[params.totalFunction] + " from its list." : ". Choose a total from its list."), wh);
      }
    }
    if (has(params, "hasColumn") && !t.columns.some((c) => loose(c.name) === loose(params.hasColumn))) return fail(t.name + ' has no "' + params.hasColumn + '" column yet.', wh);
    if (has(params, "lacksColumn") && t.columns.some((c) => loose(c.name) === loose(params.lacksColumn))) return fail(t.name + ' still has a "' + params.lacksColumn + '" column.', wh);
    const dataRows = t.range ? t.range.r2 - t.range.r1 + 1 - (t.headerRow ? 1 : 0) - (t.totalRow ? 1 : 0) : 0;
    if (has(params, "rows") && dataRows !== Number(params.rows)) return fail(t.name + " has " + count(dataRows, "data row") + "; it should have " + params.rows + ".", wh);
    if (has(params, "minRows") && dataRows < Number(params.minRows)) return fail(t.name + " has only " + count(dataRows, "data row") + "; it should have at least " + params.minRows + ". Did the table cover all the data?", wh);
    if (has(params, "columns") && t.columns.length !== Number(params.columns)) return fail(t.name + " has " + count(t.columns.length, "column") + "; it should have " + params.columns + ".", wh);
    if (truthy(params.noBlankRows) && t.range) {
      for (let r = t.range.r1 + (t.headerRow ? 1 : 0); r <= t.range.r2 - (t.totalRow ? 1 : 0); r++) {
        const row = cellsIn(found.sheet, { r1: r, c1: t.range.c1, r2: r, c2: t.range.c2 });
        if (row.every(isBlank)) return fail(t.name + " has an empty row (row " + r + "). Delete it.", whereRange(found.sheet, { r1: r, c1: t.range.c1, r2: r, c2: t.range.c2 }));
      }
    }
    return pass('The table "' + t.name + '" is set up.');
  };
  CHECKS["table.count"] = function (model, params) {
    const sheets = params.sheet ? [getSheet(model, params.sheet)].filter(Boolean) : worksheets(model);
    let tables = sheets.flatMap((s) => s.tables.map((t) => ({ t, s })));
    if (params.contains) tables = tables.filter((x) => x.t.columns.some((c) => loose(c.name).includes(loose(params.contains))) || tableText(x.s, x.t).includes(loose(params.contains)));
    const n = tables.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail(n === 0 ? "I don't see an Excel table yet." : "I see " + count(n, "table") + "; there should be at least " + min + ".");
    if (n > max) return fail(max === 0 ? "That's still an Excel table" + (tables[0] ? " (" + tables[0].t.name + ")" : "") + ". Convert it to a range." : "I see " + count(n, "table") + "; there should be no more than " + max + ".", tables[0] ? whereRange(tables[0].s, tables[0].t.range) : null);
    return pass(max === 0 ? "It's a normal range now." : "The tables are there.");
  };

  // ---------- sort.is: is the data in order? ----------
  CHECKS["sort.is"] = function (model, params) {
    const area = dataArea(model, params);
    if (area.error) return fail(area.error);
    const { sheet, range } = area;
    const by = headerColumn(sheet, range, params.by);
    if (!by) return fail('I can\'t find a "' + params.by + '" column in ' + area.label + ".");
    const then = has(params, "then") ? headerColumn(sheet, range, params.then) : null;
    if (has(params, "then") && !then) return fail('I can\'t find a "' + params.then + '" column in ' + area.label + ".");
    const dir1 = params.order === "desc" ? -1 : 1, dir2 = params.thenOrder === "desc" ? -1 : 1;
    // Only the rows that show count when a filter is on.
    const rows = [];
    for (let r = range.r1 + 1; r <= range.r2; r++) if (!(sheet.rows[r] && sheet.rows[r].hidden)) rows.push(r);
    for (let i = 1; i < rows.length; i++) {
      const a = cellAt(sheet, rows[i - 1], by).value, b = cellAt(sheet, rows[i], by).value;
      const cmp = compareValues(a, b) * dir1;
      if (cmp > 0) return fail("The data isn't sorted by " + params.by + (params.order === "desc" ? " (largest or newest first)" : " (A to Z or smallest first)") + ": row " + rows[i] + " should come before row " + rows[i - 1] + ".", whereRange(sheet, { r1: rows[i - 1], c1: range.c1, r2: rows[i], c2: range.c2 }));
      if (cmp === 0 && then) {
        const c = compareValues(cellAt(sheet, rows[i - 1], then).value, cellAt(sheet, rows[i], then).value) * dir2;
        if (c > 0) return fail("Within each " + params.by + ", the rows aren't sorted by " + params.then + (params.thenOrder === "desc" ? " (largest or newest first)" : "") + " yet. Add a second level in Data > Sort.", whereRange(sheet, { r1: rows[i - 1], c1: range.c1, r2: rows[i], c2: range.c2 }));
      }
    }
    return pass("The data is sorted by " + params.by + (then ? ", then " + params.then : "") + ".");
  };

  // ---------- filter.is ----------
  CHECKS["filter.is"] = function (model, params) {
    const area = dataArea(model, params);
    if (area.error) return fail(area.error);
    const filter = area.table ? area.table.autoFilter : area.sheet.autoFilter;
    const active = filter && filter.columns.some((c) => c.values.length || c.custom.length || c.other);
    if (has(params, "active") && !truthy(params.active)) return active ? fail("A filter is still hiding rows. Clear it (Data > Clear).", whereRange(area.sheet, area.range)) : pass("No filter is hiding rows.");
    if (!filter) return fail("Filtering isn't turned on for " + area.label + ". Use Data > Filter (or Sort & Filter > Filter).");
    if (!has(params, "column")) {
      if (has(params, "active") && truthy(params.active) && !active) return fail("The filter buttons are there, but no filter is applied yet.");
      return pass("Filtering is on.");
    }
    const col = headerColumn(area.sheet, parseRange(filter.ref) || area.range, params.column);
    const fr = parseRange(filter.ref) || area.range;
    if (!col) return fail('I can\'t find a "' + params.column + '" column in the filtered data.');
    const fc = filter.columns.find((c) => fr.c1 + c.col === col);
    if (!fc || !(fc.values.length || fc.custom.length || fc.other)) return fail("The " + params.column + " column isn't filtered yet. Use its arrow button.", whereRange(area.sheet, { r1: fr.r1, c1: col, r2: fr.r1, c2: col }));
    if (has(params, "shows")) {
      const want = listOf(params.shows).map(loose);
      const have = fc.values.map(loose).concat(fc.custom.map((x) => loose(x.val)));
      const missing = want.filter((w) => !have.some((h) => h === w || h.includes(w)));
      if (missing.length) return fail("The " + params.column + " filter should show " + listOf(params.shows).join(" and ") + ".");
      if (fc.values.length && fc.values.length !== want.length && has(params, "only")) return fail("The " + params.column + " filter shows other values too. Show only " + listOf(params.shows).join(" and ") + ".");
    }
    return pass("The " + params.column + " filter is set.");
  };

  // ---------- cf.is / cf.count: conditional formatting ----------
  const CF_WORDS = { dataBar: "Data Bars", colorScale: "Color Scales", iconSet: "Icon Sets", cellIs: "Highlight Cells Rules", top10: "Top/Bottom Rules", aboveAverage: "Above/Below Average", duplicateValues: "Duplicate Values", containsText: "Text That Contains", timePeriod: "A Date Occurring", expression: "a formula rule" };
  function cfRulesFor(model, params) {
    const t = target(model, params);
    if (t.error) return t;
    const rules = t.sheet.cf.filter((r) => !t.range || r.ranges.some((g) => overlaps(g, t.range)));
    return { sheet: t.sheet, range: t.range, label: t.range ? t.label : sheetLabel(t.sheet), rules: rules };
  }
  CHECKS["cf.is"] = function (model, params) {
    const x = cfRulesFor(model, params);
    if (x.error) return fail(x.error);
    let rules = x.rules;
    if (has(params, "type") && params.type !== "any") rules = rules.filter((r) => r.type === params.type);
    if (has(params, "operator")) rules = rules.filter((r) => r.operator === params.operator);
    if (has(params, "value")) rules = rules.filter((r) => r.formulas.some((f) => loose(f).replace(/"/g, "") === loose(params.value).replace(/"/g, "")) || loose(r.text) === loose(params.value));
    if (truthy(params.covers) && x.range) rules = rules.filter((r) => r.ranges.some((g) => contains(g, x.range)));
    const what = has(params, "type") && params.type !== "any" ? CF_WORDS[params.type] || params.type : "conditional formatting";
    if (has(params, "exists") && !truthy(params.exists)) return rules.length ? fail(x.label + " still has " + what + ". Remove it (Conditional Formatting > Clear Rules).", whereRange(x.sheet, x.range)) : pass("That rule is gone.");
    if (!rules.length) {
      if (x.rules.length) return fail(x.label + " has conditional formatting, but not " + what + (has(params, "value") ? " with " + params.value : "") + ".", whereRange(x.sheet, x.range));
      return fail(x.label + " has no " + what + " yet.", whereRange(x.sheet, x.range));
    }
    if (truthy(params.covers) === false) { /* nothing */ }
    return pass(what.charAt(0).toUpperCase() + what.slice(1) + " is applied to " + x.label + ".");
  };
  CHECKS["cf.count"] = function (model, params) {
    const x = cfRulesFor(model, params);
    if (x.error) return fail(x.error);
    const n = x.rules.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail(x.label + " has " + count(n, "conditional formatting rule") + "; it should have at least " + min + ".", whereRange(x.sheet, x.range));
    if (n > max) return fail(max === 0 ? x.label + " still has conditional formatting. Clear the rules." : x.label + " has " + count(n, "rule") + "; it should have no more than " + max + ".", whereRange(x.sheet, x.range || (x.rules[0] && x.rules[0].ranges[0])));
    return pass(max === 0 ? "The old rules are gone from " + x.label + "." : x.label + " has its conditional formatting.");
  };

  // ---------- sparkline.is ----------
  CHECKS["sparkline.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    let list = sheet.sparklines;
    const loc = has(params, "location") ? parseRange(params.location) : null;
    if (loc) list = list.filter((s) => parseSqref(s.location).some((g) => overlaps(g, loc)));
    if (!list.length) return fail("I don't see sparklines" + (loc ? " in " + rangeName(loc) : "") + " on " + sheetLabel(sheet) + " yet. Use Insert > Sparklines.", loc ? whereRange(sheet, loc) : null);
    const min = has(params, "min") ? Number(params.min) : 1;
    if (list.length < min) return fail("I see " + count(list.length, "sparkline") + "; there should be " + min + ". Did you select every location cell?", loc ? whereRange(sheet, loc) : null);
    if (has(params, "type")) {
      const want = params.type === "winloss" ? "stacked" : params.type;
      const wrong = list.find((s) => s.type !== want);
      if (wrong) return fail("The sparklines are " + (wrong.type === "stacked" ? "Win/Loss" : wrong.type) + " sparklines. They should be " + (want === "stacked" ? "Win/Loss" : params.type) + ".", loc ? whereRange(sheet, loc) : null);
    }
    if (has(params, "source") && !list.some((s) => normalRef(s.source).includes(normalRef(params.source)))) return fail("The sparklines don't use the data in " + params.source + ".");
    if (has(params, "high") && !list.every((s) => s.high === truthy(params.high))) return fail(truthy(params.high) ? "Turn on High Point in Sparkline > Show." : "High Point is still on.");
    if (has(params, "markers") && !list.every((s) => s.markers === truthy(params.markers))) return fail(truthy(params.markers) ? "Turn on Markers in Sparkline > Show." : "Markers are still on.");
    return pass("The sparklines are there.");
  };

  // ---------- chart.is / chart.count ----------
  /** The charts a check means: on one sheet (worksheet or chart sheet), or anywhere. */
  function chartsFor(model, params) {
    const sheets = has(params, "sheet") ? [getSheet(model, params.sheet)].filter(Boolean) : model.sheets;
    let list = sheets.flatMap((s) => s.charts.map((c) => Object.assign({ sheetRef: s }, c)));
    if (truthy(params.chartSheet)) list = list.filter((c) => c.chartSheet);
    return { list: list, sheet: sheets[0] || null, missingSheet: has(params, "sheet") && !sheets.length };
  }
  const CHART_WORDS = { column: "column", bar: "bar", line: "line", pie: "pie", doughnut: "doughnut", area: "area", scatter: "scatter (XY)", combo: "combo", radar: "radar", bubble: "bubble" };
  CHECKS["chart.is"] = function (model, params) {
    const x = chartsFor(model, params);
    if (x.missingSheet) return fail('I can\'t find a sheet named "' + params.sheet + '".' + (truthy(params.chartSheet) ? " Move the chart with Chart Design > Move Chart > New sheet, and type that name." : ""));
    let list = x.list;
    if (!list.length) return fail(truthy(params.chartSheet) ? "The chart isn't on its own chart sheet yet. Use Chart Design > Move Chart > New sheet." : "I don't see a chart" + (has(params, "sheet") ? " on the " + params.sheet + " sheet" : "") + " yet.");
    let chart;
    if (has(params, "index")) chart = list[Number(params.index) - 1];
    else if (has(params, "type") && params.type !== "any") chart = list.find((c) => c.type === params.type) || list[0];
    else chart = list[0];
    if (!chart) return fail("There are only " + count(list.length, "chart") + ".");
    const wh = chart.sheetRef && chart.cell ? { sheet: chart.sheetRef.name, range: chart.cell, chart: true } : (chart.sheetRef ? { sheet: chart.sheetRef.name } : null);
    const name = "The chart";
    if (has(params, "type") && params.type !== "any" && chart.type !== params.type) return fail(name + " is a " + (CHART_WORDS[chart.type] || chart.type || "different") + " chart. It should be a " + CHART_WORDS[params.type] + " chart (Chart Design > Change Chart Type).", wh);
    if (has(params, "grouping") && chart.grouping !== params.grouping) return fail(name + " is " + (chart.grouping || "standard") + ". It should be " + params.grouping + " (for example Clustered Column).", wh);
    if (has(params, "series") && chart.series.length !== Number(params.series)) return fail(name + " shows " + count(chart.series.length, "data series", "data series") + "; it should show " + params.series + ".", wh);
    if (has(params, "minSeries") && chart.series.length < Number(params.minSeries)) return fail(name + " shows " + count(chart.series.length, "data series", "data series") + "; it should show at least " + params.minSeries + ". Use Select Data to add one.", wh);
    if (has(params, "seriesName") && !chart.series.some((s) => loose(s.name).includes(loose(params.seriesName)) || loose(s.nameRef).includes(loose(params.seriesName)))) return fail(name + ' has no "' + params.seriesName + '" series yet.', wh);
    if (has(params, "plotBy") && chart.plotBy !== params.plotBy) return fail(name + "'s series come from " + (chart.plotBy || "an unusual range") + ". Use Chart Design > Switch Row/Column so they come from " + params.plotBy + ".", wh);
    if (has(params, "hasTitle") && chart.hasTitle !== truthy(params.hasTitle)) return fail(truthy(params.hasTitle) ? name + " has no title yet." : name + " still has a title.", wh);
    if (has(params, "title") && !(chart.hasTitle && loose(chart.title).includes(loose(params.title)))) return fail(chart.hasTitle ? name + "'s title says \"" + snippet(chart.title || "(linked to a cell)") + '". It should say "' + params.title + '".' : name + ' has no title yet. It should say "' + params.title + '".', wh);
    if (has(params, "axisTitles")) {
      const n = chart.axisTitles.filter((a) => !a.deleted && a.text.trim()).length;
      if (truthy(params.axisTitles) && n < 2) return fail(name + " has " + count(n, "axis title") + "; it should have one on each axis (Add Chart Element > Axis Titles).", wh);
      if (!truthy(params.axisTitles) && n) return fail(name + " still has axis titles.", wh);
    }
    if (has(params, "dataLabels") && chart.dataLabels !== truthy(params.dataLabels)) return fail(truthy(params.dataLabels) ? name + " has no data labels yet." : name + " still has data labels.", wh);
    if (has(params, "legend") && chart.legend !== params.legend) return fail(chart.legend === "none" ? name + " has no legend." + (params.legend !== "none" ? " Add one at the " + params.legend + "." : "") : name + "'s legend is on the " + chart.legend + ". It should be " + (params.legend === "none" ? "removed" : "at the " + params.legend) + ".", wh);
    if (has(params, "dataTable") && chart.dataTable !== truthy(params.dataTable)) return fail(truthy(params.dataTable) ? name + " has no data table under it. Quick Layout 5 adds one." : name + " still has a data table.", wh);
    if (has(params, "styleChanged")) {
      const def = DEFAULT_CHART_STYLE[chart.type];
      const changed = chart.styleId !== null && def !== undefined && chart.styleId !== def;
      if (truthy(params.styleChanged) && !changed) return fail(name + " still has its default chart style. Pick another in Chart Design > Chart Styles.", wh);
    }
    if (has(params, "alt") && !!chart.alt !== truthy(params.alt) && !chart.decorative) return fail(truthy(params.alt) ? name + " has no alternative text yet (right-click > Edit Alt Text)." : name + " still has alt text.", wh);
    if (has(params, "altSays") && !loose(chart.alt).includes(loose(params.altSays))) return fail(name + "'s alt text should mention " + params.altSays + ".", wh);
    if (has(params, "values") && !chart.series.some((s) => normalRef(s.values).includes(normalRef(params.values)))) return fail(name + " doesn't use the data in " + params.values + ".", wh);
    if (has(params, "covers")) {
      // Every cell of the given range must be plotted by some series (either way round).
      const want = parseRange(params.covers);
      const ranges = chart.series.map((x) => parseRange(splitSheet(x.values.split(",")[0]).ref)).filter(Boolean);
      const covered = want && cellsAllIn(want, ranges);
      if (!covered) return fail(name + " doesn't show all of " + params.covers + " yet. Use Chart Design > Select Data to include it.", wh);
    }
    return pass(name + " is set up.");
  };
  CHECKS["chart.count"] = function (model, params) {
    const x = chartsFor(model, params);
    if (x.missingSheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
    const n = x.list.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail(n === 0 ? "I don't see a chart yet." : "I see " + count(n, "chart") + "; there should be at least " + min + ".");
    if (n > max) return fail("I see " + count(n, "chart") + (has(params, "sheet") ? " on the " + params.sheet + " sheet" : "") + "; there should be no more than " + max + ".");
    return pass("The charts are there.");
  };

  // ---------- object.is / object.count: pictures, icons, shapes, SmartArt ----------
  const OBJECT_WORDS = { picture: "picture", icon: "icon", shape: "shape", textbox: "text box", smartart: "SmartArt graphic", any: "graphic" };
  function objectsFor(model, params) {
    const sheets = has(params, "sheet") ? [getSheet(model, params.sheet)].filter(Boolean) : model.sheets;
    let list = sheets.flatMap((s) => s.drawings.filter((d) => d.kind !== "chart" && d.kind !== "other").map((d) => Object.assign({ sheetRef: s }, d)));
    if (params.kind && params.kind !== "any") list = list.filter((d) => d.kind === params.kind || (params.kind === "shape" && d.kind === "textbox"));
    return list;
  }
  CHECKS["object.is"] = function (model, params) {
    let list = objectsFor(model, params);
    const word = OBJECT_WORDS[params.kind] || "graphic";
    if (has(params, "name")) list = list.filter((d) => loose(d.name).includes(loose(params.name)));
    if (has(params, "says")) list = list.filter((d) => loose(d.text).includes(loose(params.says)));
    if (!list.length) return fail("I don't see " + (has(params, "says") ? 'a ' + word + ' that says "' + params.says + '"' : (/^[aeiou]/i.test(word) ? "an " : "a ") + word) + (has(params, "sheet") ? " on the " + params.sheet + " sheet" : "") + " yet.");
    const obj = has(params, "index") ? list[Number(params.index) - 1] : list[0];
    if (!obj) return fail("There are only " + count(list.length, word) + ".");
    const wh = obj.cell ? { sheet: obj.sheetRef.name, range: obj.cell } : null;
    const label = obj.kind === "smartart" ? "The SmartArt graphic" : 'The ' + (OBJECT_WORDS[obj.kind] || "graphic") + (obj.name ? ' "' + obj.name + '"' : "");
    if (has(params, "alt") && truthy(params.alt) && !obj.alt && !obj.decorative) return fail(label + " has no alternative text yet. Right-click it > Edit Alt Text, or mark it decorative.", wh);
    if (has(params, "decorative") && obj.decorative !== truthy(params.decorative)) return fail(truthy(params.decorative) ? label + " isn't marked as decorative yet." : label + " is marked decorative.", wh);
    if (has(params, "items") && obj.kind === "smartart" && obj.items < Number(params.items)) return fail(label + " has " + count(obj.items, "item") + "; it should have " + params.items + ".", wh);
    return pass(label + " is there.");
  };
  CHECKS["object.count"] = function (model, params) {
    const list = objectsFor(model, params);
    const word = OBJECT_WORDS[params.kind] || "graphic";
    const n = list.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail(n === 0 ? "I don't see " + (/^[aeiou]/i.test(word) ? "an " : "a ") + word + " yet." : "I see " + count(n, word) + "; there should be at least " + min + ".");
    if (n > max) return fail("I see " + count(n, word) + "; there should be no more than " + max + ".");
    return pass("The " + word + "s are there.");
  };

  // ---------- scenario.is / datatable.is: what-if analysis ----------
  CHECKS["scenario.is"] = function (model, params) {
    const sheets = has(params, "sheet") ? [getSheet(model, params.sheet)].filter(Boolean) : worksheets(model);
    const scenarios = sheets.flatMap((s) => s.scenarios);
    if (!scenarios.length) return fail("I don't see any scenarios yet. Use Data > What-If Analysis > Scenario Manager > Add.");
    const missing = listOf(params.names).filter((n) => !scenarios.some((s) => loose(s.name) === loose(n)));
    if (missing.length) return fail("I don't see a scenario named " + missing.join(" or ") + " yet. The scenarios are: " + scenarios.map((s) => s.name).join(", ") + ".");
    if (has(params, "min") && scenarios.length < Number(params.min)) return fail("There are " + count(scenarios.length, "scenario") + "; there should be at least " + params.min + ".");
    if (has(params, "changing")) {
      const want = parseCell(params.changing);
      const bad = scenarios.find((s) => !s.inputs.some((i) => { const g = parseRange(i.ref); return g && want && g.r1 === want.r && g.c1 === want.c; }));
      if (bad) return fail('The scenario "' + bad.name + '" doesn\'t change ' + params.changing + ".");
    }
    return pass("The scenarios are set up.");
  };
  CHECKS["datatable.is"] = function (model, params) {
    // With no sheet or range (a project preflight), look on every worksheet.
    if (!has(params, "sheet") && !has(params, "range") && !has(params, "cell")) {
      const s = worksheets(model).find((ws) => Array.from(ws.cells.values()).some((c) => c.dataTable));
      if (!s) return fail("I don't see a data table yet. Select the whole table area, then Data > What-If Analysis > Data Table.");
      params = Object.assign({}, params, { sheet: s.name });
    }
    const t = target(model, params);
    if (t.error) return fail(t.error);
    const cells = t.range ? cellsIn(t.sheet, t.range) : Array.from(t.sheet.cells.values());
    const dt = cells.find((c) => c.dataTable);
    if (!dt) return fail("I don't see a data table" + (t.range ? " in " + t.label : "") + " yet. Select the whole table area, then Data > What-If Analysis > Data Table.", t.range ? whereRange(t.sheet, t.range) : null);
    if (has(params, "input")) {
      const inputs = [dt.dataTable.r1, dt.dataTable.r2].filter(Boolean).map((x) => x.replace(/\$/g, "").toUpperCase());
      if (!inputs.includes(String(params.input).replace(/\$/g, "").toUpperCase())) return fail("The data table's input cell is " + (inputs.join(" and ") || "not set") + ". It should be " + params.input + ".", whereRange(t.sheet, t.range));
    }
    const results = cellsIn(t.sheet, parseRange(dt.dataTable.ref) || t.range || { r1: dt.r, c1: dt.c, r2: dt.r, c2: dt.c });
    if (results.some((c) => c.type === "e")) return fail("Some data table results show errors. Check the formula at the top of the table.", whereRange(t.sheet, parseRange(dt.dataTable.ref)));
    return pass("The data table is there.");
  };

  // ---------- consolidate.is: Data > Consolidate ----------
  CHECKS["consolidate.is"] = function (model, params) {
    const sheet = getSheet(model, params.sheet);
    if (!sheet) return fail(params.sheet ? 'I can\'t find a sheet named "' + params.sheet + '".' : "This workbook has no worksheets.");
    const c = sheet.consolidate;
    if (!c || !c.refs.length) return fail("I don't see a consolidation on " + sheetLabel(sheet) + " yet. Click where the combined data should start, then Data > Consolidate.", { sheet: sheet.name });
    if (has(params, "function") && loose(c.fn) !== loose(params.function)) return fail("The consolidation uses " + c.fn.toUpperCase() + ". It should use " + String(params.function).toUpperCase() + " (the Function box in Data > Consolidate).", { sheet: sheet.name });
    if (has(params, "sources") && c.refs.length < Number(params.sources)) return fail("The consolidation combines " + count(c.refs.length, "range") + "; it should combine " + params.sources + ". Add each one in Data > Consolidate.", { sheet: sheet.name });
    if (has(params, "labels") && truthy(params.labels) && !(c.topLabels && c.leftLabels)) return fail("Tick both Top row and Left column under Use labels in, so Excel matches the rows and columns by name.", { sheet: sheet.name });
    if (has(params, "links") && c.links !== truthy(params.links)) return fail(truthy(params.links) ? "Tick Create links to source data, so the combined numbers update." : "The consolidation links to its sources; this step doesn't want links.", { sheet: sheet.name });
    return pass("The data is consolidated.");
  };

  // ---------- comment.is / comment.count: comments and notes ----------
  function commentsFor(model, params) {
    const sheets = has(params, "sheet") || has(params, "cell") ? [target(model, params).sheet].filter(Boolean) : worksheets(model);
    const t = has(params, "cell") ? target(model, params) : null;
    const inRange = (ref) => !t || !t.range || (parseCell(ref) && overlaps(t.range, { r1: parseCell(ref).r, c1: parseCell(ref).c, r2: parseCell(ref).r, c2: parseCell(ref).c }));
    const comments = sheets.flatMap((s) => s.comments.filter((c) => inRange(c.ref)).map((c) => Object.assign({ sheetRef: s, kind: "comment" }, c)));
    const notes = sheets.flatMap((s) => s.notes.filter((c) => inRange(c.ref)).map((c) => Object.assign({ sheetRef: s, kind: "note" }, c)));
    return { comments, notes, label: t && t.range ? t.label : "" };
  }
  CHECKS["comment.is"] = function (model, params) {
    const x = commentsFor(model, params);
    const isNote = params.kind === "note";
    let list = isNote ? x.notes : x.comments;
    const word = isNote ? "note" : "comment";
    if (has(params, "says")) list = list.filter((c) => loose(c.text).includes(loose(params.says)) || (c.replies || []).some((r) => loose(r.text).includes(loose(params.says))));
    if (has(params, "author")) list = list.filter((c) => loose(c.author).includes(loose(params.author)));
    const desc = has(params, "says") ? 'the ' + word + ' "' + params.says + '"' : "a " + word + (x.label ? " in " + x.label : "");
    const state = params.state || "present";
    if (state === "absent") {
      if (!list.length) return pass("That " + word + " is gone.");
      return fail(desc.charAt(0).toUpperCase() + desc.slice(1) + " is still there" + (x.label ? "" : ", in " + list[0].ref) + ". Delete it.", { sheet: list[0].sheetRef.name, range: list[0].ref });
    }
    if (!list.length) {
      const other = isNote ? x.comments : x.notes;
      if (other.length && x.label) return fail("There's a " + (isNote ? "comment" : "note") + " in " + x.label + ", but this step needs a " + word + ". " + (isNote ? "Use Review > Notes > New Note." : "Use Review > New Comment."));
      return fail("I don't see " + desc + " yet.");
    }
    const c = list[0];
    const wh = { sheet: c.sheetRef.name, range: c.ref };
    if (state === "replied" && !(c.replies || []).length) return fail("Nobody has replied to " + desc + " yet.", wh);
    if (state === "resolved" && !c.resolved) return fail(desc.charAt(0).toUpperCase() + desc.slice(1) + " isn't resolved yet. Use the ... menu on the comment > Resolve thread.", wh);
    if (state === "open" && c.resolved) return fail(desc.charAt(0).toUpperCase() + desc.slice(1) + " is resolved. Reopen it.", wh);
    return pass(state === "replied" ? "There's a reply." : state === "resolved" ? "That thread is resolved." : "The " + word + " is there.");
  };
  CHECKS["comment.count"] = function (model, params) {
    const x = commentsFor(model, params);
    const list = params.kind === "note" ? x.notes : (params.kind === "comment" ? x.comments : x.notes.concat(x.comments));
    const word = params.kind === "note" ? "note" : (params.kind === "comment" ? "comment" : "comment or note");
    const n = list.length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail("I see " + count(n, word, word === "comment or note" ? "comments and notes" : undefined) + "; there should be at least " + min + ".");
    if (n > max) return fail(max === 0 ? "There's still " + (n === 1 ? "a " + word : n + " " + word + "s") + (list[0] ? " (in " + list[0].ref + " on " + list[0].sheetRef.name + ")" : "") + "." : "There are " + n + "; there should be no more than " + max + ".", list[0] ? { sheet: list[0].sheetRef.name, range: list[0].ref } : null);
    return pass(max === 0 ? "No " + word + "s are left." : "The " + word + "s are there.");
  };

  // ---------- protect.is ----------
  CHECKS["protect.is"] = function (model, params) {
    if (has(params, "structure")) {
      if (model.workbook.structureLocked !== truthy(params.structure)) return fail(truthy(params.structure) ? "The workbook structure isn't protected yet. Use Review > Protect Workbook." : "The workbook structure is still protected.");
      if (truthy(params.structurePassword) && !model.workbook.structurePassword) return fail("The workbook structure is protected, but without a password.");
    }
    if (has(params, "sheet") || has(params, "sheetProtected") || has(params, "unlocked")) {
      const sheet = getSheet(model, params.sheet);
      if (!sheet) return fail('I can\'t find a sheet named "' + params.sheet + '".');
      if (has(params, "unlocked")) {
        const g = parseRange(params.unlocked);
        const locked = g ? cellsIn(sheet, g).find((c) => cellXf(model, sheet, c).locked) : null;
        if (locked) return fail(locked.ref + " is still locked, so nobody could type in it once the sheet is protected. Unlock the input cells first (Format Cells > Protection).", whereRange(sheet, { r1: locked.r, c1: locked.c, r2: locked.r, c2: locked.c }));
      }
      if (has(params, "sheetProtected")) {
        const on = !!sheet.protection;
        if (on !== truthy(params.sheetProtected)) return fail(truthy(params.sheetProtected) ? sheetLabel(sheet).replace(/^t/, "T") + " isn't protected yet. Use Review > Protect Sheet." : sheetLabel(sheet).replace(/^t/, "T") + " is still protected.");
        if (on && has(params, "password") && sheet.protection.password !== truthy(params.password)) return fail(truthy(params.password) ? "The sheet is protected, but without a password." : "The sheet has a password; this step doesn't want one.");
      }
    }
    return pass("The protection is set.");
  };

  // ---------- validation.is: data validation ----------
  CHECKS["validation.is"] = function (model, params) {
    const t = target(model, params);
    if (t.error) return fail(t.error);
    const rules = t.sheet.validations.filter((v) => !t.range || v.ranges.some((g) => overlaps(g, t.range)));
    if (has(params, "exists") && !truthy(params.exists)) return rules.length ? fail(t.label + " still has data validation.") : pass("No data validation is left.");
    if (!rules.length) return fail((t.range ? t.label : sheetLabel(t.sheet)) + " has no data validation yet. Use Data > Data Validation.", whereRange(t.sheet, t.range));
    const v = rules[0];
    if (has(params, "type") && params.type !== "any" && v.type !== params.type) return fail("The data validation on " + t.label + " allows " + v.type + ". It should allow a " + params.type + ".", whereRange(t.sheet, t.range));
    if (has(params, "list")) {
      const items = v.formula1.replace(/^"|"$/g, "").split(",").map(loose);
      const missing = listOf(params.list).filter((w) => !items.includes(loose(w)) && !loose(v.formula1).includes(loose(w)));
      if (missing.length) return fail("The list on " + t.label + " doesn't include " + missing.join(", ") + ".", whereRange(t.sheet, t.range));
    }
    if (t.range && !rules.some((r) => r.ranges.some((g) => contains(g, t.range)))) return fail("The data validation doesn't cover all of " + t.label + ".", whereRange(t.sheet, t.range));
    return pass("The data validation is set.");
  };

  // ---------- doc.is: what Inspect Document changes ----------
  CHECKS["doc.is"] = function (model, params) {
    if (has(params, "personalInfo")) {
      const removed = model.workbook.filterPrivacy || (!model.props.author.trim() && !model.props.lastModifiedBy.trim());
      if (params.personalInfo === "removed" && !removed) return fail("The workbook still carries personal information" + (model.props.author ? " (Author: " + snippet(model.props.author) + ")" : "") + ". Run Inspect Document and remove Document Properties and Personal Information.");
      if (params.personalInfo === "kept" && removed) return fail("The personal information was removed.");
    }
    if (has(params, "hiddenSheets")) {
      const hidden = model.sheets.filter((s) => s.state !== "visible");
      if (hidden.length !== Number(params.hiddenSheets)) return fail(hidden.length ? "The workbook still has " + count(hidden.length, "hidden sheet") + " (" + hidden.map((s) => s.name).join(", ") + "). Remove or unhide " + (hidden.length === 1 ? "it" : "them") + "." : "There should be " + count(Number(params.hiddenSheets), "hidden sheet") + ".");
    }
    return pass("The workbook is cleaned up.");
  };

  // ---------- access.is: what the accessibility checker looks for ----------
  const RULE_WORDS = { alt: "alt text", sheetNames: "sheet names", tableNames: "table names", title: "the Title property", tableHeaders: "table header rows", merged: "merged cells in tables", links: "link text" };
  CHECKS["access.is"] = function (model, params) {
    const rules = has(params, "rules") ? listOf(params.rules) : ["alt", "sheetNames", "tableNames", "title", "tableHeaders", "merged", "links"];
    const problems = [];
    let where = null;
    if (rules.includes("alt")) {
      for (const s of model.sheets) {
        const missing = s.drawings.filter((d) => ["chart", "picture", "icon", "shape", "smartart", "textbox"].includes(d.kind) && !d.alt && !d.decorative && !(d.kind === "textbox" || (d.kind === "shape" && d.text)));
        if (missing.length) { problems.push(count(missing.length, missing[0].kind === "chart" ? "chart" : "graphic") + " on the " + s.name + " sheet " + (missing.length === 1 ? "has" : "have") + " no alt text"); where = where || (missing[0].cell ? { sheet: s.name, range: missing[0].cell } : { sheet: s.name }); }
      }
    }
    if (rules.includes("sheetNames")) {
      const defaults = model.sheets.filter((s) => s.state === "visible" && /^(Sheet|Chart)\d+$/i.test(s.name));
      if (defaults.length) problems.push("the " + defaults.map((s) => s.name).join(" and ") + " sheet" + (defaults.length === 1 ? " still has its" : "s still have their") + " default name");
    }
    if (rules.includes("tableNames")) {
      const defaults = worksheets(model).flatMap((s) => s.tables.filter((t) => /^Table\d+$/i.test(t.name)).map((t) => t.name));
      if (defaults.length) problems.push("the table " + defaults.join(" and ") + " still has a default name");
    }
    if (rules.includes("title") && !model.props.title.trim()) problems.push("the workbook has no Title property (File > Info)");
    if (rules.includes("tableHeaders")) {
      const bad = worksheets(model).flatMap((s) => s.tables.filter((t) => !t.headerRow).map((t) => t.name));
      if (bad.length) problems.push("the table " + bad.join(" and ") + " has no header row");
    }
    if (rules.includes("merged")) {
      for (const s of worksheets(model)) {
        const inTable = s.merges.find((m) => s.tables.some((t) => t.range && overlaps(m, t.range)));
        if (inTable) { problems.push("cells " + rangeName(inTable) + " on the " + s.name + " sheet are merged inside a table"); where = where || whereRange(s, inTable); }
      }
    }
    if (rules.includes("links")) {
      for (const s of worksheets(model)) {
        const bad = s.hyperlinks.find((l) => /^\s*(click here|here|link)\s*$/i.test(l.label) || /^\s*(https?:\/\/|www\.)/i.test(l.label));
        if (bad) { problems.push('the link in ' + bad.ref + ' on the ' + s.name + ' sheet says "' + snippet(bad.label, 30) + '"'); where = where || whereRange(s, bad.range); }
      }
    }
    if (problems.length) return fail("Accessibility: " + problems.join("; ") + ".", where);
    return pass(has(params, "rules") ? "No problems with " + rules.map((r) => RULE_WORDS[r] || r).join(", ") + "." : "No accessibility problems found.");
  };

  // ---------- data.is: the shape of a block of data (after an import or cleanup) ----------
  CHECKS["data.is"] = function (model, params) {
    let t;
    if (has(params, "table")) {
      // table=* means "the table whose first header is the first of headers=" (useful after an import, when you can't know its name).
      let found = null;
      if (params.table === "*") {
        const first = listOf(params.headers)[0];
        for (const s of worksheets(model)) for (const tb of s.tables) if (!found && tb.columns[0] && loose(tb.columns[0].name) === loose(first)) found = { sheet: s, table: tb };
      } else {
        found = findTable(model, { table: params.table });
      }
      if (!found) return fail(params.table === "*" ? 'I don\'t see a table that starts with a "' + listOf(params.headers)[0] + '" column yet.' : 'I can\'t find a table named "' + params.table + '".');
      const g = Object.assign({}, found.table.range);
      if (found.table.totalRow) g.r2 -= 1;
      t = { sheet: found.sheet, range: g, label: found.table.name };
    } else {
      t = target(model, params);
    }
    if (t.error) return fail(t.error);
    const g = t.range || usedRange(t.sheet);
    if (!g) return fail(sheetLabel(t.sheet).replace(/^t/, "T") + " is empty.");
    if (has(params, "headers")) {
      const want = listOf(params.headers);
      for (let i = 0; i < want.length; i++) {
        const cell = cellAt(t.sheet, g.r1, g.c1 + i);
        if (loose(displayText(cell)) !== loose(want[i])) return fail("The headers on " + sheetLabel(t.sheet) + " should start " + want.join(", ") + ". " + cell.ref + ' says "' + snippet(displayText(cell)) + '".', whereRange(t.sheet, { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c1 + want.length - 1 }));
      }
    }
    // Data rows: every row under the header row that has anything in it.
    const rows = [];
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      const vals = cellsIn(t.sheet, { r1: r, c1: g.c1, r2: r, c2: g.c2 }).map((c) => loose(displayText(c)));
      if (vals.some((v) => v !== "")) rows.push({ r: r, key: vals.join("|") });
    }
    if (has(params, "rows") && rows.length !== Number(params.rows)) return fail(sheetLabel(t.sheet).replace(/^t/, "T") + " has " + count(rows.length, "data row") + "; it should have " + params.rows + ".");
    if (has(params, "minRows") && rows.length < Number(params.minRows)) return fail(sheetLabel(t.sheet).replace(/^t/, "T") + " has " + count(rows.length, "data row") + "; it should have at least " + params.minRows + ".");
    if (has(params, "maxRows") && rows.length > Number(params.maxRows)) return fail(sheetLabel(t.sheet).replace(/^t/, "T") + " has " + count(rows.length, "data row") + "; it should have no more than " + params.maxRows + ".");
    if (truthy(params.noDuplicates)) {
      const seen = new Map();
      for (const row of rows) {
        if (seen.has(row.key)) return fail("Row " + row.r + " is a duplicate of row " + seen.get(row.key) + ". Use Data > Remove Duplicates.", whereRange(t.sheet, { r1: row.r, c1: g.c1, r2: row.r, c2: g.c2 }));
        seen.set(row.key, row.r);
      }
    }
    return pass("The data on " + sheetLabel(t.sheet) + " is in shape.");
  };

  // ---------- connection.count: Get Data imports ----------
  CHECKS["connection.count"] = function (model, params) {
    const n = model.connections + worksheets(model).flatMap((s) => s.tables.filter((t) => t.queryTable)).length;
    const min = has(params, "min") ? Number(params.min) : 0, max = has(params, "max") ? Number(params.max) : Infinity;
    if (n < min) return fail("This workbook has no data connection yet. Import the data with Data > Get Data, so it can be refreshed later.");
    if (n > max) return fail("This workbook still has " + count(n, "data connection") + ".");
    return pass("The data is imported with a connection.");
  };

  // ---------- theme.is ----------
  CHECKS["theme.is"] = function (model, params) {
    if (has(params, "name") && !loose(model.theme.name).includes(loose(params.name))) return fail("The workbook uses the " + (model.theme.name || "default") + " theme. It should use " + params.name + " (Page Layout > Themes).");
    if (has(params, "changed") && truthy(params.changed) && /^office( theme)?$/i.test(model.theme.name || "Office Theme")) return fail("The workbook still uses the default Office theme. Pick another in Page Layout > Themes.");
    if (has(params, "colors") && !loose(model.theme.colorsName).includes(loose(params.colors))) return fail("The theme colors are " + (model.theme.colorsName || "the defaults") + ". They should be " + params.colors + ".");
    return pass("The theme is set.");
  };

  // ---------- cellstyle.is: a named cell style exists ----------
  CHECKS["cellstyle.is"] = function (model, params) {
    const st = model.styles.cellStyles.find((c) => loose(c.name) === loose(params.name));
    if (has(params, "exists") && !truthy(params.exists)) return st ? fail('The cell style "' + st.name + '" still exists.') : pass("That style is gone.");
    if (!st) return fail('There\'s no cell style named "' + params.name + '" yet. Use Home > Cell Styles > New Cell Style.');
    if (has(params, "custom") && truthy(params.custom) && st.builtinId !== null && BUILTIN_STYLES[st.builtinId] === st.name) return fail('"' + st.name + '" is a built-in style. Make your own with New Cell Style.');
    return pass('The cell style "' + st.name + '" is there.');
  };

  // =====================================================================
  // PART 5: RUNNING THE CHECKS
  // =====================================================================

  /** Run one check. Never throws: a broken check fails with a message. */
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
      const messages = Array.from(new Set(results.map((r) => r.message)));
      return { pass: true, messages: messages, results: results };
    }
    // Two checks can fail for the same reason (an empty cell); say it once.
    return { pass: false, messages: Array.from(new Set(failed.map((r) => r.message))), where: failed[0].where, results: results };
  }

  const api = {
    VERSION, readXlsx, read: readXlsx, runCheck, runStep, CHECKS: Object.keys(CHECKS),
    // exposed for tests and for future tools:
    _internal: { shiftFormula, relativeFormula, functionsIn, cleanFormula, formatKind, formatDecimals, parseRange, rangeName, headerParts, compareValues, splitSheet },
  };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.ExcelChecker = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
