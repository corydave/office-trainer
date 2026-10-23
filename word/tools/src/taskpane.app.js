/* =============================================================================
   OFFICE TRAINER - the add-in's own code
   -----------------------------------------------------------------------------
   build-taskpane.py copies this file into taskpane.html. It runs after the
   shared libraries, so window.TrainerMD (reads and draws lessons),
   window.WordChecker (checks a Word document), window.ExcelChecker
   (checks an Excel workbook) and window.PptChecker (checks a PowerPoint
   presentation) are ready.

   ONE ADD-IN, THREE COURSES: the same panel opens in Word, Excel and
   PowerPoint. At startup (PART 14) Office tells us which app opened us, and
   chooseApp() picks that course: its folder, its checker and its saved
   progress. In a plain browser tab, ?app=excel (or ?app=powerpoint, or the
   learner's last choice) decides.

   The parts, in order:
     PART 1  Settings for this copy of the trainer
     PART 2  Small helpers (making elements safely, fetching files)
     PART 3  Saved progress (kept only in this browser)
     PART 4  Word, Excel, PowerPoint, or a plain browser? (getting the file, loading starters)
     PART 5  Loading the course (outline, lessons, settings)
     PART 6  Moving between screens
     PART 7  The course home screen
     PART 8  A lesson: drawing it, and making drills, quizzes and videos work
     PART 9  Badges and certificates
     PART 10 Skill checks (timed)
     PART 11 Projects (preflight and the rubric self-check)
     PART 12 The badges screen (my record)
     PART 13 Settings screen
     PART 14 Starting up

   SECURITY RULES this file follows everywhere:
     - Lesson content is drawn ONLY by TrainerMD.render, which escapes all text
       first and allows a short list of safe tags. Nothing else uses innerHTML.
     - Anything from the learner's document (check messages quote it) is shown
       with textContent, so it can never run as code.
     - We only fetch files from this site, and only paths that look like
       course files (content/....md, starters/....docx).
     - Videos load only after the learner presses Play, from approved hosts,
       in a sandboxed frame.
   ============================================================================= */
(function () {
  "use strict";

  const TM = window.TrainerMD;

  // =====================================================================
  // PART 1: SETTINGS FOR THIS COPY OF THE TRAINER
  // =====================================================================

  // Shown at the bottom of the panel. Bump it whenever you upload a new
  // taskpane.html, so you can SEE whether Word picked up the new version.
  const APP_VERSION = "0.8";

  // The three courses this add-in can show. Each has its own folder (next to
  // this page), its own checker, and its own saved progress, so a Word
  // badge and an Excel badge with the same activity id ("D1.03") never mix.
  const APPS = {
    word: {
      id: "word", name: "Word", fullName: "Microsoft Word", course: "word/", certCode: "WORD",
      storeKey: "officeTrainer.word.v1",
      file: "document",                       // what we call the learner's file in messages
      checker: function () { return window.WordChecker; },
      read: function (bytes, opts) { return window.WordChecker.readDocx(bytes, opts); },
      starter: /^starters\/[A-Za-z0-9_\-/]+\.(docx|dotx)$/,
      accept: ".docx,.dotx,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.openxmlformats-officedocument.wordprocessingml.template",
      fileTest: /\.(docx|dotx)$/i,
      wrongFile: "Please choose a Word file that ends in .docx (or .dotx for a template).",
    },
    excel: {
      id: "excel", name: "Excel", fullName: "Microsoft Excel", course: "excel/", certCode: "EXCEL",
      storeKey: "officeTrainer.excel.v1",
      file: "workbook",
      checker: function () { return window.ExcelChecker; },
      read: function (bytes, opts) { return window.ExcelChecker.readXlsx(bytes, opts); },
      starter: /^starters\/[A-Za-z0-9_\-/]+\.(xlsx|xltx)$/,
      accept: ".xlsx,.xltx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.openxmlformats-officedocument.spreadsheetml.template",
      fileTest: /\.(xlsx|xltx|xlsm)$/i,
      wrongFile: "Please choose an Excel file that ends in .xlsx (or .xltx for a template).",
    },
    powerpoint: {
      id: "powerpoint", name: "PowerPoint", fullName: "Microsoft PowerPoint", course: "powerpoint/", certCode: "PPT",
      storeKey: "officeTrainer.powerpoint.v1",
      file: "presentation",
      checker: function () { return window.PptChecker; },
      read: function (bytes, opts) { return window.PptChecker.readPptx(bytes, opts); },
      starter: /^starters\/[A-Za-z0-9_\-/]+\.(pptx|potx)$/,
      accept: ".pptx,.potx,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.presentationml.template",
      fileTest: /\.(pptx|potx)$/i,
      wrongFile: "Please choose a PowerPoint file that ends in .pptx (or .potx for a template).",
    },
  };

  // The course this panel is showing. chooseApp() sets these at startup;
  // the rest of the code just reads them.
  let APP = APPS.word;
  let COURSE = APP.course;     // where the course files are, next to this page
  let STORE_KEY = APP.storeKey; // the name progress is saved under in this browser

  /** Switch to one course ("word", "excel" or "powerpoint"). Called once, before anything is drawn. */
  function chooseApp(id) {
    APP = isApp(id) ? APPS[id] : APPS.word;
    COURSE = APP.course;
    STORE_KEY = APP.storeKey;
  }

  /** Is this one of our three course names? (Not just any property: "constructor" isn't a course.) */
  function isApp(id) { return typeof id === "string" && Object.prototype.hasOwnProperty.call(APPS, id); }

  /** The checker for this course (WordChecker, ExcelChecker or PptChecker). */
  function checker() { return APP.checker(); }

  // In a browser, the last course the learner picked is remembered here.
  const APP_CHOICE_KEY = "officeTrainer.app";

  // Only these kinds of paths are ever fetched (no "..", no other sites).
  // (Each course's practice-file pattern is APP.starter, above.)
  const SAFE_LESSON_PATH = /^content\/[A-Za-z0-9_\-/]+\.md$/;
  const SAFE_EXTRA_PATH = /^starters\/[A-Za-z0-9_\-/]+\.(docx|dotx|xlsx|xltx|pptx|potx|csv|txt|png|jpg|jpeg|svg|glb|mp4|m4a|mp3|wav)$/i;

  // Skill check time choices (practice mode). "0" means no time limit.
  const TIME_CHOICES = [
    { value: "1", label: "Standard time" },
    { value: "1.5", label: "Time and a half" },
    { value: "2", label: "Double time" },
    { value: "0", label: "No time limit" },
  ];

  // =====================================================================
  // PART 2: SMALL HELPERS
  // =====================================================================

  /** Shortcut for document.getElementById. */
  function $(id) { return document.getElementById(id); }

  /**
   * Make an element. Children can be text (added as plain text, always safe),
   * other elements, or lists of either. Attributes are set one by one.
   *   el("button", { class: "btn", type: "button" }, "Check my work")
   */
  function el(tag, attrs, ...kids) {
    const node = document.createElement(tag);
    for (const [key, value] of Object.entries(attrs || {})) {
      if (value === null || value === undefined || value === false) continue;
      if (key === "onclick") node.addEventListener("click", value);
      else node.setAttribute(key, value === true ? "" : String(value));
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid === null || kid === undefined || kid === false) continue;
      node.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
    }
    return node;
  }


  /**
   * Put children into a node, skipping null, undefined and false.
   * (The browser's own append() and replaceChildren() would write the
   * word "null" on the page for a skipped optional piece.)
   */
  function add(node, ...kids) {
    node.append(...kids.flat(Infinity).filter((k) => k !== null && k !== undefined && k !== false));
    return node;
  }
  /** Like add(), but first removes whatever the node held. */
  function fill(node, ...kids) {
    node.replaceChildren();
    return add(node, ...kids);
  }

  /**
   * Badge artwork. TrainerMD.badgeSVG escapes every value it draws, so its
   * SVG text is safe to turn into an element.
   */
  function badgeArt(options) {
    const holder = document.createElement("span");
    holder.innerHTML = TM.badgeSVG(options);
    return holder.firstChild;
  }

  /** A small inline icon (check mark, arrow, play). Fixed SVG text written here, no outside input. */
  const ICONS = {
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    chevron: '<svg class="chev" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 5l7 7-7 7"/></svg>',
    play: '<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="32" cy="32" r="30" fill="rgba(255,255,255,.15)" stroke="#fff" stroke-width="3"/><path d="M26 20l20 12-20 12z" fill="#fff"/></svg>',
  };
  function icon(name) {
    const holder = document.createElement("span");
    holder.innerHTML = ICONS[name];
    return holder.firstChild;
  }

  /** Tell screen reader users something (check results, badges). */
  function announce(text) {
    const region = $("announce");
    region.textContent = "";
    // A tiny delay makes screen readers notice the same message twice in a row.
    setTimeout(function () { region.textContent = text; }, 50);
  }

  /** Today's date as "2026-10-18". */
  function today() { return new Date().toISOString().slice(0, 10); }

  /** "October 18, 2026" from "2026-10-18". */
  function longDate(iso) {
    return new Date(iso + "T12:00:00").toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  }

  /** Fetch one course file as text. Paths are relative to the course folder. */
  async function fetchCourseText(path) {
    const response = await fetch(COURSE + path, { cache: "no-cache" });
    if (!response.ok) throw new Error("Couldn't load " + path + " (" + response.status + ").");
    return response.text();
  }

  /** Move keyboard focus somewhere without the page jumping about. */
  function focusOn(node) {
    if (!node) return;
    if (!node.hasAttribute("tabindex") && !/^(A|BUTTON|INPUT|SELECT|TEXTAREA)$/.test(node.tagName)) node.setAttribute("tabindex", "-1");
    node.focus({ preventScroll: true });
    node.scrollIntoView({ block: "start", behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
  }

  // =====================================================================
  // PART 3: SAVED PROGRESS
  // ---------------------------------------------------------------------
  // Progress is kept in this browser's localStorage, and nowhere else. It
  // never holds anything from the document: only which steps passed, which
  // badges and certificates were earned, settings, and the name the learner
  // types for their certificates. Browsers can clear it (private windows,
  // "clear site data"), so the Badges screen offers "Save my record".
  // =====================================================================

  function blankProgress() {
    return {
      format: "office-trainer-progress/1",
      settings: { name: "", headingFont: "", bodyFont: "", textSize: "m", platform: "auto", theme: "system" },
      visited: {},      // station id -> date first opened or read
      steps: {},        // activity id -> { step number: date it first passed }
      ticks: {},        // activity id -> { step number: date the learner ticked it (screenshot steps) }
      quizzes: {},      // quiz id -> date answered correctly
      badges: {},       // activity id -> the badge record
      certificates: {}, // station id -> { id, issuedAt }
      skillchecks: {},  // skill check id -> { attempts: [...], current: {...} or null }
      rubric: {},       // project id -> { criterion id: true }
      last: null,       // the station opened most recently
      sectionJoin: "",  // the section key (OTK1....) or code (PP-7K3Q) the learner joined, or "" (see PART 5b)
    };
  }

  let progress = blankProgress();

  /** Read saved progress. If there's none, or the browser blocks storage, start fresh. */
  function loadProgress() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORE_KEY));
      if (saved && saved.format === "office-trainer-progress/1") progress = cleanProgress(saved);
      else copySettingsFromOtherCourse();
    } catch (e) { /* storage blocked or damaged: keep the fresh copy */ }
  }

  /**
   * First time in this course? Bring over the name and display settings from
   * the learner's other course, so they don't have to set them twice.
   * (Only settings: badges and steps always belong to one course.)
   */
  function copySettingsFromOtherCourse() {
    for (const other of Object.values(APPS)) {
      if (other.storeKey === STORE_KEY) continue;
      const saved = JSON.parse(localStorage.getItem(other.storeKey) || "null");
      if (saved && saved.format === "office-trainer-progress/1") {
        progress.settings = cleanProgress(saved).settings;
        return;
      }
    }
  }

  /** Save progress. Failing to save is never an error the learner has to deal with. */
  function saveProgress() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(progress)); } catch (e) { /* blocked: ignore */ }
  }

  /**
   * Keep only the shapes we expect, as plain values. Used for saved progress
   * and for records the learner imports, so nothing odd reaches the page.
   */
  function cleanProgress(raw) {
    const fresh = blankProgress();
    const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);
    const str = (v, max) => String(v === undefined || v === null ? "" : v).slice(0, max || 80);
    // Names that would change how an object behaves (a saved file could hold one). Never used as keys.
    const UNSAFE = ["__proto__", "constructor", "prototype"];
    const entries = (v) => Object.entries(v).filter(([k]) => !UNSAFE.includes(k));
    const dateMap = (v) => {
      const out = {};
      if (isObj(v)) for (const [k, d] of entries(v)) out[str(k, 40)] = str(d, 10);
      return out;
    };
    const nested = (v) => {
      const out = {};
      if (isObj(v)) for (const [k, inner] of entries(v)) out[str(k, 40)] = dateMap(inner);
      return out;
    };
    if (isObj(raw.settings)) {
      for (const key of Object.keys(fresh.settings)) if (raw.settings[key] !== undefined) fresh.settings[key] = str(raw.settings[key], 60);
    }
    fresh.visited = dateMap(raw.visited);
    fresh.steps = nested(raw.steps);
    fresh.ticks = nested(raw.ticks);
    fresh.quizzes = dateMap(raw.quizzes);
    fresh.rubric = {};
    if (isObj(raw.rubric)) for (const [k, v] of entries(raw.rubric)) {
      fresh.rubric[str(k, 40)] = {};
      if (isObj(v)) for (const [c, on] of entries(v)) fresh.rubric[str(k, 40)][str(c, 40)] = on === true;
    }
    if (isObj(raw.badges)) for (const [k, b] of entries(raw.badges)) {
      if (!isObj(b)) continue;
      fresh.badges[str(k, 40)] = { id: str(b.id, 40), activity: str(b.activity, 40), name: str(b.name, 80), kind: str(b.kind, 20),
        week: Number(b.week) || null, assessedBy: Array.isArray(b.assessedBy) ? b.assessedBy.map((x) => str(x, 20)).filter((x) => TM.ASSESS_KINDS.includes(x)) : [],
        earnedAt: str(b.earnedAt, 10) };
    }
    if (isObj(raw.certificates)) for (const [k, c] of entries(raw.certificates)) {
      if (!isObj(c)) continue;
      fresh.certificates[str(k, 60)] = { id: str(c.id, 40), issuedAt: str(c.issuedAt, 10) };
      // A course certificate earned in a section remembers which section, and which stops it left out.
      if (isObj(c.section)) fresh.certificates[str(k, 60)].section = {
        title: TM.cleanText(c.section.title, 80), instructor: TM.cleanText(c.section.instructor, 60),
        skip: Array.isArray(c.section.skip) ? c.section.skip.slice(0, 200).map((x) => String(x)).filter((x) => TM.STOP_ID.test(x)) : [],
      };
    }
    if (isObj(raw.skillchecks)) for (const [k, s] of entries(raw.skillchecks)) {
      if (!isObj(s)) continue;
      fresh.skillchecks[str(k, 40)] = {
        attempts: Array.isArray(s.attempts) ? s.attempts.slice(0, 20).map((a) => ({ score: Number(a.score) || 0, passed: a.passed === true, at: str(a.at, 10), scored: a.scored !== false })) : [],
        current: isObj(s.current) ? { startedAt: Number(s.current.startedAt) || Date.now(), minutes: Number(s.current.minutes) || 0, scored: s.current.scored !== false } : null,
      };
    }
    fresh.last = raw.last ? str(raw.last, 60) : null;
    // A section key can be long, so it gets its own limit. It's checked again (TrainerMD.decodeSectionKey) every time it's used.
    fresh.sectionJoin = typeof raw.sectionJoin === "string" ? raw.sectionJoin.replace(/\s+/g, "").slice(0, 20000) : "";
    return fresh;
  }

  /** The step-number map for an activity (made on first use). */
  function stepMap(kind, activityId) {
    if (!progress[kind][activityId]) progress[kind][activityId] = {};
    return progress[kind][activityId];
  }

  // =====================================================================
  // PART 4: WORD, EXCEL, POWERPOINT, OR A PLAIN BROWSER?
  // ---------------------------------------------------------------------
  // Inside Word, Excel or PowerPoint (the add-in) we read the open file directly. In a
  // normal browser tab (a Chromebook without the add-in, a phone) learners
  // download the practice file, work on it anywhere, and upload it to check.
  // Either way the checker gets the same thing: the bytes of a .docx,
  // .xlsx or .pptx file.
  // =====================================================================

  const Host = {
    inOffice: false,         // true inside Word, Excel or PowerPoint (APP says which)
    detected: "all",        // "web", "windows", "mac" or "all" (a browser: show every platform's notes)

    /** Which platform's instructions to show: the learner's choice, or what we detected. */
    platform: function () {
      const chosen = progress.settings.platform;
      return chosen && chosen !== "auto" ? chosen : Host.detected;
    },

    /** Is this desktop Word (Windows or Mac)? */
    isDesktop: function () { return ["windows", "mac", "desktop"].includes(Host.platform()); },

    /**
     * Get the learner's file: { bytes, fileName }.
     * In Word, Excel or PowerPoint: the open file. In a browser: a file they choose.
     */
    readDocument: function () {
      return Host.inOffice ? readFromOffice() : pickFile();
    },
  };

  /**
   * Inside Word, Excel or PowerPoint: ask for a copy of the open file (.docx, .xlsx or .pptx),
   * in 64 KB slices. Each app gives the last saved version plus any changes
   * since, so the learner doesn't have to save first.
   */
  function readFromOffice() {
    return new Promise(function (resolve, reject) {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (fileResult) {
        if (fileResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(APP.name + " couldn't share the " + APP.file + ": " + fileResult.error.message));
          return;
        }
        const file = fileResult.value;
        const slices = [];
        function getSlice(index) {
          file.getSliceAsync(index, function (sliceResult) {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync(); // always close, or Word may refuse the next request
              reject(new Error(APP.name + " couldn't share part of the " + APP.file + ". Try again."));
              return;
            }
            slices.push(sliceResult.value.data);
            if (index + 1 < file.sliceCount) getSlice(index + 1);
            else {
              file.closeAsync();
              resolve({ bytes: joinSlices(slices), fileName: Office.context.document.url || "" });
            }
          });
        }
        getSlice(0);
      });
    });
  }

  /** Glue a list of byte arrays into one. */
  function joinSlices(slices) {
    let total = 0;
    slices.forEach(function (s) { total += s.length; });
    const bytes = new Uint8Array(total);
    let offset = 0;
    slices.forEach(function (s) { bytes.set(s, offset); offset += s.length; });
    return bytes;
  }

  /** In a browser: let the learner choose their saved .docx, .xlsx or .pptx file. */
  function pickFile() {
    return new Promise(function (resolve, reject) {
      const input = el("input", { type: "file", accept: APP.accept, class: "sr-only", tabindex: "-1" });
      add(document.body, input);
      input.addEventListener("change", async function () {
        const file = input.files[0];
        input.remove();
        if (!file) { reject(new Error("cancelled")); return; }
        if (!APP.fileTest.test(file.name)) { reject(new Error(APP.wrongFile)); return; }
        resolve({ bytes: new Uint8Array(await file.arrayBuffer()), fileName: file.name });
      });
      input.addEventListener("cancel", function () { input.remove(); reject(new Error("cancelled")); });
      input.click();
    });
  }

  /** Does the open Word document have any text in it? (So we can warn before replacing it.) */
  async function wordHasText() {
    let hasText = false;
    await Word.run(async function (context) {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      hasText = body.text.trim() !== "";
    });
    return hasText;
  }

  /**
   * Put a starter document into the open Word document, replacing what's there.
   * Newer Word (WordApi 1.5) copies headers, footers and page setup too; older
   * Word copies the body only, so we fall back to that.
   */
  async function loadStarterIntoWord(path) {
    const response = await fetch(COURSE + path, { cache: "no-cache" });
    if (!response.ok) throw new Error("The practice file couldn't be downloaded.");
    const blob = await response.blob();
    const base64 = await blobToBase64(blob);
    // The starter's own Title, Tags and so on, so the properties start where the lesson expects.
    let starterProps = null;
    try { starterProps = (await window.WordChecker.readDocx(new Uint8Array(await blob.arrayBuffer()))).props; } catch (e) { /* not important */ }

    await Word.run(async function (context) {
      const canCopyEverything = Office.context.requirements.isSetSupported("WordApi", "1.5");
      if (canCopyEverything) {
        context.document.insertFileFromBase64(base64, "Replace", {
          importTheme: true, importStyles: true, importParagraphSpacing: true,
          importPageColor: true, importDifferentOddEvenPages: true,
        });
      } else {
        context.document.body.insertFileFromBase64(base64, "Replace");
      }
      if (starterProps && Office.context.requirements.isSetSupported("WordApi", "1.3")) {
        context.document.properties.title = starterProps.title || "";
        context.document.properties.keywords = starterProps.keywords || "";
      }
      await context.sync();
    });
  }

  /**
   * Excel: open the practice workbook as a NEW workbook (Excel.createWorkbook,
   * ExcelApi 1.8). We don't paste its sheets into the open workbook, because
   * a workbook carries things sheets don't: properties, names, links, the
   * theme, protection. Opening the whole file keeps every one of them.
   * Desktop Excel opens a new window; Excel for the web opens a new tab.
   */
  async function openStarterInExcel(path) {
    if (!(Office.context.requirements && Office.context.requirements.isSetSupported("ExcelApi", "1.8"))) {
      throw new Error("This version of Excel can't open files for the trainer. Use Download instead.");
    }
    const response = await fetch(COURSE + path, { cache: "no-cache" });
    if (!response.ok) throw new Error("The practice file couldn't be downloaded.");
    const base64 = await blobToBase64(await response.blob());
    await Excel.createWorkbook(base64);
  }

  /**
   * PowerPoint: open the practice presentation as a NEW presentation
   * (PowerPoint.createPresentation, PowerPointApi 1.1), for the same reason as
   * Excel: a presentation carries its masters, layouts, sections and show
   * settings, and opening the whole file keeps every one of them.
   */
  async function openStarterInPowerPoint(path) {
    if (!(Office.context.requirements && Office.context.requirements.isSetSupported("PowerPointApi", "1.1"))) {
      throw new Error("This version of PowerPoint can't open files for the trainer. Use Download instead.");
    }
    const response = await fetch(COURSE + path, { cache: "no-cache" });
    if (!response.ok) throw new Error("The practice file couldn't be downloaded.");
    const base64 = await blobToBase64(await response.blob());
    await PowerPoint.createPresentation(base64);
  }

  /** Word wants base64 text. FileReader makes "data:...;base64,XXXX"; we keep the XXXX. */
  function blobToBase64(blob) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();
      reader.onload = function () { resolve(String(reader.result).split(",")[1]); };
      reader.onerror = function () { reject(new Error("Couldn't read the practice file.")); };
      reader.readAsDataURL(blob);
    });
  }

  /** "Show me": select the spot a failed check is about, in Word, Excel or PowerPoint. */
  function showMe(where) {
    if (APP.id === "powerpoint") return showMeInPowerPoint(where);
    return APP.id === "excel" ? showMeInExcel(where) : showMeInWord(where);
  }

  /**
   * PowerPoint: go to the slide a failed check is about, and select the shape
   * if the check named one. where = { slide: 3, shape: "Logo" } (shape is optional).
   * Newer PowerPoint (PowerPointApi 1.5) can select slides and shapes; older
   * versions can only go to the slide, which we do with goToByIdAsync.
   * Returns true if we got there.
   */
  async function showMeInPowerPoint(where) {
    const n = Number(where.slide);
    if (!n || n < 1) return false;
    const req = Office.context.requirements;
    if (req && req.isSetSupported("PowerPointApi", "1.5")) {
      let found = false;
      await PowerPoint.run(async function (context) {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        const slide = slides.items[n - 1];
        if (!slide) return;
        context.presentation.setSelectedSlides([slide.id]);
        if (where.shape) {
          const shapes = slide.shapes;
          shapes.load("items/id,items/name");
          await context.sync();
          const shape = shapes.items.find((x) => x.name === String(where.shape));
          if (shape) slide.setSelectedShapes([shape.id]);
        }
        await context.sync();
        found = true;
      });
      return found;
    }
    // Older PowerPoint: just go to the slide (by its position, 1 = first).
    return new Promise(function (resolve) {
      Office.context.document.goToByIdAsync(n, Office.GoToType.Index, function (result) {
        resolve(result.status === Office.AsyncResultStatus.Succeeded);
      });
    });
  }

  /**
   * Excel: go to the sheet and select the cells a failed check is about.
   * where = { sheet: "June", range: "B2:B9" } (range is optional).
   * Returns true if something was selected.
   */
  async function showMeInExcel(where) {
    let found = false;
    await Excel.run(async function (context) {
      const sheet = context.workbook.worksheets.getItemOrNullObject(String(where.sheet || ""));
      await context.sync();
      if (sheet.isNullObject) return;
      sheet.activate();
      // Only plain addresses (A1 or A1:B9) are passed on to Excel.
      if (where.range && /^\$?[A-Z]{1,3}\$?\d{1,7}(:\$?[A-Z]{1,3}\$?\d{1,7})?$/i.test(where.range)) sheet.getRange(where.range).select();
      await context.sync();
      found = true;
    });
    return found;
  }

  /**
   * Word: select the spot a failed check is about, right in the document.
   * where = { search: "text", matchCase } or { emptyParagraph: true }.
   * Returns true if something was selected.
   */
  async function showMeInWord(where) {
    let found = false;
    await Word.run(async function (context) {
      const body = context.document.body;
      if (where.emptyParagraph) {
        const paragraphs = body.paragraphs;
        paragraphs.load("items/text");
        await context.sync();
        // Skip the last paragraph: Word always keeps one, and the checker allows it.
        const items = paragraphs.items.slice(0, -1);
        const empty = items.find(function (p) { return p.text.trim() === ""; });
        if (empty) { empty.select(); found = true; }
      } else if (where.search) {
        const hits = body.search(String(where.search).slice(0, 255), { matchCase: !!where.matchCase });
        hits.load("items");
        await context.sync();
        if (hits.items.length) { hits.items[0].select(); found = true; }
      }
      await context.sync();
    });
    return found;
  }

  /** Open a page in a new browser window (desktop Word opens the learner's browser). */
  function openPage(relativeUrl) {
    const url = new URL(relativeUrl, window.location.href).href;
    try {
      if (Host.inOffice && Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (e) { /* fall back to a normal new window */ }
    window.open(url, "_blank", "noopener");
  }

  // =====================================================================
  // PART 5: LOADING THE COURSE
  // ---------------------------------------------------------------------
  // At startup we read course.config.md (fonts, video hosts, mode), then
  // content/outline.md (every week and stop), then every lesson file the
  // outline names. TrainerMD.courseMap turns those into weeks and stations;
  // the roadmap and certificate pages use the very same function.
  // =====================================================================

  const Course = {
    config: {},       // course.config.md settings
    outline: null,    // the parsed outline
    docs: {},         // lesson path -> parsed lesson
    weeks: [],        // from TrainerMD.courseMap
    stations: [],     // every station, in order, each with .week (number)
    outcomes: [], mosNames: {},
    byActivity: {},   // activity id ("D1.03") -> { station, block }
  };

  async function loadCourse() {
    Course.config = TM.parse(await fetchCourseText("course.config.md")).meta;
    Course.outline = TM.parse(await fetchCourseText("content/outline.md"));

    // Every file the outline names, fetched together. A missing file just means "planned".
    const files = [];
    Course.outline.blocks.filter((b) => b.name === "week").forEach((w) => w.steps.forEach((item) => {
      const f = item.extra.file;
      if (f && SAFE_LESSON_PATH.test(f) && !f.includes("..")) files.push(f);
    }));
    const results = await Promise.allSettled(files.map((f) => fetchCourseText(f)));
    results.forEach((r, i) => { if (r.status === "fulfilled") Course.docs[files[i]] = TM.parse(r.value); });

    const map = TM.courseMap(Course.outline, Course.config, Course.docs);
    Course.weeks = map.weeks;
    Course.outcomes = map.outcomes;
    Course.mosNames = map.mosNames;
    Course.stations = [];
    Course.weeks.forEach((w) => w.stations.forEach((s) => { s.week = w.n; Course.stations.push(s); }));

    // Find each activity's station and its block in the lesson file.
    for (const st of Course.stations) {
      const doc = st.file && Course.docs[st.file];
      if (!doc) continue;
      for (const b of doc.blocks) {
        if (b.kind === "directive" && ["drill", "skillcheck", "rubric", "preflight", "quiz"].includes(b.name)) {
          Course.byActivity[b.arg] = { station: st, block: b };
        }
      }
    }
  }

  /** The course's mode: "practice" (no section) unless the course says otherwise. */
  function courseMode() { return Course.config.mode === "section" ? "section" : "practice"; }

  // =====================================================================
  // PART 5b: SECTIONS (an instructor's choice of lessons and videos)
  // ---------------------------------------------------------------------
  // An instructor sets up a section on instructor/sections.html and gives
  // the class either a SECTION KEY (one long line starting "OTK1.", with
  // everything inside it) or a short SECTION CODE such as PP-7K3Q (which
  // names a file, sections/PP-7K3Q.md, on this website).
  // Joining a section changes only what this panel SHOWS: the stops the
  // instructor left out are marked "Not in your section" (and don't count
  // toward the course certificate), and the instructor's videos replace or
  // join the course videos. Nothing is sent to anyone: progress still stays
  // on this device. TrainerMD checks every key and file before we use it
  // (lengths, ids, approved video sites), because anyone can make a key.
  // =====================================================================

  const Section = {
    active: null,    // the joined section (TrainerMD's section shape), or null
    doc: null,       // the same section as replace-video / add-video blocks, for TrainerMD.applyOverrides
    problems: [],    // parts of the key that were skipped (shown in Settings)
    error: "",       // why a saved section couldn't be loaded this time (shown on Home)
    pending: null,   // a section offered by a ?join= link, waiting for the learner to say yes
  };

  /** What a section is checked against: this course, its approved video sites, and its stop, video and activity ids. */
  function sectionContext() {
    const index = TM.courseIndex(Object.values(Course.docs));
    const hosts = Array.isArray(Course.config.videoHosts) ? Course.config.videoHosts : TM.DEFAULT_VIDEO_HOSTS;
    return { app: APP.id, videoHosts: hosts, stopIds: index.stopIds, videoIds: index.videoIds, anchorIds: index.anchorIds };
  }

  /**
   * Turn what the learner typed (a key or a code) into a section.
   * Returns { section, problems, join } where join is the tidy text to save.
   * section is null (and problems says why) when it can't be used.
   */
  async function resolveSection(typed) {
    const text = String(typed || "").replace(/\s+/g, "");
    if (!text) return { section: null, problems: ["Paste the section key or type the code your instructor gave you."], join: "" };
    if (TM.isSectionKey(text)) return Object.assign(TM.decodeSectionKey(text, sectionContext()), { join: text });

    // A short code names a section file on this website: sections/<CODE>.md
    const code = text.toUpperCase();
    if (!TM.SECTION_CODE.test(code)) {
      return { section: null, problems: ["That isn't a section key (keys start with OTK1.) or a section code (such as " + (APP.id === "word" ? "WRD" : APP.id === "excel" ? "XL" : "PP") + "-7K3Q)."], join: "" };
    }
    let response;
    try { response = await fetch(COURSE + "sections/" + code + ".md", { cache: "no-cache" }); }
    catch (e) { return { section: null, problems: ["Couldn't reach the course website to look up " + code + ". Check your internet connection and try again."], join: code }; }
    if (!response.ok) return { section: null, problems: ["There's no " + APP.name + " section with the code " + code + ". Check the code with your instructor."], join: "" };
    const doc = TM.parse(await response.text());
    if (String(doc.meta.id || "") !== code) return { section: null, problems: ["The section file for " + code + " has a different code inside it. Let your instructor know."], join: "" };
    return Object.assign(TM.sectionFromFile(doc, sectionContext()), { join: code });
  }

  /** Start using a section (or stop, with null). Saves the choice. */
  function useSection(result) {
    Section.active = result && result.section ? result.section : null;
    Section.doc = Section.active ? TM.sectionDoc(Section.active) : null;
    Section.problems = Section.active ? result.problems : [];
    Section.error = "";
    progress.sectionJoin = Section.active ? result.join : "";
    // Leaving stops out can complete the course: check the course certificate again.
    if (Course.stations.some((s) => s.kind === "course")) maybeCourseCertificate();
    saveProgress();
  }

  /** At startup: load the section saved last time, if there is one. A failure keeps the key and says so on Home. */
  async function loadSavedSection() {
    if (!progress.sectionJoin) return;
    const result = await resolveSection(progress.sectionJoin);
    if (result.section) {
      Section.active = result.section;
      Section.doc = TM.sectionDoc(result.section);
      Section.problems = result.problems;
    } else {
      Section.error = result.problems[0] || "Your section couldn't be loaded.";
    }
  }

  /** Is this stop part of the learner's section? (Everything is, with no section.) */
  function inSection(st) {
    return !Section.active || !st.docId || !Section.active.skip.includes(st.docId);
  }

  /** The stops that count: every stop, minus the ones the section leaves out. */
  function countedStations() { return Course.stations.filter(inSection); }

  // =====================================================================
  // PART 6: MOVING BETWEEN SCREENS
  // ---------------------------------------------------------------------
  // Four kinds of screen: home, a station (lesson / skill check / project /
  // info page), record (badges), settings. go() draws one into <main>,
  // moves keyboard focus to its heading, and remembers it in the address
  // (so a link like taskpane.html?lesson=W1-L2 opens that lesson).
  // =====================================================================

  let current = { screen: "home", station: null };
  let skillTimer = null; // the running skill check clock, if any
  let firstScreen = true; // the first screen after the panel opens

  function go(screen, station, focusTarget) {
    if (skillTimer) { clearInterval(skillTimer); skillTimer = null; }
    current = { screen: screen, station: station || null };
    const view = $("view");
    view.replaceChildren();
    for (const b of document.querySelectorAll("[data-nav]")) {
      b.setAttribute("aria-current", b.dataset.nav === screen ? "page" : "false");
    }
    if (screen === "home") drawHome(view);
    else if (screen === "station") drawStation(view, station);
    else if (screen === "record") drawRecord(view);
    else if (screen === "settings") drawSettings(view);
    rememberInAddress(screen, station);
    window.scrollTo(0, 0);
    const heading = view.querySelector("h1");
    if (focusTarget) setTimeout(function () { focusOn(focusTarget()); }, 30);
    else if (heading && !firstScreen) { heading.setAttribute("tabindex", "-1"); heading.focus({ preventScroll: true }); }
    firstScreen = false; // when the panel first opens, leave focus where Word put it
    updateBadgeCount();
  }

  /** Put the open lesson in the address (?lesson=W1-L2). office.js sometimes turns history off, so check first. */
  function rememberInAddress(screen, station) {
    try {
      const url = new URL(window.location.href);
      url.searchParams.delete("lesson");
      url.searchParams.delete("activity");
      if (screen === "station" && station && station.docId) url.searchParams.set("lesson", station.docId);
      if (window.history && typeof window.history.replaceState === "function") window.history.replaceState(null, "", url);
    } catch (e) { /* not important enough to show an error */ }
  }

  function updateBadgeCount() {
    $("badge-count").textContent = String(Object.keys(progress.badges).length);
  }

  // =====================================================================
  // PART 7: THE COURSE HOME SCREEN
  // =====================================================================

  const KIND_WORDS = { start: "Start here", lesson: "Lesson", skillcheck: "Skill check", project: "Project", capstone: "Capstone", course: "Course certificate" };

  /**
   * How far along a station is, in words and as a state for the styles:
   * "done", "partial", "todo" or "planned". Always words + symbol, never color alone.
   */
  function stationState(st) {
    // "skipped": the learner's section leaves this stop out (see PART 5b)
    if (st.kind === "course") {
      if (progress.certificates[st.sid]) return { state: "done", text: "✓ Certificate earned" };
      if (!sectionAllowsCourseCertificate()) return { state: "todo", text: "Your section leaves out too much of the course for a course certificate" };
      return { state: "todo", text: Section.active ? "Finish every stop in your section to earn it" : "Finish every stop to earn it" };
    }
    if (!st.file || !Course.docs[st.file]) return { state: "planned", text: "Coming soon" };
    if (!inSection(st)) {
      // Left out by the section: still open for practice, but it doesn't count.
      const got = st.activities.filter((a) => progress.badges[a.id]).length;
      return { state: "skipped", text: "Not in your section" + (got ? " · " + got + " of " + st.activities.length + " badges" : "") };
    }
    const acts = st.activities;
    const earned = acts.filter((a) => progress.badges[a.id]).length;
    if (st.kind === "skillcheck") {
      const sc = acts[0] && progress.skillchecks[acts[0].id];
      if (earned) return { state: "done", text: "✓ Passed" };
      if (sc && sc.attempts.length) return { state: "partial", text: "Best score " + Math.max.apply(null, sc.attempts.map((a) => a.score)) + "%" };
      return { state: "todo", text: "Not started" };
    }
    if (!acts.length) return progress.visited[st.sid] ? { state: "done", text: "✓ Read" } : { state: "todo", text: "Not started" };
    if (earned === acts.length) return { state: "done", text: st.kind === "lesson" ? "✓ Complete · certificate earned" : "✓ Complete" };
    if (earned) return { state: "partial", text: earned + " of " + acts.length + " badges" };
    return progress.visited[st.sid] ? { state: "partial", text: "Started · 0 of " + acts.length + " badges" } : { state: "todo", text: "Not started" };
  }

  /** The next stop to work on: the first written one that isn't done. */
  function nextStation() {
    const open = Course.stations.find((s) => s.kind !== "course" && !["done", "planned", "skipped"].includes(stationState(s).state));
    return open || Course.stations.find((s) => s.sid === progress.last) || Course.stations[0];
  }

  /** One line under a station's title: "Lesson · 50 min · 4 badges". */
  function stationMeta(st) {
    const bits = [KIND_WORDS[st.kind] || "Stop"];
    if (st.minutes) bits.push(st.minutes + " min");
    if (st.activities.length) bits.push(st.activities.length + (st.activities.length === 1 ? " badge" : " badges"));
    return bits.join(" · ");
  }

  function drawHome(view) {
    const next = nextStation();
    const started = Object.keys(progress.visited).length > 0;
    // Only the stops that count (a section can leave some out).
    const allActivities = countedStations().flatMap((s) => s.activities);
    const earnedCount = allActivities.filter((a) => progress.badges[a.id]).length;
    const certStations = countedStations().filter((s) => s.certificate && s.kind !== "course" && s.activities.length);
    const certCount = certStations.filter((s) => progress.certificates[s.sid]).length;
    const startStation = Course.stations.find((s) => s.kind === "start");
    const finale = Course.stations.find((s) => s.kind === "course");
    // Are some lessons still unwritten? Then the counts are "so far".
    const morePlanned = countedStations().some((s) => s.kind !== "course" && stationState(s).state === "planned");

    // --- Heading and mode ---
    add(view, el("section", { class: "home-head", "aria-labelledby": "home-title" },
      el("p", { class: "eyebrow" }, Course.config.college || ""),
      el("h1", { id: "home-title" }, Course.config.title || APP.fullName),
      el("p", {},
        el("span", { class: "mode-chip" }, courseMode() === "practice" ? "Practice mode" : "Class section"), " ",
        el("span", { class: "small muted" }, courseMode() === "practice"
          ? "Nothing is graded, and your progress stays on this device."
          : "Your instructor sees your results."),
        startStation ? [" ", el("button", { type: "button", class: "linkish small", onclick: () => go("station", startStation) }, "How this course works")] : null),
      sectionLine(),
      Host.inOffice ? null : courseSwitcher()));

    // --- A ?join= link offered a section: ask first ---
    if (Section.pending) add(view, drawJoinOffer());
    // --- The saved section couldn't load this time ---
    if (Section.error) {
      add(view, el("div", { class: "notice warn", role: "status" },
        el("p", {}, el("b", {}, "Your section couldn't be loaded. "), Section.error),
        el("p", { class: "small" }, "You're seeing the whole course for now. ",
          el("button", { type: "button", class: "linkish", onclick: () => go("settings", null, () => $("section-input")) }, "Section settings"))));
    }

    // --- The course certificate, once it's earned ---
    if (finale && progress.certificates[finale.sid]) {
      add(view, el("section", { class: "card continue", "aria-labelledby": "course-done-title" },
        el("p", { class: "eyebrow" }, "Course complete"),
        el("h2", { id: "course-done-title" }, "You finished every stop"),
        el("p", { class: "small" }, "Certificate no. ", el("span", { class: "mono" }, progress.certificates[finale.sid].id), " · issued " + longDate(progress.certificates[finale.sid].issuedAt)),
        el("button", { type: "button", class: "btn primary wide", onclick: () => viewCertificate(finale) }, "View and print your course certificate")));
    }

    // --- Up next ---
    if (next) {
      const state = stationState(next);
      add(view, el("section", { class: "card continue", "data-week": next.week, "aria-labelledby": "next-title" },
        el("p", { class: "eyebrow" }, started ? "Pick up where you left off" : "Start here"),
        el("h2", { id: "next-title" }, next.title),
        el("p", { class: "small muted" }, "Week " + next.week + " · " + stationMeta(next) + " · " + state.text),
        el("button", { type: "button", class: "btn primary wide", onclick: () => go("station", next) }, started ? "Continue" : "Start the course")));
    }

    // --- Progress ---
    add(view, el("section", { class: "card", "aria-labelledby": "progress-title" },
      el("h2", { id: "progress-title" }, "Your progress"),
      el("div", { class: "meter" },
        el("label", { for: "badge-meter", class: "small" }, earnedCount + " of " + allActivities.length + (morePlanned ? " badges available so far" : " badges")),
        el("progress", { id: "badge-meter", max: Math.max(allActivities.length, 1), value: earnedCount })),
      el("p", { class: "small muted", style: "margin:6px 0 0" }, certCount + " of " + certStations.length + " lesson certificates" + (morePlanned ? " · more lessons are on the way" : "")),
      el("p", { class: "small", style: "margin:6px 0 0" },
        el("button", { type: "button", class: "linkish", onclick: () => openBadgesPage() }, "See all your badges on a full page"),
        el("span", { class: "muted" }, " (opens in your browser)"))));

    // --- The weeks, each a line of stations ---
    const openWeek = next ? next.week : 1;
    const weeks = el("section", { "aria-labelledby": "weeks-title" }, el("h2", { id: "weeks-title", class: "sr-only" }, "Weeks"));
    for (const w of Course.weeks) add(weeks, drawWeek(w, w.n === openWeek));
    add(view, weeks);

    add(view, el("p", { class: "small" },
      el("button", { type: "button", class: "linkish", onclick: () => openPage(COURSE + "roadmap.html") }, "Open the full roadmap"),
      el("span", { class: "muted" }, " (opens in your browser)")));
  }

  /**
   * In a browser tab, learners pick a course here (inside Word, Excel or PowerPoint the
   * app decides). Choosing one reloads the page with ?app=..., so everything
   * starts fresh from that course's files and progress.
   */
  function courseSwitcher() {
    const links = Object.values(APPS).map((a) => {
      const here = a.id === APP.id;
      const url = new URL(window.location.href);
      url.search = "";
      url.searchParams.set("app", a.id);
      return el("a", { href: url.pathname + url.search, class: "chip", "aria-current": here ? "page" : null,
        onclick: () => { try { localStorage.setItem(APP_CHOICE_KEY, a.id); } catch (e) { /* not important */ } } },
        here ? "✓ " : "", a.name);
    });
    return el("p", { class: "small course-switch" }, el("span", { class: "muted" }, "Course: "), links.map((l, i) => [i ? " " : "", l]));
  }

  /** Under the course title: the learner's section, or a way to join one. */
  function sectionLine() {
    if (Section.active) {
      const s = Section.active;
      return el("p", { class: "small section-line" },
        el("span", { class: "mode-chip" }, "Section"), " ",
        el("b", {}, s.title), el("span", { class: "muted" }, " · " + s.instructor + (s.skip.length ? " · " + s.skip.length + (s.skip.length === 1 ? " stop" : " stops") + " left out" : "")), " ",
        el("button", { type: "button", class: "linkish small", onclick: () => go("settings", null, () => $("section-card")) }, "Section settings"));
    }
    return el("p", { class: "small section-line" },
      el("button", { type: "button", class: "linkish small", onclick: () => go("settings", null, () => $("section-input")) }, "Have a section key from your instructor?"));
  }

  /** A link with ?join=... offered a section. Show what it is and let the learner choose. */
  function drawJoinOffer() {
    const offer = Section.pending;
    const card = el("section", { class: "card continue", id: "join-offer", "aria-labelledby": "join-title", tabindex: "-1" });
    if (!offer.section) {
      fill(card, el("h2", { id: "join-title" }, "That section link didn't work"),
        el("p", { class: "small" }, offer.problems[0] || "The link is damaged."),
        el("button", { type: "button", class: "btn slim quiet", onclick: () => { Section.pending = null; go("home"); } }, "OK"));
      return card;
    }
    const s = offer.section;
    const yes = el("button", { type: "button", class: "btn primary" }, "Join this section");
    const no = el("button", { type: "button", class: "btn quiet" }, "No thanks");
    yes.addEventListener("click", () => { useSection(offer); Section.pending = null; go("home"); announce("You joined " + s.title + "."); });
    no.addEventListener("click", () => { Section.pending = null; go("home"); });
    fill(card,
      el("p", { class: "eyebrow" }, "Join a section?"),
      el("h2", { id: "join-title" }, s.title),
      el("p", { class: "small" }, "From " + s.instructor + (s.term ? " · " + s.term : "") + ". " + describeSection(s)),
      Section.active ? el("p", { class: "small" }, "This replaces your current section, " + Section.active.title + ".") : null,
      el("p", { class: "small muted" }, "Joining changes which lessons and videos you see. Your progress still stays on this device, and nothing is sent to your instructor."),
      el("div", { class: "btn-row" }, yes, no));
    return card;
  }

  /** "It leaves out 2 stops and adds 3 instructor videos." */
  function describeSection(s) {
    const bits = [];
    if (s.skip.length) bits.push("leaves out " + s.skip.length + (s.skip.length === 1 ? " stop" : " stops"));
    if (s.videos.length) bits.push("has " + s.videos.length + " video" + (s.videos.length === 1 ? "" : "s") + " from your instructor");
    return bits.length ? "It " + bits.join(" and ") + "." : "It uses the whole course as it is.";
  }

  function drawWeek(w, open) {
    const written = w.stations.filter((s) => s.file && Course.docs[s.file]).length;
    const counted = w.stations.filter(inSection);
    const done = counted.filter((s) => stationState(s).state === "done").length;
    const listId = "week-" + w.n + "-stops";
    const toggle = el("button", { type: "button", class: "week-toggle", "aria-expanded": String(open), "aria-controls": listId },
      el("span", { class: "week-swatch", "aria-hidden": "true" }),
      el("span", { class: "week-name" },
        el("b", {}, "Week " + w.n + ": " + w.title),
        el("span", {}, written ? done + " of " + counted.length + " stops done" : "Coming soon · " + w.stations.length + " stops")),
      icon("chevron"));
    const list = el("ol", { class: "line", id: listId, hidden: !open });
    for (const st of w.stations) add(list, drawStationRow(st));
    toggle.addEventListener("click", function () {
      const isOpen = toggle.getAttribute("aria-expanded") === "true";
      toggle.setAttribute("aria-expanded", String(!isOpen));
      list.hidden = isOpen;
    });
    return el("section", { class: "week", "data-week": w.n }, el("h3", { class: "sr-only" }, "Week " + w.n), toggle, list);
  }

  function drawStationRow(st) {
    const state = stationState(st);
    const stop = el("span", { class: "stop", "aria-hidden": "true" }, state.state === "done" ? icon("check") : null);
    if (state.state === "planned" || st.kind === "course") {
      return el("li", { class: "station", "data-state": state.state }, stop,
        el("div", { class: "station-planned" },
          el("span", { class: "station-title" }, st.title),
          el("span", { class: "station-meta" }, stationMeta(st)),
          el("span", { class: "station-status", "data-state": state.state }, state.text)));
    }
    return el("li", { class: "station", "data-state": state.state }, stop,
      el("button", { type: "button", class: "station-btn", onclick: () => go("station", st) },
        el("span", { class: "station-title" }, st.title),
        el("span", { class: "station-meta" }, stationMeta(st)),
        el("span", { class: "station-status", "data-state": state.state }, state.text)));
  }

  // =====================================================================
  // PART 8: A LESSON
  // ---------------------------------------------------------------------
  // TrainerMD.render draws the lesson exactly as the editor previews it.
  // Then we "wire it up": videos get a Play button, quizzes can be answered,
  // and drills get a practice-file bar, live step results and badges.
  // =====================================================================

  /** Options for TrainerMD.render, from the course settings and this learner's platform. */
  function renderOptions() {
    const hosts = Array.isArray(Course.config.videoHosts) ? Course.config.videoHosts : TM.DEFAULT_VIDEO_HOSTS;
    return { platform: Host.platform(), showChecks: false, videoHosts: hosts, mode: courseMode(), assetBase: COURSE, appName: APP.name };
  }

  /** Would TrainerMD.render show something marked for this platform? (Same rule as render's.) */
  function shows(p) {
    const now = Host.platform();
    return now === "all" || p === now || (p === "desktop" && (now === "windows" || now === "mac"));
  }

  /** The steps the learner can see on this platform, with their number in the lesson file. */
  function visibleSteps(block) {
    return block.steps.map((step, index) => ({ step: step, index: index }))
      .filter((x) => !x.step.extra.platform || shows(x.step.extra.platform));
  }

  /** What kind of step is this? check (the trainer checks), attest, answer, shot (screenshot), or info. */
  function stepKind(step) {
    if (step.checks.length && step.checks.every((c) => c.name === "self.attest")) return "attest";
    if (step.checks.length) return "check";
    if (step.extra.answer !== undefined) return "answer";
    if (step.extra.screenshot !== undefined) return "shot";
    return "info";
  }

  /** Steps that must be done for the badge (everything except instruction-only steps). */
  function requiredSteps(block) {
    return visibleSteps(block).filter((x) => stepKind(x.step) !== "info").map((x) => x.index);
  }

  /** Has this step been done (passed a check, or ticked by the learner)? */
  function stepDone(activityId, index) {
    return !!((progress.steps[activityId] || {})[index] || (progress.ticks[activityId] || {})[index]);
  }

  function drawStation(view, st) {
    // The instructor's videos (if the learner is in a section) replace or join the course videos.
    const doc = Section.doc ? TM.applyOverrides(Course.docs[st.file], Section.doc) : Course.docs[st.file];
    if (!progress.visited[st.sid]) progress.visited[st.sid] = today();
    progress.last = st.sid;
    saveProgress();

    const page = el("div", { "data-week": st.week });

    // --- Where am I? ---
    const sameKind = Course.weeks.find((w) => w.n === st.week).stations.filter((s) => s.kind === st.kind);
    const where = st.kind === "lesson"
      ? "Week " + st.week + " · Lesson " + (sameKind.indexOf(st) + 1) + " of " + sameKind.length
      : "Week " + st.week + " · " + (KIND_WORDS[st.kind] || "");
    add(page, el("nav", { class: "crumbs", "aria-label": "Where you are" },
      el("button", { type: "button", class: "btn quiet slim", onclick: () => go("home") }, "‹ Course"),
      el("span", { class: "where" }, where)));
    add(page, el("div", { class: "lesson-bar", "aria-hidden": "true" }));
    if (!inSection(st)) {
      add(page, el("p", { class: "notice small" }, el("b", {}, "Not in your section. "),
        Section.active.title + " leaves this stop out. You can still do it for practice; it just doesn't count toward your course certificate."));
    }

    // --- The lesson itself, drawn by TrainerMD ---
    // We render into a <template> first: nothing inside a template loads
    // (no videos, no images) until we've finished wiring it up.
    const template = document.createElement("template");
    template.innerHTML = TM.render(doc, renderOptions()); // TrainerMD escapes all lesson text
    const article = el("article", { class: "lesson", "aria-labelledby": "lesson-h1" });
    add(article, template.content);
    const h1 = article.querySelector("h1");
    if (h1) h1.id = "lesson-h1";

    // Activity chips under the title, so it's easy to jump to each drill.
    if (st.activities.length > 1 && st.kind !== "skillcheck") {
      const strip = el("ul", { class: "activity-strip", "aria-label": "Activities in this lesson", id: "activity-strip" });
      article.querySelector(".tm-head").after(strip);
      fillActivityStrip(strip, st);
    }

    wireVideos(article);
    wireLessonLinks(article, st);
    wireQuizzes(article, doc);
    wireActivities(article, doc, st);
    if (doc.blocks.some((b) => b.name === "rubric")) wireRubric(article, doc, st); // projects and the capstone
    add(page, article);

    // --- The end of the lesson: certificate, or what's left to do ---
    if (st.activities.length && st.kind === "lesson") add(page, drawEndCard(st));
    add(page, drawStationNav(st));
    add(view, page);
  }

  function fillActivityStrip(strip, st) {
    fill(strip, ...st.activities.map((a) => {
      const done = !!progress.badges[a.id];
      return el("li", {}, el("button", { type: "button", class: "chip", "data-done": String(done),
        onclick: () => focusOn(document.getElementById("h-" + a.id) || document.getElementById("act-" + a.id)) },
        done ? "✓ " : "", a.badge.name, el("span", { class: "sr-only" }, done ? ", badge earned" : ", not done yet")));
    }));
  }

  /** Previous and next stops. */
  function drawStationNav(st) {
    // Previous and next skip the stops the section leaves out (unless this is one of them).
    let open = Course.stations.filter((s) => s.file && Course.docs[s.file] && inSection(s));
    if (!open.includes(st)) open = Course.stations.filter((s) => s.file && Course.docs[s.file]);
    const i = open.indexOf(st);
    const prev = open[i - 1], next = open[i + 1];
    return el("nav", { class: "station-nav", "aria-label": "Previous and next" },
      prev ? el("button", { type: "button", class: "btn quiet slim", onclick: () => go("station", prev) }, "‹ " + prev.title) : el("span"),
      next ? el("button", { type: "button", class: "btn slim", onclick: () => go("station", next) }, next.title + " ›")
        : el("button", { type: "button", class: "btn slim", onclick: () => go("home") }, "Back to the course"));
  }

  // ---------- Videos: nothing loads until Play is pressed ----------
  function wireVideos(root) {
    for (const frame of root.querySelectorAll(".tm-video-box iframe")) {
      const box = frame.parentNode;
      const title = frame.getAttribute("title") || "Video";
      frame.remove(); // so it never loads by itself
      const play = el("button", { type: "button", class: "play-btn", "aria-label": "Play video: " + title },
        icon("play"), el("span", {}, "Play video"), el("small", {}, "Loads from the video site when you press play"));
      play.addEventListener("click", function () {
        const src = new URL(frame.getAttribute("src"));
        src.searchParams.set("autoplay", "1"); // they pressed play, so start playing
        frame.setAttribute("src", src.href);
        fill(box, frame);
        frame.focus();
      });
      add(box, play);
    }
  }

  // ---------- Links from one lesson to another (02-clean-text.md) ----------
  function wireLessonLinks(root, st) {
    for (const link of root.querySelectorAll("a.tm-lesson-link")) {
      link.addEventListener("click", function (event) {
        event.preventDefault();
        const folder = st.file.split("/").slice(0, -1);
        const parts = folder.concat(link.getAttribute("href").split("#")[0].split("/"));
        const resolved = [];
        parts.forEach((p) => { if (p === "..") resolved.pop(); else if (p && p !== ".") resolved.push(p); });
        const target = Course.stations.find((s) => s.file === resolved.join("/"));
        if (target) go("station", target);
      });
    }
  }

  // ---------- Quizzes ----------
  function wireQuizzes(root, doc) {
    const quizzes = doc.blocks.filter((b) => b.name === "quiz");
    root.querySelectorAll(".tm-quiz").forEach(function (box, i) {
      const block = quizzes[i];
      if (!block) return;
      const button = box.querySelector(".tm-quiz-check");
      const feedback = box.querySelector(".tm-feedback");
      const explain = box.querySelector(".tm-explain");
      let tries = 0;
      if (progress.quizzes[block.arg]) {
        feedback.className = "tm-feedback right";
        feedback.textContent = "✓ You answered this one correctly before.";
      }
      button.addEventListener("click", function () {
        let right, answered;
        const numberBox = box.querySelector("input[type=number]");
        if (numberBox) {
          answered = numberBox.value.trim() !== "";
          right = Number(numberBox.value) === Number(numberBox.dataset.answer);
        } else {
          const options = Array.from(box.querySelectorAll("input[data-correct]"));
          answered = options.some((o) => o.checked);
          right = options.every((o) => o.checked === (o.dataset.correct === "true"));
        }
        if (!answered) {
          feedback.className = "tm-feedback wrong";
          feedback.textContent = "Choose an answer first.";
          return;
        }
        tries++;
        if (right) {
          feedback.className = "tm-feedback right";
          feedback.textContent = "✓ Correct.";
          if (explain) explain.hidden = false;
          progress.quizzes[block.arg] = progress.quizzes[block.arg] || today();
          saveProgress();
        } else {
          feedback.className = "tm-feedback wrong";
          feedback.textContent = tries >= 2 ? "Not quite. Here's the explanation:" : "Not quite. Try again.";
          if (explain && tries >= 2) explain.hidden = false;
        }
      });
    });
  }

  // ---------- Drills, skill checks and preflights ----------
  function wireActivities(root, doc, st) {
    const blocks = doc.blocks.filter((b) => b.name === "drill" || b.name === "skillcheck" || b.name === "preflight");
    root.querySelectorAll("section.tm-drill").forEach(function (section, i) {
      const block = blocks[i];
      if (!block) return;
      section.id = "act-" + block.arg;
      if (block.name === "skillcheck") wireSkillCheck(section, block, st, doc);
      else wireDrill(section, block, st, doc);
    });
  }

  /** The practice-file bar: open the starter in Word, Excel or PowerPoint, or download it. */
  function starterBar(path, st, extraFiles) {
    // Other files the drill needs (a photo to insert, a recipient list, another version to compare)
    const extras = TM.splitListPublic(extraFiles).filter((f) => SAFE_EXTRA_PATH.test(f) && !f.includes(".."));
    const hasStarter = path && path !== "none" && APP.starter.test(path) && !path.includes("..");
    if (!hasStarter) {
      // No practice file (a project's preflight), but maybe files to download.
      return extras.length ? add(el("div", { class: "starter-bar" }), extrasLine(extras, "Download: ")) : null;
    }
    const fileName = path.split("/").pop();
    const bar = el("div", { class: "starter-bar" });
    const info = el("div", { class: "file" }, el("b", {}, "Practice file"), el("span", { class: "mono" }, fileName));
    add(bar, info);
    if (Host.inOffice && (APP.id === "excel" || APP.id === "powerpoint")) {
      // Excel and PowerPoint open the practice file as its own workbook or
      // presentation (see openStarterInExcel and openStarterInPowerPoint).
      const open = el("button", { type: "button", class: "btn slim" }, "Open it in " + APP.name);
      const status = el("span", { class: "small", role: "status" });
      open.addEventListener("click", async function () {
        open.disabled = true;
        status.textContent = "Opening…";
        try {
          await (APP.id === "excel" ? openStarterInExcel(path) : openStarterInPowerPoint(path));
          status.textContent = "✓ Opened in a new " + APP.name + " window. In that window, open Office Trainer (Home tab) to check your work there.";
          announce("Practice file opened in a new window.");
        } catch (e) {
          status.textContent = "Couldn't open it: " + e.message;
        } finally { open.disabled = false; }
      });
      add(bar, open, el("a", { class: "btn slim quiet", href: COURSE + path, download: fileName }, "Download"), status);
    } else if (Host.inOffice) {
      const loaded = sessionLoaded[path];
      const open = el("button", { type: "button", class: "btn slim" + (loaded ? " quiet" : "") }, loaded ? "Load a fresh copy" : "Open it here");
      const status = el("span", { class: "small", role: "status" }, loaded ? "✓ Open in this " + APP.file : "");
      open.addEventListener("click", () => startLoad(path, bar, open, status));
      add(bar, open, status);
    } else {
      add(bar, el("a", { class: "btn slim", href: COURSE + path, download: fileName }, "Download"),
        el("span", { class: "small muted", style: "flex-basis:100%" }, "Open it in " + APP.name + ", do the steps, save, then use Upload and check."));
    }
    if (extras.length) add(bar, extrasLine(extras, "Also download: "));
    return bar;
  }

  /** "Also download: logo.png · rentals.csv" - links to the extra files a drill needs. */
  function extrasLine(extras, label) {
    return el("div", { class: "extras small", style: "flex-basis:100%" },
      el("b", {}, label),
      extras.map((f, i) => [i ? " · " : "", el("a", { href: COURSE + f, download: f.split("/").pop() }, f.split("/").pop())]));
  }
  const sessionLoaded = {}; // starter path -> true once loaded into Word in this session

  /** Loading a starter replaces the document, so ask first if it has text in it. */
  async function startLoad(path, bar, button, status) {
    status.textContent = "";
    const old = bar.querySelector(".confirm");
    if (old) old.remove();
    try {
      if (await wordHasText()) {
        const yes = el("button", { type: "button", class: "btn slim danger" }, "Yes, replace it");
        const no = el("button", { type: "button", class: "btn slim quiet" }, "Cancel");
        const box = el("div", { class: "confirm", role: "group", "aria-label": "Replace the document?" },
          el("p", {}, "This replaces everything in the open document. Continue?"), el("div", { class: "btn-row" }, yes, no));
        add(bar, box);
        yes.focus();
        no.addEventListener("click", () => { box.remove(); button.focus(); });
        yes.addEventListener("click", () => { box.remove(); doLoad(path, button, status); });
      } else {
        await doLoad(path, button, status);
      }
    } catch (e) {
      status.textContent = "Couldn't read the document: " + e.message;
    }
  }
  async function doLoad(path, button, status) {
    button.disabled = true;
    status.textContent = "Loading…";
    try {
      await loadStarterIntoWord(path);
      sessionLoaded[path] = true;
      button.textContent = "Load a fresh copy";
      button.classList.add("quiet");
      status.textContent = "✓ Practice file loaded. Start with step 1.";
      announce("Practice file loaded.");
    } catch (e) {
      status.textContent = "Couldn't load the practice file: " + e.message;
    } finally {
      button.disabled = false;
      button.focus();
    }
  }

  function wireDrill(section, block, st, doc) {
    const id = block.arg;
    const practiceOnly = block.name === "preflight"; // preflights give feedback, no badge
    const head = section.querySelector(".tm-drill-head");

    // Desktop-only drill on the web: say so up front.
    if (block.header.platform === "desktop" && Host.platform() === "web") {
      head.after(el("div", { class: "notice warn small" },
        el("p", {}, el("b", {}, "This drill is for desktop " + APP.name + ". "),
          APP.name + " for the web can't do all of it. Read along and try what you can, or do it in desktop " + APP.name + " (campus labs have it). Screenshot steps work anywhere.")));
    }
    const bar = starterBar(block.header.starter, st, block.header.files);
    if (bar) section.querySelector(".tm-drill-head").after(bar);

    // Each step: show its state, add a tick box for screenshot steps.
    const items = section.querySelectorAll(":scope > ol.tm-steps > li.tm-step");
    visibleSteps(block).forEach(function (x, k) {
      const li = items[k];
      if (!li) return;
      li.dataset.step = String(x.index);
      const status = li.querySelector(".tm-step-status");
      const kind = stepKind(x.step);
      if (kind === "check" || kind === "answer") {
        status.textContent = kind === "answer" ? "Type your answer, then select Check my work." : "The trainer checks this step.";
        if (!practiceOnly && stepDone(id, x.index)) showStepResult(li, { pass: true, messages: ["Done earlier. Select Check my work to check again."] });
      } else if (kind === "shot" || kind === "attest") {
        status.textContent = "";
        const label = kind === "shot"
          ? (courseMode() === "section" ? "I added this screenshot to my evidence document." : "I took the screenshot (or wrote my description) and compared it with the example.")
          : x.step.checks[0].params.label;
        const box = el("input", { type: "checkbox" });
        box.checked = stepDone(id, x.index);
        li.dataset.result = box.checked ? "pass" : "";
        box.addEventListener("change", function () {
          const ticks = stepMap("ticks", id);
          if (box.checked) ticks[x.index] = today(); else delete ticks[x.index];
          li.dataset.result = box.checked ? "pass" : "";
          saveProgress();
          if (box.checked) maybeAward(block, section, st, doc);
        });
        add(status, el("label", { class: "tick" }, box, el("span", {}, label)));
      } else {
        status.textContent = "Do this step. The trainer doesn't check it.";
      }
    });

    // The action row: Check my work, and a summary screen readers hear.
    const actions = section.querySelector(":scope > .tm-actions");
    actions.replaceChildren();
    const checkable = visibleSteps(block).filter((x) => ["check", "answer"].includes(stepKind(x.step)));
    if (checkable.length) {
      const button = el("button", { type: "button", class: "btn primary" }, Host.inOffice ? "Check my work" : "Upload my file and check");
      const summary = el("p", { class: "drill-summary small", role: "status", "aria-live": "polite" });
      button.addEventListener("click", () => checkDrill(block, section, st, doc, button, summary, practiceOnly));
      add(actions, button, summary);
    } else if (!practiceOnly) {
      add(actions, el("p", { class: "small muted" }, "Tick each screenshot step when you've done it to earn the badge."));
    }
    if (!practiceOnly) showAward(block, section, st, doc, false);
  }

  /** Run every checked step in a drill (or preflight) against the learner's document. */
  async function checkDrill(block, section, st, doc, button, summary, practiceOnly) {
    const id = block.arg;
    const label = button.textContent;
    button.disabled = true;
    button.textContent = "Checking…";
    summary.textContent = "";
    try {
      const file = await Host.readDocument();
      const model = await APP.read(file.bytes, { fileName: file.fileName });
      let checked = 0, passed = 0;
      for (const x of visibleSteps(block)) {
        const kind = stepKind(x.step);
        if (kind !== "check" && kind !== "answer") continue;
        const li = section.querySelector(':scope > ol.tm-steps > li[data-step="' + x.index + '"]');
        const result = kind === "answer" ? checkAnswer(li, x.step) : checker().runStep(model, x.step);
        checked++;
        if (result.pass) {
          passed++;
          if (!practiceOnly) { const steps = stepMap("steps", id); steps[x.index] = steps[x.index] || today(); }
        }
        showStepResult(li, result);
      }
      if (!practiceOnly) saveProgress();
      const fileNote = Host.inOffice ? "" : " (checked " + file.fileName + ")";
      summary.textContent = passed === checked
        ? "✓ All " + checked + " checked " + (checked === 1 ? "step is" : "steps are") + " done" + fileNote + "."
        : passed + " of " + checked + " checked steps done" + fileNote + ". See the notes under each step.";
      announce(summary.textContent);
      if (!practiceOnly) maybeAward(block, section, st, doc);
    } catch (e) {
      if (e.message !== "cancelled") summary.textContent = e.message; // plain text, never code
    } finally {
      button.disabled = false;
      button.textContent = label;
    }
  }

  /** An "answer" step: compare what the learner typed with the answer in the lesson file. */
  function checkAnswer(li, step) {
    const input = li && li.querySelector(".tm-answer input");
    const typed = input ? input.value.trim() : "";
    if (!typed) return { pass: false, messages: ["Type your answer first."] };
    const want = String(step.extra.answer).trim();
    const same = (!isNaN(Number(typed)) && !isNaN(Number(want))) ? Number(typed) === Number(want) : typed.toLowerCase() === want.toLowerCase();
    return same ? { pass: true, messages: ["That's right."] } : { pass: false, messages: ["That's not it yet. Count again."] };
  }

  /** Show a step's result under it: a word, the messages, and (in Word, Excel or PowerPoint) a Show me button. */
  function showStepResult(li, result, options) {
    if (!li) return;
    const opts = options || {};
    li.dataset.result = result.pass ? "pass" : "fail";
    let box = li.querySelector(":scope > .result");
    if (!box) { box = el("div", { class: "result" }); add(li, box); }
    box.dataset.state = result.pass ? "pass" : "fail";
    fill(box, el("b", {}, result.pass ? "✓ Done" : "Not yet"));
    if (!opts.hideMessages) {
      // Messages can quote the document, so they go in as plain text.
      if (result.messages.length === 1) add(box, el("span", {}, result.messages[0]));
      else add(box, el("ul", {}, result.messages.map((m) => el("li", {}, m))));
    }
    if (!result.pass && result.where && Host.inOffice && !opts.hideMessages) {
      const show = el("button", { type: "button", class: "btn slim quiet" }, "Show me where");
      const note = el("span", { class: "small", role: "status" });
      show.addEventListener("click", async function () {
        try {
          const found = await showMe(result.where);
          note.textContent = found ? " Selected in the " + APP.file + "." : " I couldn't find that spot. Check again after your next change.";
        } catch (e) { note.textContent = " Couldn't select it: " + e.message; }
      });
      add(box, el("div", {}, show, note));
    }
  }

  // =====================================================================
  // PART 9: BADGES AND CERTIFICATES
  // ---------------------------------------------------------------------
  // A badge is earned when every step that counts is done: checked steps
  // passed, screenshot steps ticked, answers right. A lesson's certificate
  // is issued when all of its badges are earned. Numbers come from
  // TrainerMD.makeId: unique, with a check character that catches typos.
  // In practice mode they are NOT verifiable (see badges-and-certificates.md).
  // =====================================================================

  /** The station's summary of an activity (its badge name, kind, how it's assessed). */
  function activityInfo(st, id) { return st.activities.find((a) => a.id === id); }

  function maybeAward(block, section, st, doc) {
    const id = block.arg;
    if (progress.badges[id]) return;
    const required = requiredSteps(block);
    if (!required.length || !required.every((i) => stepDone(id, i))) return;
    awardBadge(st, id);
    showAward(block, section, st, doc, true);
  }

  /** Record a badge. Returns the record. */
  function awardBadge(st, id) {
    const info = activityInfo(st, id);
    const badge = {
      id: TM.makeId("OTB-" + id.replace(/[^A-Za-z0-9]/g, "")),
      activity: id, name: info.badge.name, kind: info.badge.kind, week: info.badge.week || st.week,
      assessedBy: info.assess, earnedAt: today(),
    };
    progress.badges[id] = badge;
    maybeCertificate(st);
    maybeCourseCertificate(); // the last badge might be a skill check or the capstone, not a lesson
    saveProgress();
    updateBadgeCount();
    announce("Badge earned: " + badge.name + ".");
    const strip = $("activity-strip");
    if (strip) fillActivityStrip(strip, st);
    refreshEndCard(st);
    return badge;
  }

  /** The "Badge earned" banner at the bottom of an activity. isNew = just earned (animate it). */
  function showAward(block, section, st, doc, isNew) {
    const badge = progress.badges[block.arg];
    const old = section.querySelector(":scope > .award-banner");
    if (old) old.remove();
    if (!badge) return;
    section.classList.add("is-earned");
    const award = section.querySelector(".tm-award span");
    if (award) award.textContent = "Badge earned: " + badge.name;

    // What's next: the next activity in this lesson, or the certificate.
    const ids = st.activities.map((a) => a.id);
    const nextId = ids.slice(ids.indexOf(block.arg) + 1).find((x) => !progress.badges[x]) || ids.find((x) => !progress.badges[x]);
    let nextButton = null;
    if (nextId) {
      nextButton = el("button", { type: "button", class: "btn slim", onclick: () => focusOn(document.getElementById("h-" + nextId) || document.getElementById("act-" + nextId)) },
        "Next: " + nextId + " ›");
    } else if ($("end-card")) {
      nextButton = el("button", { type: "button", class: "btn slim primary", onclick: () => focusOn($("end-card")) }, "See your certificate ›");
    }
    const banner = el("div", { class: "award-banner" + (isNew ? " new" : "") },
      badgeArt({ week: badge.week, kind: badge.kind, code: badge.activity, label: badge.name + " badge", size: 64 }),
      el("div", {},
        el("b", {}, (isNew ? "Badge earned: " : "✓ Badge earned: ") + badge.name),
        el("span", { class: "small" }, "No. ", el("span", { class: "mono" }, badge.id)),
        nextButton ? el("div", { style: "margin-top:6px" }, nextButton) : null));
    add(section, banner);
  }

  /** Issue the lesson certificate once every badge in the lesson is earned. */
  function maybeCertificate(st) {
    if (!st.certificate || !st.activities.length || progress.certificates[st.sid]) return;
    if (!st.activities.every((a) => progress.badges[a.id])) return;
    const code = (st.docId || st.sid).replace(/[^A-Za-z0-9]/g, "");
    progress.certificates[st.sid] = { id: TM.makeId("OT-" + code), issuedAt: today() };
    announce("Lesson complete. Your certificate is ready.");
    maybeCourseCertificate();
  }

  /**
   * Can a course certificate be earned in this section? A section key can come
   * from anyone, so one that leaves out most of the course can't hand out a
   * course certificate: the section must keep at least half of the stops that
   * have badges.
   */
  function sectionAllowsCourseCertificate() {
    if (!Section.active || !Section.active.skip.length) return true;
    const withBadges = Course.stations.filter((s) => s.kind !== "course" && s.activities.length);
    const kept = withBadges.filter(inSection);
    return kept.length >= 1 && kept.length * 2 >= withBadges.length;
  }

  /** The course certificate: every stop written and finished (not reachable until all weeks exist). */
  function maybeCourseCertificate() {
    const finale = Course.stations.find((s) => s.kind === "course");
    if (!finale || progress.certificates[finale.sid]) return;
    if (!sectionAllowsCourseCertificate()) return;
    const others = countedStations().filter((s) => s !== finale); // a section's left-out stops don't count
    if (others.some((s) => stationState(s).state === "planned")) return;
    if (!others.every((s) => s.activities.every((a) => progress.badges[a.id]))) return;
    // The course certificate's number names the course: OT-WORD-..., OT-EXCEL-... or OT-PPT-...
    const cert = { id: TM.makeId("OT-" + APP.certCode), issuedAt: today() };
    // Earned in a section that left stops out: remember that, so the certificate says so for good.
    if (Section.active && Section.active.skip.length) cert.section = { title: Section.active.title, instructor: Section.active.instructor, skip: Section.active.skip.slice() };
    progress.certificates[finale.sid] = cert;
  }

  /** The certificate record, in the format certificate.html reads (office-trainer-certificate/1). */
  function certificateRecord(st) {
    const cert = progress.certificates[st.sid];
    const isCourse = st.kind === "course";
    // A course certificate earned in a section that left stops out says so, and lists only what
    // counted. That comes from the certificate itself (saved when it was issued), not from
    // whatever section the learner is in today.
    const snap = isCourse && cert.section && cert.section.skip.length ? cert.section : null;
    const partial = !!snap;
    const counted = Course.stations.filter((s) => s.kind !== "course" && !(snap && s.docId && snap.skip.includes(s.docId)));
    const acts = isCourse ? counted.flatMap((s) => s.activities) : st.activities;
    const union = (key) => Array.from(new Set(counted.flatMap((s) => s[key] || []))).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    return {
      format: "office-trainer-certificate/1",
      id: cert.id, issuedAt: cert.issuedAt,
      issuer: Course.config.issuer || Course.config.college || "",
      mode: courseMode(),
      learner: { name: progress.settings.name || "Learner" },
      section: null,
      achievement: { kind: isCourse ? "course" : "lesson", id: st.sid, title: partial ? st.title + ": " + snap.title : st.title, week: isCourse ? null : st.week },
      badges: acts.map((a) => progress.badges[a.id]).filter(Boolean),
      outcomes: partial ? union("clos") : st.clos, mos: partial ? union("mos") : st.mos,
    };
  }

  /** Open certificate.html with the record in the part of the address after #, which never leaves the browser. */
  function viewCertificate(st) {
    const json = JSON.stringify(certificateRecord(st));
    const bytes = new TextEncoder().encode(json);
    let binary = "";
    bytes.forEach((b) => { binary += String.fromCharCode(b); });
    const base64url = btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    openPage(COURSE + "certificate.html#record=" + base64url);
  }

  /** Save one certificate's record as a .json file (certificate.html can open it again later). */
  function downloadCertificate(st) {
    const record = certificateRecord(st);
    downloadJson(record, record.id + ".json");
  }

  function downloadJson(data, fileName) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = el("a", { href: url, download: fileName });
    add(document.body, a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  /** The card at the end of a lesson: the certificate, or the badges still to earn. */
  function drawEndCard(st) {
    const card = el("section", { class: "card end-card", id: "end-card", "aria-labelledby": "end-title", tabindex: "-1" });
    fillEndCard(card, st);
    return card;
  }
  function refreshEndCard(st) {
    const card = $("end-card");
    if (card) fillEndCard(card, st);
  }
  function fillEndCard(card, st) {
    const cert = progress.certificates[st.sid];
    const badgeRow = el("div", { class: "badge-row" }, st.activities.map((a) => {
      const earned = !!progress.badges[a.id];
      const art = badgeArt({ week: st.week, kind: a.badge.kind, code: a.id, label: a.badge.name + (earned ? " badge, earned" : " badge, not earned yet"), size: 52 });
      if (!earned) { art.style.filter = "grayscale(1)"; art.style.opacity = ".5"; } // not earned yet: shown faded
      return el("button", { type: "button", class: "linkish", title: a.badge.name, onclick: () => focusOn(document.getElementById("h-" + a.id) || document.getElementById("act-" + a.id)) }, art);
    }));
    if (!cert) {
      const left = st.activities.filter((a) => !progress.badges[a.id]);
      fill(card, 
        el("h2", { id: "end-title" }, "Finish the lesson"),
        el("p", { class: "small" }, "Earn all " + st.activities.length + " badges to get this lesson's certificate. Still to do: " + left.map((a) => a.id).join(", ") + "."),
        badgeRow);
      return;
    }
    const nameInput = el("input", { type: "text", id: "cert-name", maxlength: "60", autocomplete: "name", value: progress.settings.name });
    const saveName = el("button", { type: "button", class: "btn slim quiet" }, "Save name");
    const nameNote = el("span", { class: "small", role: "status" });
    saveName.addEventListener("click", function () {
      progress.settings.name = nameInput.value.trim().slice(0, 60);
      saveProgress();
      nameNote.textContent = progress.settings.name ? " Saved." : " Name cleared.";
    });
    fill(card, 
      el("p", { class: "eyebrow" }, "Lesson complete"),
      el("h2", { id: "end-title" }, "Your certificate is ready"),
      badgeRow,
      el("p", { class: "small" }, "Certificate no. ", el("span", { class: "mono" }, cert.id), " · issued " + longDate(cert.issuedAt)),
      el("div", { class: "field" },
        el("label", { for: "cert-name" }, "Name on your certificates"),
        el("div", { class: "btn-row" }, nameInput, saveName, nameNote),
        el("small", {}, "Only stored on this device. Use any name you like.")),
      el("div", { class: "btn-row" },
        el("button", { type: "button", class: "btn primary", onclick: () => viewCertificate(st) }, "View and print"),
        el("button", { type: "button", class: "btn quiet", onclick: () => downloadCertificate(st) }, "Save record (.json)")),
      courseMode() === "practice" ? el("p", { class: "small muted", style: "margin:8px 0 0" }, "Practice record: made on your device, not kept or verified by the college.") : null);
  }

  // =====================================================================
  // PART 10: SKILL CHECKS
  // ---------------------------------------------------------------------
  // Like a MOS exam task: one starter, a set of tasks, a time limit, no
  // hints. "Check my work" shows only Done / Not yet for each task; the
  // full notes appear after Submit. Your score is what the document shows
  // when you submit (or when time runs out).
  // In practice mode the learner can choose extra time, or no limit:
  // timed tests must offer that (WCAG 2.2.1), and practice is never graded.
  // =====================================================================

  function wireSkillCheck(section, block, st, doc) {
    const id = block.arg;
    const record = progress.skillchecks[id] || (progress.skillchecks[id] = { attempts: [], current: null });
    const list = section.querySelector(":scope > ol.tm-steps");
    const actions = section.querySelector(":scope > .tm-actions");
    const head = section.querySelector(".tm-drill-head");

    // Number the tasks for the code below, and set each task's status line.
    const items = list.querySelectorAll(":scope > li.tm-step");
    visibleSteps(block).forEach(function (x, k) {
      const li = items[k];
      if (!li) return;
      li.dataset.step = String(x.index);
      const status = li.querySelector(".tm-step-status");
      status.textContent = stepKind(x.step) === "answer" ? "Type your answer here." : "";
    });
    const hidden = block.steps.length - visibleSteps(block).length;
    const panel = el("div", { class: "sc-start" });
    head.after(panel);

    function attemptsUsed() { return record.attempts.filter((a) => a.scored).length; }
    function attemptsAllowed() { return Number(block.header.attempts) || 0; } // 0 = no limit
    function limitMinutes() { return Number(block.header.timeLimit) || 0; }

    // ---------- Before starting ----------
    function drawStart() {
      list.hidden = true;
      actions.replaceChildren();
      const used = attemptsUsed(), allowed = attemptsAllowed();
      const left = allowed ? allowed - used : Infinity;
      const best = record.attempts.length ? Math.max.apply(null, record.attempts.map((a) => a.score)) : null;
      const bits = [];
      if (allowed) bits.push("Attempts used: " + used + " of " + allowed);
      if (best !== null) bits.push("Best score: " + best + "%");
      if (progress.badges[id]) bits.push("✓ Passed");

      const choices = el("fieldset", {}, el("legend", {}, "Time"),
        TIME_CHOICES.map((c, i) => {
          const minutes = Math.round(Number(c.value) * limitMinutes());
          const input = el("input", { type: "radio", name: "sc-time", value: c.value, checked: i === 0 });
          return el("label", {}, input, c.value === "0" ? c.label : c.label + " (" + minutes + " min)");
        }),
        el("p", { class: "small muted", style: "margin:4px 0 0" }, "Extra time is always fine in practice mode. In a class section, your instructor sets the time."));

      const start = el("button", { type: "button", class: "btn primary wide" }, left > 0 ? "Start the skill check" : "Start a practice round (not scored)");
      start.addEventListener("click", function () {
        const factor = courseMode() === "practice" ? Number((panel.querySelector("input[name=sc-time]:checked") || {}).value || 1) : 1;
        record.current = { startedAt: Date.now(), minutes: Math.round(limitMinutes() * factor), scored: left > 0 };
        saveProgress();
        drawAttempt();
        focusOn(section.querySelector(".timer"));
      });

      fill(panel, 
        bits.length ? el("p", { class: "small" }, bits.join(" · ")) : null,
        hidden ? el("p", { class: "small muted" }, hidden === 1 ? "One task needs desktop " + APP.name + ", so it isn't counted here." : hidden + " tasks need desktop " + APP.name + ", so they aren't counted here.") : null,
        starterBar(block.header.starter, st, block.header.files),
        el("p", { class: "small" }, "Open the practice file first. The clock starts when you select Start."),
        limitMinutes() && courseMode() === "practice" ? choices : null,
        start);
    }

    // ---------- During an attempt ----------
    function drawAttempt() {
      list.hidden = false;
      list.querySelectorAll(":scope > li").forEach((li) => { li.dataset.result = ""; const r = li.querySelector(":scope > .result"); if (r) r.remove(); });
      const cur = record.current;
      const time = el("span", { class: "time", "aria-hidden": "true" }, "--:--");
      const spoken = el("span", { class: "sr-only" });
      const submit = el("button", { type: "button", class: "btn slim primary" }, "Submit");
      const timer = el("div", { class: "timer", role: "timer", "aria-label": "Skill check clock", tabindex: "-1" },
        time, spoken, el("span", { class: "label" }, cur.minutes ? (cur.scored ? "Time left" : "Time left · practice round") : "No time limit · time so far"), submit);
      fill(panel, timer, starterBar(block.header.starter, st, block.header.files));

      const check = el("button", { type: "button", class: "btn" }, Host.inOffice ? "Check my work" : "Upload and check");
      const summary = el("p", { class: "drill-summary small", role: "status", "aria-live": "polite" });
      fill(actions, check, summary);
      check.addEventListener("click", () => runSkill(false, check, summary));
      submit.addEventListener("click", () => confirmSubmit(submit, summary));

      // The clock. We announce only at 5 minutes and 1 minute, so screen readers aren't flooded.
      let said5 = false, said1 = false;
      function tick() {
        const now = Date.now();
        if (!cur.minutes) { time.textContent = clock(now - cur.startedAt); return; }
        const left = cur.startedAt + cur.minutes * 60000 - now;
        time.textContent = clock(Math.max(left, 0));
        spoken.textContent = Math.ceil(Math.max(left, 0) / 60000) + " minutes left";
        timer.dataset.low = String(left < 5 * 60000);
        if (!said5 && left < 5 * 60000 && left > 4 * 60000) { said5 = true; announce("5 minutes left."); }
        if (!said1 && left < 60000 && left > 0) { said1 = true; announce("1 minute left."); }
        if (left <= 0) {
          clearInterval(skillTimer); skillTimer = null;
          announce("Time's up.");
          if (Host.inOffice) runSkill(true, check, summary);
          else { summary.textContent = "Time's up. Upload your saved file to score it."; submit.textContent = "Upload and submit"; check.hidden = true; }
        }
      }
      if (skillTimer) clearInterval(skillTimer);
      skillTimer = setInterval(tick, 1000);
      tick();
    }

    function confirmSubmit(submit, summary) {
      const old = panel.querySelector(".confirm");
      if (old) { old.remove(); return; }
      const yes = el("button", { type: "button", class: "btn slim primary" }, "Yes, submit");
      const no = el("button", { type: "button", class: "btn slim quiet" }, "Keep working");
      const box = el("div", { class: "confirm" }, el("p", {}, "Submit now? Your score is what your " + APP.file + " shows right now."), el("div", { class: "btn-row" }, yes, no));
      add(panel, box);
      yes.focus();
      no.addEventListener("click", () => { box.remove(); submit.focus(); });
      yes.addEventListener("click", () => { box.remove(); runSkill(true, submit, summary); });
    }

    /** Check every task. final = this is the Submit (score it and show full notes). */
    async function runSkill(final, button, summary) {
      const label = button.textContent;
      button.disabled = true;
      button.textContent = "Checking…";
      try {
        const file = await Host.readDocument();
        const model = await APP.read(file.bytes, { fileName: file.fileName });
        let earned = 0, total = 0, done = 0, counted = 0;
        for (const x of visibleSteps(block)) {
          const kind = stepKind(x.step);
          if (kind !== "check" && kind !== "answer") continue;
          const li = list.querySelector(':scope > li[data-step="' + x.index + '"]');
          const result = kind === "answer" ? checkAnswer(li, x.step) : checker().runStep(model, x.step);
          const points = Number(x.step.extra.points) || 1;
          total += points; counted++;
          if (result.pass) { earned += points; done++; }
          showStepResult(li, result, { hideMessages: !final });
        }
        if (!final) {
          summary.textContent = done + " of " + counted + " tasks done so far.";
          announce(summary.textContent);
          return;
        }
        finishAttempt(Math.round(earned / Math.max(total, 1) * 100));
      } catch (e) {
        if (e.message !== "cancelled") summary.textContent = e.message;
      } finally {
        button.disabled = false;
        button.textContent = label;
      }
    }

    function finishAttempt(score) {
      if (skillTimer) { clearInterval(skillTimer); skillTimer = null; }
      const cur = record.current || { scored: true };
      const passing = Number(block.header.passing) || 70;
      const passed = score >= passing;
      record.attempts.push({ score: score, passed: passed, at: today(), scored: cur.scored });
      record.current = null;
      if (passed && cur.scored && !progress.badges[id]) awardBadge(st, id);
      saveProgress();

      const allowed = attemptsAllowed();
      const left = allowed ? allowed - attemptsUsed() : Infinity;
      const again = el("button", { type: "button", class: "btn" }, left > 0 ? "Try again" : "Practice round (not scored)");
      again.addEventListener("click", () => { drawStart(); focusOn(panel); });
      fill(panel, el("div", { class: "card", tabindex: "-1", id: "sc-result" },
        el("p", { class: "eyebrow" }, cur.scored ? "Your score" : "Practice round score"),
        el("p", { class: "score" }, score + "%"),
        el("p", {}, passed ? "✓ Passed. Passing is " + passing + "%." : "Not passed yet. Passing is " + passing + "%. Read the notes under each task, then try again."),
        allowed ? el("p", { class: "small muted" }, "Attempts used: " + attemptsUsed() + " of " + allowed) : null,
        el("div", { class: "btn-row" }, again)));
      actions.replaceChildren();
      showAward(block, section, st, doc, passed && cur.scored);
      announce("Your score is " + score + " percent. " + (passed ? "Passed." : "Not passed yet."));
      focusOn($("sc-result"));
    }

    // Is an attempt already running (the panel was closed and reopened)?
    if (record.current) {
      drawAttempt();
      const cur = record.current;
      if (cur.minutes && Date.now() > cur.startedAt + cur.minutes * 60000) {
        panel.prepend(el("p", { class: "notice warn small" }, "Time ran out while this panel was closed. Select Submit to score what's in your " + APP.file + " now."));
      }
    } else {
      drawStart();
    }
    showAward(block, section, st, doc, false);
  }

  /** Milliseconds to "mm:ss" (or "h:mm:ss"). */
  function clock(ms) {
    const total = Math.round(ms / 1000);
    const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
    const two = (n) => String(n).padStart(2, "0");
    return h ? h + ":" + two(m) + ":" + two(s) : two(m) + ":" + two(s);
  }

  // =====================================================================
  // PART 11: PROJECTS
  // ---------------------------------------------------------------------
  // The preflight is wired like a drill (feedback only, no badge). The
  // rubric becomes a self-check in practice mode: tick every line your
  // work meets, then mark the project complete for its badge. In a class
  // section, the instructor grades it instead.
  // =====================================================================

  function wireRubric(article, doc, st) {
    const block = doc.blocks.find((b) => b.name === "rubric");
    const table = article.querySelector(".tm-rubric");
    if (!block || !table) return;
    const id = block.arg;
    const holder = table.closest(".tm-table");
    const box = el("section", { class: "card", id: "act-" + id, "aria-labelledby": "h-" + id });
    holder.after(box);

    if (courseMode() === "section") {
      add(box, el("h2", { id: "h-" + id }, "How it's graded"),
        el("p", {}, "Your instructor grades this project with the rubric above. Hand in your files where your instructor asks (for example, in Brightspace)."));
      return;
    }
    const ticks = progress.rubric[id] || (progress.rubric[id] = {});
    const boxes = block.items.map((item) => {
      const input = el("input", { type: "checkbox" });
      input.checked = !!ticks[item.id];
      input.addEventListener("change", function () {
        ticks[item.id] = input.checked;
        saveProgress();
        complete.disabled = !allTicked() || !!progress.badges[id];
      });
      return el("label", { class: "tick" }, input, el("span", {}, el("b", {}, item.id + " (" + item.points + " points). "), item.text));
    });
    const allTicked = () => block.items.every((item) => ticks[item.id]);
    const complete = el("button", { type: "button", class: "btn primary" }, "Mark the project complete");
    complete.disabled = !allTicked() || !!progress.badges[id];
    complete.addEventListener("click", function () {
      if (!allTicked()) return;
      awardBadge(st, id);
      complete.disabled = true;
      showAward({ arg: id }, box, st, doc, true);
    });
    add(box, 
      el("h2", { id: "h-" + id }, "Self-check"),
      el("p", { class: "small" }, "Practice mode: nobody grades this. Compare your work with each line of the rubric, and tick the ones it meets. Your files stay with you."),
      el("fieldset", { style: "border:0;padding:0;margin:0 0 10px" }, el("legend", { class: "sr-only" }, "Rubric self-check"), boxes),
      complete);
    showAward({ arg: id }, box, st, doc, false);
  }

  // =====================================================================
  // PART 12: THE BADGES SCREEN (MY RECORD)
  // =====================================================================

  // ---------- The badges page (badges.html, in a browser tab) ----------
  // The panel's saved progress lives in the panel's own storage, which a
  // browser tab can't read. So we pack the badges and certificates from
  // every course on this device into the part of the page address after
  // "#". Browsers never send that part to a server: the page unpacks it
  // in the tab. Nothing else (steps, settings, sections) goes along.

  /** Every course's badges and certificates on this device, in a compact form (format "otb1"). */
  function badgeHandoff() {
    const courses = {};
    for (const app of Object.values(APPS)) {
      let p = null;
      if (app.id === APP.id) p = progress;
      else {
        try {
          const raw = JSON.parse(localStorage.getItem(app.storeKey) || "null");
          if (raw && raw.format === "office-trainer-progress/1") p = cleanProgress(raw);
        } catch (e) { /* storage blocked or damaged: skip that course */ }
      }
      if (!p) continue;
      // [activity id, badge number, date earned] and [stop id, certificate number, date issued]
      const b = Object.values(p.badges).map((x) => [x.activity, x.id, x.earnedAt]);
      // (a course certificate earned in a section adds the section's name and the stops it left out)
      const k = Object.entries(p.certificates).map(([sid, c]) => c.section ? [sid, c.id, c.issuedAt, c.section.title, c.section.skip] : [sid, c.id, c.issuedAt]);
      if (b.length || k.length) courses[app.id] = { b: b, k: k };
    }
    return { f: "otb1", n: progress.settings.name || "", d: today(), c: courses };
  }

  /** Bytes to base64url (letters, digits, - and _: safe in an address). */
  function bytesToBase64Url(bytes) {
    let binary = "";
    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  /**
   * Pack an object for the address: "z." + squeezed (deflate) base64url when
   * the browser can squeeze, else "j." + plain base64url. badges.html reads both.
   */
  async function packForAddress(obj) {
    const json = JSON.stringify(obj);
    if (typeof CompressionStream === "function") {
      try {
        const squeezed = new Blob([json]).stream().pipeThrough(new CompressionStream("deflate-raw"));
        return "z." + bytesToBase64Url(new Uint8Array(await new Response(squeezed).arrayBuffer()));
      } catch (e) { /* fall back to plain */ }
    }
    return "j." + TM.toBase64Url(json);
  }

  /** Open badges.html with this device's badges. */
  async function openBadgesPage() {
    // In a browser tab, open the new tab right away (while the click still
    // counts, so pop-up blockers allow it), then fill in the address.
    const early = Host.inOffice ? null : window.open("", "_blank");
    const url = new URL("badges.html#r=" + (await packForAddress(badgeHandoff())), window.location.href).href;
    if (early) {
      early.opener = null; // the badges page can't reach back into this one
      early.location.replace(url);
    } else openPage(url);
  }

  function drawRecord(view) {
    // Counts are for the stops that count (a section can leave some out),
    // but a badge or certificate already earned is always listed.
    const allActs = countedStations().flatMap((s) => s.activities.map((a) => ({ a: a, st: s })));
    const earned = allActs.filter((x) => progress.badges[x.a.id]);
    const certStations = countedStations().filter((s) => s.certificate && (s.activities.length || s.kind === "course"));
    const certs = Course.stations.filter((s) => s.certificate && progress.certificates[s.sid]);

    add(view, el("h1", { style: "font-family:var(--tm-heading);font-size:1.4rem;margin:0 0 4px" }, "Your badges"),
      el("p", { class: "muted small" }, earned.length + " of " + allActs.length + " badges · " +
        certs.filter((s) => s.kind !== "course" && inSection(s)).length + " of " + certStations.filter((s) => s.kind !== "course").length + " lesson certificates" +
        (certs.some((s) => s.kind === "course") ? " · course certificate earned" : "")));

    // --- The full-page view (a browser tab) ---
    add(view, el("section", { class: "card", "aria-labelledby": "page-title" },
      el("h2", { id: "page-title" }, "Your badges page"),
      el("p", { class: "small" }, "See every badge and certificate from all your Office Trainer courses on one page, with room to explore, print or save them."),
      el("button", { type: "button", class: "btn primary", onclick: () => openBadgesPage() }, "Open my badges page"),
      el("p", { class: "small muted", style: "margin:6px 0 0" }, "Opens in your browser. Your badges travel inside the page's address and aren't sent to any server.")));

    // --- Badges, week by week ---
    for (const w of Course.weeks) {
      const acts = w.stations.filter((s) => inSection(s) || s.activities.some((a) => progress.badges[a.id]))
        .flatMap((s) => s.activities.filter((a) => inSection(s) || progress.badges[a.id]).map((a) => ({ a: a, st: s })));
      if (!acts.length) continue;
      add(view, el("h2", { style: "font-size:1.05rem;margin:16px 0 8px" }, "Week " + w.n + ": " + w.title),
        el("ul", { class: "badge-grid" }, acts.map(function (x) {
          const b = progress.badges[x.a.id];
          return el("li", { class: "badge-card", "data-earned": String(!!b) },
            badgeArt({ week: w.n, kind: x.a.badge.kind, code: x.a.id, label: x.a.badge.name + " badge", size: 64 }),
            el("b", {}, x.a.badge.name),
            el("span", {}, b ? "✓ Earned " + longDate(b.earnedAt) : "Not earned yet"),
            b ? el("span", { class: "mono" }, b.id) : el("button", { type: "button", class: "linkish small", onclick: () => go("station", x.st, () => document.getElementById("h-" + x.a.id) || document.getElementById("act-" + x.a.id)) }, "Go to " + x.a.id));
        })));
    }

    // --- Certificates ---
    add(view, el("h2", { style: "font-size:1.05rem;margin:16px 0 8px" }, "Certificates"));
    if (!certs.length) add(view, el("p", { class: "small muted" }, "Finish every badge in a lesson to get its certificate."));
    add(view, el("ul", { class: "cert-list" }, certs.map((s) => el("li", {},
      el("div", {}, el("b", {}, s.title), el("br"), el("span", { class: "small" }, "No. ", el("span", { class: "mono" }, progress.certificates[s.sid].id), " · " + longDate(progress.certificates[s.sid].issuedAt))),
      el("button", { type: "button", class: "btn slim", onclick: () => viewCertificate(s) }, "View"),
      el("button", { type: "button", class: "btn slim quiet", onclick: () => downloadCertificate(s) }, "Save .json")))));

    // --- Keeping your record ---
    const status = el("p", { class: "small", role: "status", "aria-live": "polite" });
    const openInput = el("input", { type: "file", accept: "application/json,.json", class: "sr-only", id: "open-record", tabindex: "-1" });
    openInput.addEventListener("change", async function () {
      const f = openInput.files[0];
      openInput.value = "";
      if (!f) return;
      try {
        if (f.size > 2 * 1024 * 1024) throw new Error("too big");
        const raw = JSON.parse(await f.text());
        if (!raw || raw.format !== "office-trainer-progress/1") throw new Error("wrong format");
        if (raw.app && raw.app !== APP.id && isApp(raw.app)) { status.textContent = "That record is from the " + APPS[raw.app].name + " course. Open it in Office Trainer for " + APPS[raw.app].name + "."; return; }
        askToReplace(cleanProgress(raw));
      } catch (e) { status.textContent = "That file isn't a saved Office Trainer record. Choose a file made with Save my record."; }
    });
    function askToReplace(incoming) {
      const yes = el("button", { type: "button", class: "btn slim danger" }, "Replace my progress");
      const no = el("button", { type: "button", class: "btn slim quiet" }, "Cancel");
      const box = el("div", { class: "confirm" }, el("p", {}, "Replace the progress on this device with the saved record (" + Object.keys(incoming.badges).length + " badges)?"), el("div", { class: "btn-row" }, yes, no));
      fill(status, box);
      yes.focus();
      no.addEventListener("click", () => { status.textContent = ""; });
      yes.addEventListener("click", () => { progress = incoming; saveProgress(); applySettings(); go("record"); announce("Your saved record is back."); });
    }
    const reset = el("button", { type: "button", class: "btn slim danger" }, "Start over");
    reset.addEventListener("click", function () {
      const yes = el("button", { type: "button", class: "btn slim danger" }, "Yes, erase my progress");
      const no = el("button", { type: "button", class: "btn slim quiet" }, "Cancel");
      fill(status, el("div", { class: "confirm" }, el("p", {}, "Erase every badge, certificate and step on this device? Save your record first if you might want it back."), el("div", { class: "btn-row" }, yes, no)));
      yes.focus();
      no.addEventListener("click", () => { status.textContent = ""; reset.focus(); });
      yes.addEventListener("click", () => {
        const settings = progress.settings; // keep fonts and name
        progress = blankProgress();
        progress.settings = settings;
        saveProgress();
        go("record");
        announce("Progress erased.");
      });
    });
    add(view, el("section", { class: "card", "aria-labelledby": "keep-title" },
      el("h2", { id: "keep-title" }, "Keep your record"),
      el("p", { class: "small" }, courseMode() === "practice"
        ? "Your badges live only in this browser. Clearing browsing data, or using a different computer, starts you over. Save your record to keep a copy, and open it on any computer."
        : "Your instructor sees your results. You can also keep your own copy."),
      el("div", { class: "btn-row" },
        el("button", { type: "button", class: "btn slim", onclick: () => downloadJson(Object.assign({ app: APP.id }, progress), "office-trainer-" + APP.id + "-record-" + today() + ".json") }, "Save my record"),
        el("button", { type: "button", class: "btn slim quiet", onclick: () => openInput.click() }, "Open a saved record"),
        reset),
      openInput, status));
  }

  // =====================================================================
  // PART 13: SETTINGS
  // =====================================================================

  /** Apply the saved settings: theme, text size and the two lesson fonts. */
  function applySettings() {
    const s = progress.settings;
    const root = document.documentElement;
    if (s.theme === "light" || s.theme === "dark") root.dataset.theme = s.theme; else delete root.dataset.theme;
    const sizes = { s: "14px", m: "15px", l: "17px", xl: "19px" };
    root.style.setProperty("--text-size", sizes[s.textSize] || sizes.m);
    const allowChange = Course.config.learnerCanChangeFonts !== false && Course.config.learnerCanChangeFonts !== "false";
    const headingName = (allowChange && s.headingFont) || Course.config.headingFont || TM.DEFAULT_FONTS.heading;
    const bodyName = (allowChange && s.bodyFont) || Course.config.bodyFont || TM.DEFAULT_FONTS.body;
    root.style.setProperty("--tm-heading", TM.fontStack(headingName, TM.DEFAULT_FONTS.heading));
    root.style.setProperty("--tm-body", TM.fontStack(bodyName, TM.DEFAULT_FONTS.body));
  }

  function drawSettings(view) {
    const s = progress.settings;
    const allowFonts = Course.config.learnerCanChangeFonts !== false && Course.config.learnerCanChangeFonts !== "false";
    const status = el("p", { class: "small", role: "status", "aria-live": "polite" });
    const saved = () => { saveProgress(); applySettings(); status.textContent = "Saved."; };

    /** A labelled drop-down that saves as soon as it changes. */
    function choice(id, label, help, options, key, disabled) {
      const select = el("select", { id: id, disabled: disabled }, options.map((o) => el("option", { value: o.value, selected: s[key] === o.value }, o.label)));
      select.addEventListener("change", () => { s[key] = select.value; saved(); });
      return el("div", { class: "field" }, el("label", { for: id }, label), select, help ? el("small", {}, help) : null);
    }
    const fonts = Object.keys(TM.FONT_STACKS);
    const detected = { web: APP.name + " for the web", windows: "Windows", mac: "Mac", all: "a web browser (all notes shown)" }[Host.detected] || Host.detected;

    const name = el("input", { type: "text", id: "set-name", maxlength: "60", autocomplete: "name", value: s.name });
    name.addEventListener("change", () => { s.name = name.value.trim().slice(0, 60); saved(); });

    add(view, 
      el("h1", { style: "font-family:var(--tm-heading);font-size:1.4rem;margin:0 0 12px" }, "Settings"),
      drawSectionCard(),
      el("div", { class: "field" }, el("label", { for: "set-name" }, "Name on your certificates"), name,
        el("small", {}, "Stored only on this device. A nickname is fine.")),
      choice("set-heading", "Heading font", allowFonts ? "" : "Your course sets the fonts.",
        [{ value: "", label: "Course default (" + (Course.config.headingFont || TM.DEFAULT_FONTS.heading) + ")" }].concat(fonts.map((f) => ({ value: f, label: f }))), "headingFont", !allowFonts),
      choice("set-body", "Text font", "",
        [{ value: "", label: "Course default (" + (Course.config.bodyFont || TM.DEFAULT_FONTS.body) + ")" }].concat(fonts.map((f) => ({ value: f, label: f }))), "bodyFont", !allowFonts),
      el("div", { class: "sample", "aria-hidden": "true" }, el("h3", {}, "Format Text and Paragraphs"), el("p", {}, "Space After is how professionals separate paragraphs.")),
      choice("set-size", "Text size", "", [{ value: "s", label: "Small" }, { value: "m", label: "Medium (default)" }, { value: "l", label: "Large" }, { value: "xl", label: "Extra large" }], "textSize"),
      choice("set-theme", "Colors", "", [{ value: "system", label: "Match my computer" }, { value: "light", label: "Light" }, { value: "dark", label: "Dark" }], "theme"),
      choice("set-platform", "Show instructions for", "Detected: " + detected + ". This changes which notes you see, not what's checked.",
        [{ value: "auto", label: "Automatic" }, { value: "web", label: APP.name + " for the web" }, { value: "windows", label: "Desktop " + APP.name + " on Windows" }, { value: "mac", label: "Desktop " + APP.name + " on a Mac" }], "platform"),
      status,
      el("section", { class: "card", "aria-labelledby": "about-title" },
        el("h2", { id: "about-title" }, "About"),
        el("p", { class: "small" }, "Office Trainer " + APP_VERSION + " · " + APP.name + " checker " + checker().VERSION + " · lessons format " + TM.VERSION + "."),
        el("p", { class: "small" }, Host.inOffice ? "Running inside " + APP.name + ": the trainer checks the open " + APP.file + "." : "Running in a browser: download each practice file, work on it in " + APP.name + ", then upload it to check."),
        el("p", { class: "small" }, "Practice mode: your progress stays on this device. A section changes which lessons and videos you see; it doesn't send anything to your instructor."),
        el("p", { class: "small" }, "Keyboard: F6 (Ctrl+F6 in " + APP.name + " for the web) moves between the " + APP.file + ", the ribbon and this panel.")));
  }

  /** Settings: the section card (join with a key or code, or leave). */
  function drawSectionCard() {
    const card = el("section", { class: "card", id: "section-card", "aria-labelledby": "section-title", tabindex: "-1" });
    const status = el("div", { class: "small", role: "status", "aria-live": "polite" });
    const input = el("input", { type: "text", id: "section-input", autocomplete: "off", spellcheck: "false", autocapitalize: "off",
      maxlength: "20000", "aria-describedby": "section-help" });
    const join = el("button", { type: "button", class: "btn slim primary" }, "Join");

    async function tryJoin() {
      join.disabled = true;
      status.textContent = "Checking…";
      const result = await resolveSection(input.value);
      join.disabled = false;
      if (!result.section) { status.textContent = result.problems[0]; input.focus(); return; }
      useSection(result);
      go("settings", null, () => $("section-card"));
      announce("You joined " + result.section.title + ".");
    }
    join.addEventListener("click", tryJoin);
    input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); tryJoin(); } });

    const joinField = el("div", { class: "field" },
      el("label", { for: "section-input" }, Section.active ? "Switch to another section" : "Section key or code"),
      el("div", { class: "btn-row" }, input, join),
      el("small", { id: "section-help" }, "Your instructor gives you this. A key is one long line that starts with OTK1. A code is short, such as " +
        (APP.id === "word" ? "WRD" : APP.id === "excel" ? "XL" : "PP") + "-7K3Q."));

    if (!Section.active) {
      fill(card, el("h2", { id: "section-title" }, "Your section"),
        el("p", { class: "small" }, Section.error ? "Your saved section couldn't be loaded: " + Section.error : "You're not in a section, so you see the whole " + APP.name + " course."),
        joinField, status,
        el("p", { class: "small muted" }, "Joining changes which lessons and videos you see. Your progress still stays on this device, and nothing is sent to your instructor."));
      return card;
    }
    const s = Section.active;
    const leave = el("button", { type: "button", class: "btn slim danger" }, "Leave section");
    leave.addEventListener("click", function () {
      const yes = el("button", { type: "button", class: "btn slim danger" }, "Yes, leave it");
      const no = el("button", { type: "button", class: "btn slim quiet" }, "Cancel");
      fill(status, el("div", { class: "confirm" }, el("p", {}, "Leave " + s.title + "? You'll see the whole course again. Your badges stay."), el("div", { class: "btn-row" }, yes, no)));
      yes.focus();
      no.addEventListener("click", () => { status.textContent = ""; leave.focus(); });
      yes.addEventListener("click", () => { useSection(null); go("settings", null, () => $("section-card")); announce("You left the section."); });
    });
    const skipped = Course.stations.filter((st) => !inSection(st));
    fill(card, el("h2", { id: "section-title" }, "Your section"),
      el("p", {}, el("b", {}, s.title), el("br"), el("span", { class: "small" }, s.instructor + (s.term ? " · " + s.term : ""))),
      el("p", { class: "small" }, describeSection(s)),
      skipped.length ? el("details", { class: "small" }, el("summary", {}, "Stops left out (" + skipped.length + ")"),
        el("ul", {}, skipped.map((st) => el("li", {}, "Week " + st.week + ": " + st.title)))) : null,
      Section.problems.length ? el("details", { class: "small" }, el("summary", {}, "Parts of the key that were skipped (" + Section.problems.length + ")"),
        el("ul", {}, Section.problems.map((p) => el("li", {}, p)))) : null,
      el("div", { class: "btn-row", style: "margin:8px 0" }, leave),
      joinField, status);
    return card;
  }

  // =====================================================================
  // PART 14: STARTING UP
  // ---------------------------------------------------------------------
  // Work out where we are (Word, Excel, or a plain browser), load the progress
  // and the course, then open the lesson named in the link
  // (taskpane.html?lesson=W1-L2&activity=D1.03) or the course home.
  // =====================================================================

  let started = false;

  function detectPlatform() {
    try {
      const p = Office.context.diagnostics && Office.context.diagnostics.platform;
      if (p === "PC" || p === Office.PlatformType.PC) return "windows";
      if (p === "Mac" || p === Office.PlatformType.Mac) return "mac";
    } catch (e) { /* older office.js: assume the web */ }
    return "web";
  }

  /** In a browser tab: which course? ?app=excel in the address, else the last one picked, else Word. */
  function browserApp() {
    const asked = new URLSearchParams(window.location.search).get("app");
    if (isApp(asked)) {
      try { localStorage.setItem(APP_CHOICE_KEY, asked); } catch (e) { /* not important */ }
      return asked;
    }
    try { const last = localStorage.getItem(APP_CHOICE_KEY); if (isApp(last)) return last; } catch (e) { /* blocked */ }
    return "word";
  }

  /**
   * hostApp is "word", "excel" or "powerpoint" inside Office, or null in a plain browser tab.
   */
  async function start(hostApp) {
    if (started) {
      // We started as a browser page, and Office answered late: start again
      // inside the app (a different course needs a fresh page).
      if (hostApp && !Host.inOffice) {
        if (hostApp !== APP.id) { window.location.replace(window.location.pathname); return; }
        Host.inOffice = true;
        Host.detected = detectPlatform();
        go(current.screen, current.station);
      }
      return;
    }
    started = true;
    chooseApp(hostApp || browserApp());
    Host.inOffice = !!hostApp;
    Host.detected = hostApp ? detectPlatform() : "all";
    loadProgress();
    applySettings();
    $("brand-course").textContent = APP.fullName;
    document.title = "Office Trainer: " + APP.name;
    $("version-label").textContent = "Version " + APP_VERSION + (hostApp ? "" : " · browser (upload) mode");

    try {
      await loadCourse();
    } catch (e) {
      fill($("view"), el("div", { class: "notice warn" },
        el("p", {}, el("b", {}, "The course couldn't load.")),
        el("p", {}, "Check your internet connection, then close this panel and open it again."),
        el("p", { class: "small" }, e.message)));
      return;
    }
    applySettings(); // the course may set default fonts
    await loadSavedSection(); // the learner's section (PART 5b), before anything counts badges
    // Catch up: a learner who finished everything before the last lesson was
    // published gets the course certificate as soon as the course is complete.
    if (Course.stations.some((s) => s.kind === "course")) {
      const before = Object.keys(progress.certificates).length;
      maybeCourseCertificate();
      if (Object.keys(progress.certificates).length !== before) saveProgress();
    }
    $("brand-course").textContent = Course.config.title || APP.fullName;

    const params = new URLSearchParams(window.location.search);
    // A link with ?join=<key or code> offers a section. We never join without asking.
    const joinParam = params.get("join");
    if (joinParam && joinParam.replace(/\s+/g, "") !== progress.sectionJoin) {
      Section.pending = await resolveSection(joinParam.slice(0, 20000));
      try { const url = new URL(window.location.href); url.searchParams.delete("join"); window.history.replaceState(null, "", url); } catch (e) { /* not important */ }
    }
    const wanted = params.get("lesson");
    const station = wanted && Course.stations.find((s) => s.docId === wanted || s.sid === wanted);
    const activity = params.get("activity");
    if (Section.pending) {
      go("home", null, () => $("join-offer"));
    } else if (station && Course.docs[station.file]) {
      go("station", station, activity ? () => document.getElementById("h-" + activity) || document.getElementById("act-" + activity) : null);
    } else {
      go("home");
    }
  }

  // The top bar.
  $("nav-brand").addEventListener("click", () => go("home"));
  for (const b of document.querySelectorAll("[data-nav]")) b.addEventListener("click", () => go(b.dataset.nav));

  // Ask office.js where we are. If it didn't load, or doesn't answer in 3 seconds, assume a browser tab.
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      const host = info && info.host;
      start(host === Office.HostType.Excel ? "excel" : host === Office.HostType.Word ? "word" : host === Office.HostType.PowerPoint ? "powerpoint" : null);
    });
    setTimeout(function () { start(null); }, 3000);
  } else {
    start(null);
  }
})();
