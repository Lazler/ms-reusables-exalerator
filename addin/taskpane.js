/* Deck Components – task pane (Office.js / PowerPoint) */
/* global Office, PowerPoint, DCCore, DCI18n */
"use strict";

const C = DCCore;
const I = DCI18n;
const t = I.t;
const VERSION = "0.5.3";

const STATE_KEY = "deckComponents.v1";
const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder"];
const COMP_TYPES = ["GeometricShape", "TextBox"];
const TAG = { TPL: "DC_TPL", OUT: "DC_OUT", FILL: "DC_FILL", LINE: "DC_LINE", FONT: "DC_FONT", OVR: "DC_OVR" };

/** Persisted inside the .pptx (document settings). */
let state = { vars: [], components: {}, showRaw: false };
/** Runtime only. */
const ui = {
  compInfo: {},
  selection: null,
  lastSelection: null,
  autoSync: false,
  addingVar: false, // inline "new variable" row open
  inspectorForm: null, // null | "create" | "variant"
  inspectorFor: null, // selection signature the inline form belongs to
  confirmDelete: null,
};

// ================================================================ helpers

const $ = (id) => document.getElementById(id);

function el(tag, attrs, ...children) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === "class") n.className = v;
    else if (k === "html") n.innerHTML = v; // only used with static icon SVG
    else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
    else if (k in n && k !== "list") n[k] = v;
    else n.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat()) if (c != null && c !== false) n.append(c.nodeType ? c : String(c));
  return n;
}

/** Renders text and shows {{…}} and ◆ … ◆ as inline code. */
function rich(text) {
  const frag = document.createDocumentFragment();
  String(text)
    .split(/(\{\{[^{}]+\}\}|◆[^◆]+◆)/u)
    .forEach((part, i) => frag.append(i % 2 ? el("code", {}, part) : part));
  return frag;
}

// Small inline icon set (16×16, stroke = currentColor)
const ICONS = {
  plus: '<path d="M8 3v10M3 8h10"/>',
  more: '<circle cx="3.5" cy="8" r="1.1" fill="currentColor" stroke="none"/><circle cx="8" cy="8" r="1.1" fill="currentColor" stroke="none"/><circle cx="12.5" cy="8" r="1.1" fill="currentColor" stroke="none"/>',
  gear: '<circle cx="8" cy="8" r="2.2"/><path d="M8 1.8v1.6M8 12.6v1.6M1.8 8h1.6M12.6 8h1.6M3.6 3.6l1.1 1.1M11.3 11.3l1.1 1.1M3.6 12.4l1.1-1.1M11.3 4.7l1.1-1.1"/>',
  insert: '<path d="M3 4h7M3 8h4M3 12h7"/><path d="M10.5 6.5 13 8l-2.5 1.5"/>',
  paint: '<path d="M3 10.5 8.5 5l2.5 2.5L5.5 13H3z"/><path d="M9.5 4l1-1 2.5 2.5-1 1"/>',
  sync: '<path d="M12.5 6A4.7 4.7 0 0 0 4 5.2M3.5 10A4.7 4.7 0 0 0 12 10.8"/><path d="M4 2.5v2.8h2.8M12 13.5v-2.8H9.2"/>',
  diamond: '<path d="M8 1.8 14.2 8 8 14.2 1.8 8z"/>',
  diamondOpen: '<path d="M8 2.5 13.5 8 8 13.5 2.5 8z"/><path d="M8 5.5 10.5 8 8 10.5 5.5 8z"/>',
  text: '<path d="M3.5 4h9M8 4v8.5"/>',
  close: '<path d="M4 4l8 8M12 4l-8 8"/>',
  target: '<circle cx="8" cy="8" r="5"/><circle cx="8" cy="8" r="1.5"/>',
  cursor: '<path d="M4 2.5 12 8l-3.6.8L10.5 13l-1.7.8-2.1-4.3L4 12z"/>',
  check: '<path d="M3.5 8.5 6.5 11.5 12.5 4.5"/>',
};
function icon(name) {
  return el("span", {
    class: "ico",
    "aria-hidden": "true",
    html: '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">' + ICONS[name] + "</svg>",
  });
}

function iconButton(name, title, onclick, extra) {
  return el("button", Object.assign({ type: "button", class: "icon-btn", title, "aria-label": title, onclick }, extra || {}), icon(name));
}

// ---------------------------------------------------------------- popover menu

let openMenu = null;
function closeMenu() {
  if (openMenu) openMenu.remove();
  openMenu = null;
}
document.addEventListener("click", (e) => {
  if (openMenu && !openMenu.contains(e.target)) closeMenu();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeMenu();
});

/**
 * A "⋯" (or other icon) button that opens a small menu.
 * items: { label, onclick, danger, checked, keepOpen } | "sep" | { heading }
 */
function menuButton(items, opts) {
  opts = opts || {};
  const btn = iconButton(opts.icon || "more", opts.title || t("ui.more"), (e) => {
    e.stopPropagation();
    if (openMenu && openMenu._owner === btn) return closeMenu();
    closeMenu();
    const m = el("div", { class: "menu", role: "menu" });
    m._owner = btn;
    for (const it of typeof items === "function" ? items() : items) {
      if (!it) continue;
      if (it === "sep") m.append(el("div", { class: "menu-sep" }));
      else if (it.heading) m.append(el("div", { class: "menu-heading" }, it.heading));
      else {
        m.append(
          el(
            "button",
            {
              type: "button",
              role: it.checked === undefined ? "menuitem" : "menuitemcheckbox",
              "aria-checked": it.checked === undefined ? null : String(!!it.checked),
              class: "menu-item" + (it.danger ? " danger" : ""),
              disabled: it.disabled,
              onclick: (ev) => {
                ev.stopPropagation();
                if (!it.keepOpen) closeMenu();
                it.onclick();
              },
            },
            el("span", { class: "menu-check" }, it.checked ? icon("check") : null),
            el("span", {}, it.label)
          )
        );
      }
    }
    document.body.append(m);
    const r = btn.getBoundingClientRect();
    const w = m.offsetWidth;
    m.style.top = Math.min(r.bottom + 4, window.innerHeight - m.offsetHeight - 8) + "px";
    m.style.left = Math.max(8, Math.min(r.right - w, window.innerWidth - w - 8)) + "px";
    openMenu = m;
  });
  btn.classList.add("menu-btn");
  return btn;
}

// ---------------------------------------------------------------- toast / status

let toastTimer = null;
/** kind: undefined (info) | "ok" | "warn" | "err". Info/ok hide by themselves. */
function setStatus(msg, kind) {
  const f = $("status");
  clearTimeout(toastTimer);
  $("status-text").textContent = msg || "";
  f.className = "toast " + (kind || "info") + (msg ? " show" : "");
  if (!kind || kind === "ok") toastTimer = setTimeout(() => f.classList.remove("show"), 3500);
}

function errText(e) {
  if (!e) return t("err.unknown");
  let s = e.message || String(e);
  if (e.code) s += " [" + e.code + "]";
  if (e.debugInfo && e.debugInfo.errorLocation) s += "\n" + t("err.location") + ": " + e.debugInfo.errorLocation;
  return s;
}

function debounce(fn, ms) {
  let tm;
  return (...a) => {
    clearTimeout(tm);
    tm = setTimeout(() => fn(...a), ms);
  };
}

// ---------------------------------------------------------------- diagnostics log

const LOG = [];
/** Keeps the last 100 events for the Diagnostics view. */
function log(kind, msg) {
  const d = new Date();
  const time = d.toTimeString().slice(0, 8) + "." + String(d.getMilliseconds()).padStart(3, "0");
  LOG.push(time + "  " + kind.padEnd(5) + " " + msg);
  if (LOG.length > 100) LOG.shift();
  if (kind === "ERR") console.warn(msg); // shown in Diagnostics; not a script error
}

const QUEUE_TIMEOUT_MS = 20000;
/**
 * Runs all PowerPoint calls one after another and reports errors.
 * A call that never returns is abandoned after 20 s so later updates still run.
 */
let chain = Promise.resolve();
function queue(label, fn) {
  const run = () => {
    const started = Date.now();
    let timer;
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(t("err.timeout"))), QUEUE_TIMEOUT_MS);
    });
    return Promise.race([Promise.resolve().then(fn), timeout])
      .then((v) => {
        log("ok", label + " (" + (Date.now() - started) + " ms)");
        return v;
      })
      .finally(() => clearTimeout(timer));
  };
  const p = chain.then(run).catch((e) => {
    log("ERR", label + ": " + errText(e).replace(/\n/g, " | "));
    setStatus(t("err.during", { action: t(label) }) + ": " + errText(e), "err");
  });
  chain = p;
  return p;
}

function saveState() {
  const s = Office.context.document.settings;
  s.set(STATE_KEY, JSON.stringify(state));
  updateCounts();
  return new Promise((resolve) =>
    s.saveAsync((r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus(t("err.save", { error: r.error && r.error.message }), "err");
      }
      resolve();
    })
  );
}

function loadState() {
  try {
    const raw = Office.context.document.settings.get(STATE_KEY);
    if (raw) {
      const s = typeof raw === "string" ? JSON.parse(raw) : raw;
      state = { vars: s.vars || [], components: s.components || {}, showRaw: !!s.showRaw };
    }
  } catch (e) {
    setStatus(t("err.state", { error: errText(e) }), "err");
  }
}

function lsGet(k) {
  try {
    return localStorage.getItem(k);
  } catch (_) {
    return null;
  }
}
function lsSet(k, v) {
  try {
    localStorage.setItem(k, v);
  } catch (_) {
    /* storage unavailable – ignore */
  }
}

// ================================================================ PowerPoint access

/** Loads all shapes on the given slides (default: all), group contents excluded. */
async function loadShapes(ctx, slideIds) {
  if (slideIds && slideIds.length) {
    try {
      const picked = slideIds.map((id) => {
        const sl = ctx.presentation.slides.getItemOrNullObject(id);
        sl.load("id");
        sl.shapes.load("items/id,items/name,items/type");
        return sl;
      });
      await ctx.sync();
      const out = [];
      picked.forEach((sl) => {
        if (sl.isNullObject) return;
        sl.shapes.items.forEach((sh) => out.push({ slideId: sl.id, slideIndex: null, shape: sh, id: sh.id, name: sh.name, type: sh.type }));
      });
      return out;
    } catch (e) {
      /* fall back to the full slide list */
    }
  }
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const wanted = slides.items
    .map((sl, i) => ({ sl, i }))
    .filter(({ sl }) => !slideIds || slideIds.includes(sl.id));
  wanted.forEach(({ sl }) => sl.shapes.load("items/id,items/name,items/type"));
  await ctx.sync();
  const out = [];
  wanted.forEach(({ sl, i }) =>
    sl.shapes.items.forEach((sh) =>
      out.push({ slideId: sl.id, slideIndex: i, shape: sh, id: sh.id, name: sh.name, type: sh.type })
    )
  );
  return out;
}

/**
 * Shapes whose text frame cannot be read (pictures, empty picture placeholders …).
 * Remembered so later batches skip their text instead of failing and falling back
 * to slow one-by-one loading on every click.
 */
const noText = new Set();
const shapeKey = (s) => s.slideId + ":" + s.id;
const wantsText = (s) => TEXT_TYPES.includes(s.type) && !noText.has(shapeKey(s));

function loadTag(shape, key) {
  const tg = shape.tags.getItemOrNullObject(key);
  tg.load("value");
  return tg;
}
const tagVal = (tg) => (tg && !tg.isNullObject ? tg.value : null);

/**
 * Loads data for many items in one batch. If that fails (e.g. an exotic shape has
 * no text frame), it retries item by item and skips the failing ones.
 */
async function loadEach(ctx, items, loader) {
  if (!items.length) return items;
  items.forEach(loader);
  try {
    await ctx.sync();
    return items;
  } catch (batchError) {
    log("warn", "batch load failed (" + errText(batchError).split("\n")[0] + "), loading one by one");
    const ok = [];
    for (const it of items) {
      try {
        loader(it);
        await ctx.sync();
        ok.push(it);
      } catch (e) {
        // Retry without the text frame; if that works, remember the shape as text-less
        const key = it.slideId != null ? shapeKey(it) : null;
        if (key && TEXT_TYPES.includes(it.type) && !noText.has(key)) {
          noText.add(key);
          try {
            loader(it);
            await ctx.sync();
            ok.push(it);
            log("info", "“" + it.name + "” has no readable text – treated as a plain shape from now on");
            continue;
          } catch (_) {
            noText.delete(key);
          }
        }
        log("warn", "skipped “" + it.name + "” (" + it.type + "): " + errText(e).split("\n")[0]);
      }
    }
    return ok;
  }
}

/**
 * Applies queued writes for many items in one sync. If that fails, it retries
 * item by item so a single problematic shape does not block all others.
 * @param writers  array of { name, write: () => void }
 * @returns number of items that could not be written
 */
async function writeEach(ctx, writers, what) {
  if (!writers.length) return 0;
  writers.forEach((w) => w.write());
  try {
    await ctx.sync();
    return 0;
  } catch (batchError) {
    log("warn", what + ": batch write failed (" + errText(batchError).split("\n")[0] + "), writing one by one");
    let failed = 0;
    for (const w of writers) {
      try {
        w.write();
        await ctx.sync();
      } catch (e) {
        failed++;
        log("ERR", what + ": could not write “" + w.name + "”: " + errText(e).split("\n")[0]);
      }
    }
    return failed;
  }
}

/**
 * Reads the current selection. PowerPoint on the web throws a GeneralException in
 * some states (focus in the thumbnail pane, text in edit mode, groups/placeholders),
 * so this falls back in stages instead of failing.
 */
async function getSelection(ctx) {
  // Fast path: slide + shapes in a single round trip
  try {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    const sid = slides.items.length ? slides.items[0].id : null;
    return { slideId: sid, shapes: shapes.items.map((sh) => ({ shape: sh, id: sh.id, name: sh.name, type: sh.type, slideId: sid })) };
  } catch (e) {
    /* fall through to the staged, more tolerant path */
  }
  let slideId = null;
  try {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    await ctx.sync();
    slideId = slides.items.length ? slides.items[0].id : null;
  } catch (e) {
    console.warn("Slide selection not readable", e);
  }
  const wrap = (sh, type) => ({ shape: sh, id: sh.id, name: sh.name, type, slideId });

  // Stage 1: everything in one batch
  try {
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/id,items/name,items/type");
    await ctx.sync();
    return { slideId, shapes: shapes.items.map((sh) => wrap(sh, sh.type)) };
  } catch (e) {
    console.warn("Selection stage 1 failed", e);
  }

  // Stage 2: shape by shape (first all at once by index, then one by one)
  const list = [];
  try {
    const shapes = ctx.presentation.getSelectedShapes();
    const count = shapes.getCount();
    await ctx.sync();
    try {
      const byIndex = [];
      for (let i = 0; i < count.value; i++) {
        const sh = shapes.getItemAt(i);
        sh.load("id,name,type");
        byIndex.push(sh);
      }
      await ctx.sync();
      return { slideId, shapes: byIndex.map((sh) => wrap(sh, sh.type)), partial: true };
    } catch (_) {
      /* one of them is odd – read them one by one */
    }
    for (let i = 0; i < count.value; i++) {
      try {
        const sh = shapes.getItemAt(i);
        sh.load("id,name");
        await ctx.sync();
        let type = "Unsupported";
        try {
          sh.load("type");
          await ctx.sync();
          type = sh.type;
        } catch (_) {
          /* type unknown */
        }
        list.push(wrap(sh, type));
      } catch (e) {
        console.warn("Selected shape " + i + " not readable", e);
      }
    }
  } catch (e) {
    console.warn("Selection stage 2 failed", e);
  }
  return { slideId, shapes: list, partial: true };
}

function shapeById(ctx, s) {
  return ctx.presentation.slides.getItem(s.slideId).shapes.getItem(s.id);
}

// ================================================================ variables

const textVars = () => state.vars.filter((v) => v.type !== "color");
const colorOf = (name) => {
  const v = state.vars.find((x) => x.name === name && x.type === "color");
  return v ? C.normalizeColor(v.value) : null;
};

/** Default values set on a component's master. */
function compDefaults(shapeName) {
  const key = C.componentNameOf(shapeName);
  const comp = key && state.components[key];
  return (comp && comp.snapshot && comp.snapshot.defaults) || {};
}

/** Placeholders in a component's master text – they belong to the component, not to the deck. */
function componentPlaceholders(shapeName) {
  const key = C.componentNameOf(shapeName);
  const comp = key && state.components[key];
  const src = comp && comp.snapshot && comp.snapshot.text && comp.snapshot.text.source;
  return src ? C.extractNames(src) : [];
}

/** Friendly name for a PowerPoint shape type. */
function typeLabel(type) {
  const k = "shape." + type;
  const s = t(k);
  return s === k ? type : s;
}

/** Global → master default → the shape's own override. */
function effectiveVars(shapeName, ovr) {
  return C.mergeVars(textVars(), compDefaults(shapeName), ovr.vars);
}

/**
 * Renders text boxes and color bindings.
 * @param opts.slideIds  only these slides (default: whole deck)
 * @param opts.skip      shapes to leave alone, e.g. the one the user is typing in
 */
async function applyVariables(ctx, opts) {
  opts = opts || {};
  const skip = opts.skip || [];
  const all = (opts.shapes || (await loadShapes(ctx, opts.slideIds))).filter(
    (s) => !skip.some((k) => k.slideId === s.slideId && k.id === s.id)
  );
  const items = await loadEach(ctx, all, (it) => {
    it.tags = {};
    it.tr = null;
    for (const k of Object.values(TAG)) it.tags[k] = loadTag(it.shape, k);
    if (wantsText(it)) {
      it.tr = it.shape.textFrame.textRange;
      it.tr.load("text");
    }
  });

  const missing = new Set();
  let changed = 0;
  let fields = 0;
  let detached = 0;
  const writers = [];

  for (const it of items) {
    const ovr = C.parseOverrides(tagVal(it.tags[TAG.OVR]));
    const vars = effectiveVars(it.name, ovr);
    const local = componentPlaceholders(it.name);
    const ops = [];
    if (it.tr) {
      const current = it.tr.text;
      const oldTpl = tagVal(it.tags[TAG.TPL]);
      const oldOut = tagVal(it.tags[TAG.OUT]);
      const tpl = C.resolveTemplate(current, oldTpl, oldOut, vars);
      if (tpl) {
        fields++;
        const r = C.render(tpl, vars);
        r.missing.filter((m) => !local.includes(m)).forEach((m) => missing.add(m));
        const out = state.showRaw ? tpl : r.text;
        if (!C.sameText(out, current)) {
          ops.push(() => (it.tr.text = out));
          changed++;
        }
        if (tpl !== oldTpl) ops.push(() => it.shape.tags.add(TAG.TPL, tpl));
        if (out !== oldOut) ops.push(() => it.shape.tags.add(TAG.OUT, out));
      } else if (oldTpl != null) {
        // The text no longer contains the variable's value: the user replaced it.
        detached++;
        log("info", "“" + it.name + "” was edited by hand – no longer linked to " + C.extractNames(oldTpl).map((n) => "{{" + n + "}}").join(", "));
        ops.push(() => it.shape.tags.delete(TAG.TPL));
        if (oldOut != null) ops.push(() => it.shape.tags.delete(TAG.OUT));
      }
    }
    // Colors: instance override > color variable
    const fill = ovr.fill || colorOf(tagVal(it.tags[TAG.FILL]));
    if (fill) ops.push(() => it.shape.fill.setSolidColor(fill));
    const line = ovr.line || colorOf(tagVal(it.tags[TAG.LINE]));
    if (line) {
      ops.push(() => {
        it.shape.lineFormat.visible = true;
        it.shape.lineFormat.color = line;
      });
    }
    const font = ovr.font || colorOf(tagVal(it.tags[TAG.FONT]));
    if (font && it.tr) ops.push(() => (it.tr.font.color = font));
    if (ops.length) writers.push({ name: it.name, write: () => ops.forEach((op) => op()) });
  }
  const failed = await writeEach(ctx, writers, "update");
  log(
    "info",
    "update: " + all.length + " shape(s) on " + (opts.slideIds ? opts.slideIds.length + " slide(s)" : "all slides") +
      ", " + (all.length - items.length) + " unreadable, " + fields + " with variables, " + changed + " changed" +
      (failed ? ", " + failed + " failed" : "") + (detached ? ", " + detached + " unlinked" : "") +
      (missing.size ? ", missing: " + [...missing].join(", ") : "")
  );
  return { changed, fields, missing: [...missing], failed };
}

function reportApply(r, always) {
  if (r.failed) {
    setStatus(t("msg.applyFailed", { n: r.failed }), "err");
  } else if (r.missing.length) {
    setStatus(t("msg.missing", { names: r.missing.join(", ") }), "warn");
  } else if (always || r.changed) {
    setStatus(t("msg.applied", { fields: r.fields, changed: r.changed }) + (state.showRaw ? " " + t("msg.rawView") : ""), "ok");
  }
}

function runApplyVariables() {
  return queue("act.apply", async () => {
    const r = await PowerPoint.run(applyVariables);
    reportApply(r, true);
  });
}
const applyVariablesSoon = debounce(() => queue("act.apply", async () => reportApply(await PowerPoint.run(applyVariables), false)), 600);

const hasRenderables = () => state.vars.length > 0 || Object.keys(state.components).length > 0;


/**
 * Automatic "push": reads the masters on the given slides (all slides if unknown),
 * compares them with their last snapshot and pushes the ones that changed.
 * Works without knowing which shape the user just left, so it also works when
 * PowerPoint on the web cannot report the selection.
 * @param opts.slideIds  slides whose masters are checked (null = all)
 * @param opts.skip      masters to leave alone because they are still selected
 */
async function autoPushIn(ctx, all, skip) {
  const candidates = Object.keys(state.components).filter((name) => {
    const m = state.components[name].master;
    return !skip.some((k) => k.slideId === m.slideId && k.id === m.shapeId);
  });
  const masters = [];
  for (const name of candidates) {
    const m = state.components[name].master;
    const s = all.find((x) => x.slideId === m.slideId && x.id === m.shapeId);
    if (s) masters.push(Object.assign({ comp: name }, s));
  }
  if (!masters.length) return null;
  // All masters in one round trip
  const loaded = await loadEach(ctx, masters, (it) => (it.h = queueStyle(it.shape, wantsText(it))));
  const dirty = loaded
    .filter((it) => JSON.stringify(buildStyle(it.h)) !== JSON.stringify(state.components[it.comp].snapshot))
    .map((it) => it.comp);
  log("info", "auto-push: checked " + loaded.length + " master(s), changed: " + (dirty.join(", ") || "none"));
  return dirty.length ? syncComponents(ctx, dirty) : null;
}

/**
 * Automatic update after a click: push changed masters, then fill in placeholders.
 * Calls arriving while one is waiting are merged into it, so fast clicking does
 * not pile up work. Everything runs in one PowerPoint.run to save round trips.
 */
let pendingUpdate = null;
function autoUpdate(slideIds, skip) {
  if (!hasRenderables()) return Promise.resolve();
  if (pendingUpdate) {
    // merge: unknown slides (null) win, otherwise union
    pendingUpdate.slideIds = pendingUpdate.slideIds && slideIds ? [...new Set([...pendingUpdate.slideIds, ...slideIds])] : null;
    pendingUpdate.skip = skip;
    return pendingUpdate.promise;
  }
  pendingUpdate = { slideIds, skip };
  pendingUpdate.promise = queue("act.apply", async () => {
    const job = pendingUpdate;
    pendingUpdate = null;
    const res = await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx, job.slideIds);
      const pushed = ui.autoSync && Object.keys(state.components).length ? await autoPushIn(ctx, all, job.skip) : null;
      const rendered = await applyVariables(ctx, { shapes: all, slideIds: job.slideIds, skip: job.skip });
      return { pushed, rendered };
    });
    if (res.pushed) {
      await saveState();
      renderComps();
      if (res.pushed.warnings.length) setStatus(t("msg.synced", { n: res.pushed.synced }) + "\n" + res.pushed.warnings.join("\n"), "warn");
      else if (res.pushed.synced) setStatus(t("msg.synced", { n: res.pushed.synced }), "ok");
    }
    if (res.rendered.changed || res.rendered.failed) reportApply(res.rendered, false);
  });
  return pendingUpdate.promise;
}

function scanPlaceholders() {
  return queue("act.scan", async () => {
    const names = await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      const items = await loadEach(ctx, all.filter((s) => TEXT_TYPES.includes(s.type)), (it) => {
        it.tpl = loadTag(it.shape, TAG.TPL);
        it.ovr = loadTag(it.shape, TAG.OVR);
        it.tr = it.shape.textFrame.textRange;
        it.tr.load("text");
      });
      const found = new Set();
      for (const it of items) {
        const local = { ...compDefaults(it.name), ...C.parseOverrides(tagVal(it.ovr)).vars };
        const add = (n) => {
          if (!(n in local)) found.add(n);
        };
        C.extractNames(tagVal(it.tpl) || "").forEach(add);
        C.extractNames(it.tr.text).forEach(add);
      }
      return [...found];
    });
    const known = new Set(state.vars.map((v) => v.name));
    const fresh = names.filter((n) => !known.has(n));
    fresh.forEach((n) => state.vars.push({ name: n, type: "text", value: "" }));
    if (fresh.length) {
      await saveState();
      renderVars();
      setStatus(t("msg.scanNew", { names: fresh.join(", ") }), "ok");
    } else {
      setStatus(names.length ? t("msg.scanAll", { n: names.length }) : t("msg.scanNone"), "ok");
    }
  });
}

function insertPlaceholder(name) {
  return queue("act.insert", async () => {
    await PowerPoint.run(async (ctx) => {
      const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
      await ctx.sync();
      if (tr.isNullObject) throw new Error(t("err.noCursor"));
      tr.text = "{{" + name + "}}";
      await ctx.sync();
    });
    setStatus(t("msg.inserted", { name }), "ok");
  });
}

function bindColor(name, kind) {
  const key = { fill: TAG.FILL, line: TAG.LINE, font: TAG.FONT }[kind];
  return queue("act.bind", async () => {
    const n = await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      if (!sel.shapes.length) throw new Error(t("err.noSelection"));
      sel.shapes.forEach((s) => s.shape.tags.add(key, name));
      await ctx.sync();
      return sel.shapes.length;
    });
    setStatus(t("msg.bound", { n, name }), "ok");
  }).then(() => runApplyVariables());
}

function unbindColors() {
  return queue("act.unbind", async () => {
    const n = await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      const items = await loadEach(ctx, sel.shapes, (it) => {
        it.t = [TAG.FILL, TAG.LINE, TAG.FONT].map((k) => [k, loadTag(it.shape, k)]);
      });
      let c = 0;
      items.forEach((it) =>
        it.t.forEach(([k, tg]) => {
          if (!tg.isNullObject) {
            it.shape.tags.delete(k);
            c++;
          }
        })
      );
      await ctx.sync();
      return c;
    });
    setStatus(n ? t("msg.unbound", { n }) : t("msg.noBindings"), "ok");
  });
}

// ================================================================ components

/** Queues the loads needed to snapshot a master's style. Call buildStyle after ctx.sync(). */
function queueStyle(shape, withText) {
  shape.load("width,height");
  shape.fill.load("type,foregroundColor,transparency");
  shape.lineFormat.load("visible,color,weight,dashStyle");
  const h = { shape, text: null };
  if (withText) {
    const tf = shape.textFrame;
    const tr = tf.textRange;
    tr.load("text");
    tr.font.load("name,size,bold,italic,underline,color");
    tr.paragraphFormat.load("horizontalAlignment");
    tf.load("verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,autoSizeSetting");
    h.text = { tf, tr, tpl: loadTag(shape, TAG.TPL), out: loadTag(shape, TAG.OUT), ovr: loadTag(shape, TAG.OVR) };
  }
  return h;
}

function buildStyle(h) {
  const shape = h.shape;
  const st = {
    width: shape.width,
    height: shape.height,
    fill: { type: shape.fill.type, color: shape.fill.foregroundColor, transparency: shape.fill.transparency },
    line: {
      visible: shape.lineFormat.visible,
      color: shape.lineFormat.color,
      weight: shape.lineFormat.weight,
      dashStyle: shape.lineFormat.dashStyle,
    },
    text: null,
  };
  if (!h.text) return st;
  const { tf, tr, tpl, out, ovr } = h.text;
  st.defaults = C.parseOverrides(tagVal(ovr)).vars;
  const vars = C.mergeVars(textVars(), st.defaults);
  const source = C.resolveTemplate(tr.text, tagVal(tpl), tagVal(out), vars) || tr.text;
  st.text = {
    source,
    font: {
      name: tr.font.name,
      size: tr.font.size,
      bold: tr.font.bold,
      italic: tr.font.italic,
      underline: tr.font.underline,
      color: tr.font.color,
    },
    hAlign: tr.paragraphFormat.horizontalAlignment,
    vAlign: tf.verticalAlignment,
    margins: [tf.leftMargin, tf.rightMargin, tf.topMargin, tf.bottomMargin],
    wordWrap: tf.wordWrap,
    autoSize: tf.autoSizeSetting,
  };
  return st;
}

/** Snapshot of one shape's style (one round trip; text is optional if unreadable). */
async function readStyle(ctx, shape, type) {
  if (TEXT_TYPES.includes(type)) {
    try {
      const h = queueStyle(shape, true);
      await ctx.sync();
      return buildStyle(h);
    } catch (e) {
      log("warn", "text format not readable: " + errText(e).split("\n")[0]);
    }
  }
  const h = queueStyle(shape, false);
  await ctx.sync();
  return buildStyle(h);
}

const isSet = (v) => v !== null && v !== undefined;

function applyStyle(inst, st, comp) {
  const sh = inst.shape;
  const ovr = C.parseOverrides(tagVal(inst.ovr));
  const vars = C.mergeVars(textVars(), st.defaults, ovr.vars);
  if (st.fill) {
    if (st.fill.type === "NoFill") sh.fill.clear();
    else if (st.fill.type === "Solid" && st.fill.color) {
      sh.fill.setSolidColor(st.fill.color);
      if (isSet(st.fill.transparency)) sh.fill.transparency = st.fill.transparency;
    }
  }
  if (st.line) {
    if (st.line.visible === false) sh.lineFormat.visible = false;
    else {
      sh.lineFormat.visible = true;
      if (st.line.color) sh.lineFormat.color = st.line.color;
      if (isSet(st.line.weight)) sh.lineFormat.weight = st.line.weight;
      if (st.line.dashStyle) sh.lineFormat.dashStyle = st.line.dashStyle;
    }
  }
  const tx = st.text;
  if (tx && inst.tr) {
    const tf = sh.textFrame;
    if (tx.autoSize) tf.autoSizeSetting = tx.autoSize;
    if (comp.syncText !== false) {
      if (C.hasPlaceholders(tx.source)) {
        const out = state.showRaw ? tx.source : C.render(tx.source, vars).text;
        inst.tr.text = out;
        sh.tags.add(TAG.TPL, tx.source);
        sh.tags.add(TAG.OUT, out);
      } else {
        inst.tr.text = tx.source;
        if (!inst.tpl.isNullObject) sh.tags.delete(TAG.TPL);
        if (!inst.out.isNullObject) sh.tags.delete(TAG.OUT);
      }
    }
    const f = tx.font || {};
    for (const k of ["name", "size", "bold", "italic", "underline", "color"]) {
      if (isSet(f[k]) && f[k] !== "") inst.tr.font[k] = f[k];
    }
    if (tx.hAlign) inst.tr.paragraphFormat.horizontalAlignment = tx.hAlign;
    if (tx.vAlign) tf.verticalAlignment = tx.vAlign;
    if (tx.margins) {
      const [l, r, tp, b] = tx.margins;
      if (isSet(l)) tf.leftMargin = l;
      if (isSet(r)) tf.rightMargin = r;
      if (isSet(tp)) tf.topMargin = tp;
      if (isSet(b)) tf.bottomMargin = b;
    }
    if (isSet(tx.wordWrap)) tf.wordWrap = tx.wordWrap;
  }
  // The instance's own colors win over the master
  if (ovr.fill) sh.fill.setSolidColor(ovr.fill);
  if (ovr.line) {
    sh.lineFormat.visible = true;
    sh.lineFormat.color = ovr.line;
  }
  if (ovr.font && inst.tr) inst.tr.font.color = ovr.font;
  // Size last, so auto-size does not immediately override it
  if (comp.syncSize !== false) {
    if (isSet(st.width)) sh.width = st.width;
    if (isSet(st.height)) sh.height = st.height;
  }
}

function infoFor(all, name, comp) {
  const members = all.filter((s) => C.componentNameOf(s.name) === name);
  const master = members.find((s) => s.slideId === comp.master.slideId && s.id === comp.master.shapeId);
  return {
    members,
    master,
    instances: members.filter((s) => s !== master),
    info: { instances: members.length - (master ? 1 : 0), masterFound: !!master, masterSlide: master ? master.slideIndex + 1 : null },
  };
}

/** Syncs the given (or all) components. Returns statistics. */
async function syncComponents(ctx, names) {
  const all = await loadShapes(ctx);
  const result = { synced: 0, warnings: [] };
  ui.compInfo = {};
  for (const name of Object.keys(state.components)) {
    const comp = state.components[name];
    const { master, instances, info } = infoFor(all, name, comp);
    ui.compInfo[name] = info;
    if (names && !names.includes(name)) continue;

    let st;
    if (master) {
      st = await readStyle(ctx, master.shape, master.type);
      comp.snapshot = st;
    } else if (comp.snapshot) {
      st = comp.snapshot;
      result.warnings.push(t("warn.masterSnapshot", { name: label(name) }));
    } else {
      result.warnings.push(t("warn.masterMissing", { name: label(name) }));
      continue;
    }

    const targets = instances.filter((s) => COMP_TYPES.includes(s.type) || s.type === "Placeholder");
    const skipped = instances.length - targets.length;
    if (skipped) result.warnings.push(t("warn.skipped", { name: label(name), n: skipped }));

    const loaded = await loadEach(ctx, targets, (it) => {
      it.tpl = loadTag(it.shape, TAG.TPL);
      it.out = loadTag(it.shape, TAG.OUT);
      it.ovr = loadTag(it.shape, TAG.OVR);
      it.tr = st.text ? it.shape.textFrame.textRange : null;
    });

    const failed = await writeEach(
      ctx,
      loaded.map((it) => ({ name: it.name + " (slide " + (it.slideIndex + 1) + ")", write: () => applyStyle(it, st, comp) })),
      "push " + name
    );
    result.synced += loaded.length - failed;
    if (failed) result.warnings.push(t("warn.failed", { name: label(name), n: failed }));
    log("info", "push “" + name + "”: master " + (master ? "slide " + (master.slideIndex + 1) : "missing") + ", " + loaded.length + " instance(s), " + failed + " failed");
  }
  return result;
}

function runSync(names, quiet) {
  return queue("act.sync", async () => {
    const r = await PowerPoint.run((ctx) => syncComponents(ctx, names));
    await saveState();
    renderComps();
    if (r.warnings.length) setStatus(t("msg.synced", { n: r.synced }) + "\n" + r.warnings.join("\n"), "warn");
    else if (!quiet) setStatus(t("msg.synced", { n: r.synced }), "ok");
  });
}

function refreshCompInfo() {
  return queue("act.count", async () => {
    await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      ui.compInfo = {};
      for (const [name, comp] of Object.entries(state.components)) ui.compInfo[name] = infoFor(all, name, comp).info;
    });
    renderComps();
  });
}

/** Display name: "Button" or "Button · Secondary". */
const label = (key) => {
  const k = C.splitKey(key);
  return k.variant === C.DEFAULT_VARIANT ? k.base : k.base + " · " + k.variant;
};
const variantLabel = (key) => {
  const v = C.splitKey(key).variant;
  return v === C.DEFAULT_VARIANT ? t("variant.default") : v;
};
const variantsOf = (base) => Object.keys(state.components).filter((k) => C.splitKey(k).base === base).sort();

/** key = "Button" (new component) or "Button / Secondary" (variant). */
function createComponent(key) {
  const parts = C.splitKey(key || "");
  if (!C.isValidName(parts.base) || !C.isValidName(parts.variant)) return setStatus(t("err.invalidName"), "err");
  const name = C.joinKey(parts.base, parts.variant);
  if (state.components[name]) return setStatus(t("err.exists", { name: label(name) }), "err");
  return queue("act.create", async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      if (sel.shapes.length !== 1) throw new Error(t("err.exactlyOne"));
      const s = sel.shapes[0];
      if (!COMP_TYPES.includes(s.type)) throw new Error(t("err.typeUnsupported", { type: typeLabel(s.type) }));
      const existing = C.componentNameOf(s.name);
      if (existing && state.components[existing] && isMasterOf(s)) {
        throw new Error(t("err.alreadyMaster", { name: label(existing) }));
      }
      s.shape.name = C.componentShapeName(name);
      await ctx.sync();
      const snapshot = await readStyle(ctx, s.shape, s.type);
      state.components[name] = { master: { slideId: sel.slideId, shapeId: s.id }, snapshot, syncText: true, syncSize: true };
      ui.compInfo[name] = { instances: 0, masterFound: true, masterSlide: null };
    });
    ui.inspectorForm = null;
    await saveState();
    renderComps();
    setStatus(parts.variant === C.DEFAULT_VARIANT ? t("msg.compCreated", { name }) : t("msg.variantCreated", { name: label(name) }), "ok");
  }).then(() => {
    refreshCompInfo();
    refreshSelection();
  });
}

function linkSelection(name) {
  return queue("act.link", async () => {
    const n = await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      const comp = state.components[name];
      const targets = sel.shapes.filter((s) => !(s.slideId === comp.master.slideId && s.id === comp.master.shapeId));
      if (!targets.length) throw new Error(t("err.selectTargets"));
      targets.forEach((s) => (s.shape.name = C.componentShapeName(name)));
      await ctx.sync();
      return targets.length;
    });
    setStatus(t("msg.linked", { n, name: label(name) }), "ok");
  }).then(() => runSync([name], true).then(refreshSelection));
}

function detachSelection() {
  return queue("act.detach", async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      sel.shapes.forEach((s) => {
        if (isMasterOf(s)) throw new Error(t("err.detachMaster"));
      });
      sel.shapes.forEach((s) => (s.shape.name = C.stripComponentName(s.name)));
      await ctx.sync();
    });
    setStatus(t("msg.detached"), "ok");
  }).then(() => {
    refreshCompInfo();
    refreshSelection();
  });
}

/** Changes a shape's overrides (tag DC_OVR). */
function setOverride(s, key, mutate) {
  return queue("act.override", async () => {
    await PowerPoint.run(async (ctx) => {
      const sh = shapeById(ctx, s);
      const tg = loadTag(sh, TAG.OVR);
      await ctx.sync();
      const o = C.parseOverrides(tagVal(tg));
      mutate(o);
      const raw = C.serializeOverrides(o);
      if (raw) sh.tags.add(TAG.OVR, raw);
      else if (!tg.isNullObject) sh.tags.delete(TAG.OVR);
      await ctx.sync();
      s.ovr = C.parseOverrides(raw);
    });
  })
    .then(() => runSync([key], true))
    .then(() => applyVariablesSoon())
    .then(() => renderInspector());
}

function switchVariant(s, newKey) {
  return queue("act.variant", async () => {
    await PowerPoint.run(async (ctx) => {
      shapeById(ctx, s).name = C.componentShapeName(newKey);
      await ctx.sync();
    });
  })
    .then(() => runSync([newKey], true))
    .then(() => {
      refreshCompInfo();
      refreshSelection();
    });
}

function makeMaster(name, s) {
  state.components[name].master = { slideId: s.slideId, shapeId: s.id };
  return saveState().then(() => runSync([name])).then(refreshSelection);
}

function deleteComponent(name) {
  return queue("act.delete", async () => {
    await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      all.filter((s) => C.componentNameOf(s.name) === name).forEach((s) => (s.shape.name = C.stripComponentName(s.name)));
      await ctx.sync();
    });
    delete state.components[name];
    delete ui.compInfo[name];
    await saveState();
    renderComps();
    setStatus(t("msg.compDeleted", { name: label(name) }), "ok");
  });
}

function showMaster(name) {
  const m = state.components[name].master;
  return queue("act.showMaster", async () => {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.setSelectedSlides([m.slideId]);
      await ctx.sync();
      ctx.presentation.slides.getItem(m.slideId).setSelectedShapes([m.shapeId]);
      await ctx.sync();
    });
  });
}

// ================================================================ selection

function isMasterOf(s) {
  const name = C.componentNameOf(s.name);
  const comp = name && state.components[name];
  return !!comp && comp.master.slideId === s.slideId && comp.master.shapeId === s.id;
}

const selSignature = (sel) => (sel ? sel.shapes.map((s) => s.slideId + ":" + s.id).join(",") : "");

let selectionQueued = false;
/** Selection events arrive in bursts (even when only the cursor moves): read once per burst. */
function onSelectionChanged() {
  if (selectionQueued) return;
  refreshSelection();
}

function refreshSelection() {
  selectionQueued = true;
  return queue("act.selection", async () => {
    selectionQueued = false;
    let sel;
    try {
      sel = await PowerPoint.run(async (ctx) => {
        const sel = await getSelection(ctx);
        const one = sel.shapes.length === 1 ? sel.shapes[0] : null;
        const prevOne = ui.selection && ui.selection.shapes.length === 1 ? ui.selection.shapes[0] : null;
        if (one && !state.components[C.componentNameOf(one.name)]) {
          one.ovr = C.parseOverrides(null); // not a component: nothing to read
        } else if (one && prevOne && prevOne.ovr && shapeKey(prevOne) === shapeKey(one) && prevOne.name === one.name) {
          one.ovr = prevOne.ovr; // same shape as before (cursor moved): reuse
        } else if (one) {
          try {
            const tg = loadTag(sel.shapes[0].shape, TAG.OVR);
            await ctx.sync();
            sel.shapes[0].ovr = C.parseOverrides(tagVal(tg));
          } catch (_) {
            sel.shapes[0].ovr = C.parseOverrides(null);
          }
        }
        return sel;
      });
    } catch (e) {
      // Reading the selection is passive – show a hint, not a red error
      console.warn("Selection not readable", e);
      sel = { slideId: null, shapes: [], unreadable: true };
    }
    log("info", "selection: " + (sel.unreadable ? "unreadable" : sel.shapes.length + " shape(s)" + (sel.partial ? " (partial)" : "") + (sel.slideId ? "" : ", slide unknown")));
    ui.lastSelection = ui.selection;
    ui.selection = sel;
    if (selSignature(sel) !== ui.inspectorFor) {
      ui.inspectorForm = null;
      ui.inspectorFor = selSignature(sel);
    }
    renderInspector();

    // Automatic updates after every click: fill in placeholders the user just typed
    // and push changed masters. The shapes selected now are skipped (the user may
    // still be typing in them). If the slides are unknown – PowerPoint on the web
    // sometimes cannot report the selection – the whole deck is checked.
    const prev = ui.lastSelection;
    const unsure = (x) => !x || x.unreadable || x.partial;
    // Still in the same shape (e.g. only the text cursor moved): nothing to update
    if (!unsure(sel) && !unsure(prev) && selSignature(sel) === selSignature(prev)) return;
    const known = [prev && prev.slideId, sel.slideId].filter(Boolean);
    const slideIds = known.length && !unsure(sel) && !unsure(prev) ? [...new Set(known)] : null;
    autoUpdate(slideIds, sel.shapes);
  });
}

// ================================================================ rendering: variables

function updateCounts() {
  $("count-vars").textContent = state.vars.length || "";
  $("count-comps").textContent = Object.keys(state.components).length || "";
}

function renderVars() {
  const list = $("var-list");
  list.replaceChildren();
  updateCounts();

  if (ui.addingVar) list.append(renderAddVarRow());
  $("var-empty").hidden = state.vars.length > 0 || ui.addingVar;
  $("vars-footnote").hidden = state.vars.length === 0;

  state.vars.forEach((v, idx) => {
    const remove = () => {
      state.vars.splice(idx, 1);
      saveState();
      renderVars();
      runApplyVariables();
    };
    let lead, value, action;
    if (v.type === "color") {
      const hex = el("input", { type: "text", class: "cell-input mono", value: v.value || "", placeholder: "#RRGGBB", "aria-label": t("vars.hex") });
      const pick = el("input", { type: "color", value: C.normalizeColor(v.value) || "#000000", "aria-label": t("vars.pick") });
      const commit = (val) => {
        const c = C.normalizeColor(val);
        if (!c) return setStatus(t("err.invalidColor", { value: val }), "err");
        v.value = c;
        hex.value = c;
        pick.value = c;
        sw.style.background = c;
        saveState();
        applyVariablesSoon();
      };
      pick.addEventListener("input", () => commit(pick.value));
      hex.addEventListener("change", () => commit(hex.value));
      const sw = el("label", { class: "swatch", title: t("vars.pick") }, pick);
      sw.style.background = C.normalizeColor(v.value) || "#000000";
      lead = sw;
      value = hex;
      action = menuButton(
        [
          { heading: t("vars.applyTo") },
          { label: t("vars.bind.fill"), onclick: () => bindColor(v.name, "fill") },
          { label: t("vars.bind.line"), onclick: () => bindColor(v.name, "line") },
          { label: t("vars.bind.font"), onclick: () => bindColor(v.name, "font") },
          "sep",
          { label: t("vars.unbind"), onclick: unbindColors },
        ],
        { icon: "paint", title: t("vars.applyTo") }
      );
    } else {
      lead = el("span", { class: "type-glyph", title: t("vars.type.text") }, icon("text"));
      value = el("input", { type: "text", class: "cell-input", value: v.value || "", placeholder: t("vars.valuePh"), "aria-label": t("vars.valueOf", { name: v.name }) });
      value.addEventListener("input", () => {
        v.value = value.value;
        saveState();
        applyVariablesSoon();
      });
      action = iconButton("insert", t("vars.insertTip", { name: v.name }), () => insertPlaceholder(v.name), { class: "icon-btn insert-btn" });
    }
    const more = menuButton([
      v.type !== "color" ? { label: t("vars.insert", { name: v.name }), onclick: () => insertPlaceholder(v.name) } : null,
      { label: t("vars.delete"), danger: true, onclick: remove },
    ]);
    list.append(
      el(
        "li",
        { class: "row var-row", "data-name": v.name },
        lead,
        el("span", { class: "var-name mono", title: "{{" + v.name + "}}" }, v.name),
        value,
        el("span", { class: "row-actions" }, action, more)
      )
    );
  });
}

function renderAddVarRow() {
  let type = "text";
  const name = el("input", { id: "var-name", class: "cell-input", placeholder: t("vars.namePh"), autocomplete: "off", "aria-label": t("vars.name") });
  const seg = el(
    "div",
    { class: "seg", role: "group", "aria-label": t("vars.type") },
    ...["text", "color"].map((ty) =>
      el(
        "button",
        {
          type: "button",
          "data-type": ty,
          class: ty === type ? "active" : "",
          onclick: (e) => {
            type = ty;
            seg.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b === e.currentTarget));
            name.focus();
          },
        },
        t("vars.type." + ty)
      )
    )
  );
  const submit = () => {
    const n = name.value.trim();
    if (!C.isValidName(n)) return setStatus(t("err.invalidName"), "err");
    if (state.vars.some((v) => v.name === n)) return setStatus(t("err.exists", { name: n }), "err");
    state.vars.unshift({ name: n, type, value: type === "color" ? "#C43E1C" : "" });
    ui.addingVar = false;
    saveState();
    renderVars();
    const input = document.querySelector('.var-row[data-name="' + CSS.escape(n) + '"] .cell-input');
    if (input) input.focus();
    setStatus(type === "color" ? t("msg.colorAdded", { name: n }) : t("msg.varAdded", { name: n }), "ok");
  };
  name.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submit();
    }
    if (e.key === "Escape") {
      ui.addingVar = false;
      renderVars();
    }
  });
  setTimeout(() => name.focus(), 0);
  return el(
    "li",
    { class: "row add-row", id: "var-add" },
    seg,
    name,
    el("button", { type: "button", class: "btn primary sm", id: "var-add-submit", onclick: submit }, t("vars.add")),
    iconButton("close", t("ui.cancel"), () => {
      ui.addingVar = false;
      renderVars();
    })
  );
}

// ================================================================ rendering: components

function compSelect() {
  return el(
    "select",
    { class: "select", "aria-label": t("comp.select") },
    Object.keys(state.components).sort().map((n) => el("option", { value: n }, label(n)))
  );
}

/** Inline name form used in the inspector (create component / save as variant). */
function inlineForm(placeholder, buttonText, onSubmit, prefix) {
  const input = el("input", { class: "input", id: "insp-name", placeholder, autocomplete: "off", "aria-label": placeholder });
  const submit = () => onSubmit(input.value.trim());
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      submit();
    }
    if (e.key === "Escape") {
      ui.inspectorForm = null;
      renderInspector();
    }
  });
  setTimeout(() => input.focus(), 0);
  return el(
    "div",
    { class: "inline-form" },
    prefix ? el("span", { class: "prefix" }, prefix) : null,
    input,
    el("button", { type: "button", class: "btn primary sm", id: "insp-submit", onclick: submit }, buttonText),
    iconButton("close", t("ui.cancel"), () => {
      ui.inspectorForm = null;
      renderInspector();
    })
  );
}

function inspectorHead(kind, title, sub, menu) {
  return el(
    "div",
    { class: "insp-head" },
    el("span", { class: "insp-icon " + kind }, icon(kind === "master" ? "diamond" : kind === "instance" ? "diamondOpen" : "cursor")),
    el("div", { class: "insp-title" }, el("div", { class: "insp-name" }, title), sub ? el("div", { class: "insp-sub" }, sub) : null),
    menu || null
  );
}

/** Property fields: placeholders of the master text + (instances) colors. */
function renderProps(s, key, isMaster) {
  const comp = state.components[key];
  const snap = comp.snapshot || {};
  const names = C.extractNames((snap.text && snap.text.source) || "");
  const ovr = s.ovr || C.parseOverrides(null);
  const defaults = snap.defaults || {};
  const box = el("div", { class: "props" });

  box.append(el("div", { class: "props-title" }, isMaster ? t("ovr.defaults") : t("ovr.instance")));
  if (!names.length) box.append(el("p", { class: "hint" }, rich(t("ovr.tip"))));

  for (const n of names) {
    const g = state.vars.find((v) => v.name === n && v.type !== "color");
    const inherited = isMaster ? (g ? g.value : "") : defaults[n] != null ? defaults[n] : g ? g.value : "";
    const own = !!ovr.vars[n];
    const input = el("input", { type: "text", class: "input", value: ovr.vars[n] || "", placeholder: inherited || t("ovr.empty"), "aria-label": n });
    input.addEventListener("change", () => setOverride(s, key, (o) => (o.vars[n] = input.value)));
    box.append(
      el(
        "label",
        { class: "prop" },
        el("span", { class: "prop-label mono" + (own ? " own" : ""), title: own ? t("ovr.src.own") : inherited ? t("ovr.inherited") : "" }, n),
        input
      )
    );
  }

  if (!isMaster) {
    const masterColor = {
      fill: snap.fill && snap.fill.color,
      font: snap.text && snap.text.font && snap.text.font.color,
      line: snap.line && snap.line.color,
    };
    const colors = el("div", { class: "color-props" });
    for (const k of ["fill", "font", "line"]) {
      const own = !!ovr[k];
      const val = ovr[k] || C.normalizeColor(masterColor[k]) || "#000000";
      const pick = el("input", { type: "color", value: val, "aria-label": t("ovr." + k) });
      pick.addEventListener("change", () => setOverride(s, key, (o) => (o[k] = pick.value)));
      const sw = el("label", { class: "swatch sm" }, pick);
      sw.style.background = val;
      colors.append(
        el(
          "div",
          { class: "color-prop" + (own ? " own" : "") },
          sw,
          el("span", { class: "color-label" }, t("ovr." + k)),
          own ? iconButton("close", t("ovr.reset"), () => setOverride(s, key, (o) => (o[k] = null)), { class: "icon-btn tiny reset-color" }) : null
        )
      );
    }
    box.append(colors);
  }
  return box;
}

function renderInspector() {
  const box = $("inspector");
  box.replaceChildren();
  const sel = ui.selection;
  const hasComps = Object.keys(state.components).length > 0;

  if (!sel || !sel.shapes.length) {
    box.className = "inspector empty";
    box.append(
      el("span", { class: "insp-icon" }, icon("cursor")),
      el("p", {}, sel && (sel.unreadable || sel.partial) ? t("sel.unreadable") : t("sel.none"))
    );
    return;
  }
  box.className = "inspector";

  if (sel.shapes.length > 1) {
    box.append(inspectorHead("plain", t("sel.multi", { n: sel.shapes.length }), null, menuButton([{ label: t("sel.detachAll"), onclick: detachSelection }])));
    if (hasComps) {
      const s = compSelect();
      box.append(el("div", { class: "insp-actions" }, s, el("button", { type: "button", class: "btn primary sm", onclick: () => linkSelection(s.value) }, t("sel.link"))));
    }
    return;
  }

  const s = sel.shapes[0];
  const key = C.componentNameOf(s.name);
  const comp = key && state.components[key];

  if (comp && isMasterOf(s)) {
    const info = ui.compInfo[key];
    const n = info ? info.instances : 0;
    box.append(
      inspectorHead("master", label(key), t("sel.masterSub"), menuButton([
        { label: t("sel.addVariant"), onclick: () => ((ui.inspectorForm = "variant"), renderInspector()) },
      ]))
    );
    if (ui.inspectorForm === "variant") {
      box.append(inlineForm(t("new.variantPh"), t("new.create"), (v) => createComponent(C.joinKey(C.splitKey(key).base, v)), C.splitKey(key).base + " /"));
      box.append(el("p", { class: "hint" }, t("sel.variantFromMasterHint")));
    }
    box.append(
      el("div", { class: "insp-actions" }, el("button", { type: "button", class: "btn primary block", id: "push-btn", onclick: () => runSync([key]) }, icon("sync"), t("sel.push", { n })))
    );
    box.append(renderProps(s, key, true));
    return;
  }

  if (comp) {
    const base = C.splitKey(key).base;
    box.append(
      inspectorHead("instance", label(key), t("sel.instanceSub"), menuButton([
        { label: t("sel.showMaster"), onclick: () => showMaster(key) },
        { label: t("sel.saveVariant"), onclick: () => ((ui.inspectorForm = "variant"), renderInspector()) },
        { label: t("sel.makeMaster"), onclick: () => makeMaster(key, s) },
        "sep",
        { label: t("ovr.resetAll"), disabled: !C.serializeOverrides(s.ovr), onclick: () => setOverride(s, key, (o) => { o.vars = {}; o.fill = o.font = o.line = null; }) },
        { label: t("sel.detach"), danger: true, onclick: detachSelection },
      ]))
    );
    if (ui.inspectorForm === "variant") {
      box.append(inlineForm(t("new.variantPh"), t("new.create"), (v) => createComponent(C.joinKey(base, v)), base + " /"));
      box.append(el("p", { class: "hint" }, t("sel.saveVariantHint")));
    }
    const variants = variantsOf(base);
    if (variants.length > 1) {
      const vs = el(
        "select",
        { class: "select", id: "variant-select", "aria-label": t("sel.variant"), onchange: () => switchVariant(s, vs.value) },
        variants.map((k) => el("option", { value: k, selected: k === key }, variantLabel(k)))
      );
      box.append(el("label", { class: "prop" }, el("span", { class: "prop-label" }, t("sel.variant")), vs));
    }
    box.append(renderProps(s, key, false));
    return;
  }

  if (key) {
    box.append(inspectorHead("plain", s.name, t("sel.unregisteredSub")));
    box.append(el("div", { class: "insp-actions" }, el("button", { type: "button", class: "btn primary block", onclick: () => createComponent(key) }, t("sel.register"))));
    return;
  }

  // Plain shape
  box.append(inspectorHead("plain", s.name, typeLabel(s.type)));
  if (!COMP_TYPES.includes(s.type)) {
    box.append(el("p", { class: "hint" }, t("err.typeUnsupported", { type: typeLabel(s.type) })));
    return;
  }
  if (ui.inspectorForm === "create") {
    box.append(inlineForm(t("new.namePh"), t("new.create"), (n) => createComponent(n)));
  } else {
    box.append(
      el(
        "div",
        { class: "insp-actions" },
        el("button", { type: "button", class: "btn primary block", id: "create-btn", onclick: () => ((ui.inspectorForm = "create"), renderInspector()) }, icon("diamond"), t("sel.create"))
      )
    );
  }
  if (hasComps) {
    const sl = compSelect();
    box.append(
      el("div", { class: "insp-actions secondary" }, el("span", { class: "or" }, t("sel.orLink")), sl, el("button", { type: "button", class: "btn sm", id: "link-btn", onclick: () => linkSelection(sl.value) }, t("sel.link")))
    );
  }
}

function renderComps() {
  const list = $("comp-list");
  list.replaceChildren();
  updateCounts();
  const names = Object.keys(state.components).sort();
  $("comp-empty").hidden = names.length > 0;
  $("comps-footnote").hidden = names.length === 0;

  for (const name of names) {
    const comp = state.components[name];
    const info = ui.compInfo[name];
    const k = C.splitKey(name);
    const isVariant = k.variant !== C.DEFAULT_VARIANT;
    let meta = "";
    if (info) {
      meta = info.masterFound
        ? info.masterSlide
          ? t("comps.masterOn", { slide: info.masterSlide, n: info.instances })
          : t("comps.master", { n: info.instances })
        : t("comps.instances", { n: info.instances });
    }
    const opt = (prop, text) => ({
      label: text,
      checked: comp[prop] !== false,
      onclick: () => {
        comp[prop] = comp[prop] === false;
        saveState();
      },
    });
    const confirming = ui.confirmDelete === name;
    list.append(
      el(
        "li",
        { class: "row comp-row" + (isVariant ? " variant" : ""), "data-name": name },
        el("span", { class: "comp-glyph" }, icon("diamond")),
        el(
          "div",
          { class: "comp-text" },
          el("div", { class: "comp-name" }, isVariant ? k.variant : k.base, info && !info.masterFound ? el("span", { class: "badge warn" }, t("comps.masterMissing")) : null),
          el("div", { class: "comp-meta" }, meta)
        ),
        el(
          "span",
          { class: "row-actions" },
          iconButton("sync", t("comps.sync"), () => runSync([name]), { class: "icon-btn sync-btn" }),
          iconButton("target", t("sel.showMaster"), () => showMaster(name), { disabled: info && !info.masterFound }),
          menuButton(() => [
            opt("syncText", t("comps.syncText")),
            opt("syncSize", t("comps.syncSize")),
            "sep",
            {
              label: confirming ? t("comps.confirmDelete") : t("comps.delete"),
              danger: true,
              keepOpen: !confirming,
              onclick: () => {
                if (confirming) {
                  ui.confirmDelete = null;
                  deleteComponent(name);
                } else {
                  ui.confirmDelete = name;
                  closeMenu();
                  renderComps();
                  const btn = document.querySelector('.comp-row[data-name="' + CSS.escape(name) + '"] .menu-btn');
                  if (btn) btn.click();
                  setTimeout(() => {
                    if (ui.confirmDelete === name) ui.confirmDelete = null;
                  }, 5000);
                }
              },
            },
          ])
        )
      )
    );
  }
  renderInspector();
}

// ================================================================ shell

function switchTab(tab) {
  document.querySelectorAll(".tab").forEach((b) => {
    const on = b.dataset.tab === tab;
    b.classList.toggle("active", on);
    b.setAttribute("aria-selected", on);
  });
  $("tab-vars").hidden = tab !== "vars";
  $("tab-comps").hidden = tab !== "comps";
  lsSet("dc.tab", tab);
  if (tab === "comps") {
    refreshCompInfo();
    refreshSelection();
  }
}

/** Applies the current language to all static texts (data-i18n attributes). */
function applyStaticTexts() {
  document.documentElement.lang = I.getLang();
  document.querySelectorAll("[data-i18n]").forEach((n) => n.replaceChildren(rich(t(n.dataset.i18n))));
  document.querySelectorAll("[data-i18n-title]").forEach((n) => {
    n.title = t(n.dataset.i18nTitle);
    n.setAttribute("aria-label", n.title);
  });
}

function setLanguage(lang) {
  I.setLang(lang);
  lsSet("dc.lang", lang);
  applyStaticTexts();
  renderVars();
  renderComps();
}

function showVersion() {
  const meta = document.querySelector('meta[name="dc-build"]');
  const build = meta && !meta.content.startsWith("__") ? meta.content : "dev";
  $("version").textContent = "v" + VERSION + " · " + build;
}

function diagnosticsText() {
  const d = (Office.context && Office.context.diagnostics) || {};
  const req = (v) => {
    try {
      return Office.context.requirements.isSetSupported("PowerPointApi", v) ? "yes" : "no";
    } catch (_) {
      return "?";
    }
  };
  const sel = ui.selection;
  return [
    "Deck Components " + $("version").textContent,
    "Office: " + (d.host || "?") + " · " + (d.platform || "?") + " · " + (d.version || "?") + " · " + (Office.context.displayLanguage || "?"),
    "PowerPointApi 1.3: " + req("1.3") + " · 1.4: " + req("1.4") + " · 1.5: " + req("1.5") + " · 1.8: " + req("1.8"),
    "Variables: " + state.vars.map((v) => v.name + "=" + JSON.stringify(v.value)).join(", "),
    "Components: " + (Object.keys(state.components).join(", ") || "none"),
    "Placeholder view: " + state.showRaw + " · auto-update instances: " + ui.autoSync,
    "Selection: " + (sel ? (sel.unreadable ? "unreadable" : sel.shapes.map((s) => s.name + " [" + s.type + "]").join(", ") || "nothing") : "not read yet"),
    "",
    "Recent events (newest last):",
    ...LOG,
  ].join("\n");
}

function openDiagnostics() {
  const close = () => overlay.remove();
  const pre = el("pre", { class: "diag-log", id: "diag-log", tabindex: "0" }, diagnosticsText());
  const overlay = el(
    "div",
    { class: "overlay", id: "diag", onclick: (e) => e.target === overlay && close() },
    el(
      "div",
      { class: "dialog", role: "dialog", "aria-label": t("diag.title") },
      el("div", { class: "dialog-head" }, el("strong", {}, t("diag.title")), iconButton("close", t("ui.close"), close)),
      el("p", { class: "hint" }, t("diag.hint")),
      pre,
      el(
        "div",
        { class: "dialog-actions" },
        el("button", { type: "button", class: "btn sm", onclick: () => { pre.textContent = diagnosticsText(); } }, t("diag.refresh")),
        el("button", { type: "button", class: "btn sm", onclick: runApplyVariables }, t("vars.apply")),
        el(
          "button",
          {
            type: "button",
            class: "btn primary sm",
            onclick: () => {
              const range = document.createRange();
              range.selectNodeContents(pre);
              const s = window.getSelection();
              s.removeAllRanges();
              s.addRange(range);
              let ok = false;
              try {
                ok = document.execCommand("copy");
              } catch (_) {
                ok = false;
              }
              setStatus(ok ? t("diag.copied") : t("diag.copyManual"), "ok");
            },
          },
          t("diag.copy")
        )
      )
    )
  );
  document.body.append(overlay);
}

function mountToolbarMenus() {
  $("settings-slot").replaceChildren(
    menuButton(
      () => [
        { heading: t("settings.language") },
        { label: "English", checked: I.getLang() === "en", onclick: () => setLanguage("en") },
        { label: "Deutsch", checked: I.getLang() === "de", onclick: () => setLanguage("de") },
        "sep",
        { heading: t("settings.behavior") },
        {
          label: t("vars.showRaw"),
          checked: state.showRaw,
          onclick: () => {
            state.showRaw = !state.showRaw;
            saveState();
            runApplyVariables();
          },
        },
        {
          label: t("comps.autoSync"),
          checked: ui.autoSync,
          onclick: () => {
            ui.autoSync = !ui.autoSync;
            lsSet("dc.autoSync", ui.autoSync ? "1" : "0");
          },
        },
        "sep",
        { label: t("diag.title"), onclick: openDiagnostics },
      ],
      { icon: "gear", title: t("settings.title") }
    )
  );
  $("vars-menu-slot").replaceChildren(
    menuButton(() => [
      { label: t("vars.apply"), onclick: runApplyVariables },
      { label: t("vars.scan"), onclick: scanPlaceholders },
      "sep",
      {
        label: t("vars.showRaw"),
        checked: state.showRaw,
        onclick: () => {
          state.showRaw = !state.showRaw;
          saveState();
          runApplyVariables();
        },
      },
    ])
  );
  $("comps-menu-slot").replaceChildren(
    menuButton(() => [
      { label: t("comps.syncAll"), onclick: () => runSync(null) },
      { label: t("comps.recount"), onclick: refreshCompInfo },
      "sep",
      {
        label: t("comps.autoSync"),
        checked: ui.autoSync,
        onclick: () => {
          ui.autoSync = !ui.autoSync;
          lsSet("dc.autoSync", ui.autoSync ? "1" : "0");
        },
      },
    ])
  );
}

function startAddVar() {
  ui.addingVar = true;
  switchTab("vars");
  renderVars();
}

// ================================================================ start

Office.onReady((info) => {
  I.setLang(I.detect(Office.context.displayLanguage, lsGet("dc.lang")));
  applyStaticTexts();
  showVersion();
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus(t("hostOnly"), "err");
    return;
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) $("unsupported").hidden = false;
  loadState();
  log("info", "pane opened, " + state.vars.length + " variable(s), " + Object.keys(state.components).length + " component(s)");
  ui.autoSync = lsGet("dc.autoSync") !== "0"; // on unless the user turned it off

  document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => switchTab(b.dataset.tab)));
  $("var-new").addEventListener("click", startAddVar);
  $("var-empty-btn").addEventListener("click", startAddVar);
  $("status-close").addEventListener("click", () => $("status").classList.remove("show"));
  mountToolbarMenus();

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, debounce(onSelectionChanged, 300));

  renderVars();
  renderComps();
  switchTab(lsGet("dc.tab") === "comps" ? "comps" : "vars");
  // Fill in anything typed while the pane was closed (not the shape being edited)
  PowerPoint.run(getSelection)
    .catch(() => ({ shapes: [] }))
    .then((sel) => autoUpdate(null, sel.shapes))
    .then(() => refreshSelection());
});
