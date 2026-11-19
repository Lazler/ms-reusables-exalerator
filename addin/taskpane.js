/* Deck Components – task pane (Office.js / PowerPoint) */
/* global Office, PowerPoint, DCCore, DCI18n */
"use strict";

const C = DCCore;
const I = DCI18n;
const t = I.t;
const VERSION = "0.4.0";
const STATE_KEY = "deckComponents.v1";
const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder"];
const COMP_TYPES = ["GeometricShape", "TextBox"];
const TAG = { TPL: "DC_TPL", OUT: "DC_OUT", FILL: "DC_FILL", LINE: "DC_LINE", FONT: "DC_FONT", OVR: "DC_OVR" };

/** Persisted inside the .pptx (document settings). */
let state = { vars: [], components: {}, showRaw: false };
/** Runtime only. */
const ui = { compInfo: {}, selection: null, lastSelection: null, pendingDelete: null };

// ---------------------------------------------------------------- helpers

const $ = (id) => document.getElementById(id);

function el(tag, attrs, ...children) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === "class") n.className = v;
    else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
    else if (k in n && k !== "list") n[k] = v;
    else n.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat()) if (c != null && c !== false) n.append(c.nodeType ? c : String(c));
  return n;
}

function setStatus(msg, kind) {
  const f = $("status");
  f.textContent = msg || "";
  f.className = kind || "";
}

/** Renders text and shows {{…}} and ◆ … ◆ as inline code. */
function rich(text) {
  const frag = document.createDocumentFragment();
  String(text)
    .split(/(\{\{[^{}]+\}\}|◆[^◆]+◆)/u)
    .forEach((part, i) => frag.append(i % 2 ? el("code", {}, part) : part));
  return frag;
}

function errText(e) {
  if (!e) return t("err.unknown");
  let s = e.message || String(e);
  if (e.code) s += " [" + e.code + "]";
  if (e.debugInfo && e.debugInfo.errorLocation) s += "\n" + t("err.location") + ": " + e.debugInfo.errorLocation;
  return s;
}

function debounce(fn, ms) {
  let t;
  return (...a) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}

/** Runs all PowerPoint calls one after another and reports errors in the status bar. */
let chain = Promise.resolve();
function queue(label, fn) {
  if (typeof label === "function") {
    fn = label;
    label = null;
  }
  const p = chain.then(fn).catch((e) => {
    console.error(label, e);
    setStatus((label ? t("err.during", { action: t(label) }) : t("err.prefix")) + ": " + errText(e), "err");
  });
  chain = p;
  return p;
}

function saveState() {
  const s = Office.context.document.settings;
  s.set(STATE_KEY, JSON.stringify(state));
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

// ---------------------------------------------------------------- PowerPoint access

/** Loads all shapes on all slides (group contents excluded). */
async function loadShapes(ctx) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  slides.items.forEach((sl) => sl.shapes.load("items/id,items/name,items/type"));
  await ctx.sync();
  const out = [];
  slides.items.forEach((sl, i) =>
    sl.shapes.items.forEach((sh) =>
      out.push({ slideId: sl.id, slideIndex: i, shape: sh, id: sh.id, name: sh.name, type: sh.type })
    )
  );
  return out;
}

function loadTag(shape, key) {
  const t = shape.tags.getItemOrNullObject(key);
  t.load("value");
  return t;
}
const tagVal = (t) => (t && !t.isNullObject ? t.value : null);

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
  } catch (_) {
    const ok = [];
    for (const it of items) {
      try {
        loader(it);
        await ctx.sync();
        ok.push(it);
      } catch (e) {
        console.warn("Skipped:", it.name, e);
      }
    }
    return ok;
  }
}

/**
 * Reads the current selection. PowerPoint on the web throws a GeneralException in
 * some states (focus in the thumbnail pane, text in edit mode, groups/placeholders),
 * so this falls back in stages instead of failing.
 */
async function getSelection(ctx) {
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

  // Stage 2: shape by shape, type optional
  const list = [];
  try {
    const shapes = ctx.presentation.getSelectedShapes();
    const count = shapes.getCount();
    await ctx.sync();
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

// ---------------------------------------------------------------- variables

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

/** Global → master default → the shape's own override. */
function effectiveVars(shapeName, ovr) {
  return C.mergeVars(textVars(), compDefaults(shapeName), ovr.vars);
}

/** Renders all text boxes and color bindings in the deck. */
async function applyVariables(ctx) {
  const all = await loadShapes(ctx);
  const items = await loadEach(ctx, all, (it) => {
    it.tags = {};
    for (const k of Object.values(TAG)) it.tags[k] = loadTag(it.shape, k);
    if (TEXT_TYPES.includes(it.type)) {
      it.tr = it.shape.textFrame.textRange;
      it.tr.load("text");
    }
  });

  const missing = new Set();
  let changed = 0;
  let fields = 0;

  for (const it of items) {
    const ovr = C.parseOverrides(tagVal(it.tags[TAG.OVR]));
    const vars = effectiveVars(it.name, ovr);
    // Text
    if (it.tr) {
      const current = it.tr.text;
      const oldTpl = tagVal(it.tags[TAG.TPL]);
      const tpl = C.resolveTemplate(current, oldTpl, tagVal(it.tags[TAG.OUT]), vars);
      if (tpl) {
        fields++;
        const r = C.render(tpl, vars);
        r.missing.forEach((m) => missing.add(m));
        const out = state.showRaw ? tpl : r.text;
        if (!C.sameText(out, current)) {
          it.tr.text = out;
          changed++;
        }
        it.shape.tags.add(TAG.TPL, tpl);
        it.shape.tags.add(TAG.OUT, out);
      } else if (oldTpl != null) {
        it.shape.tags.delete(TAG.TPL);
        if (tagVal(it.tags[TAG.OUT]) != null) it.shape.tags.delete(TAG.OUT);
      }
    }
    // Colors: instance override > color variable
    const fill = ovr.fill || colorOf(tagVal(it.tags[TAG.FILL]));
    if (fill) it.shape.fill.setSolidColor(fill);
    const line = ovr.line || colorOf(tagVal(it.tags[TAG.LINE]));
    if (line) {
      it.shape.lineFormat.visible = true;
      it.shape.lineFormat.color = line;
    }
    const font = ovr.font || colorOf(tagVal(it.tags[TAG.FONT]));
    if (font && it.tr) it.tr.font.color = font;
  }
  await ctx.sync();
  return { changed, fields, missing: [...missing] };
}

function runApplyVariables() {
  return queue("act.apply", async () => {
    setStatus(t("msg.updating"));
    const r = await PowerPoint.run(applyVariables);
    let msg = t("msg.applied", { fields: r.fields, changed: r.changed });
    if (state.showRaw) msg += " " + t("msg.rawView");
    if (r.missing.length) {
      setStatus(msg + "\n" + t("msg.missing", { names: r.missing.join(", ") }), "warn");
    } else setStatus(msg, "ok");
  });
}
const applyVariablesSoon = debounce(runApplyVariables, 700);

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
  }).then(() => runApplyVariables());
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
        it.t.forEach(([k, t]) => {
          if (!t.isNullObject) {
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

// ---------------------------------------------------------------- components

async function readStyle(ctx, shape, type) {
  shape.load("width,height");
  shape.fill.load("type,foregroundColor,transparency");
  shape.lineFormat.load("visible,color,weight,dashStyle");
  await ctx.sync();
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
  if (!TEXT_TYPES.includes(type)) return st;
  try {
    const tf = shape.textFrame;
    const tr = tf.textRange;
    tr.load("text");
    tr.font.load("name,size,bold,italic,underline,color");
    tr.paragraphFormat.load("horizontalAlignment");
    tf.load("verticalAlignment,leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,autoSizeSetting");
    const tpl = loadTag(shape, TAG.TPL);
    const out = loadTag(shape, TAG.OUT);
    const ovr = loadTag(shape, TAG.OVR);
    await ctx.sync();
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
  } catch (e) {
    console.warn("Text format not readable", e);
  }
  return st;
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
  const t = st.text;
  if (t && inst.tr) {
    const tf = sh.textFrame;
    if (t.autoSize) tf.autoSizeSetting = t.autoSize;
    if (comp.syncText !== false) {
      if (C.hasPlaceholders(t.source)) {
        const out = state.showRaw ? t.source : C.render(t.source, vars).text;
        inst.tr.text = out;
        sh.tags.add(TAG.TPL, t.source);
        sh.tags.add(TAG.OUT, out);
      } else {
        inst.tr.text = t.source;
        if (!inst.tpl.isNullObject) sh.tags.delete(TAG.TPL);
        if (!inst.out.isNullObject) sh.tags.delete(TAG.OUT);
      }
    }
    const f = t.font || {};
    for (const k of ["name", "size", "bold", "italic", "underline", "color"]) {
      if (isSet(f[k]) && f[k] !== "") inst.tr.font[k] = f[k];
    }
    if (t.hAlign) inst.tr.paragraphFormat.horizontalAlignment = t.hAlign;
    if (t.vAlign) tf.verticalAlignment = t.vAlign;
    if (t.margins) {
      const [l, r, tp, b] = t.margins;
      if (isSet(l)) tf.leftMargin = l;
      if (isSet(r)) tf.rightMargin = r;
      if (isSet(tp)) tf.topMargin = tp;
      if (isSet(b)) tf.bottomMargin = b;
    }
    if (isSet(t.wordWrap)) tf.wordWrap = t.wordWrap;
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

/** Syncs the given (or all) components. Returns statistics. */
async function syncComponents(ctx, names) {
  const all = await loadShapes(ctx);
  const result = { synced: 0, warnings: [] };
  ui.compInfo = {};
  for (const name of Object.keys(state.components)) {
    const comp = state.components[name];
    const members = all.filter((s) => C.componentNameOf(s.name) === name);
    const master = members.find((s) => s.slideId === comp.master.slideId && s.id === comp.master.shapeId);
    const instances = members.filter((s) => s !== master);
    ui.compInfo[name] = {
      instances: instances.length,
      masterFound: !!master,
      masterSlide: master ? master.slideIndex + 1 : null,
    };
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

    loaded.forEach((it) => applyStyle(it, st, comp));
    try {
      await ctx.sync();
      result.synced += loaded.length;
    } catch (e) {
      // Retry one by one to find the failing instance
      for (const it of loaded) {
        try {
          applyStyle(it, st, comp);
          await ctx.sync();
          result.synced++;
        } catch (e2) {
          result.warnings.push(t("warn.slide", { name: label(name), slide: it.slideIndex + 1, error: errText(e2) }));
        }
      }
    }
  }
  return result;
}

function runSync(names) {
  return queue("act.sync", async () => {
    setStatus(t("msg.syncing"));
    const r = await PowerPoint.run((ctx) => syncComponents(ctx, names));
    await saveState();
    renderComps();
    const msg = t("msg.synced", { n: r.synced });
    setStatus(r.warnings.length ? msg + "\n" + r.warnings.join("\n") : msg, r.warnings.length ? "warn" : "ok");
  });
}

function refreshCompInfo() {
  return queue("act.count", async () => {
    await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      ui.compInfo = {};
      for (const [name, comp] of Object.entries(state.components)) {
        const members = all.filter((s) => C.componentNameOf(s.name) === name);
        const master = members.find((s) => s.slideId === comp.master.slideId && s.id === comp.master.shapeId);
        ui.compInfo[name] = {
          instances: members.length - (master ? 1 : 0),
          masterFound: !!master,
          masterSlide: master ? master.slideIndex + 1 : null,
        };
      }
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

function shapeById(ctx, s) {
  return ctx.presentation.slides.getItem(s.slideId).shapes.getItem(s.id);
}

/** key = "Button" (new component) or "Button / Secondary" (variant). */
function createComponent(key) {
  const parts = C.splitKey(key || "");
  if (!C.isValidName(parts.base) || !C.isValidName(parts.variant)) {
    return setStatus(t("err.invalidName"), "err");
  }
  const name = C.joinKey(parts.base, parts.variant);
  if (state.components[name]) return setStatus(t("err.exists", { name: label(name) }), "err");
  return queue("act.create", async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      if (sel.shapes.length !== 1) throw new Error(t("err.exactlyOne"));
      const s = sel.shapes[0];
      if (!COMP_TYPES.includes(s.type)) {
        throw new Error(t("err.typeUnsupported", { type: s.type }));
      }
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
    await saveState();
    $("comp-name").value = "";
    renderComps();
    setStatus(
      parts.variant === C.DEFAULT_VARIANT
        ? t("msg.compCreated", { name })
        : t("msg.variantCreated", { name: label(name) }),
      "ok"
    );
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
      const targets = sel.shapes.filter(
        (s) => !(s.slideId === comp.master.slideId && s.id === comp.master.shapeId)
      );
      if (!targets.length) throw new Error(t("err.selectTargets"));
      targets.forEach((s) => (s.shape.name = C.componentShapeName(name)));
      await ctx.sync();
      return targets.length;
    });
    setStatus(t("msg.linked", { n, name: label(name) }), "ok");
  }).then(() => runSync([name]).then(refreshSelection));
}

function detachSelection() {
  return queue("act.detach", async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      sel.shapes.forEach((s) => {
        const comp = state.components[C.componentNameOf(s.name)];
        if (comp && comp.master.slideId === s.slideId && comp.master.shapeId === s.id) {
          throw new Error(t("err.detachMaster"));
        }
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
      const t = loadTag(sh, TAG.OVR);
      await ctx.sync();
      const o = C.parseOverrides(tagVal(t));
      mutate(o);
      const raw = C.serializeOverrides(o);
      if (raw) sh.tags.add(TAG.OVR, raw);
      else if (!t.isNullObject) sh.tags.delete(TAG.OVR);
      await ctx.sync();
      s.ovr = C.parseOverrides(raw);
    });
  })
    .then(() => runSync([key]))
    .then(() => runApplyVariables())
    .then(() => renderSelection());
}

function switchVariant(s, newKey) {
  return queue("act.variant", async () => {
    await PowerPoint.run(async (ctx) => {
      shapeById(ctx, s).name = C.componentShapeName(newKey);
      await ctx.sync();
    });
  })
    .then(() => runSync([newKey]))
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

// ---------------------------------------------------------------- selection

function isMasterOf(s) {
  const name = C.componentNameOf(s.name);
  const comp = name && state.components[name];
  return !!comp && comp.master.slideId === s.slideId && comp.master.shapeId === s.id;
}

function refreshSelection() {
  return queue("act.selection", async () => {
    let sel;
    try {
      sel = await PowerPoint.run(async (ctx) => {
        const sel = await getSelection(ctx);
        if (sel.shapes.length === 1) {
          try {
            const t = loadTag(sel.shapes[0].shape, TAG.OVR);
            await ctx.sync();
            sel.shapes[0].ovr = C.parseOverrides(tagVal(t));
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
    ui.lastSelection = ui.selection;
    ui.selection = sel;
    renderSelection();

    // Auto-sync: the selection has left a master
    if ($("auto-sync").checked && ui.lastSelection) {
      const left = ui.lastSelection.shapes.filter(
        (p) => isMasterOf(p) && !sel.shapes.some((s) => s.slideId === p.slideId && s.id === p.id)
      );
      const names = [...new Set(left.map((p) => C.componentNameOf(p.name)))];
      if (names.length) runSync(names);
    }
  });
}

function compSelect() {
  return el("select", { "aria-label": t("comp.select") }, Object.keys(state.components).sort().map((n) => el("option", { value: n }, label(n))));
}

/** "Overrides" box (instance) or "Defaults" box (master). */
function renderOverrides(s, key, isMaster) {
  const comp = state.components[key];
  const snap = comp.snapshot || {};
  const src = (snap.text && snap.text.source) || "";
  const names = C.extractNames(src);
  const ovr = s.ovr || C.parseOverrides(null);
  const defaults = snap.defaults || {};
  const box = el("div", { class: "ovr" }, el("h3", {}, isMaster ? t("ovr.defaults") : t("ovr.instance")));

  if (!names.length) box.append(el("p", { class: "hint" }, rich(t("ovr.tip"))));
  for (const n of names) {
    const g = state.vars.find((v) => v.name === n && v.type !== "color");
    const fallback = isMaster ? (g ? g.value : "") : defaults[n] != null ? defaults[n] : g ? g.value : "";
    const input = el("input", { type: "text", value: ovr.vars[n] || "", placeholder: fallback || t("ovr.empty"), "aria-label": n });
    input.addEventListener("change", () => setOverride(s, key, (o) => (o.vars[n] = input.value)));
    const src = isMaster
      ? g ? t("ovr.src.global", { value: g.value }) : t("ovr.src.local")
      : ovr.vars[n] ? t("ovr.src.own") : defaults[n] != null ? t("ovr.src.master") : g ? t("ovr.src.globalShort") : t("ovr.src.none");
    box.append(el("label", { class: "field" }, el("span", { class: "meta" }, "{{" + n + "}} · " + src), input));
  }

  if (!isMaster) {
    const masterColor = {
      fill: snap.fill && snap.fill.color,
      font: snap.text && snap.text.font && snap.text.font.color,
      line: snap.line && snap.line.color,
    };
    for (const [k, lbl] of [["fill", t("ovr.fill")], ["font", t("ovr.font")], ["line", t("ovr.line")]]) {
      const own = !!ovr[k];
      const pick = el("input", { type: "color", value: ovr[k] || C.normalizeColor(masterColor[k]) || "#000000", "aria-label": lbl });
      pick.addEventListener("change", () => setOverride(s, key, (o) => (o[k] = pick.value)));
      box.append(
        el(
          "div",
          { class: "row" },
          el("span", { class: "meta grow" }, lbl + " · " + (own ? t("ovr.ownColor") : t("ovr.fromMaster"))),
          pick,
          el("button", { class: "small", disabled: !own, onclick: () => setOverride(s, key, (o) => (o[k] = null)) }, t("ovr.reset"))
        )
      );
    }
    if (C.serializeOverrides(ovr)) {
      box.append(
        el("div", { class: "row" }, el("button", { onclick: () => setOverride(s, key, (o) => {
          o.vars = {};
          o.fill = o.font = o.line = null;
        }) }, t("ovr.resetAll")))
      );
    }
  }
  return box;
}

function renderSelection() {
  const info = $("sel-info");
  const act = $("sel-actions");
  act.replaceChildren();
  const sel = ui.selection;
  if (!sel || !sel.shapes.length) {
    info.textContent = sel && (sel.unreadable || sel.partial) ? t("sel.unreadable") : t("sel.none");
    return;
  }
  const hasComps = Object.keys(state.components).length > 0;
  if (sel.shapes.length > 1) {
    info.textContent = t("sel.multi", { n: sel.shapes.length });
    if (hasComps) {
      const s = compSelect();
      act.append(el("div", { class: "row" }, s, el("button", { onclick: () => linkSelection(s.value) }, t("sel.linkAll"))));
    }
    act.append(el("div", { class: "row" }, el("button", { onclick: detachSelection }, t("sel.detachAll"))));
    return;
  }
  const s = sel.shapes[0];
  const cname = C.componentNameOf(s.name);
  const comp = cname && state.components[cname];
  if (comp && isMasterOf(s)) {
    info.replaceChildren(el("b", {}, t("sel.masterOf", { name: label(cname) })));
    act.append(el("div", { class: "row" }, el("button", { class: "primary", onclick: () => runSync([cname]) }, t("sel.updateInstances"))));
    act.append(renderOverrides(s, cname, true));
  } else if (comp) {
    info.replaceChildren(el("b", {}, t("sel.instanceOf", { name: label(cname) })));
    act.append(
      el(
        "div",
        { class: "row" },
        el("button", { onclick: () => showMaster(cname) }, t("sel.showMaster")),
        el("button", { onclick: () => makeMaster(cname, s) }, t("sel.makeMaster")),
        el("button", { onclick: detachSelection }, t("sel.detach"))
      )
    );
    const variants = variantsOf(C.splitKey(cname).base);
    if (variants.length > 1) {
      const vs = el(
        "select",
        { "aria-label": t("sel.variant"), onchange: () => switchVariant(s, vs.value) },
        variants.map((k) => el("option", { value: k, selected: k === cname }, variantLabel(k)))
      );
      act.append(el("div", { class: "row" }, el("span", { class: "meta" }, t("sel.variant")), vs));
    }
    act.append(renderOverrides(s, cname, false));
  } else if (cname) {
    info.textContent = t("sel.unregistered", { shape: s.name, name: label(cname) });
    act.append(
      el("div", { class: "row" }, el("button", { onclick: () => createComponent(cname) }, t("sel.register")))
    );
  } else {
    info.textContent = t("sel.plain", { shape: s.name, type: s.type });
    if (hasComps) {
      const sl = compSelect();
      act.append(el("div", { class: "row" }, sl, el("button", { onclick: () => linkSelection(sl.value) }, t("sel.link"))));
    }
  }
}

// ---------------------------------------------------------------- rendering

function renderVars() {
  const list = $("var-list");
  list.replaceChildren();
  $("var-empty").hidden = state.vars.length > 0;
  state.vars.forEach((v, idx) => {
    const remove = el("button", {
      class: "small danger",
      title: t("vars.delete"),
      onclick: () => {
        state.vars.splice(idx, 1);
        saveState();
        renderVars();
        runApplyVariables();
      },
    }, "✕");

    const head = el("div", { class: "head" }, el("span", { class: "name" }, v.name), el("span", { class: "badge" }, v.type === "color" ? t("vars.type.color") : t("vars.type.text")), remove);

    let value;
    let controls;
    if (v.type === "color") {
      const hex = el("input", { type: "text", value: v.value || "", placeholder: "#RRGGBB", "aria-label": t("vars.hex") });
      const pick = el("input", { type: "color", value: C.normalizeColor(v.value) || "#000000", "aria-label": t("vars.pick") });
      const commit = (val) => {
        const c = C.normalizeColor(val);
        if (!c) return setStatus(t("err.invalidColor", { value: val }), "err");
        v.value = c;
        hex.value = c;
        pick.value = c;
        saveState();
        applyVariablesSoon();
      };
      pick.addEventListener("input", () => commit(pick.value));
      hex.addEventListener("change", () => commit(hex.value));
      value = el("div", { class: "value" }, pick, hex);
      controls = el(
        "div",
        { class: "controls" },
        el("span", { class: "meta" }, t("vars.bindTo")),
        el("button", { class: "small", onclick: () => bindColor(v.name, "fill") }, t("vars.bind.fill")),
        el("button", { class: "small", onclick: () => bindColor(v.name, "line") }, t("vars.bind.line")),
        el("button", { class: "small", onclick: () => bindColor(v.name, "font") }, t("vars.bind.font")),
        el("button", { class: "small", onclick: unbindColors, title: t("vars.unbindTip") }, t("vars.unbind"))
      );
    } else {
      const input = el("input", { type: "text", value: v.value || "", placeholder: t("vars.valuePh"), "aria-label": t("vars.valueOf", { name: v.name }) });
      input.addEventListener("input", () => {
        v.value = input.value;
        saveState();
        applyVariablesSoon();
      });
      value = el("div", { class: "value" }, input);
      controls = el(
        "div",
        { class: "controls" },
        el("button", { class: "small", onclick: () => insertPlaceholder(v.name), title: t("vars.insertTip") }, t("vars.insert", { name: v.name }))
      );
    }
    list.append(el("li", {}, head, value, controls));
  });
}

function renderComps() {
  const list = $("comp-list");
  list.replaceChildren();
  const names = Object.keys(state.components).sort();
  $("comp-empty").hidden = names.length > 0;
  // Form "New component / Variant of …"
  const of = $("comp-of");
  const prev = of.value;
  const bases = [...new Set(names.map((k) => C.splitKey(k).base))];
  of.replaceChildren(el("option", { value: "" }, t("new.component")), ...bases.map((b) => el("option", { value: b }, t("new.variantOf", { name: b }))));
  of.value = bases.includes(prev) ? prev : "";
  for (const name of names) {
    const comp = state.components[name];
    const info = ui.compInfo[name];
    let meta = "…";
    if (info) {
      meta = info.masterFound
        ? info.masterSlide
          ? t("comps.masterOn", { slide: info.masterSlide, n: info.instances })
          : t("comps.master", { n: info.instances })
        : t("comps.instances", { n: info.instances });
    }
    const opt = (key, label) => {
      const cb = el("input", {
        type: "checkbox",
        checked: comp[key] !== false,
        onchange: () => {
          comp[key] = cb.checked;
          saveState();
        },
      });
      return el("label", { class: "check meta" }, cb, label);
    };
    const confirmDel = ui.pendingDelete === name;
    list.append(
      el(
        "li",
        {},
        el(
          "div",
          { class: "head" },
          el("span", { class: "name" }, "◆ " + label(name)),
          info && !info.masterFound ? el("span", { class: "badge warn" }, t("comps.masterMissing")) : null
        ),
        el("div", { class: "meta" }, meta),
        el("div", { class: "controls" }, opt("syncText", t("comps.syncText")), opt("syncSize", t("comps.syncSize"))),
        el(
          "div",
          { class: "controls" },
          el("button", { class: "small primary", onclick: () => runSync([name]) }, t("comps.sync")),
          el("button", { class: "small", onclick: () => showMaster(name), disabled: info && !info.masterFound }, t("sel.showMaster")),
          el(
            "button",
            {
              class: "small danger",
              onclick: () => {
                if (confirmDel) {
                  ui.pendingDelete = null;
                  deleteComponent(name);
                } else {
                  ui.pendingDelete = name;
                  renderComps();
                  setTimeout(() => {
                    if (ui.pendingDelete === name) {
                      ui.pendingDelete = null;
                      renderComps();
                    }
                  }, 4000);
                }
              },
            },
            confirmDel ? t("comps.confirmDelete") : t("comps.delete")
          )
        )
      )
    );
  }
  renderSelection();
}

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
  document.querySelectorAll("[data-i18n-ph]").forEach((n) => (n.placeholder = t(n.dataset.i18nPh)));
  document.querySelectorAll("[data-i18n-aria]").forEach((n) => n.setAttribute("aria-label", t(n.dataset.i18nAria)));
  document.querySelectorAll(".lang button").forEach((b) => b.classList.toggle("active", b.dataset.lang === I.getLang()));
  updateCompNamePlaceholder();
}

function updateCompNamePlaceholder() {
  $("comp-name").placeholder = $("comp-of").value ? t("new.variantPh") : t("new.namePh");
}

function setLanguage(lang) {
  I.setLang(lang);
  lsSet("dc.lang", lang);
  applyStaticTexts();
  renderVars();
  renderComps();
  setStatus(t("ready"));
}

function showVersion() {
  const meta = document.querySelector('meta[name="dc-build"]');
  const build = meta && !meta.content.startsWith("__") ? meta.content : "dev";
  $("version").textContent = "v" + VERSION + " · " + build;
}

// ---------------------------------------------------------------- start

Office.onReady((info) => {
  I.setLang(I.detect(Office.context.displayLanguage, lsGet("dc.lang")));
  applyStaticTexts();
  showVersion();
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus(t("hostOnly"), "err");
    return;
  }
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    $("unsupported").hidden = false;
  }
  loadState();

  document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => switchTab(b.dataset.tab)));

  $("var-add").addEventListener("submit", (e) => {
    e.preventDefault();
    const name = $("var-name").value.trim();
    if (!C.isValidName(name)) return setStatus(t("err.invalidName"), "err");
    if (state.vars.some((v) => v.name === name)) return setStatus(t("err.exists", { name }), "err");
    const type = $("var-type").value;
    state.vars.push({ name, type, value: type === "color" ? "#C43E1C" : "" });
    $("var-name").value = "";
    saveState();
    renderVars();
    setStatus(type === "color" ? t("msg.colorAdded", { name }) : t("msg.varAdded", { name }), "ok");
  });

  $("show-raw").checked = state.showRaw;
  $("show-raw").addEventListener("change", () => {
    state.showRaw = $("show-raw").checked;
    saveState();
    runApplyVariables();
  });
  $("apply-vars").addEventListener("click", runApplyVariables);
  $("scan-vars").addEventListener("click", scanPlaceholders);

  $("comp-add").addEventListener("submit", (e) => {
    e.preventDefault();
    const base = $("comp-of").value;
    const name = $("comp-name").value.trim();
    createComponent(base ? C.joinKey(base, name) : name);
  });
  $("comp-of").addEventListener("change", updateCompNamePlaceholder);
  document.querySelectorAll(".lang button").forEach((b) => b.addEventListener("click", () => setLanguage(b.dataset.lang)));
  $("sync-all").addEventListener("click", () => runSync(null));
  $("refresh-comps").addEventListener("click", refreshCompInfo);

  $("auto-sync").checked = lsGet("dc.autoSync") === "1";
  $("auto-sync").addEventListener("change", () => lsSet("dc.autoSync", $("auto-sync").checked ? "1" : "0"));

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, debounce(refreshSelection, 250));

  renderVars();
  renderComps();
  switchTab(lsGet("dc.tab") === "comps" ? "comps" : "vars");
  setStatus(t("ready"));
});
