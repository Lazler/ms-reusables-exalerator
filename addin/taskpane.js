/* Deck Components – Seitenbereich (Office.js / PowerPoint) */
/* global Office, PowerPoint, DCCore */
"use strict";

const C = DCCore;
const STATE_KEY = "deckComponents.v1";
const TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder"];
const COMP_TYPES = ["GeometricShape", "TextBox"];
const TAG = { TPL: "DC_TPL", OUT: "DC_OUT", FILL: "DC_FILL", LINE: "DC_LINE", FONT: "DC_FONT" };

/** Persistiert in der .pptx (Document Settings). */
let state = { vars: [], components: {}, showRaw: false };
/** Nur zur Laufzeit. */
const ui = { compInfo: {}, selection: null, lastSelection: null, pendingDelete: null };

// ---------------------------------------------------------------- Hilfen

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

function errText(e) {
  if (!e) return "Unbekannter Fehler";
  let s = e.message || String(e);
  if (e.code) s += " [" + e.code + "]";
  if (e.debugInfo && e.debugInfo.errorLocation) s += "\nOrt: " + e.debugInfo.errorLocation;
  return s;
}

function debounce(fn, ms) {
  let t;
  return (...a) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}

/** Alle PowerPoint-Aufrufe nacheinander ausführen, Fehler im Statusbereich zeigen. */
let chain = Promise.resolve();
function queue(fn) {
  const p = chain.then(fn).catch((e) => {
    console.error(e);
    setStatus("Fehler: " + errText(e), "err");
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
        setStatus("Speichern im Dokument fehlgeschlagen: " + (r.error && r.error.message), "err");
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
    setStatus("Gespeicherter Zustand ist beschädigt: " + errText(e), "err");
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
    /* egal */
  }
}

// ---------------------------------------------------------------- PowerPoint-Zugriff

/** Lädt alle Formen aller Folien (ohne Gruppeninhalte). */
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
 * Lädt etwas für viele Elemente in einem Rutsch. Schlägt das fehl (z. B. weil eine
 * exotische Form kein Textfeld hat), wird Element für Element geladen und
 * fehlerhafte Elemente werden übersprungen.
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
        console.warn("Übersprungen:", it.name, e);
      }
    }
    return ok;
  }
}

async function getSelection(ctx) {
  const shapes = ctx.presentation.getSelectedShapes();
  shapes.load("items/id,items/name,items/type");
  const slides = ctx.presentation.getSelectedSlides();
  slides.load("items/id");
  await ctx.sync();
  const slideId = slides.items.length ? slides.items[0].id : null;
  return {
    slideId,
    shapes: shapes.items.map((sh) => ({ shape: sh, id: sh.id, name: sh.name, type: sh.type, slideId })),
  };
}

// ---------------------------------------------------------------- Variablen

const textVars = () => state.vars.filter((v) => v.type !== "color");
const colorOf = (name) => {
  const v = state.vars.find((x) => x.name === name && x.type === "color");
  return v ? C.normalizeColor(v.value) : null;
};

/** Rendert alle Textfelder und Farbbindungen im Deck. */
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

  const vars = textVars();
  const missing = new Set();
  let changed = 0;
  let fields = 0;

  for (const it of items) {
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
    // Farben
    const fill = colorOf(tagVal(it.tags[TAG.FILL]));
    if (fill) it.shape.fill.setSolidColor(fill);
    const line = colorOf(tagVal(it.tags[TAG.LINE]));
    if (line) {
      it.shape.lineFormat.visible = true;
      it.shape.lineFormat.color = line;
    }
    const font = colorOf(tagVal(it.tags[TAG.FONT]));
    if (font && it.tr) it.tr.font.color = font;
  }
  await ctx.sync();
  return { changed, fields, missing: [...missing] };
}

function runApplyVariables() {
  return queue(async () => {
    setStatus("Aktualisiere …");
    const r = await PowerPoint.run(applyVariables);
    let msg = `${r.fields} Textfeld(er) mit Variablen, ${r.changed} geändert.`;
    if (state.showRaw) msg += " (Platzhalter-Ansicht)";
    if (r.missing.length) {
      setStatus(msg + "\nUnbekannte Platzhalter: " + r.missing.join(", ") + " – unten anlegen oder „Platzhalter im Deck finden“.", "warn");
    } else setStatus(msg, "ok");
  });
}
const applyVariablesSoon = debounce(runApplyVariables, 700);

function scanPlaceholders() {
  return queue(async () => {
    const names = await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      const items = await loadEach(ctx, all.filter((s) => TEXT_TYPES.includes(s.type)), (it) => {
        it.tpl = loadTag(it.shape, TAG.TPL);
        it.tr = it.shape.textFrame.textRange;
        it.tr.load("text");
      });
      const found = new Set();
      for (const it of items) {
        C.extractNames(tagVal(it.tpl) || "").forEach((n) => found.add(n));
        C.extractNames(it.tr.text).forEach((n) => found.add(n));
      }
      return [...found];
    });
    const known = new Set(state.vars.map((v) => v.name));
    const fresh = names.filter((n) => !known.has(n));
    fresh.forEach((n) => state.vars.push({ name: n, type: "text", value: "" }));
    if (fresh.length) {
      await saveState();
      renderVars();
      setStatus("Neu angelegt: " + fresh.join(", ") + ". Werte eintragen.", "ok");
    } else {
      setStatus(names.length ? "Alle Platzhalter sind bereits angelegt (" + names.length + ")." : "Keine Platzhalter im Deck gefunden.", "ok");
    }
  });
}

function insertPlaceholder(name) {
  return queue(async () => {
    await PowerPoint.run(async (ctx) => {
      const tr = ctx.presentation.getSelectedTextRangeOrNullObject();
      await ctx.sync();
      if (tr.isNullObject) throw new Error("Bitte zuerst in ein Textfeld klicken (Cursor setzen).");
      tr.text = "{{" + name + "}}";
      await ctx.sync();
    });
    setStatus("{{" + name + "}} eingefügt.", "ok");
  }).then(() => runApplyVariables());
}

function bindColor(name, kind) {
  const key = { fill: TAG.FILL, line: TAG.LINE, font: TAG.FONT }[kind];
  return queue(async () => {
    const n = await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      if (!sel.shapes.length) throw new Error("Bitte zuerst eine oder mehrere Formen markieren.");
      sel.shapes.forEach((s) => s.shape.tags.add(key, name));
      await ctx.sync();
      return sel.shapes.length;
    });
    setStatus(`${n} Form(en) an „${name}“ gebunden.`, "ok");
  }).then(() => runApplyVariables());
}

function unbindColors() {
  return queue(async () => {
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
    setStatus(n ? `${n} Farbbindung(en) gelöst.` : "Auswahl hat keine Farbbindungen.", "ok");
  });
}

// ---------------------------------------------------------------- Komponenten

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
    await ctx.sync();
    const source = C.resolveTemplate(tr.text, tagVal(tpl), tagVal(out), textVars()) || tr.text;
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
    console.warn("Textformat nicht lesbar", e);
  }
  return st;
}

const isSet = (v) => v !== null && v !== undefined;

function applyStyle(inst, st, comp) {
  const sh = inst.shape;
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
        const out = state.showRaw ? t.source : C.render(t.source, textVars()).text;
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
  // Größe zuletzt, damit AutoSize sie nicht direkt wieder überschreibt
  if (comp.syncSize !== false) {
    if (isSet(st.width)) sh.width = st.width;
    if (isSet(st.height)) sh.height = st.height;
  }
}

/** Synchronisiert die angegebenen (oder alle) Komponenten. Gibt Statistik zurück. */
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
      result.warnings.push(`„${name}“: Master nicht gefunden – letzter bekannter Stand verwendet. Wähle eine Instanz und klicke „Zum Master machen“.`);
    } else {
      result.warnings.push(`„${name}“: Master nicht gefunden.`);
      continue;
    }

    const targets = instances.filter((s) => COMP_TYPES.includes(s.type) || s.type === "Placeholder");
    const skipped = instances.length - targets.length;
    if (skipped) result.warnings.push(`„${name}“: ${skipped} Instanz(en) mit nicht unterstütztem Typ übersprungen (Gruppe/Bild/Tabelle).`);

    const loaded = await loadEach(ctx, targets, (it) => {
      it.tpl = loadTag(it.shape, TAG.TPL);
      it.out = loadTag(it.shape, TAG.OUT);
      it.tr = st.text ? it.shape.textFrame.textRange : null;
    });

    loaded.forEach((it) => applyStyle(it, st, comp));
    try {
      await ctx.sync();
      result.synced += loaded.length;
    } catch (e) {
      // Einzeln wiederholen, um die fehlerhafte Instanz zu finden
      for (const it of loaded) {
        try {
          applyStyle(it, st, comp);
          await ctx.sync();
          result.synced++;
        } catch (e2) {
          result.warnings.push(`„${name}“ auf Folie ${it.slideIndex + 1}: ${errText(e2)}`);
        }
      }
    }
  }
  return result;
}

function runSync(names) {
  return queue(async () => {
    setStatus("Synchronisiere …");
    const r = await PowerPoint.run((ctx) => syncComponents(ctx, names));
    await saveState();
    renderComps();
    const msg = `${r.synced} Instanz(en) synchronisiert.`;
    setStatus(r.warnings.length ? msg + "\n" + r.warnings.join("\n") : msg, r.warnings.length ? "warn" : "ok");
  });
}

function refreshCompInfo() {
  return queue(async () => {
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

function createComponent(name) {
  name = (name || "").trim();
  if (!C.isValidName(name)) return setStatus("Ungültiger Name. Erlaubt: Buchstaben, Ziffern, Leerzeichen, _ . -", "err");
  if (state.components[name]) return setStatus(`„${name}“ existiert bereits.`, "err");
  return queue(async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      if (sel.shapes.length !== 1) throw new Error("Bitte genau eine Form markieren.");
      const s = sel.shapes[0];
      if (!COMP_TYPES.includes(s.type)) {
        throw new Error(`Typ „${s.type}“ wird im Prototyp nicht unterstützt. Nur Formen und Textfelder.`);
      }
      const existing = C.componentNameOf(s.name);
      if (existing && state.components[existing]) throw new Error(`Die Form gehört schon zu „${existing}“.`);
      s.shape.name = C.componentShapeName(name);
      await ctx.sync();
      const snapshot = await readStyle(ctx, s.shape, s.type);
      state.components[name] = { master: { slideId: sel.slideId, shapeId: s.id }, snapshot, syncText: true, syncSize: true };
      ui.compInfo[name] = { instances: 0, masterFound: true, masterSlide: null };
    });
    await saveState();
    $("comp-name").value = "";
    renderComps();
    setStatus(`Komponente „${name}“ angelegt. Jetzt den Master kopieren und auf anderen Folien einfügen.`, "ok");
  }).then(() => {
    refreshCompInfo();
    refreshSelection();
  });
}

function linkSelection(name) {
  return queue(async () => {
    const n = await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      const comp = state.components[name];
      const targets = sel.shapes.filter(
        (s) => !(s.slideId === comp.master.slideId && s.id === comp.master.shapeId)
      );
      if (!targets.length) throw new Error("Bitte die Form(en) markieren, die Instanz werden sollen.");
      targets.forEach((s) => (s.shape.name = C.componentShapeName(name)));
      await ctx.sync();
      return targets.length;
    });
    setStatus(`${n} Form(en) mit „${name}“ verknüpft.`, "ok");
  }).then(() => runSync([name]).then(refreshSelection));
}

function detachSelection() {
  return queue(async () => {
    await PowerPoint.run(async (ctx) => {
      const sel = await getSelection(ctx);
      sel.shapes.forEach((s) => {
        const comp = state.components[C.componentNameOf(s.name)];
        if (comp && comp.master.slideId === s.slideId && comp.master.shapeId === s.id) {
          throw new Error("Der Master kann nicht gelöst werden. Komponente stattdessen löschen oder anderen Master wählen.");
        }
      });
      sel.shapes.forEach((s) => (s.shape.name = C.stripComponentName(s.name)));
      await ctx.sync();
    });
    setStatus("Gelöst. Die Form wird nicht mehr synchronisiert.", "ok");
  }).then(() => {
    refreshCompInfo();
    refreshSelection();
  });
}

function makeMaster(name, s) {
  state.components[name].master = { slideId: s.slideId, shapeId: s.id };
  return saveState().then(() => runSync([name])).then(refreshSelection);
}

function deleteComponent(name) {
  return queue(async () => {
    await PowerPoint.run(async (ctx) => {
      const all = await loadShapes(ctx);
      all.filter((s) => C.componentNameOf(s.name) === name).forEach((s) => (s.shape.name = C.stripComponentName(s.name)));
      await ctx.sync();
    });
    delete state.components[name];
    delete ui.compInfo[name];
    await saveState();
    renderComps();
    setStatus(`Komponente „${name}“ gelöscht. Die Formen bleiben unverändert auf den Folien.`, "ok");
  });
}

function showMaster(name) {
  const m = state.components[name].master;
  return queue(async () => {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.setSelectedSlides([m.slideId]);
      await ctx.sync();
      ctx.presentation.slides.getItem(m.slideId).setSelectedShapes([m.shapeId]);
      await ctx.sync();
    });
  });
}

// ---------------------------------------------------------------- Auswahl

function isMasterOf(s) {
  const name = C.componentNameOf(s.name);
  const comp = name && state.components[name];
  return !!comp && comp.master.slideId === s.slideId && comp.master.shapeId === s.id;
}

function refreshSelection() {
  return queue(async () => {
    const sel = await PowerPoint.run(getSelection);
    ui.lastSelection = ui.selection;
    ui.selection = sel;
    renderSelection();

    // Auto-Sync: Auswahl hat einen Master verlassen
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
  return el("select", { "aria-label": "Komponente" }, Object.keys(state.components).map((n) => el("option", { value: n }, n)));
}

function renderSelection() {
  const info = $("sel-info");
  const act = $("sel-actions");
  act.replaceChildren();
  const sel = ui.selection;
  if (!sel || !sel.shapes.length) {
    info.textContent = "Nichts ausgewählt.";
    return;
  }
  const hasComps = Object.keys(state.components).length > 0;
  if (sel.shapes.length > 1) {
    info.textContent = sel.shapes.length + " Formen ausgewählt.";
    if (hasComps) {
      const s = compSelect();
      act.append(el("div", { class: "row" }, s, el("button", { onclick: () => linkSelection(s.value) }, "Alle als Instanz verknüpfen")));
    }
    act.append(el("div", { class: "row" }, el("button", { onclick: detachSelection }, "Alle lösen")));
    return;
  }
  const s = sel.shapes[0];
  const cname = C.componentNameOf(s.name);
  const comp = cname && state.components[cname];
  if (comp && isMasterOf(s)) {
    info.replaceChildren(el("b", {}, "Master"), " von „" + cname + "“");
    act.append(el("div", { class: "row" }, el("button", { class: "primary", onclick: () => runSync([cname]) }, "Instanzen aktualisieren")));
  } else if (comp) {
    info.replaceChildren(el("b", {}, "Instanz"), " von „" + cname + "“");
    act.append(
      el(
        "div",
        { class: "row" },
        el("button", { onclick: () => showMaster(cname) }, "Master zeigen"),
        el("button", { onclick: () => makeMaster(cname, s) }, "Zum Master machen"),
        el("button", { onclick: detachSelection }, "Lösen")
      )
    );
  } else if (cname) {
    info.textContent = `„${s.name}“ – Komponente „${cname}“ ist in diesem Deck nicht registriert.`;
    act.append(
      el("div", { class: "row" }, el("button", { onclick: () => createComponent(cname) }, "Als Master registrieren"))
    );
  } else {
    info.textContent = `„${s.name}“ (${s.type})`;
    if (hasComps) {
      const sl = compSelect();
      act.append(el("div", { class: "row" }, sl, el("button", { onclick: () => linkSelection(sl.value) }, "Als Instanz verknüpfen")));
    }
  }
}

// ---------------------------------------------------------------- Rendering

function renderVars() {
  const list = $("var-list");
  list.replaceChildren();
  $("var-empty").hidden = state.vars.length > 0;
  state.vars.forEach((v, idx) => {
    const remove = el("button", {
      class: "small danger",
      title: "Variable löschen",
      onclick: () => {
        state.vars.splice(idx, 1);
        saveState();
        renderVars();
        runApplyVariables();
      },
    }, "✕");

    const head = el("div", { class: "head" }, el("span", { class: "name" }, v.name), el("span", { class: "badge" }, v.type === "color" ? "Farbe" : "Text"), remove);

    let value;
    let controls;
    if (v.type === "color") {
      const hex = el("input", { type: "text", value: v.value || "", placeholder: "#RRGGBB", "aria-label": "Hex-Wert" });
      const pick = el("input", { type: "color", value: C.normalizeColor(v.value) || "#000000", "aria-label": "Farbe wählen" });
      const commit = (val) => {
        const c = C.normalizeColor(val);
        if (!c) return setStatus("Ungültige Farbe: " + val, "err");
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
        el("span", { class: "meta" }, "Auswahl binden:"),
        el("button", { class: "small", onclick: () => bindColor(v.name, "fill") }, "Füllung"),
        el("button", { class: "small", onclick: () => bindColor(v.name, "line") }, "Linie"),
        el("button", { class: "small", onclick: () => bindColor(v.name, "font") }, "Schrift"),
        el("button", { class: "small", onclick: unbindColors, title: "Alle Farbbindungen der Auswahl lösen" }, "Lösen")
      );
    } else {
      const input = el("input", { type: "text", value: v.value || "", placeholder: "Wert", "aria-label": "Wert von " + v.name });
      input.addEventListener("input", () => {
        v.value = input.value;
        saveState();
        applyVariablesSoon();
      });
      value = el("div", { class: "value" }, input);
      controls = el(
        "div",
        { class: "controls" },
        el("button", { class: "small", onclick: () => insertPlaceholder(v.name), title: "An der Cursorposition einfügen" }, "{{" + v.name + "}} einfügen")
      );
    }
    list.append(el("li", {}, head, value, controls));
  });
}

function renderComps() {
  const list = $("comp-list");
  list.replaceChildren();
  const names = Object.keys(state.components);
  $("comp-empty").hidden = names.length > 0;
  for (const name of names) {
    const comp = state.components[name];
    const info = ui.compInfo[name];
    let meta = "…";
    if (info) {
      meta = info.masterFound
        ? `Master${info.masterSlide ? " auf Folie " + info.masterSlide : ""} · ${info.instances} Instanz(en)`
        : `${info.instances} Instanz(en)`;
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
          el("span", { class: "name" }, "◆ " + name),
          info && !info.masterFound ? el("span", { class: "badge warn" }, "Master fehlt") : null
        ),
        el("div", { class: "meta" }, meta),
        el("div", { class: "controls" }, opt("syncText", "Text übernehmen"), opt("syncSize", "Größe übernehmen")),
        el(
          "div",
          { class: "controls" },
          el("button", { class: "small primary", onclick: () => runSync([name]) }, "Synchronisieren"),
          el("button", { class: "small", onclick: () => showMaster(name), disabled: info && !info.masterFound }, "Master zeigen"),
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
            confirmDel ? "Wirklich löschen?" : "Löschen"
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

// ---------------------------------------------------------------- Start

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Dieses Add-in läuft nur in PowerPoint.", "err");
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
    if (!C.isValidName(name)) return setStatus("Ungültiger Name. Erlaubt: Buchstaben, Ziffern, Leerzeichen, _ . -", "err");
    if (state.vars.some((v) => v.name === name)) return setStatus(`„${name}“ existiert bereits.`, "err");
    const type = $("var-type").value;
    state.vars.push({ name, type, value: type === "color" ? "#C43E1C" : "" });
    $("var-name").value = "";
    saveState();
    renderVars();
    setStatus(type === "color" ? `Farbe „${name}“ angelegt. Formen markieren und binden.` : `Jetzt {{${name}}} in ein Textfeld schreiben oder über den Button einfügen.`, "ok");
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
    createComponent($("comp-name").value);
  });
  $("sync-all").addEventListener("click", () => runSync(null));
  $("refresh-comps").addEventListener("click", refreshCompInfo);

  $("auto-sync").checked = lsGet("dc.autoSync") === "1";
  $("auto-sync").addEventListener("change", () => lsSet("dc.autoSync", $("auto-sync").checked ? "1" : "0"));

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, debounce(refreshSelection, 250));

  renderVars();
  renderComps();
  switchTab(lsGet("dc.tab") === "comps" ? "comps" : "vars");
  setStatus("Bereit.");
});
