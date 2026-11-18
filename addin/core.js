/*
 * Deck Components – reine Logik ohne Office-Abhängigkeit.
 * Läuft im Browser (window.DCCore) und in Node (für Tests).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DCCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // {{Name}} – Name darf Buchstaben, Ziffern, Leerzeichen, _ . - enthalten
  const VAR_RE = /\{\{\s*([\p{L}\p{N}_.\- ]+?)\s*\}\}/gu;
  const VAR_TEST_RE = /\{\{\s*([\p{L}\p{N}_.\- ]+?)\s*\}\}/u; // ohne g: kein lastIndex-Zustand
  const NAME_RE = /^[\p{L}\p{N}_.\- ]+$/u;

  // Komponenten-Namenskonvention im Auswahlbereich: "◆ Button ◆"
  const COMP_MARK = "◆";
  const COMP_RE = /^\s*◆\s*(.+?)\s*◆/u;

  function isValidName(name) {
    return typeof name === "string" && name.trim().length > 0 && NAME_RE.test(name.trim());
  }

  function hasPlaceholders(text) {
    return !!text && VAR_TEST_RE.test(text);
  }

  function extractNames(text) {
    const out = [];
    if (!text) return out;
    for (const m of text.matchAll(VAR_RE)) {
      if (!out.includes(m[1])) out.push(m[1]);
    }
    return out;
  }

  function varMap(vars) {
    const map = new Map();
    for (const v of vars || []) map.set(v.name, v.value == null ? "" : String(v.value));
    return map;
  }

  /** Ersetzt {{Name}} durch Werte. Unbekannte Namen bleiben stehen. */
  function render(tpl, vars) {
    const map = varMap(vars);
    const missing = [];
    const text = (tpl || "").replace(VAR_RE, (whole, name) => {
      if (map.has(name)) return map.get(name);
      if (!missing.includes(name)) missing.push(name);
      return whole;
    });
    return { text, missing };
  }

  /** PowerPoint liefert \r bzw. \v statt \n – nur für Vergleiche normalisieren. */
  function normalize(t) {
    return (t == null ? "" : String(t)).replace(/\r\n|\r|\v|\u000b/g, "\n").replace(/\s+$/, "");
  }

  function sameText(a, b) {
    return normalize(a) === normalize(b);
  }

  /**
   * Wurde ein gerenderter Text vom Nutzer bearbeitet, versuchen wir, die
   * Werte der Variablen, die vorher im Template standen, wieder in
   * Platzhalter zurückzuverwandeln ("Hallo ACME" -> "Hallo {{Kunde}}").
   */
  function reabsorb(current, oldTpl, vars) {
    const map = varMap(vars);
    const names = extractNames(oldTpl)
      .filter((n) => map.has(n) && map.get(n).length >= 2)
      .sort((a, b) => map.get(b).length - map.get(a).length);
    let text = current || "";
    // Erst Platzhalter maskieren, damit wir sie nicht doppelt ersetzen
    const tokens = [];
    const mask = (s) =>
      s.replace(VAR_RE, (w) => {
        tokens.push(w);
        return "\u0000" + (tokens.length - 1) + "\u0000";
      });
    text = mask(text);
    for (const n of names) {
      const val = map.get(n);
      if (text.includes(val)) {
        text = text.split(val).join("{{" + n + "}}");
        text = mask(text);
      }
    }
    return text.replace(/\u0000(\d+)\u0000/g, (_, i) => tokens[Number(i)]);
  }

  /**
   * Entscheidet, welches Template für ein Textfeld gilt.
   * @param current  aktueller Text in PowerPoint
   * @param tpl      gespeichertes Template (Tag) oder null
   * @param lastOut  zuletzt von uns geschriebener Text (Tag) oder null
   * @returns Template-String oder null (= Feld hat keine Variablen mehr)
   */
  /**
   * Leitet aus Template + zuletzt geschriebenem Text die damals eingesetzten
   * Werte ab ("Hallo {{Kunde}}" + "Hallo ACME" -> {Kunde: "ACME"}).
   */
  function extractValues(tpl, out) {
    if (tpl == null || out == null) return null;
    const parts = normalize(tpl).split(new RegExp(VAR_RE.source, "u"));
    // parts: [lit, name, lit, name, lit, ...]
    const esc = (t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    let src = "^";
    const names = [];
    for (let i = 0; i < parts.length; i++) {
      if (i % 2 === 0) src += esc(parts[i]);
      else {
        src += "([\\s\\S]*?)";
        names.push(parts[i]);
      }
    }
    const m = new RegExp(src + "$", "u").exec(normalize(out));
    if (!m) return null;
    const vals = [];
    names.forEach((n, i) => {
      if (!vals.some((v) => v.name === n)) vals.push({ name: n, value: m[i + 1] });
    });
    return vals;
  }

  function resolveTemplate(current, tpl, lastOut, vars) {
    if (tpl) {
      if (lastOut == null || sameText(current, lastOut)) return tpl;
      // Nutzer hat den Text geändert. Zurückverwandeln mit den Werten, die
      // damals eingesetzt waren (nicht den aktuellen – die können neu sein).
      const before = extractValues(tpl, lastOut);
      let candidate = before ? reabsorb(current, tpl, before) : current;
      candidate = reabsorb(candidate, tpl, vars);
      return hasPlaceholders(candidate) ? candidate : null;
    }
    return hasPlaceholders(current) ? current : null;
  }

  function normalizeColor(input) {
    if (typeof input !== "string") return null;
    let s = input.trim().replace(/^#/, "");
    if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map((c) => c + c).join("");
    if (!/^[0-9a-f]{6}$/i.test(s)) return null;
    return "#" + s.toUpperCase();
  }

  function componentNameOf(shapeName) {
    if (!shapeName) return null;
    const m = COMP_RE.exec(shapeName);
    return m ? m[1] : null;
  }

  function componentShapeName(compName) {
    return COMP_MARK + " " + compName.trim() + " " + COMP_MARK;
  }

  function stripComponentName(shapeName) {
    const n = componentNameOf(shapeName);
    return n ? n : shapeName;
  }

  return {
    VAR_RE,
    isValidName,
    hasPlaceholders,
    extractNames,
    render,
    normalize,
    sameText,
    reabsorb,
    resolveTemplate,
    extractValues,
    normalizeColor,
    componentNameOf,
    componentShapeName,
    stripComponentName,
  };
});
