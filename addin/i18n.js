/*
 * Deck Components – UI strings.
 * English is the primary language; German is a secondary translation.
 * Missing German keys fall back to English.
 */
(function (root) {
  "use strict";

  const STRINGS = {
    en: {
      // header / tabs
      "tab.vars": "Variables",
      "tab.comps": "Components",
      "unsupported": "This PowerPoint version does not support the required API (PowerPointApi 1.5). Please use PowerPoint on the web or a current Microsoft 365 version.",
      "ready": "Ready.",
      "hostOnly": "This add-in only runs in PowerPoint.",

      // variables tab
      "vars.title": "Variables",
      "vars.hint": "Type {{Name}} in any text box. The value is filled in everywhere in the deck.",
      "vars.empty": "No variables yet. Add one below, e.g. Client, Project or Date.",
      "vars.namePh": "Name, e.g. Client",
      "vars.type": "Type",
      "vars.type.text": "Text",
      "vars.type.color": "Color",
      "vars.add": "Add",
      "vars.view": "View",
      "vars.showRaw": "Show placeholders instead of values (for editing)",
      "vars.apply": "Update deck",
      "vars.scan": "Find placeholders in deck",
      "vars.delete": "Delete variable",
      "vars.valuePh": "Value",
      "vars.valueOf": "Value of {name}",
      "vars.insert": "Insert {{{name}}}",
      "vars.insertTip": "Insert at the cursor position",
      "vars.hex": "Hex value",
      "vars.pick": "Pick color",
      "vars.bindTo": "Apply to selection:",
      "vars.bind.fill": "Fill",
      "vars.bind.line": "Line",
      "vars.bind.font": "Text",
      "vars.unbind": "Unbind",
      "vars.unbindTip": "Remove all color bindings from the selection",

      // variables messages
      "msg.updating": "Updating…",
      "msg.applied": "{fields} text box(es) with variables, {changed} changed.",
      "msg.rawView": "(placeholder view)",
      "msg.missing": "Unknown placeholders: {names}. Add them below or use “Find placeholders in deck”.",
      "msg.scanNew": "Added: {names}. Please fill in the values.",
      "msg.scanAll": "All placeholders already exist ({n}).",
      "msg.scanNone": "No placeholders found in the deck.",
      "msg.inserted": "{{{name}}} inserted.",
      "msg.bound": "{n} shape(s) bound to “{name}”.",
      "msg.unbound": "{n} color binding(s) removed.",
      "msg.noBindings": "The selection has no color bindings.",
      "msg.varAdded": "Now type {{{name}}} in a text box or use the insert button.",
      "msg.colorAdded": "Color “{name}” added. Select shapes and apply it.",
      "err.invalidName": "Invalid name. Allowed: letters, digits, spaces, _ . -",
      "err.exists": "“{name}” already exists.",
      "err.invalidColor": "Invalid color: {value}",
      "err.noCursor": "Click into a text box first (place the cursor).",
      "err.noSelection": "Select one or more shapes first.",

      // components tab
      "sel.title": "Selection",
      "sel.none": "Nothing selected.",
      "sel.unreadable": "Could not read the selection. Click directly on a shape on the slide once.",
      "sel.multi": "{n} shapes selected.",
      "sel.linkAll": "Link all as instances",
      "sel.detachAll": "Detach all",
      "sel.masterOf": "Master of “{name}”",
      "sel.instanceOf": "Instance of “{name}”",
      "sel.updateInstances": "Update instances",
      "sel.showMaster": "Show master",
      "sel.makeMaster": "Make master",
      "sel.detach": "Detach",
      "sel.variant": "Variant:",
      "sel.unregistered": "“{shape}” – component “{name}” is not registered in this deck.",
      "sel.register": "Register as master",
      "sel.plain": "“{shape}” ({type})",
      "sel.link": "Link as instance",
      "comp.select": "Component",

      "new.title": "New component / variant",
      "new.hint": "Select exactly one shape or text box. It becomes the master. For a variant (e.g. a red button), select a modified copy and choose “Variant of …”.",
      "new.kind": "Kind",
      "new.component": "New component",
      "new.variantOf": "Variant of {name}",
      "new.namePh": "Name, e.g. Button",
      "new.variantPh": "Variant, e.g. Secondary",
      "new.create": "Create",

      "comps.title": "Components",
      "comps.hint": "To create instances, copy the master (Ctrl/Cmd + C, V or D) and paste it on any slide. Every shape named ◆ Name ◆ in the Selection Pane is an instance.",
      "comps.empty": "No components yet.",
      "comps.syncAll": "Sync all",
      "comps.recount": "Recount",
      "comps.autoSync": "Auto-sync when the selection leaves a master",
      "comps.masterOn": "Master on slide {slide} · {n} instance(s)",
      "comps.master": "Master · {n} instance(s)",
      "comps.instances": "{n} instance(s)",
      "comps.masterMissing": "Master missing",
      "comps.syncText": "Sync text",
      "comps.syncSize": "Sync size",
      "comps.sync": "Sync",
      "comps.delete": "Delete",
      "comps.confirmDelete": "Really delete?",
      "variant.default": "Default",

      // overrides
      "ovr.defaults": "Component defaults",
      "ovr.instance": "Overrides for this instance",
      "ovr.tip": "Tip: write e.g. {{Label}} in the master’s text and sync. Then the text can be changed per instance.",
      "ovr.empty": "(empty)",
      "ovr.src.global": "global: {value}",
      "ovr.src.local": "this component only",
      "ovr.src.own": "own value",
      "ovr.src.master": "from master",
      "ovr.src.globalShort": "global",
      "ovr.src.none": "empty",
      "ovr.fill": "Fill",
      "ovr.font": "Text color",
      "ovr.line": "Line",
      "ovr.ownColor": "own color",
      "ovr.fromMaster": "from master",
      "ovr.reset": "Reset",
      "ovr.resetAll": "Reset all overrides",

      // component messages
      "msg.syncing": "Syncing…",
      "msg.synced": "{n} instance(s) synced.",
      "warn.masterSnapshot": "“{name}”: master not found – using last known state. Select an instance and click “Make master”.",
      "warn.masterMissing": "“{name}”: master not found.",
      "warn.skipped": "“{name}”: skipped {n} instance(s) of an unsupported type (group/picture/table).",
      "warn.slide": "“{name}” on slide {slide}: {error}",
      "msg.compCreated": "Component “{name}” created. Now copy the master and paste it on other slides.",
      "msg.variantCreated": "Variant “{name}” created. Instances can now be switched to it.",
      "msg.linked": "{n} shape(s) linked to “{name}”.",
      "msg.detached": "Detached. The shape is no longer synced.",
      "msg.compDeleted": "Component “{name}” deleted. The shapes stay unchanged on the slides.",
      "err.exactlyOne": "Select exactly one shape.",
      "err.typeUnsupported": "Type “{type}” is not supported yet. Only shapes and text boxes.",
      "err.alreadyMaster": "This shape is already the master of “{name}”.",
      "err.selectTargets": "Select the shape(s) that should become instances.",
      "err.detachMaster": "The master cannot be detached. Delete the component or choose another master instead.",

      // errors
      "err.prefix": "Error",
      "err.during": "Error during “{action}”",
      "err.unknown": "Unknown error",
      "err.location": "Location",
      "err.save": "Saving to the document failed: {error}",
      "err.state": "Stored state is corrupted: {error}",

      // action labels (used in error messages)
      "act.apply": "Update deck",
      "act.scan": "Find placeholders",
      "act.insert": "Insert placeholder",
      "act.bind": "Apply color",
      "act.unbind": "Unbind color",
      "act.sync": "Sync",
      "act.count": "Count components",
      "act.create": "Create component",
      "act.link": "Link as instance",
      "act.detach": "Detach",
      "act.override": "Save override",
      "act.variant": "Switch variant",
      "act.delete": "Delete component",
      "act.showMaster": "Show master",
      "act.selection": "Read selection",
    },

    de: {
      "tab.vars": "Variablen",
      "tab.comps": "Komponenten",
      "unsupported": "Diese PowerPoint-Version unterstützt die benötigte API (PowerPointApi 1.5) nicht. Bitte PowerPoint im Web oder eine aktuelle Microsoft-365-Version nutzen.",
      "ready": "Bereit.",
      "hostOnly": "Dieses Add-in läuft nur in PowerPoint.",

      "vars.title": "Variablen",
      "vars.hint": "Schreib {{Name}} in ein beliebiges Textfeld. Der Wert wird überall im Deck eingesetzt.",
      "vars.empty": "Noch keine Variablen. Unten anlegen, z. B. Kunde, Projekt oder Datum.",
      "vars.namePh": "Name, z. B. Kunde",
      "vars.type": "Typ",
      "vars.type.text": "Text",
      "vars.type.color": "Farbe",
      "vars.add": "Hinzufügen",
      "vars.view": "Ansicht",
      "vars.showRaw": "Platzhalter statt Werte anzeigen (zum Bearbeiten)",
      "vars.apply": "Deck aktualisieren",
      "vars.scan": "Platzhalter im Deck finden",
      "vars.delete": "Variable löschen",
      "vars.valuePh": "Wert",
      "vars.valueOf": "Wert von {name}",
      "vars.insert": "{{{name}}} einfügen",
      "vars.insertTip": "An der Cursorposition einfügen",
      "vars.hex": "Hex-Wert",
      "vars.pick": "Farbe wählen",
      "vars.bindTo": "Auf Auswahl anwenden:",
      "vars.bind.fill": "Füllung",
      "vars.bind.line": "Linie",
      "vars.bind.font": "Schrift",
      "vars.unbind": "Lösen",
      "vars.unbindTip": "Alle Farbbindungen der Auswahl lösen",

      "msg.updating": "Aktualisiere …",
      "msg.applied": "{fields} Textfeld(er) mit Variablen, {changed} geändert.",
      "msg.rawView": "(Platzhalter-Ansicht)",
      "msg.missing": "Unbekannte Platzhalter: {names}. Unten anlegen oder „Platzhalter im Deck finden“.",
      "msg.scanNew": "Neu angelegt: {names}. Bitte Werte eintragen.",
      "msg.scanAll": "Alle Platzhalter sind bereits angelegt ({n}).",
      "msg.scanNone": "Keine Platzhalter im Deck gefunden.",
      "msg.inserted": "{{{name}}} eingefügt.",
      "msg.bound": "{n} Form(en) an „{name}“ gebunden.",
      "msg.unbound": "{n} Farbbindung(en) gelöst.",
      "msg.noBindings": "Die Auswahl hat keine Farbbindungen.",
      "msg.varAdded": "Jetzt {{{name}}} in ein Textfeld schreiben oder über den Button einfügen.",
      "msg.colorAdded": "Farbe „{name}“ angelegt. Formen markieren und anwenden.",
      "err.invalidName": "Ungültiger Name. Erlaubt: Buchstaben, Ziffern, Leerzeichen, _ . -",
      "err.exists": "„{name}“ existiert bereits.",
      "err.invalidColor": "Ungültige Farbe: {value}",
      "err.noCursor": "Bitte zuerst in ein Textfeld klicken (Cursor setzen).",
      "err.noSelection": "Bitte zuerst eine oder mehrere Formen markieren.",

      "sel.title": "Auswahl",
      "sel.none": "Nichts ausgewählt.",
      "sel.unreadable": "Auswahl konnte nicht gelesen werden. Klicke einmal direkt auf eine Form auf der Folie.",
      "sel.multi": "{n} Formen ausgewählt.",
      "sel.linkAll": "Alle als Instanz verknüpfen",
      "sel.detachAll": "Alle lösen",
      "sel.masterOf": "Master von „{name}“",
      "sel.instanceOf": "Instanz von „{name}“",
      "sel.updateInstances": "Instanzen aktualisieren",
      "sel.showMaster": "Master zeigen",
      "sel.makeMaster": "Zum Master machen",
      "sel.detach": "Lösen",
      "sel.variant": "Variante:",
      "sel.unregistered": "„{shape}“ – Komponente „{name}“ ist in diesem Deck nicht registriert.",
      "sel.register": "Als Master registrieren",
      "sel.plain": "„{shape}“ ({type})",
      "sel.link": "Als Instanz verknüpfen",
      "comp.select": "Komponente",

      "new.title": "Neue Komponente / Variante",
      "new.hint": "Markiere genau eine Form oder ein Textfeld. Sie wird zum Master. Für eine Variante (z. B. Button in Rot) eine abgewandelte Kopie markieren und „Variante von …“ wählen.",
      "new.kind": "Art",
      "new.component": "Neue Komponente",
      "new.variantOf": "Variante von {name}",
      "new.namePh": "Name, z. B. Button",
      "new.variantPh": "Variante, z. B. Sekundär",
      "new.create": "Anlegen",

      "comps.title": "Komponenten",
      "comps.hint": "Instanzen anlegen: Master kopieren (Strg/Cmd + C, V oder D) und auf beliebigen Folien einfügen. Jede Form mit dem Namen ◆ Name ◆ im Auswahlbereich ist eine Instanz.",
      "comps.empty": "Noch keine Komponenten.",
      "comps.syncAll": "Alle synchronisieren",
      "comps.recount": "Neu zählen",
      "comps.autoSync": "Auto-Sync, wenn die Auswahl einen Master verlässt",
      "comps.masterOn": "Master auf Folie {slide} · {n} Instanz(en)",
      "comps.master": "Master · {n} Instanz(en)",
      "comps.instances": "{n} Instanz(en)",
      "comps.masterMissing": "Master fehlt",
      "comps.syncText": "Text übernehmen",
      "comps.syncSize": "Größe übernehmen",
      "comps.sync": "Synchronisieren",
      "comps.delete": "Löschen",
      "comps.confirmDelete": "Wirklich löschen?",
      "variant.default": "Standard",

      "ovr.defaults": "Standardwerte der Komponente",
      "ovr.instance": "Anpassungen dieser Instanz",
      "ovr.tip": "Tipp: Schreib z. B. {{Label}} in den Text des Masters und synchronisiere. Dann lässt sich der Text pro Instanz anpassen.",
      "ovr.empty": "(leer)",
      "ovr.src.global": "global: {value}",
      "ovr.src.local": "nur in dieser Komponente",
      "ovr.src.own": "eigener Wert",
      "ovr.src.master": "vom Master",
      "ovr.src.globalShort": "global",
      "ovr.src.none": "leer",
      "ovr.fill": "Füllung",
      "ovr.font": "Schriftfarbe",
      "ovr.line": "Linie",
      "ovr.ownColor": "eigene Farbe",
      "ovr.fromMaster": "vom Master",
      "ovr.reset": "Zurücksetzen",
      "ovr.resetAll": "Alle Anpassungen zurücksetzen",

      "msg.syncing": "Synchronisiere …",
      "msg.synced": "{n} Instanz(en) synchronisiert.",
      "warn.masterSnapshot": "„{name}“: Master nicht gefunden – letzter bekannter Stand verwendet. Wähle eine Instanz und klicke „Zum Master machen“.",
      "warn.masterMissing": "„{name}“: Master nicht gefunden.",
      "warn.skipped": "„{name}“: {n} Instanz(en) mit nicht unterstütztem Typ übersprungen (Gruppe/Bild/Tabelle).",
      "warn.slide": "„{name}“ auf Folie {slide}: {error}",
      "msg.compCreated": "Komponente „{name}“ angelegt. Jetzt den Master kopieren und auf anderen Folien einfügen.",
      "msg.variantCreated": "Variante „{name}“ angelegt. Instanzen können jetzt darauf umgestellt werden.",
      "msg.linked": "{n} Form(en) mit „{name}“ verknüpft.",
      "msg.detached": "Gelöst. Die Form wird nicht mehr synchronisiert.",
      "msg.compDeleted": "Komponente „{name}“ gelöscht. Die Formen bleiben unverändert auf den Folien.",
      "err.exactlyOne": "Bitte genau eine Form markieren.",
      "err.typeUnsupported": "Typ „{type}“ wird noch nicht unterstützt. Nur Formen und Textfelder.",
      "err.alreadyMaster": "Die Form ist schon Master von „{name}“.",
      "err.selectTargets": "Bitte die Form(en) markieren, die Instanz werden sollen.",
      "err.detachMaster": "Der Master kann nicht gelöst werden. Komponente stattdessen löschen oder anderen Master wählen.",

      "err.prefix": "Fehler",
      "err.during": "Fehler bei „{action}“",
      "err.unknown": "Unbekannter Fehler",
      "err.location": "Ort",
      "err.save": "Speichern im Dokument fehlgeschlagen: {error}",
      "err.state": "Gespeicherter Zustand ist beschädigt: {error}",

      "act.apply": "Deck aktualisieren",
      "act.scan": "Platzhalter finden",
      "act.insert": "Platzhalter einfügen",
      "act.bind": "Farbe anwenden",
      "act.unbind": "Farbbindung lösen",
      "act.sync": "Synchronisieren",
      "act.count": "Komponenten zählen",
      "act.create": "Komponente anlegen",
      "act.link": "Als Instanz verknüpfen",
      "act.detach": "Lösen",
      "act.override": "Anpassung speichern",
      "act.variant": "Variante wechseln",
      "act.delete": "Komponente löschen",
      "act.showMaster": "Master zeigen",
      "act.selection": "Auswahl lesen",
    },
  };

  const LANGS = Object.keys(STRINGS);
  let lang = "en";

  /** Picks "de" for German Office UIs, English for everything else. */
  function detect(officeLanguage, stored) {
    if (stored && LANGS.includes(stored)) return stored;
    const l = String(officeLanguage || "").toLowerCase();
    return l.startsWith("de") ? "de" : "en";
  }

  function setLang(l) {
    lang = LANGS.includes(l) ? l : "en";
  }

  function getLang() {
    return lang;
  }

  /**
   * Translate a key. {param} is replaced from params.
   * Note: "{{{name}}}" renders as "{{Name}}" (literal braces around a param).
   */
  function t(key, params) {
    let s = (STRINGS[lang] && STRINGS[lang][key]) || STRINGS.en[key] || key;
    if (params) {
      s = s.replace(/\{(\w+)\}/g, (whole, p) => (p in params ? String(params[p]) : whole));
    }
    return s;
  }

  const api = { t, setLang, getLang, detect, LANGS, STRINGS };
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.DCI18n = api;
})(typeof self !== "undefined" ? self : this);
