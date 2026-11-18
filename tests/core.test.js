const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../addin/core.js");

const vars = [
  { name: "Kunde", type: "text", value: "ACME GmbH" },
  { name: "Projekt", type: "text", value: "Relaunch" },
  { name: "Jahr", type: "text", value: "2027" },
];

test("render ersetzt bekannte Variablen und meldet unbekannte", () => {
  const r = C.render("Für {{Kunde}} – {{ Projekt }} ({{Fehlt}})", vars);
  assert.equal(r.text, "Für ACME GmbH – Relaunch ({{Fehlt}})");
  assert.deepEqual(r.missing, ["Fehlt"]);
});

test("hasPlaceholders / extractNames", () => {
  assert.equal(C.hasPlaceholders("ohne"), false);
  assert.equal(C.hasPlaceholders("mit {{Kunde}}"), true);
  assert.equal(C.hasPlaceholders("mit {{Kunde}}"), true); // lastIndex-Falle
  assert.deepEqual(C.extractNames("{{a}} {{b}} {{a}}"), ["a", "b"]);
  assert.deepEqual(C.extractNames("{{Größe}}"), ["Größe"]);
});

test("resolveTemplate: unveränderter Text behält Template", () => {
  assert.equal(C.resolveTemplate("Hallo ACME GmbH", "Hallo {{Kunde}}", "Hallo ACME GmbH", vars), "Hallo {{Kunde}}");
  // PowerPoint-Zeilenumbrüche
  assert.equal(C.resolveTemplate("A\rACME GmbH", "A\n{{Kunde}}", "A\nACME GmbH", vars), "A\n{{Kunde}}");
});

test("resolveTemplate: Bearbeitung mit neuem Platzhalter behält alte Variablen", () => {
  const t = C.resolveTemplate("Hallo ACME GmbH, {{Projekt}}", "Hallo {{Kunde}}", "Hallo ACME GmbH", vars);
  assert.equal(t, "Hallo {{Kunde}}, {{Projekt}}");
});

test("resolveTemplate: Bearbeitung ohne Werte -> keine Variablen mehr", () => {
  assert.equal(C.resolveTemplate("Ganz neuer Text", "Hallo {{Kunde}}", "Hallo ACME GmbH", vars), null);
});

test("resolveTemplate: Text bearbeitet, Wert noch drin -> Variable bleibt", () => {
  assert.equal(C.resolveTemplate("Hi ACME GmbH!", "Hallo {{Kunde}}", "Hallo ACME GmbH", vars), "Hi {{Kunde}}!");
});

test("resolveTemplate: neues Feld mit Platzhalter", () => {
  assert.equal(C.resolveTemplate("{{Jahr}}", null, null, vars), "{{Jahr}}");
  assert.equal(C.resolveTemplate("normal", null, null, vars), null);
});

test("reabsorb ersetzt längere Werte zuerst und nicht doppelt", () => {
  const v = [
    { name: "A", value: "ACME" },
    { name: "B", value: "ACME GmbH" },
  ];
  assert.equal(C.reabsorb("ACME GmbH und ACME", "{{A}} {{B}}", v), "{{B}} und {{A}}");
});

test("normalizeColor", () => {
  assert.equal(C.normalizeColor("#abc"), "#AABBCC");
  assert.equal(C.normalizeColor("ff0000"), "#FF0000");
  assert.equal(C.normalizeColor("rot"), null);
});

test("Komponentennamen", () => {
  assert.equal(C.componentShapeName("Button"), "◆ Button ◆");
  assert.equal(C.componentNameOf("◆ Button ◆"), "Button");
  assert.equal(C.componentNameOf("◆ Primär Button ◆ 2"), "Primär Button"); // PowerPoint-Suffix
  assert.equal(C.componentNameOf("Rechteck 3"), null);
  assert.equal(C.stripComponentName("◆ Button ◆"), "Button");
});

test("isValidName", () => {
  assert.equal(C.isValidName("Kunde"), true);
  assert.equal(C.isValidName("Kunde Name_2"), true);
  assert.equal(C.isValidName("a{b"), false);
  assert.equal(C.isValidName("  "), false);
});

test("extractValues liest damals eingesetzte Werte", () => {
  assert.deepEqual(C.extractValues("Hallo {{Kunde}} ({{Jahr}})", "Hallo ACME (2026)"), [
    { name: "Kunde", value: "ACME" },
    { name: "Jahr", value: "2026" },
  ]);
  assert.equal(C.extractValues("Hallo {{Kunde}}", "Tschüss ACME"), null);
  assert.deepEqual(C.extractValues("a.b {{X}}\nz", "a.b 1+1\rz"), [{ name: "X", value: "1+1" }]);
});

test("resolveTemplate: Bearbeitung UND neuer Wert gleichzeitig", () => {
  const neu = [{ name: "Kunde", value: "Fabrikam" }];
  assert.equal(C.resolveTemplate("Neues Angebot für Contoso", "Angebot für {{Kunde}}", "Angebot für Contoso", neu), "Neues Angebot für {{Kunde}}");
});
