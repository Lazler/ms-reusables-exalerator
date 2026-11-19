const test = require("node:test");
const assert = require("node:assert/strict");
const I = require("../addin/i18n.js");

test("German and English have the same keys", () => {
  const en = Object.keys(I.STRINGS.en).sort();
  const de = Object.keys(I.STRINGS.de).sort();
  assert.deepEqual(de.filter((k) => !en.includes(k)), [], "keys only in German");
  assert.deepEqual(en.filter((k) => !de.includes(k)), [], "keys missing in German");
});

test("placeholders match between languages", () => {
  const params = (s) => [...s.matchAll(/(?<!\{)\{(\w+)\}(?!\})|\{\{\{(\w+)\}\}\}/g)].map((m) => m[1] || m[2]).sort();
  for (const k of Object.keys(I.STRINGS.en)) {
    assert.deepEqual(params(I.STRINGS.de[k]), params(I.STRINGS.en[k]), k);
  }
});

test("language detection defaults to English", () => {
  assert.equal(I.detect("de-DE"), "de");
  assert.equal(I.detect("de-AT"), "de");
  assert.equal(I.detect("en-GB"), "en");
  assert.equal(I.detect("fr-FR"), "en");
  assert.equal(I.detect(undefined), "en");
  assert.equal(I.detect("de-DE", "en"), "en"); // manual choice wins
});

test("t() fills parameters and keeps {{…}} literal", () => {
  I.setLang("en");
  assert.equal(I.t("vars.insert", { name: "Client" }), "Insert {{Client}}");
  assert.equal(I.t("msg.synced", { n: 3 }), "3 instance(s) synced.");
  I.setLang("de");
  assert.equal(I.t("vars.insert", { name: "Kunde" }), "{{Kunde}} einfügen");
  assert.equal(I.t("unknown.key"), "unknown.key");
  I.setLang("en");
});
