const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../addin/core.js");

const vars = [
  { name: "Client", type: "text", value: "ACME Ltd" },
  { name: "Project", type: "text", value: "Relaunch" },
  { name: "Year", type: "text", value: "2027" },
];

test("render replaces known variables and reports unknown ones", () => {
  const r = C.render("For {{Client}} – {{ Project }} ({{Missing}})", vars);
  assert.equal(r.text, "For ACME Ltd – Relaunch ({{Missing}})");
  assert.deepEqual(r.missing, ["Missing"]);
});

test("hasPlaceholders / extractNames", () => {
  assert.equal(C.hasPlaceholders("plain"), false);
  assert.equal(C.hasPlaceholders("with {{Client}}"), true);
  assert.equal(C.hasPlaceholders("with {{Client}}"), true); // no lastIndex trap
  assert.deepEqual(C.extractNames("{{a}} {{b}} {{a}}"), ["a", "b"]);
  assert.deepEqual(C.extractNames("{{Größe}}"), ["Größe"]); // non-ASCII names
});

test("resolveTemplate: unchanged text keeps its template", () => {
  assert.equal(C.resolveTemplate("Hello ACME Ltd", "Hello {{Client}}", "Hello ACME Ltd", vars), "Hello {{Client}}");
  // PowerPoint line breaks
  assert.equal(C.resolveTemplate("A\rACME Ltd", "A\n{{Client}}", "A\nACME Ltd", vars), "A\n{{Client}}");
});

test("resolveTemplate: editing with a new placeholder keeps the old variables", () => {
  const tpl = C.resolveTemplate("Hello ACME Ltd, {{Project}}", "Hello {{Client}}", "Hello ACME Ltd", vars);
  assert.equal(tpl, "Hello {{Client}}, {{Project}}");
});

test("resolveTemplate: text replaced without values -> no variables left", () => {
  assert.equal(C.resolveTemplate("Completely new text", "Hello {{Client}}", "Hello ACME Ltd", vars), null);
});

test("resolveTemplate: edited text still contains the value -> variable stays", () => {
  assert.equal(C.resolveTemplate("Hi ACME Ltd!", "Hello {{Client}}", "Hello ACME Ltd", vars), "Hi {{Client}}!");
});

test("resolveTemplate: new text box with a placeholder", () => {
  assert.equal(C.resolveTemplate("{{Year}}", null, null, vars), "{{Year}}");
  assert.equal(C.resolveTemplate("plain", null, null, vars), null);
});

test("resolveTemplate: edit AND new value at the same time", () => {
  const now = [{ name: "Client", value: "Fabrikam" }];
  assert.equal(C.resolveTemplate("New offer for Contoso", "Offer for {{Client}}", "Offer for Contoso", now), "New offer for {{Client}}");
});

test("reabsorb replaces longer values first and never twice", () => {
  const v = [
    { name: "A", value: "ACME" },
    { name: "B", value: "ACME Ltd" },
  ];
  assert.equal(C.reabsorb("ACME Ltd and ACME", "{{A}} {{B}}", v), "{{B}} and {{A}}");
});

test("extractValues reads the values filled in at the time", () => {
  assert.deepEqual(C.extractValues("Hello {{Client}} ({{Year}})", "Hello ACME (2026)"), [
    { name: "Client", value: "ACME" },
    { name: "Year", value: "2026" },
  ]);
  assert.equal(C.extractValues("Hello {{Client}}", "Bye ACME"), null);
  assert.deepEqual(C.extractValues("a.b {{X}}\nz", "a.b 1+1\rz"), [{ name: "X", value: "1+1" }]);
});

test("normalizeColor", () => {
  assert.equal(C.normalizeColor("#abc"), "#AABBCC");
  assert.equal(C.normalizeColor("ff0000"), "#FF0000");
  assert.equal(C.normalizeColor("red"), null);
});

test("component shape names", () => {
  assert.equal(C.componentShapeName("Button"), "◆ Button ◆");
  assert.equal(C.componentNameOf("◆ Button ◆"), "Button");
  assert.equal(C.componentNameOf("◆ Primary Button ◆ 2"), "Primary Button"); // PowerPoint suffix
  assert.equal(C.componentNameOf("Rectangle 3"), null);
  assert.equal(C.stripComponentName("◆ Button ◆"), "Button");
});

test("isValidName", () => {
  assert.equal(C.isValidName("Client"), true);
  assert.equal(C.isValidName("Client Name_2"), true);
  assert.equal(C.isValidName("a{b"), false);
  assert.equal(C.isValidName("  "), false);
});

test("variant keys", () => {
  assert.deepEqual(C.splitKey("Button"), { base: "Button", variant: "Default" });
  assert.deepEqual(C.splitKey("Button / Secondary"), { base: "Button", variant: "Secondary" });
  assert.equal(C.joinKey("Button", "Default"), "Button");
  assert.equal(C.joinKey("Button", "Standard"), "Button"); // legacy German default
  assert.equal(C.joinKey("Button", "Red"), "Button / Red");
  assert.equal(C.componentNameOf(C.componentShapeName("Button / Red") + " 2"), "Button / Red");
});

test("overrides", () => {
  const o = C.parseOverrides('{"vars":{"Label":"Buy"},"fill":"#abc"}');
  assert.deepEqual(o, { vars: { Label: "Buy" }, fill: "#AABBCC", font: null, line: null });
  assert.deepEqual(C.parseOverrides("broken"), { vars: {}, fill: null, font: null, line: null });
  assert.equal(C.serializeOverrides({ vars: { Label: "" }, fill: null }), null);
  assert.equal(C.serializeOverrides({ vars: { Label: "X" }, font: "f00" }), '{"vars":{"Label":"X"},"font":"#FF0000"}');
});

test("mergeVars: instance > master default > global", () => {
  const m = C.mergeVars([{ name: "Client", value: "ACME" }, { name: "Label", value: "g" }], { Label: "Default" }, { Label: "Instance", Client: "" });
  assert.deepEqual(m, [{ name: "Client", value: "ACME" }, { name: "Label", value: "Instance" }]);
});
