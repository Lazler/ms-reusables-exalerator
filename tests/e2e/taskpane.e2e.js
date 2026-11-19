// End-to-end test: loads the real task pane in headless Chromium with a simulated
// PowerPoint API (office-mock.js) and drives the UI like a user would.
const { chromium } = require("playwright");
const path = require("path");
const root = path.join(__dirname, "..", "..", "addin");

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 340, height: 800 } });
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  p.on("console", (m) => m.type() === "error" && errors.push(m.text()));
  await p.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ path: path.join(__dirname, "office-mock.js"), contentType: "text/javascript" }));
  await p.goto("file://" + root + "/taskpane.html");
  await p.waitForTimeout(300);

  const wait = (ms) => p.waitForTimeout(ms);
  const st = () => p.textContent("#status-text");
  const txt = (s, i) => p.evaluate(([s, i]) => __pp.slides[s].shapes[i].textFrame.textRange.text, [s, i]);
  const shape = (sl, id) => p.evaluate(([sl, id]) => {
    const s = __pp.slides[sl].shapes.find((x) => x.id === String(id));
    return { name: s.name, fill: s.fill.foregroundColor, text: s.textFrame.textRange.text, font: s.textFrame.textRange.font.color };
  }, [sl, id]);
  const select = (slide, ids) => p.evaluate(([s, i]) => __pp.select(s, i), [slide, ids]);
  const inspector = () => p.textContent("#inspector");
  /** Opens the menu behind a button and clicks an item by its text. */
  const pick = async (buttonSel, itemText) => {
    await p.click(buttonSel);
    await p.click(`.menu .menu-item:has-text("${itemText}")`);
    await wait(150);
  };
  const addVar = async (name, type) => {
    await p.click("#var-new");
    if (type === "color") await p.click('.seg button[data-type="color"]');
    await p.fill("#var-name", name);
    await p.click("#var-add-submit");
  };
  const varInput = (name) => `.var-row[data-name="${name}"] .cell-input`;
  const setProp = async (value) => {
    await p.fill("#inspector .props .input", value);
    await p.dispatchEvent("#inspector .props .input", "change");
    await wait(900);
  };
  const check = (c, m) => { console.log((c ? "PASS " : "FAIL ") + m); if (!c) process.exitCode = 1; };

  // ---- empty state ----
  check(await p.isVisible("#var-empty"), "empty state shown for variables");
  check(!(await p.isVisible("#vars-footnote")), "footnote hidden while empty");

  // ---- variables ----
  await addVar("Client");
  check(!(await p.isVisible("#var-empty")), "empty state gone after adding");
  await p.fill(varInput("Client"), "ACME Ltd"); await wait(1000);
  check(await txt(0, 0) === "Offer for ACME Ltd", "variable rendered: " + await txt(0, 0));
  check((await st()).includes("No value for: Project"), "missing variable reported: " + await st());
  check(await txt(1, 0) === "ACME Ltd – {{Project}}", "slide 2 partially rendered: " + await txt(1, 0));
  await pick("#vars-menu-slot .menu-btn", "Find placeholders");
  await wait(200);
  check((await p.$$("#var-list .var-row")).length === 2, "Project created by scan");
  await p.fill(varInput("Project"), "Relaunch"); await wait(1000);
  check(await txt(1, 0) === "ACME Ltd – Relaunch", "second variable: " + await txt(1, 0));
  await p.fill(varInput("Client"), "Contoso"); await wait(1000);
  check(await txt(0, 0) === "Offer for Contoso", "value change applied everywhere: " + await txt(0, 0));
  // user edits the rendered text, then changes the value
  await p.evaluate(() => (__pp.slides[0].shapes[0].textFrame.textRange.text = "New offer for Contoso"));
  await p.fill(varInput("Client"), "Fabrikam"); await wait(1000);
  check(await txt(0, 0) === "New offer for Fabrikam", "manual edit kept: " + await txt(0, 0));
  await pick("#vars-menu-slot .menu-btn", "placeholders}} instead"); await wait(300);
  check(await txt(0, 0) === "New offer for {{Client}}", "placeholder view: " + await txt(0, 0));
  await pick("#vars-menu-slot .menu-btn", "placeholders}} instead"); await wait(300);
  check(await txt(0, 0) === "New offer for Fabrikam", "back to values: " + await txt(0, 0));

  // ---- color variable ----
  await addVar("Primary", "color");
  await select(0, [3]); await wait(400);
  await pick('.var-row[data-name="Primary"] .menu-btn >> nth=0', "As fill"); await wait(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#C43E1C", "color bound to fill");
  await p.fill(varInput("Primary"), "#00ff00");
  await p.dispatchEvent(varInput("Primary"), "change"); await wait(1000);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#00FF00", "color change -> shape follows");

  // ---- components ----
  await p.click(".tab[data-tab=comps]"); await wait(200);
  await select(0, [3]); await wait(400);
  check((await inspector()).includes("Create component"), "plain shape offers Create component");
  check((await inspector()).includes("Shape"), "friendly type label");
  await p.click("#create-btn"); await p.fill("#insp-name", "Button"); await p.click("#insp-submit"); await wait(500);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "◆ Button ◆", "master renamed");
  check((await inspector()).includes("Master component"), "inspector shows master");
  // simulate copy & paste (the name travels with the copy)
  await p.evaluate(() => { const s = __pp.mkShape({ id: 9, name: "◆ Button ◆", fill: "#123456", text: "old" }); s.width = 40; __pp.slides[1].shapes.push(s); });
  await p.evaluate(() => { const m = __pp.slides[0].shapes[1]; m.fill.foregroundColor = "#AA00AA"; m.width = 222; m.textFrame.textRange.text = "Buy now {{Client}}"; m.textFrame.textRange.font.bold = true; });
  await pick("#comps-menu-slot .menu-btn", "Push all components"); await wait(400);
  const inst = await p.evaluate(() => { const s = __pp.slides[1].shapes.find((x) => x.id === "9"); return { f: s.fill.foregroundColor, w: s.width, t: s.textFrame.textRange.text, b: s.textFrame.textRange.font.bold }; });
  check(inst.f === "#AA00AA" && inst.w === 222 && inst.b === true, "instance takes master style " + JSON.stringify(inst));
  check(inst.t === "Buy now Fabrikam", "instance text with variable: " + inst.t);
  check((await p.textContent("#comp-list")).includes("1 instance"), "instances counted: " + await p.textContent("#comp-list"));
  // changing a variable updates the instance text
  await p.click(".tab[data-tab=vars]"); await p.fill(varInput("Client"), "Northwind"); await wait(1000);
  check(await txt(1, 2) === "Buy now Northwind", "variable inside instance: " + await txt(1, 2));
  // select instance -> make master
  await p.click(".tab[data-tab=comps]"); await select(1, [9]); await wait(400);
  check((await inspector()).includes("Instance"), "inspector recognizes instance");
  await pick("#inspector .insp-head .menu-btn", "Make this the master"); await wait(500);
  check((await inspector()).includes("Master component"), "made master");
  check((await inspector()).includes("Push changes to 1 instance"), "push button shows instance count");
  // auto-sync
  await pick("#comps-menu-slot .menu-btn", "Auto-push");
  await p.evaluate(() => (__pp.slides[1].shapes.find((x) => x.id === "9").fill.foregroundColor = "#0000FF"));
  await select(0, [2]); await wait(700);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#0000FF", "auto-sync when leaving the master");
  await select(0, [3]); await wait(400);
  await pick("#inspector .insp-head .menu-btn", "Detach instance"); await wait(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "Button", "instance detached");
  await p.click('.comp-row[data-name="Button"] .menu-btn');
  await p.click('.menu .menu-item:has-text("Delete component")'); await wait(150);
  await p.click('.menu .menu-item:has-text("Click again")'); await wait(400);
  check(await p.evaluate(() => __pp.slides[1].shapes.find((x) => x.id === "9").name) === "Button", "component deleted, names cleaned up");
  check(await p.isVisible("#comp-empty"), "components empty state");

  // ---- instance overrides ----
  await p.evaluate(() => __pp.slides[0].shapes.push(__pp.mkShape({ id: 20, name: "Card", fill: "#336699", text: "{{Label}} for {{Client}}" })));
  await select(0, [20]); await wait(400);
  await p.click("#create-btn"); await p.fill("#insp-name", "Card"); await p.press("#insp-name", "Enter"); await wait(500);
  await p.evaluate(() => __pp.slides[1].shapes.push(__pp.mkShape({ id: 21, name: "◆ Card ◆", text: "x" })));
  await select(0, [2]); await wait(300);
  await select(0, [20]); await wait(400);
  await p.click("#push-btn"); await wait(500);
  check((await inspector()).includes("Default values"), "master shows default values");
  await setProp("Learn more");
  check((await shape(0, 20)).text === "Learn more for Northwind", "master default: " + (await shape(0, 20)).text);
  check((await shape(1, 21)).text === "Learn more for Northwind", "instance inherits default: " + (await shape(1, 21)).text);
  check(!(await st()).includes("No value for: Label"), "component placeholder not reported as missing");
  await select(1, [21]); await wait(400);
  check((await inspector()).includes("Overrides"), "instance shows overrides");
  await setProp("Buy now");
  check((await shape(1, 21)).text === "Buy now for Northwind", "instance text overridden: " + (await shape(1, 21)).text);
  check((await shape(0, 20)).text === "Learn more for Northwind", "master unchanged");
  await p.fill("#inspector .color-props input[type=color] >> nth=0", "#ff00ff"); await wait(900);
  check((await shape(1, 21)).fill === "#FF00FF", "own fill color: " + (await shape(1, 21)).fill);
  check(await p.isVisible("#inspector .reset-color"), "reset button for own color");
  await p.evaluate(() => (__pp.slides[0].shapes.find((x) => x.id === "20").fill.foregroundColor = "#111111"));
  await pick("#comps-menu-slot .menu-btn", "Push all components"); await wait(400);
  check((await shape(1, 21)).fill === "#FF00FF", "own color survives sync");
  await select(0, [2]); await wait(300);
  await select(1, [21]); await wait(400);
  await pick("#inspector .insp-head .menu-btn", "Reset all overrides"); await wait(900);
  const r = await shape(1, 21);
  check(r.fill === "#111111" && r.text === "Learn more for Northwind", "reset all: " + JSON.stringify(r));

  // ---- variants: copy an instance, change it, "Save as new variant" ----
  await p.evaluate(() => __pp.slides[0].shapes.push(__pp.mkShape({ id: 22, name: "◆ Card ◆", fill: "#008000", text: "{{Label}} for {{Client}}" })));
  await select(0, [22]); await wait(400);
  await pick("#inspector .insp-head .menu-btn", "Save as new variant");
  await p.fill("#insp-name", "Green"); await p.click("#insp-submit"); await wait(500);
  check((await shape(0, 22)).name === "◆ Card / Green ◆", "variant created: " + (await shape(0, 22)).name);
  check(await p.isVisible('.comp-row.variant[data-name="Card / Green"]'), "variant listed under its component");
  await select(1, [21]); await wait(400);
  await setProp("Go");
  await p.selectOption("#variant-select", "Card / Green"); await wait(900);
  const v = await shape(1, 21);
  check(v.name === "◆ Card / Green ◆" && v.fill === "#008000", "variant switched: " + JSON.stringify(v));
  check(v.text === "Go for Northwind", "override kept when switching variant: " + v.text);

  // ---- GeneralException while reading the selection (PowerPoint on the web) ----
  await p.evaluate(() => { window.__selFail = 1; __pp.select(1, [21]); }); await wait(500);
  check((await inspector()).includes("Instance"), "selection read via fallback: " + await inspector());
  await p.evaluate(() => { window.__selFail = 99; __pp.select(1, [21]); }); await wait(500);
  check((await inspector()).includes("read the selection"), "unreadable selection -> hint");
  check(!(await p.getAttribute("#status", "class") || "").includes("err"), "no red error for unreadable selection");
  await p.evaluate(() => { window.__selFail = 0; __pp.select(1, [21]); }); await wait(400);

  // ---- auto-render: typing {{Client}} renders once the user leaves the text box ----
  await p.evaluate(() => __pp.slides[0].shapes.push(__pp.mkShape({ id: 30, name: "Typed", type: "TextBox", text: "" })));
  await select(0, [30]); await wait(400);
  await p.evaluate(() => (__pp.slides[0].shapes.find((x) => x.id === "30").textFrame.textRange.text = "Hi {{Client}}"));
  await select(0, [30]); await wait(500); // still in the same box -> must not be touched
  check((await shape(0, 30)).text === "Hi {{Client}}", "box being edited is left alone: " + (await shape(0, 30)).text);
  await select(0, [2]); await wait(600); // leave the box
  check((await shape(0, 30)).text === "Hi Northwind", "rendered after leaving the box: " + (await shape(0, 30)).text);

  // ---- language + version ----
  check((await p.textContent('.tab[data-tab=vars] [data-i18n]')) === "Variables", "English UI by default");
  check(/^v\d+\.\d+\.\d+ · dev$/.test(await p.textContent("#version")), "version shown: " + await p.textContent("#version"));
  check((await p.$$("#vars-footnote code")).length === 1, "{{Name}} shown as code in the hint");
  await pick("#settings-slot .menu-btn", "Deutsch"); await wait(200);
  check((await p.textContent('.tab[data-tab=vars] [data-i18n]')) === "Variablen", "switch to German");
  check(((await p.getAttribute('.var-row[data-name="Client"] .insert-btn', "title")) || "").includes("einfügen"), "dynamic lists re-rendered in German");
  await pick("#settings-slot .menu-btn", "English"); await wait(200);
  check((await p.textContent('.tab[data-tab=comps] [data-i18n]')) === "Components", "switch back to English");

  check(errors.length === 0, "no JS errors " + errors.join(" | "));
  console.log("Status:", await st());
  await b.close();
})();
