// End-to-end test: loads the real task pane in headless Chromium with a simulated
// PowerPoint API (office-mock.js) and drives the UI like a user would.
const { chromium } = require("playwright");
const path = require("path");
const root = path.join(__dirname, "..", "..", "addin");

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  const errors = [];
  p.on("pageerror", (e) => errors.push(e.message));
  p.on("console", (m) => m.type() === "error" && errors.push(m.text()));
  await p.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ path: path.join(__dirname, "office-mock.js"), contentType: "text/javascript" }));
  await p.goto("file://" + root + "/taskpane.html");
  await p.waitForTimeout(200);

  const wait = (ms) => p.waitForTimeout(ms);
  const st = () => p.textContent("#status");
  const txt = (s, i) => p.evaluate(([s, i]) => __pp.slides[s].shapes[i].textFrame.textRange.text, [s, i]);
  const shape = (sl, id) => p.evaluate(([sl, id]) => {
    const s = __pp.slides[sl].shapes.find((x) => x.id === String(id));
    return { name: s.name, fill: s.fill.foregroundColor, text: s.textFrame.textRange.text, font: s.textFrame.textRange.font.color };
  }, [sl, id]);
  const select = (slide, ids) => p.evaluate(([s, i]) => __pp.select(s, i), [slide, ids]);
  const setOverrideText = async (value) => {
    await p.fill("#sel-actions .ovr input[type=text]", value);
    await p.dispatchEvent("#sel-actions .ovr input[type=text]", "change");
    await wait(800);
  };
  const check = (c, m) => { console.log((c ? "PASS " : "FAIL ") + m); if (!c) process.exitCode = 1; };

  // ---- variables ----
  await p.fill("#var-name", "Client"); await p.click("#var-add button");
  await p.fill("#var-list li input[type=text]", "ACME Ltd"); await wait(1000);
  check(await txt(0, 0) === "Offer for ACME Ltd", "variable rendered: " + await txt(0, 0));
  check((await st()).includes("Unknown placeholders: Project"), "missing variable reported: " + await st());
  check(await txt(1, 0) === "ACME Ltd – {{Project}}", "slide 2 partially rendered: " + await txt(1, 0));
  await p.click("#scan-vars"); await wait(200);
  check((await p.$$("#var-list li")).length === 2, "Project created by scan");
  await p.fill("#var-list li:nth-child(2) input[type=text]", "Relaunch"); await wait(1000);
  check(await txt(1, 0) === "ACME Ltd – Relaunch", "second variable: " + await txt(1, 0));
  await p.fill("#var-list li input[type=text]", "Contoso"); await wait(1000);
  check(await txt(0, 0) === "Offer for Contoso", "value change applied everywhere: " + await txt(0, 0));
  // user edits the rendered text, then changes the value
  await p.evaluate(() => (__pp.slides[0].shapes[0].textFrame.textRange.text = "New offer for Contoso"));
  await p.fill("#var-list li input[type=text]", "Fabrikam"); await wait(1000);
  check(await txt(0, 0) === "New offer for Fabrikam", "manual edit kept: " + await txt(0, 0));
  await p.check("#show-raw"); await wait(300);
  check(await txt(0, 0) === "New offer for {{Client}}", "placeholder view: " + await txt(0, 0));
  await p.uncheck("#show-raw"); await wait(300);

  // ---- color variable ----
  await p.fill("#var-name", "Primary"); await p.selectOption("#var-type", "color"); await p.click("#var-add button");
  await select(0, [3]); await wait(400);
  await p.click("#var-list li:nth-child(3) button:has-text('Fill')"); await wait(300);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#C43E1C", "color bound to fill");
  await p.fill("#var-list li:nth-child(3) input[type=text]", "#00ff00");
  await p.dispatchEvent("#var-list li:nth-child(3) input[type=text]", "change"); await wait(1000);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#00FF00", "color change -> shape follows");

  // ---- components ----
  await p.click(".tab[data-tab=comps]"); await wait(200);
  await select(0, [3]); await wait(400);
  await p.fill("#comp-name", "Button"); await p.click("#comp-add button"); await wait(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "◆ Button ◆", "master renamed");
  check((await p.textContent("#sel-info")).includes("Master of"), "selection shows master");
  // simulate copy & paste (the name travels with the copy)
  await p.evaluate(() => { const s = __pp.mkShape({ id: 9, name: "◆ Button ◆", fill: "#123456", text: "old" }); s.width = 40; __pp.slides[1].shapes.push(s); });
  await p.evaluate(() => { const m = __pp.slides[0].shapes[1]; m.fill.foregroundColor = "#AA00AA"; m.width = 222; m.textFrame.textRange.text = "Buy now {{Client}}"; m.textFrame.textRange.font.bold = true; });
  await p.click("#sync-all"); await wait(400);
  const inst = await p.evaluate(() => { const s = __pp.slides[1].shapes.find((x) => x.id === "9"); return { f: s.fill.foregroundColor, w: s.width, t: s.textFrame.textRange.text, b: s.textFrame.textRange.font.bold }; });
  check(inst.f === "#AA00AA" && inst.w === 222 && inst.b === true, "instance takes master style " + JSON.stringify(inst));
  check(inst.t === "Buy now Fabrikam", "instance text with variable: " + inst.t);
  check((await p.textContent("#comp-list")).includes("1 instance"), "instances counted: " + await p.textContent("#comp-list .meta"));
  // changing a variable updates the instance text
  await p.click(".tab[data-tab=vars]"); await p.fill("#var-list li input[type=text]", "Northwind"); await wait(1000);
  check(await txt(1, 2) === "Buy now Northwind", "variable inside instance: " + await txt(1, 2));
  // select instance -> make master
  await p.click(".tab[data-tab=comps]"); await select(1, [9]); await wait(400);
  check((await p.textContent("#sel-info")).includes("Instance of"), "selection recognizes instance");
  await p.click("#sel-actions button:has-text('Make master')"); await wait(500);
  check((await p.textContent("#sel-info")).includes("Master of"), "made master");
  // auto-sync
  await p.check("#auto-sync");
  await p.evaluate(() => (__pp.slides[1].shapes.find((x) => x.id === "9").fill.foregroundColor = "#0000FF"));
  await select(0, [2]); await wait(600);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#0000FF", "auto-sync when leaving the master");
  await select(0, [3]); await wait(400);
  await p.click("#sel-actions button:has-text('Detach')"); await wait(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "Button", "instance detached");
  await p.click("#comp-list button:has-text('Delete')"); await p.click("#comp-list button:has-text('Really')"); await wait(400);
  check(await p.evaluate(() => __pp.slides[1].shapes.find((x) => x.id === "9").name) === "Button", "component deleted, names cleaned up");

  // ---- instance overrides ----
  await p.evaluate(() => __pp.slides[0].shapes.push(__pp.mkShape({ id: 20, name: "Card", fill: "#336699", text: "{{Label}} for {{Client}}" })));
  await select(0, [20]); await wait(400);
  await p.selectOption("#comp-of", ""); await p.fill("#comp-name", "Card"); await p.click("#comp-add button"); await wait(500);
  await p.evaluate(() => __pp.slides[1].shapes.push(__pp.mkShape({ id: 21, name: "◆ Card ◆", text: "x" })));
  await p.click("#sync-all"); await wait(400);
  await select(0, [20]); await wait(400);
  check((await p.textContent("#sel-actions")).includes("Component defaults"), "master shows defaults");
  await setOverrideText("Learn more");
  check((await shape(0, 20)).text === "Learn more for Northwind", "master default: " + (await shape(0, 20)).text);
  check((await shape(1, 21)).text === "Learn more for Northwind", "instance inherits default: " + (await shape(1, 21)).text);
  await select(1, [21]); await wait(400);
  check((await p.textContent("#sel-actions")).includes("Overrides for this instance"), "instance shows overrides");
  await setOverrideText("Buy now");
  check((await shape(1, 21)).text === "Buy now for Northwind", "instance text overridden: " + (await shape(1, 21)).text);
  check((await shape(0, 20)).text === "Learn more for Northwind", "master unchanged");
  await p.fill("#sel-actions .ovr input[type=color] >> nth=0", "#ff00ff"); await wait(800);
  check((await shape(1, 21)).fill === "#FF00FF", "own fill color: " + (await shape(1, 21)).fill);
  await p.evaluate(() => (__pp.slides[0].shapes.find((x) => x.id === "20").fill.foregroundColor = "#111111"));
  await p.click("#sync-all"); await wait(400);
  check((await shape(1, 21)).fill === "#FF00FF", "own color survives sync");
  await select(0, [2]); await wait(300);
  await select(1, [21]); await wait(400);
  await p.click("#sel-actions button:has-text('Reset all overrides')"); await wait(800);
  const r = await shape(1, 21);
  check(r.fill === "#111111" && r.text === "Learn more for Northwind", "reset all: " + JSON.stringify(r));

  // ---- variants ----
  await p.evaluate(() => __pp.slides[0].shapes.push(__pp.mkShape({ id: 22, name: "◆ Card ◆", fill: "#008000", text: "{{Label}} for {{Client}}" })));
  await select(0, [22]); await wait(400);
  await p.selectOption("#comp-of", "Card"); await p.fill("#comp-name", "Green"); await p.click("#comp-add button"); await wait(500);
  check((await shape(0, 22)).name === "◆ Card / Green ◆", "variant created: " + (await shape(0, 22)).name);
  check((await p.textContent("#comp-list")).includes("Card · Green"), "variant listed");
  await select(1, [21]); await wait(400);
  await setOverrideText("Go");
  await p.selectOption('#sel-actions select[aria-label="Variant:"]', "Card / Green"); await wait(900);
  const v = await shape(1, 21);
  check(v.name === "◆ Card / Green ◆" && v.fill === "#008000", "variant switched: " + JSON.stringify(v));
  check(v.text === "Go for Northwind", "override kept when switching variant: " + v.text);

  // ---- GeneralException while reading the selection (PowerPoint on the web) ----
  await p.evaluate(() => { window.__selFail = 1; __pp.select(1, [21]); }); await wait(500);
  check((await p.textContent("#sel-info")).includes("Instance of"), "selection read via fallback: " + await p.textContent("#sel-info"));
  await p.evaluate(() => { window.__selFail = 99; __pp.select(1, [21]); }); await wait(500);
  check((await p.textContent("#sel-info")).includes("Could not read"), "unreadable selection -> hint");
  check(!(await p.getAttribute("#status", "class") || "").includes("err"), "no red error for unreadable selection");
  await p.evaluate(() => { window.__selFail = 0; __pp.select(1, [21]); }); await wait(400);

  // ---- language + version ----
  check((await p.textContent(".tab[data-tab=vars]")) === "Variables", "English UI by default");
  check(/^v\d+\.\d+\.\d+ · dev$/.test(await p.textContent("#version")), "version shown: " + await p.textContent("#version"));
  check((await p.$$("#tab-vars .hint code")).length === 1, "{{Name}} shown as code in the hint");
  await p.click(".lang button[data-lang=de]"); await wait(200);
  check((await p.textContent(".tab[data-tab=vars]")) === "Variablen", "switch to German");
  check((await p.textContent("#var-list")).includes("einfügen"), "dynamic lists re-rendered in German");
  await p.click(".lang button[data-lang=en]"); await wait(200);
  check((await p.textContent(".tab[data-tab=comps]")) === "Components", "switch back to English");

  check(errors.length === 0, "no JS errors " + errors.join(" | "));
  console.log("Status:", await st());
  if (process.env.SHOT) {
    await p.setViewportSize({ width: 340, height: 900 });
    await select(1, [21]); await wait(400);
    await p.screenshot({ path: process.env.SHOT });
  }
  await b.close();
})();
