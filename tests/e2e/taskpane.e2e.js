const { chromium } = require("playwright");
const path = require("path");
const root = path.join(__dirname, "..", "..", "addin");
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  const errors = []; p.on("pageerror", e => errors.push(e.message)); p.on("console", m => m.type()==="error" && errors.push(m.text()));
  await p.route("https://appsforoffice.microsoft.com/**", r => r.fulfill({ path: path.join(__dirname, "office-mock.js"), contentType: "text/javascript" }));
  await p.goto("file://" + root + "/taskpane.html");
  await p.waitForTimeout(200);
  const st = () => p.textContent("#status");
  const txt = (s, i) => p.evaluate(([s,i]) => __pp.slides[s].shapes[i].textFrame.textRange.text, [s,i]);
  const check = (c, m) => { console.log((c?"PASS ":"FAIL ")+m); if(!c) process.exitCode=1; };

  // Variablen
  await p.fill("#var-name", "Kunde"); await p.click("#var-add button");
  await p.fill("#var-list li input[type=text]", "ACME GmbH");
  await p.waitForTimeout(1000);
  check(await txt(0,0) === "Angebot für ACME GmbH", "Variable gerendert: " + await txt(0,0));
  check((await st()).includes("Projekt"), "fehlende Variable gemeldet: " + await st());
  check(await txt(1,0) === "ACME GmbH – {{Projekt}}", "Folie 2 teilweise: " + await txt(1,0));
  await p.click("#scan-vars"); await p.waitForTimeout(200);
  check((await p.$$("#var-list li")).length === 2, "Projekt per Scan angelegt");
  await p.fill("#var-list li:nth-child(2) input[type=text]", "Relaunch"); await p.waitForTimeout(1000);
  check(await txt(1,0) === "ACME GmbH – Relaunch", "zweite Variable: " + await txt(1,0));
  await p.fill("#var-list li input[type=text]", "Contoso"); await p.waitForTimeout(1000);
  check(await txt(0,0) === "Angebot für Contoso", "Wertänderung überall: " + await txt(0,0));
  // Nutzer bearbeitet gerenderten Text
  await p.evaluate(() => __pp.slides[0].shapes[0].textFrame.textRange.text = "Neues Angebot für Contoso"); 
  await p.fill("#var-list li input[type=text]", "Fabrikam"); await p.waitForTimeout(1000);
  check(await txt(0,0) === "Neues Angebot für Fabrikam", "Bearbeitung übernommen: " + await txt(0,0));
  await p.check("#show-raw"); await p.waitForTimeout(300);
  check(await txt(0,0) === "Neues Angebot für {{Kunde}}", "Platzhalter-Ansicht: " + await txt(0,0));
  await p.uncheck("#show-raw"); await p.waitForTimeout(300);
  // Farbvariable
  await p.fill("#var-name", "Primär"); await p.selectOption("#var-type", "color"); await p.click("#var-add button");
  await p.evaluate(() => __pp.select(0, [3])); await p.waitForTimeout(400);
  await p.click("#var-list li:nth-child(3) button:has-text('Füllung')"); await p.waitForTimeout(300);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#C43E1C", "Farbbindung");
  await p.fill("#var-list li:nth-child(3) input[type=text]", "#00ff00"); await p.dispatchEvent("#var-list li:nth-child(3) input[type=text]", "change"); await p.waitForTimeout(1000);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#00FF00", "Farbe ändern -> Form folgt");

  // Komponenten
  await p.click(".tab[data-tab=comps]"); await p.waitForTimeout(200);
  await p.evaluate(() => __pp.select(0, [3])); await p.waitForTimeout(400);
  await p.fill("#comp-name", "Button"); await p.click("#comp-add button"); await p.waitForTimeout(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "◆ Button ◆", "Master umbenannt");
  check((await p.textContent("#sel-info")).includes("Master"), "Auswahl zeigt Master");
  // Kopieren & Einfügen simulieren (Name wandert mit)
  await p.evaluate(() => { const s = __pp.mkShape({ id: 9, name: "◆ Button ◆", fill: "#123456", text: "alt" }); s.width = 40; __pp.slides[1].shapes.push(s); });
  await p.evaluate(() => { const m = __pp.slides[0].shapes[1]; m.fill.foregroundColor = "#AA00AA"; m.width = 222; m.textFrame.textRange.text = "Jetzt kaufen {{Kunde}}"; m.textFrame.textRange.font.bold = true; });
  await p.click("#sync-all"); await p.waitForTimeout(400);
  const inst = await p.evaluate(() => { const s = __pp.slides[1].shapes.find(x=>x.id==="9"); return { f: s.fill.foregroundColor, w: s.width, t: s.textFrame.textRange.text, b: s.textFrame.textRange.font.bold }; });
  check(inst.f === "#AA00AA" && inst.w === 222 && inst.b === true, "Instanz übernimmt Stil " + JSON.stringify(inst));
  check(inst.t === "Jetzt kaufen Fabrikam", "Instanz-Text mit Variable: " + inst.t);
  check((await p.textContent("#comp-list")).includes("1 Instanz"), "Instanzen gezählt: " + await p.textContent("#comp-list .meta"));
  // Variable ändern -> Instanz-Text folgt
  await p.click(".tab[data-tab=vars]"); await p.fill("#var-list li input[type=text]", "Northwind"); await p.waitForTimeout(1000);
  check(await txt(1,2) === "Jetzt kaufen Northwind", "Variable in Instanz: " + await txt(1,2));
  // Auswahl Instanz -> Lösen
  await p.click(".tab[data-tab=comps]"); await p.evaluate(() => __pp.select(1, [9])); await p.waitForTimeout(400);
  check((await p.textContent("#sel-info")).includes("Instanz"), "Auswahl erkennt Instanz");
  await p.click("#sel-actions button:has-text('Zum Master machen')"); await p.waitForTimeout(500);
  check((await p.textContent("#sel-info")).includes("Master"), "Zum Master gemacht");
  // Auto-Sync
  await p.check("#auto-sync");
  await p.evaluate(() => __pp.slides[1].shapes.find(x=>x.id==="9").fill.foregroundColor = "#0000FF");
  await p.evaluate(() => __pp.select(0, [2])); await p.waitForTimeout(600);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].fill.foregroundColor) === "#0000FF", "Auto-Sync beim Verlassen des Masters");
  await p.evaluate(() => __pp.select(0, [3])); await p.waitForTimeout(400);
  await p.click("#sel-actions button:has-text('Lösen')"); await p.waitForTimeout(400);
  check(await p.evaluate(() => __pp.slides[0].shapes[1].name) === "Button", "Instanz gelöst");
  await p.click("#comp-list button:has-text('Löschen')"); await p.click("#comp-list button:has-text('Wirklich')"); await p.waitForTimeout(400);
  check(await p.evaluate(() => __pp.slides[1].shapes.find(x=>x.id==="9").name) === "Button", "Komponente gelöscht, Namen bereinigt");

  // ---- Anpassungen pro Instanz (Overrides) ----
  const shape = (sl, id) => p.evaluate(([sl,id]) => { const s = __pp.slides[sl].shapes.find(x=>x.id===String(id)); return { name: s.name, fill: s.fill.foregroundColor, text: s.textFrame.textRange.text, font: s.textFrame.textRange.font.color }; }, [sl,id]);
  await p.evaluate(() => { __pp.slides[0].shapes.push(__pp.mkShape({ id: 20, name: "Karte", fill: "#336699", text: "{{Label}} für {{Kunde}}" })); });
  await p.evaluate(() => __pp.select(0, [20])); await p.waitForTimeout(400);
  await p.selectOption("#comp-of", ""); await p.fill("#comp-name", "Karte"); await p.click("#comp-add button"); await p.waitForTimeout(500);
  await p.evaluate(() => { __pp.slides[1].shapes.push(__pp.mkShape({ id: 21, name: "◆ Karte ◆", text: "x" })); });
  await p.click("#sync-all"); await p.waitForTimeout(400);
  await p.evaluate(() => __pp.select(0, [20])); await p.waitForTimeout(400);
  check((await p.textContent("#sel-actions")).includes("Standardwerte"), "Master zeigt Standardwerte");
  await p.fill("#sel-actions .ovr input[type=text]", "Mehr erfahren"); await p.dispatchEvent("#sel-actions .ovr input[type=text]", "change"); await p.waitForTimeout(800);
  check((await shape(0,20)).text === "Mehr erfahren für Northwind", "Master-Standardwert: " + (await shape(0,20)).text);
  check((await shape(1,21)).text === "Mehr erfahren für Northwind", "Instanz erbt Standardwert: " + (await shape(1,21)).text);
  await p.evaluate(() => __pp.select(1, [21])); await p.waitForTimeout(400);
  check((await p.textContent("#sel-actions")).includes("Anpassungen dieser Instanz"), "Instanz zeigt Anpassungen");
  await p.fill("#sel-actions .ovr input[type=text]", "Jetzt kaufen"); await p.dispatchEvent("#sel-actions .ovr input[type=text]", "change"); await p.waitForTimeout(800);
  check((await shape(1,21)).text === "Jetzt kaufen für Northwind", "Instanz-Text überschrieben: " + (await shape(1,21)).text);
  check((await shape(0,20)).text === "Mehr erfahren für Northwind", "Master unverändert");
  await p.fill("#sel-actions .ovr input[type=color] >> nth=0", "#ff00ff"); await p.waitForTimeout(800);
  check((await shape(1,21)).fill === "#FF00FF", "eigene Füllfarbe: " + (await shape(1,21)).fill);
  await p.evaluate(() => __pp.slides[0].shapes.find(x=>x.id==="20").fill.foregroundColor = "#111111");
  await p.click("#sync-all"); await p.waitForTimeout(400);
  check((await shape(1,21)).fill === "#FF00FF", "eigene Farbe übersteht Sync");
  await p.evaluate(() => __pp.select(0, [2])); await p.waitForTimeout(300);
  await p.evaluate(() => __pp.select(1, [21])); await p.waitForTimeout(400);
  await p.click("#sel-actions button:has-text('Alle Anpassungen zurücksetzen')"); await p.waitForTimeout(800);
  const r = await shape(1,21);
  check(r.fill === "#111111" && r.text === "Mehr erfahren für Northwind", "Zurücksetzen: " + JSON.stringify(r));

  // ---- Varianten ----
  await p.evaluate(() => { __pp.slides[0].shapes.push(__pp.mkShape({ id: 22, name: "◆ Karte ◆", fill: "#008000", text: "{{Label}} für {{Kunde}}" })); });
  await p.evaluate(() => __pp.select(0, [22])); await p.waitForTimeout(400);
  await p.selectOption("#comp-of", "Karte"); await p.fill("#comp-name", "Grün"); await p.click("#comp-add button"); await p.waitForTimeout(500);
  check((await shape(0,22)).name === "◆ Karte / Grün ◆", "Variante angelegt: " + (await shape(0,22)).name);
  check((await p.textContent("#comp-list")).includes("Karte · Grün"), "Variante in Liste");
  await p.evaluate(() => __pp.select(1, [21])); await p.waitForTimeout(400);
  await p.fill("#sel-actions .ovr input[type=text]", "Los"); await p.dispatchEvent("#sel-actions .ovr input[type=text]", "change"); await p.waitForTimeout(800);
  await p.selectOption("#sel-actions select[aria-label=Variante]", "Karte / Grün"); await p.waitForTimeout(900);
  const v = await shape(1,21);
  check(v.name === "◆ Karte / Grün ◆" && v.fill === "#008000", "Variante gewechselt: " + JSON.stringify(v));
  check(v.text === "Los für Northwind", "Anpassung bleibt beim Variantenwechsel: " + v.text);

  check(errors.length === 0, "keine JS-Fehler " + errors.join(" | "));
  console.log("Status:", await st());
  if (process.env.SHOT) { await p.setViewportSize({ width: 340, height: 900 }); await p.evaluate(() => __pp.select(1, [21])); await p.waitForTimeout(400); await p.screenshot({ path: process.env.SHOT }); }
  await b.close();
})();
