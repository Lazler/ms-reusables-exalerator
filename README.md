# Deck Components – Variablen & Komponenten für PowerPoint

Ein PowerPoint-Add-in, das zwei Ideen aus Figma nach PowerPoint bringt:

| Figma | Deck Components |
|---|---|
| **Variables** | `{{Kunde}}` in ein beliebiges Textfeld schreiben und den Wert zentral im Seitenbereich ändern. Alle Stellen im Deck werden aktualisiert. Dazu kommen **Farbvariablen**, die du an Füllung, Linie oder Schrift von Formen bindest. |
| **Components** | Eine Form wird zum **Master**. Kopien davon auf beliebigen Folien sind **Instanzen**. Änderst du den Master, übernehmen alle Instanzen Füllung, Linie, Größe, Text und Schrift. Die Position jeder Instanz bleibt. |

Alles wird **in der .pptx** gespeichert, ohne Server und ohne Datenbank. Das Add-in besteht nur aus statischen Dateien.

> Status: **Prototyp** (v0.3). Logik und Oberfläche sind gegen eine nachgebaute PowerPoint-API getestet, aber noch nicht in echtem PowerPoint. Siehe [Testplan](#erster-test-in-powerpoint-bitte-in-dieser-reihenfolge).

---

## So funktioniert's

### Variablen
1. Tab **Variablen** → Name (z. B. `Kunde`) → **Hinzufügen** → Wert eintragen.
2. `{{Kunde}}` in ein Textfeld tippen oder mit dem Button **„{{Kunde}} einfügen“** an der Cursorposition einsetzen.
3. Wert ändern: Alle Textfelder im Deck werden automatisch aktualisiert.

- **Platzhalter statt Werte anzeigen** schaltet das ganze Deck auf `{{…}}` um. So siehst und bearbeitest du die Vorlagen.
- Wenn du einen gerenderten Text direkt bearbeitest (z. B. „Angebot für ACME“ → „Neues Angebot für ACME“), erkennt das Add-in den Wert wieder und behält die Variable.
- **Platzhalter im Deck finden** legt alle `{{…}}`, die im Deck vorkommen, als Variablen an. Das ist praktisch für bestehende Vorlagen.
- **Farbvariablen:** Formen markieren und auf *Füllung*, *Linie* oder *Schrift* klicken. Ändert sich die Farbe, folgen alle gebundenen Formen.

### Komponenten
1. Eine Form oder ein Textfeld markieren → Tab **Komponenten** → Name → **Als Komponente anlegen**.
   Die Form heißt jetzt `◆ Name ◆` (sichtbar unter *Start → Anordnen → Auswahlbereich*).
2. Instanzen anlegen: den Master **kopieren und einfügen** (Strg/Cmd + C/V oder D), auch auf anderen Folien.
   Jede Form mit dem Namen `◆ Name ◆` ist eine Instanz. Du kannst eine Form auch im Auswahlbereich so benennen
   oder sie markieren und **„Als Instanz verknüpfen“** klicken.
3. Den Master ändern → **Synchronisieren**. Alternativ **Auto-Sync** einschalten: Dann wird synchronisiert, sobald du den Master abwählst.

Pro Komponente lässt sich festlegen, ob **Text** und **Größe** übernommen werden.
Eine markierte Instanz kann **gelöst** werden (Detach) oder **zum Master** werden.

### Instanzen anpassen (wie Overrides in Figma)
Markierst du eine Instanz, zeigt der Tab *Komponenten* den Bereich **„Anpassungen dieser Instanz“**:
- **Texte:** Jeder Platzhalter im Master-Text, z. B. `{{Label}}`, bekommt ein eigenes Feld. Trägst du etwas ein, gilt der Wert nur für diese Instanz. Lässt du das Feld leer, gilt der Wert vom Master bzw. der globale Wert.
  So hat jeder Button denselben Stil, aber einen eigenen Text.
- **Farben:** Füllung, Schriftfarbe und Linie lassen sich pro Instanz überschreiben. Eigene Farben bleiben beim Synchronisieren erhalten, und *Zurücksetzen* holt wieder die Farbe des Masters.
- Markierst du den **Master**, setzt du dort die **Standardwerte** (z. B. `Label = Mehr erfahren`), die alle Instanzen ohne eigenen Wert bekommen.

Reihenfolge, welcher Wert gewinnt: **Instanz → Standardwert im Master → globale Variable**.

### Varianten (Komponentenarten)
Beispiel: Button in *Standard*, *Sekundär* und *Rot*.
1. Eine Kopie des Masters anlegen und abwandeln (z. B. rot einfärben).
2. Kopie markieren → bei *Neue Komponente / Variante* **„Variante von Button“** wählen → Name `Rot` → **Anlegen**.
   Die Form heißt jetzt `◆ Button / Rot ◆` und ist Master dieser Variante.
3. Eine Instanz markieren → im Dropdown **Variante** umschalten. Sie übernimmt sofort den Stil der Variante. Eigene Texte und Farben bleiben dabei erhalten.

---

## Einrichtung (ca. 10 Minuten)

Office-Add-ins müssen über **HTTPS** ausgeliefert werden. Am einfachsten geht das mit GitHub Pages. Das Manifest zeigt bereits auf
`https://lazler.github.io/ms-reusables-exalerator/`.

1. **GitHub Pages aktivieren:** Repo → *Settings → Pages → Source: „GitHub Actions“*.
   Bei einem privaten Repo setzt Pages einen kostenpflichtigen GitHub-Plan voraus. Alternativ das Repo öffentlich machen (es enthält keine Daten) oder die [lokale Variante](#alternative-lokal-ohne-github-pages) nutzen.
2. Den Code auf `main` bringen. Der Workflow `.github/workflows/pages.yml` führt die Tests aus und veröffentlicht den Ordner `addin/`.
   Prüfen: https://lazler.github.io/ms-reusables-exalerator/taskpane.html lädt.
3. **Add-in in PowerPoint laden (Sideloading)**, nur für dich und ohne Store:
   - **PowerPoint im Web:** Präsentation öffnen → *Start → Add-Ins → Weitere Add-Ins → Meine Add-Ins → „Mein Add-In hochladen“* → `addin/manifest.xml`.
   - **Mac:** `addin/manifest.xml` nach `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` kopieren (Ordner ggf. anlegen) → PowerPoint neu starten → *Start → Add-Ins*.
   - **Windows:** Netzwerkfreigabe als „Vertrauenswürdigen Add-In-Katalog“ eintragen (*Datei → Optionen → Trust Center → Kataloge*) oder PowerPoint im Web nutzen.
4. Im Menüband unter *Start* erscheint **Deck Components → Variablen & Komponenten**.

Voraussetzung: **PowerPointApi 1.5**, also PowerPoint im Web oder ein aktuelles Microsoft 365. Ältere Versionen zeigen einen Hinweis im Seitenbereich.

### Alternative: lokal ohne GitHub Pages
```bash
npm run certs   # einmalig: lokales HTTPS-Zertifikat installieren
npm start       # https://localhost:3000
```
Dann `manifest.local.xml` statt `addin/manifest.xml` hochladen.

---

## Erster Test in PowerPoint (bitte in dieser Reihenfolge)

1. **Komponenten-Kern:** Komponente anlegen → Master mit Strg+C/Strg+V auf eine andere Folie kopieren → im Auswahlbereich prüfen: Heißt die Kopie auch `◆ Name ◆`? Davon hängt das ganze Komponenten-Konzept ab. Falls PowerPoint den Namen ändert, schick mir, wie die Kopie heißt.
2. Master-Farbe ändern → *Synchronisieren* → Folgt die Kopie?
3. `{{Kunde}}` in ein Textfeld schreiben, Wert setzen → wird ersetzt? Datei speichern, schließen, neu öffnen → sind Variablen und Komponenten noch da?

Bei Fehlern steht im grauen Statusbereich unten eine Meldung (inkl. Code und Ort). Die bitte einfach weitergeben.

## Grenzen des Prototyps
- Unterstützt nur Formen und Textfelder. **Gruppen, Bilder, Tabellen und SmartArt** werden noch nicht synchronisiert.
- **Gemischte Formatierung** innerhalb eines Textfelds (z. B. ein fettes Wort) geht verloren, wenn der Text neu gesetzt wird.
- Verlaufs- und Bildfüllungen werden nicht übertragen, nur Volltonfarben und „keine Füllung“.
- Variablen und Komponenten gelten **pro Präsentation**. Eine geteilte Bibliothek über mehrere Decks hinweg (wie Figma-Libraries) wäre der nächste große Schritt.
- Bitte vorerst nur privat testen (eigenes Microsoft-Konto, eigenes Gerät), nicht mit Kundendecks auf dem Firmenlaptop.

## Entwicklung

```
addin/
  manifest.xml      Add-in-Manifest (GitHub Pages)
  taskpane.html/.css/.js   Seitenbereich (Office.js, kein Build-Schritt)
  core.js           reine Logik: Platzhalter, Templates, Namenskonvention
  assets/           Icons
manifest.local.xml  Manifest für https://localhost:3000
tests/
  core.test.js      Unit-Tests (node --test)
  e2e/              Seitenbereich im Headless-Browser gegen eine Mock-PowerPoint-API
```

```bash
npm test          # Unit-Tests
npm run test:e2e  # braucht Playwright + Chromium
```

### Wie die Daten gespeichert werden
- Liste der Variablen und Komponenten: `Office.context.document.settings` (liegt in der .pptx).
- Pro Textfeld: unsichtbare Shape-Tags `DC_TPL` (Vorlage mit `{{…}}`) und `DC_OUT` (zuletzt geschriebener Text, um Bearbeitungen zu erkennen).
- Farbbindungen: Tags `DC_FILL`, `DC_LINE`, `DC_FONT`.
- Anpassungen pro Instanz bzw. Standardwerte im Master: Tag `DC_OVR` (JSON).
- Komponenten-Zugehörigkeit: der **Formname** `◆ Name ◆` bzw. `◆ Name / Variante ◆`. Er wandert beim Kopieren mit und ist im Auswahlbereich sichtbar und editierbar. Der Master wird über Folien- und Form-ID identifiziert. Zusätzlich gibt es einen Snapshot seines Stils, falls er gelöscht wird.
