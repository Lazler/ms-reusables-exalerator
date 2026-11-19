# Deck Components – Figma-style variables & components for PowerPoint

Consulting decks repeat the same things over and over: the client name, the project name, the date, the same callout box, the same "key takeaway" banner on 30 slides. Changing one of them means hunting through the whole deck.

**Deck Components** brings two ideas from Figma to PowerPoint:

| Figma | Deck Components |
|---|---|
| **Variables** | Type `{{Client}}` in any text box and set its value once in the side pane. Every occurrence in the deck updates. **Color variables** can be bound to the fill, line or text color of shapes. |
| **Components** | One shape becomes the **master**. Copies of it on any slide are **instances**. Change the master and all instances follow (fill, line, size, text, font). Each instance keeps its position. |
| **Instance overrides** | Each instance can have its own text values (e.g. `{{Label}}`) and its own colors. They survive syncing. |
| **Variants** | e.g. *Button / Default*, *Button / Secondary*, *Button / Red*. Switch an instance between variants from a dropdown. |

Everything is stored **inside the .pptx**: no server, no database, no account. The add-in itself is a set of static files.

The UI language follows PowerPoint (English, or German for German Office installations). You can switch it manually with **EN / DE** in the pane header.

> **Status: prototype (v0.4).** The logic and UI are covered by unit tests and by browser tests against a simulated PowerPoint API. Real-world testing in PowerPoint is ongoing; see [Known limitations](#known-limitations).

---

## How it works

### Variables
1. **Variables** tab → enter a name (e.g. `Client`) → **Add** → fill in the value.
2. Type `{{Client}}` in a text box, or click **Insert {{Client}}** to insert it at the cursor.
3. Change the value and every text box in the deck updates.

- **Show placeholders instead of values** switches the whole deck to `{{…}}` so you can edit the templates.
- If you edit a rendered text directly (e.g. "Offer for ACME" → "New offer for ACME"), the add-in recognizes the value and keeps the variable.
- **Find placeholders in deck** creates variables for every `{{…}}` already in the deck. This is handy for existing templates.
- **Color variables:** select shapes and click *Fill*, *Line* or *Text*. When the color changes, all bound shapes follow.

### Components
1. Select a shape or text box → **Components** tab → enter a name → **Create**.
   The shape is now named `◆ Name ◆` (visible in *Home → Arrange → Selection Pane*).
2. To create instances, **copy and paste** the master (Ctrl/Cmd + C/V or D), on any slide.
   Every shape named `◆ Name ◆` is an instance. You can also rename a shape in the Selection Pane, or select it and click **Link as instance**.
3. Change the master, then click **Sync**. You can also turn on **Auto-sync**, which syncs as soon as you deselect the master.

For each component you choose whether **text** and **size** are synced. You can **detach** a selected instance, or **make it the master**.

### Instance overrides
When you select an instance, the **Components** tab shows **Overrides for this instance**:
- **Text:** every placeholder in the master text (e.g. `{{Label}}`) gets its own field. A value entered here applies to this instance only. If you leave it empty, the master's default or the global value is used.
- **Colors:** fill, text color and line can be overridden per instance. *Reset* restores the master's color.
- When you select the **master**, the box shows the **component defaults** that all instances inherit.

Precedence: **instance → master default → global variable**.

### Variants
1. Copy the master and change the copy (e.g. make it red).
2. Select the copy → *New component / variant* → choose **Variant of Button** → name it `Red` → **Create**.
   The shape is renamed `◆ Button / Red ◆` and becomes the master of that variant.
3. Select any instance and switch the **Variant** dropdown. The instance keeps its overrides.

---

## Setup

Office add-ins must be served over **HTTPS**. This repository publishes the add-in through GitHub Pages:
`https://lazler.github.io/ms-reusables-exalerator/`

### One-time: enable GitHub Pages
Repository → *Settings → Pages → Build and deployment* → Source **Deploy from a branch** → branch **`gh-pages`**, folder **`/ (root)`** → *Save*.

On every push to `main`, `.github/workflows/pages.yml` runs the tests, stamps the build number into the page, and publishes `addin/` to the `gh-pages` branch.

### Load the add-in in PowerPoint (sideloading)
1. Download [`addin/manifest.xml`](https://raw.githubusercontent.com/Lazler/ms-reusables-exalerator/main/addin/manifest.xml).
2. **PowerPoint on the web:** open a presentation → *Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in* → choose the manifest.
   **Mac:** copy the manifest to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` and restart PowerPoint.
3. On the *Home* tab, click **Variables & Components**. The pane header shows the version, e.g. `v0.4.0 · a1b2c3d`.

Requirement: **PowerPointApi 1.5**, which means PowerPoint on the web or a current Microsoft 365.

> **Corporate environments:** many organizations block sideloading. For team-wide use, an admin can deploy the manifest centrally via the Microsoft 365 admin center (*Integrated apps*). The add-in only reads and writes the open presentation and sends no data anywhere.

### After an update
Office caches add-ins aggressively. If the version in the pane header doesn't change:
- **Web:** close the pane and reload the browser tab with Ctrl + F5.
- If the **manifest** version changed (see `scripts/make-manifests.py`), remove the add-in under *My Add-ins* and upload the new manifest.

### Local development (without GitHub Pages)
```bash
npm run certs   # once: install a local HTTPS certificate
npm start       # serves addin/ at https://localhost:3000
```
Then upload `manifest.local.xml` instead of `addin/manifest.xml`.

---

## Known limitations
- Only shapes and text boxes are supported. **Groups, pictures, tables and SmartArt** are not synced yet.
- **Mixed formatting** inside one text box (e.g. a single bold word) is lost when the text is rewritten.
- Gradient and picture fills are not synced, only solid colors and "no fill".
- Variables and components live **per presentation**. A shared library across decks (like Figma libraries) is the next big step.

## Development

```
addin/
  manifest.xml             add-in manifest (GitHub Pages) – generated
  taskpane.html/.css/.js   side pane (Office.js, no build step)
  core.js                  pure logic: placeholders, templates, naming, overrides
  i18n.js                  UI strings (English primary, German secondary)
  assets/                  icons
manifest.local.xml         manifest for https://localhost:3000 – generated
scripts/make-manifests.py  generates both manifests (bump VERSION here)
tests/
  core.test.js, i18n.test.js   unit tests (node --test)
  e2e/                         side pane in a headless browser against a simulated PowerPoint API
```

```bash
npm test          # unit tests
npm run test:e2e  # requires Playwright + Chromium
```

### Adding a language
Add a block to `STRINGS` in `addin/i18n.js` and extend `detect()`. `tests/i18n.test.js` checks that every language has the same keys and placeholders as English.

### How data is stored
- List of variables and components: `Office.context.document.settings` (inside the .pptx).
- Per text box: hidden shape tags `DC_TPL` (template with `{{…}}`) and `DC_OUT` (last text written, used to detect manual edits).
- Color bindings: tags `DC_FILL`, `DC_LINE`, `DC_FONT`.
- Instance overrides and master defaults: tag `DC_OVR` (JSON).
- Component membership: the **shape name** `◆ Name ◆` or `◆ Name / Variant ◆`. It travels with copy & paste and can be edited in the Selection Pane. The master is identified by slide ID and shape ID, and a snapshot of its style is kept in case it is deleted.
