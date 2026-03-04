# Cascade Tools — Excel Add-in

A lightweight Office.js add-in with two features:

1. **Cascade Sheet** — duplicate the active sheet for every value in a named range, replacing a target cell on each copy.
2. **Row Groups** — select rows, tag them with a group name (or `#hide`), and toggle visibility per group.

No VBA. Pure JavaScript. Works on Windows, Mac, and Excel Online.

---

## Quick Start (local, no hosting needed)

### 1. Install a local HTTPS server

Office.js requires HTTPS. The easiest way:

```bash
npm install -g office-addin-dev-certs http-server
npx office-addin-dev-certs install   # creates trusted localhost certs
```

### 2. Serve the files

Open a terminal in this folder and run:

```bash
npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```

Leave this running.

### 3. Sideload the add-in

**Windows (Excel desktop):**
1. Open Excel → **Insert** tab → **My Add-ins** → **Upload My Add-in**
2. Browse to `manifest.xml` in this folder → **Upload**
3. A new **Cascade Tools** tab appears in the ribbon.

**Mac (Excel desktop):**
1. Copy `manifest.xml` to:
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
   (Create the `wef` folder if it doesn't exist.)
2. Restart Excel. The tab will appear.

**Excel Online:**
1. Open a workbook → **Insert** → **Office Add-ins** → **Upload My Add-in**
2. Upload `manifest.xml`.

---

## Deploying for your team (GitHub Pages)

If you want a permanent URL so teammates don't need to run a local server:

### 1. Create a GitHub repo
Push this entire folder as a repo (e.g., `cascade-tools`).

### 2. Enable GitHub Pages
Go to **Settings → Pages** → set source to `main` branch, root folder.

### 3. Update manifest.xml
Replace every `https://localhost:3000` with your Pages URL:
```
https://YOUR_USERNAME.github.io/cascade-tools
```

### 4. Share the manifest
Send `manifest.xml` to teammates. They sideload it the same way (step 3 above). The JS/HTML/CSS loads from GitHub Pages automatically.

---

## How to use

### Cascade Sheet
1. Open the workbook that has a **named range** with your list values.
2. Go to the sheet you want to cascade.
3. Click on the cell you want each copy to vary (e.g., a company name in B2).
4. Open the **Cascade Tools** panel from the ribbon.
5. Click **Pick** (or the cell address auto-fills).
6. Select the named range from the dropdown.
7. Preview the values, then click **Cascade**.
8. One copy of the sheet is created per value, with the target cell replaced.

### Row Groups
1. Select one or more rows on your sheet.
2. Open the panel → **Row Groups** tab.
3. Type a group name (e.g., `Detail`, `#hide`, `Assumptions`).
4. Click **Tag selected**.
5. Use the eye / show / hide buttons next to each group to control visibility.
6. **Show all rows** unhides everything.
7. **Remove tags from selection** untags rows you've selected.

---

## Project structure

```
cascade-tools/
├── manifest.xml      ← Office add-in manifest (URLs point to your host)
├── taskpane.html      ← Side-panel UI
├── taskpane.css       ← Styles
├── taskpane.js        ← All logic (cascade + row groups)
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
└── README.md
```

---

## Requirements

- Excel 2019 or later (desktop), or Excel Online
- ExcelApi 1.7+ (ships with Excel 2019+)
- HTTPS for local dev (see Quick Start above)

---

## Adding more features later

The codebase is designed to be extended. To add a new feature:

1. Add a new tab button in `taskpane.html`
2. Add a `<section id="tab-yourfeature">` content block
3. Write your logic in `taskpane.js` using the same `Excel.run()` pattern
4. Wire it up in an `initYourFeature()` function called from `Office.onReady`
