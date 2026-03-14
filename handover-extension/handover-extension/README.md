# Handover Importer — Chrome Extension

Drag & drop a handover `.xlsx` onto the floating panel while the Smartsheet form is open — fields fill automatically.

---

## ⚡ One-time setup: download the Excel library

Before loading the extension, you need to add one file:

1. Go to: https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
2. Save the page as **`xlsx.min.js`**
3. Place it inside this folder (next to `manifest.json`)

---

## 🔧 Install in Chrome

1. Open Chrome and go to **chrome://extensions**
2. Enable **Developer mode** (toggle, top-right)
3. Click **Load unpacked**
4. Select this folder (`handover-extension/`)
5. The 📋 icon will appear in your Chrome toolbar

> **Pin it:** Click the puzzle piece icon in Chrome toolbar → pin "Handover Importer" for easy access.

---

## 👥 Sharing with your team (2–10 people)

**Option A — Share the folder directly (easiest)**
- Zip this folder and send it to teammates
- Each person follows the Install steps above
- Note: Chrome may show an "unpacked extension" warning — this is normal for team-shared extensions

**Option B — Publish to Chrome Web Store (private)**
- Go to https://chrome.google.com/webstore/devconsole
- Pay the one-time $5 developer fee
- Upload a zip of this folder
- Set visibility to **Unlisted** — share the link only with your team
- Updates are pushed automatically

---

## 📋 How to use

1. Open the Smartsheet form: https://app.smartsheet.com/b/form/0195fd4e00797b80a1fe7a314e5471ec
2. Click the **📋 Handover Importer** button (bottom-right of the page)
3. Drag & drop your handover `.xlsx` into the panel (or click to browse)
4. Review the extracted fields in the panel
5. Click **⚡ Fill Form Fields**
6. Fields populate automatically — review and hit **Submit** on the form

---

## 📌 Fields populated automatically

| Form Field             | Source in Handover Doc          |
|------------------------|---------------------------------|
| SO #                   | Cell B2 (strips "SO " prefix)   |
| Job Name               | Cell B3                         |
| Customer Name          | Cell B35 (before first comma)   |
| Color                  | Cell B34 (reformatted)          |
| Series                 | Cell B22 (extracts "Flow HPL")  |
| Edging                 | Cell B22 (extracts "HPL")       |
| Net USD Value          | F22 + F23 + F24                 |
| Customer Request Date  | Cell C32                        |
| Special Materials?     | Cell B40                        |

## ✏️ Fields to fill manually on the form

- Unit Stalls
- Units Urinal
- Floor Gap
- Handover Document Date

---

## 🐛 Troubleshooting

**Fields not filling?**
Smartsheet may update their form HTML. Open the browser console (F12) on the form page and check for errors. The extension uses multiple strategies to find fields by label text — if the form structure changes significantly, `content.js` may need a selector update.

**Extension not appearing?**
Make sure you're on `https://app.smartsheet.com/b/form/...` — the extension only activates on Smartsheet form URLs.
