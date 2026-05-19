# Container Arrival Tracker

> A lightweight, zero-infrastructure tool for warehouse, customs, and procurement teams to track container arrivals and share documents in real time.

**Built by Testa Wu — Eshcol Studio | Build. Model. Deliver.**

---

## The Problem

In traditional manufacturing and trading companies, container arrival information is fragmented across three departments:

```
Procurement    → has invoice and packing list
Customs        → knows the customs clearance date
Warehouse      → has no idea when the container arrives
                 and no documents to verify cargo
```

The result: warehouse staff arrive unprepared. Cargo verification is manual and error-prone. Demurrage fees accumulate from delays.

This tool closes that gap.

---

## What It Does

A single HTML file that three departments share. When procurement uploads the invoice and packing list, and customs logs the arrival date, the warehouse sees everything — in one place, in real time.

**No server required for v1. No installation. No IT department.**

---

## Features

| Feature | Description |
|---|---|
| Cargo entry form | 12 fields covering all arrival information |
| Document upload | Invoice and packing list file attachment |
| Interior photo | Container condition documentation after unloading |
| Search | Query by cargo name, container no., invoice no., project code |
| Local Excel sync | Auto read/write a local .xlsx file (Chrome / Edge) |
| Google Sheets sync | Shared database across all departments |
| Excel export | One-click export of all records |
| Local storage | Works offline, data persists in browser |
| Stats dashboard | Total, this month, pending, arrived |

---

## Three Storage Layers

Choose the one that fits your situation, or use all three together.

```
Layer 1 — Browser localStorage
Always on. No setup. Data stays in the browser.
Works offline. Single user only.

Layer 2 — Local Excel File (Chrome / Edge only)
Select or create a .xlsx file on your computer.
Every save / edit / delete writes back to the same file automatically.
No download popup. Single user or shared via NAS / network folder.

Layer 3 — Google Sheets
Multi-user shared database. Free. No server needed.
Requires Google account and one-time setup.
All departments access the same data in real time.
```

---

## Quick Start — Option A: Local Excel File

Best for: single user, or teams sharing a network folder.

Requires Chrome or Edge.

1. Download `container-arrival-tracker-v1.html`
2. Open it in Chrome or Edge
3. In the **Local Excel File** panel, click **Create New File**
4. Choose where to save the `.xlsx` file (e.g. a shared network folder)
5. Start adding records — every change is written to the Excel file automatically

**Next time you open the tool:**
1. Click **Select Excel File**
2. Choose the same `.xlsx` file
3. All records load automatically

If you save the Excel file in a shared network folder (NAS, company server), all team members can use the same file by selecting it from their own browser.

---

## Quick Start — Option B: Google Sheets

Best for: multi-user teams, remote access, real-time sharing.

### Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet
2. Rename the first tab to `Records`
3. Share the sheet with all departments

### Step 2 — Get your Spreadsheet ID

From your Google Sheet URL:
```
https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
```
Copy the `SPREADSHEET_ID` part.

### Step 3 — Create a Google Cloud API Key

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use an existing one)
3. Enable **Google Sheets API**
4. Go to **Credentials** → **Create Credentials** → **API Key**
5. Restrict the key to **Google Sheets API** only (recommended)

### Step 4 — Enable write access via Apps Script

API Keys allow read-only access. For write access (save / edit / delete), set up the Apps Script Web App below.

1. In your Google Sheet, go to **Extensions → Apps Script**
2. Paste this code:

```javascript
const SHEET_NAME = 'Records';

const HEADERS = [
  'ID', 'Cargo Name', 'Project Code', 'Invoice No.', 'Packing List No.',
  'Customs Date', 'Arrival Date', 'Yard Zone', 'Container No.',
  'Invoice File', 'Packing List File', 'Has Photo', 'Created At'
];

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === 'append') {
    sheet.appendRow([
      data.id, data.cargo, data.project, data.invoiceNo, data.plNo,
      data.customsDate, data.arrivalDate, data.zone, data.containerNo,
      data.invoiceFile, data.plFile, data.hasPhoto, data.createdAt
    ]);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'update') {
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 1, 1, 13).setValues([[
          data.id, data.cargo, data.project, data.invoiceNo, data.plNo,
          data.customsDate, data.arrivalDate, data.zone, data.containerNo,
          data.invoiceFile, data.plFile, data.hasPhoto, data.createdAt
        ]]);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'delete') {
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify({ data: values }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

3. Click **Deploy → New Deployment**
4. Type: **Web App**
5. Execute as: **Me**
6. Who has access: **Anyone**
7. Click **Deploy** and copy the **Web App URL**

### Step 5 — Configure the tracker

Open `container-arrival-tracker-v1.html` in your browser.

In the **Google Sheets Config** panel:
- Paste your **Spreadsheet ID**
- Paste your **API Key**
- Paste your **Apps Script Web App URL**
- Enter your **Sheet Name** (default: `Records`)
- Click **Test** to verify the connection

### Step 6 — Share with your team

**Option A — Company file server / NAS**
Place the HTML file in a shared network folder. All users open it from the same location.

**Option B — GitHub Pages (free)**
Fork this repository, enable GitHub Pages, and share the URL.

**Option C — Cloudflare Workers (free)**
Upload the HTML file as a static asset. See [Cloudflare Workers documentation](https://developers.cloudflare.com/workers/).

---

## Browser Compatibility

| Feature | Chrome | Edge | Firefox | Safari |
|---|---|---|---|---|
| All basic features | Yes | Yes | Yes | Yes |
| Local Excel File | Yes | Yes | No | No |
| Google Sheets | Yes | Yes | Yes | Yes |

---

## File Structure

```
container-arrival-tracker/
|
|-- v1-google-sheets/
|   |-- container-arrival-tracker-v1.html    Main application
|   `-- apps-script.gs                       Google Apps Script code
|
|-- v2-backend/                              Coming soon
|   |-- frontend/
|   |   `-- index.html
|   |-- backend/
|   |   |-- main.py                          FastAPI
|   |   |-- database.py                      SQLite / PostgreSQL
|   |   |-- models.py
|   |   `-- requirements.txt
|   |-- docker-compose.yml
|   `-- README.md
|
`-- README.md
```

---

## Record Fields

| Field | Required | Description |
|---|---|---|
| Cargo Name | Yes | Description of the goods |
| Project Code | No | Internal project reference |
| Invoice No. | Yes | Supplier invoice number |
| Invoice File | No | PDF or image upload |
| Packing List No. | No | Packing list reference number |
| Packing List File | No | PDF or image upload |
| Customs Date | No | Date customs clearance was filed |
| Arrival Date | Yes | Expected or actual container arrival date |
| Yard Zone | No | Storage area designation |
| Container No. | Yes | ISO container number (e.g. ABCU1234567) |
| Interior Photo | No | Photo of container interior after unloading |

---

## Roadmap

### v1 — Current
- Single HTML file, zero dependencies
- Three storage layers: localStorage, Local Excel, Google Sheets
- File System Access API for automatic local Excel read/write
- Google Sheets read via API Key, write via Apps Script
- Excel export
- Stats dashboard

### v2 — Backend (planned)
- FastAPI + SQLite / PostgreSQL
- Multi-user authentication
- Role-based access (procurement / customs / warehouse)
- File storage (local or S3-compatible)
- REST API for ERP integration
- Docker Compose for self-hosted deployment
- Audit log (who changed what, when)

### v3 — Notifications (future)
- Email / LINE notification when arrival date is approaching
- Slack / Teams webhook integration
- Mobile-optimized view

---

## Why This Exists

This tool was designed from direct warehouse operations experience.

The problem is not technical. It is organizational.

Procurement has the documents. Customs has the schedule. Warehouse has neither.

The standard response in traditional industries is: that is not my department. The result is warehouse staff waiting without information, cargo arriving without documentation, and containers incurring demurrage fees while teams scramble to find paperwork.

A shared record system with clear ownership at each step eliminates this entirely.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | HTML, CSS, JavaScript (vanilla) |
| File read/write | File System Access API |
| File export | SheetJS (xlsx) |
| Database (v1) | Google Sheets via Sheets API |
| Write access (v1) | Google Apps Script Web App |
| Local fallback | localStorage |
| Backend (v2) | Python + FastAPI |
| Database (v2) | SQLite (dev) / PostgreSQL (prod) |
| Deployment | Cloudflare Workers / GitHub Pages / Docker |

---

## Contributing

Issues and pull requests are welcome.

If you are using this in production and have feedback — especially from warehouse, customs, or procurement teams — please open an issue. Real operations feedback shapes the roadmap.

---

## License

MIT License — free to use, modify, and distribute.

---

## Author

**Testa Wu**
Eshcol Studio — Build. Model. Deliver.

[buildmodeldeliver.com](https://buildmodeldeliver.com) · [LinkedIn](https://linkedin.com/in/testa-wu) · [Medium](https://medium.com/@testa-wu)

> "The UI is not a report — it is a decision surface."
