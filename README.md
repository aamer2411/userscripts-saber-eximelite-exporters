# SABER and EximElite Excel Exporters

Tampermonkey userscripts that turn two paged, export-less web portals into one-click Excel downloads: the **SABER** Contact Officer portal (Saudi product conformity certificates) and the **EximElite** EDA-EX Power BI report (Indian export shipment data). Everything runs in your browser, inside your own logged-in session.

![Userscript](https://img.shields.io/badge/Userscript-Tampermonkey-00485B?logo=tampermonkey&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-Vanilla-F7DF1E?logo=javascript&logoColor=black)
![Output](https://img.shields.io/badge/Output-Excel%20.xlsx-217346?logo=microsoftexcel&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blue)

---

## Scripts

Pick **one** of the options below. The combined script contains both tools, so don't install it next to the single-site ones.

| Script | Runs on | Install |
|---|---|---|
| **SABER Conformity Certificates Exporter** | `saber.sa/ContactOfficer/*` | [Install](https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/saber-coc-detail-exporter.user.js) |
| **EximElite EDA-EX Bulk Exporter** | `eximelite.com`, `app.powerbi.com/view*` | [Install](https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/eximelite-exporter.user.js) |
| **Combined (both tools)** | All of the above, picks the right tool per site | [Install](https://raw.githubusercontent.com/aamer2411/userscripts-saber-eximelite-exporters/main/scripts/eximelite-saber-combined.user.js) |

---

## Features

### SABER Conformity Certificates Exporter

- **Every certificate, full detail:** pages the whole Conformity Certificates list, opens each certificate's detail page and exports 16 fields per certificate, including Certification Body, Certificate Number, Issue and Expire Date, Certificate Status, Product, Trade Mark, Model, HS Code, Technical Regulation and Manufacturer.
- **Download-only certificates too:** some certificates have no detail page, only a PDF download. The script fetches each PDF and reads its fields in the browser. Download-only certificates that aren't a certificate PDF (e.g. IECEE) are exported with the list's number, dates, product, model and status.
- **Handles locked certificates:** SABER hides a certificate's detail page until its review form is dealt with. The script presses the review page's **Skip** for you (default), the same action a person does by hand. An optional mode submits a 5-star review instead.
- **Clean workbook:** one sheet, a `Sl No` column, rows sorted by Expire Date (soonest first), formatted as an Excel table.
- **Safe by design:** everything is a read except the review skip. No other action link (Cancel, Delete and so on) is ever touched. If the session expires, the run stops and asks you to log in again.
- **Gentle on a slow site:** detail pages and PDFs are fetched 3 at a time.

Output file: `SABER_Conformity_Certificates_YYYYMMDD.xlsx`

### EximElite EDA-EX Bulk Exporter

- **All rows, not just what's on screen:** reuses the report's own Power BI query (with your current slicer filters) to pull every shipment-level row, paged 30,000 rows at a time.
- **Exact row counts:** shipments that are identical in every column are kept as separate rows, so the export matches the report's shipment count.
- **Ready-to-use workbook:**

  | Sheet | Contents |
  |---|---|
  | `Shipments` | One row per shipment, 13 core columns (date, exporter, HS code, consignee, contact details, product) |
  | `KPI_1` | Year, Month, Exporter, HS Code, Consignee report with export, HS code and consignee counts |
  | `KPI_2` | Exporter, Year, Month, Day, Consignee report with monthly shipment counts and the HS codes shipped each day |
  | `KPI Pivot Tables` | Native Excel pivot tables over both KPI sheets, collapsed by default |
  | `Name Mapping` | Company names merged from several spellings (e.g. `Pvt Ltd` vs `Private Limited`); same name only counts as one company when the pin code (exporters) or foreign port (consignees) also matches |
  | `Raw Data` | All 31 columns, adding ports, countries, quantities, currency and value fields |

- **Descriptive file names** built from the active filters: `Year_Month_Chapter_HSN_Country.xlsx` (e.g. `2025_All-Months_All-Chapters_8517_All-Countries.xlsx`).
- **Works in both frames:** the button sits on the EximElite page while the export runs inside the embedded Power BI report.

---

## Installation

1. Install the [Tampermonkey](https://www.tampermonkey.net/) extension for Chrome, Edge or Firefox.
2. In Chrome or Edge, allow userscripts: open the extension's details page and turn on **Allow user scripts** (or **Developer mode** on older versions).
3. Click an **Install** link in the [Scripts](#scripts) table and confirm in Tampermonkey.

Tampermonkey checks this repository for updates automatically.

---

## Usage

1. Log in to the site as usual.
2. Open the right page:
   - **SABER:** any page under Contact Officer (the script reads the Conformity Certificates list itself).
   - **EximElite:** the EDA-EX report. Set the slicers (year, month, HS code, country) to the data you want.
3. Click the round green download button in the top-right corner (hovering it shows what it does).
4. Watch progress in the status pill next to it. The `.xlsx` lands in your Downloads folder.

---

## Configuration

Settings are constants near the top of each script (Tampermonkey dashboard, then **Edit**).

| Script | Constant | Default | Purpose |
|---|---|---|---|
| SABER | `UNLOCK_MODE` | `'skip'` | `'skip'` presses Skip on locked reviews; `'submit5'` sends a 5-star review |
| SABER | `DETAIL_CONCURRENCY` | `3` | Detail pages and PDFs fetched in parallel (keep low, the site is slow) |
| SABER | `COLUMNS` | 16 fields | Exported fields, order and headers |
| EximElite | `COLUMNS` | 13 fields | Columns on the Shipments and KPI sheets |
| EximElite | `RAW_EXTRA_COLUMNS` | 18 fields | Extra columns on the Raw Data sheet only |
| EximElite | `PAGE_ROWS` | `30000` | Rows requested per query page |

Editing a script in Tampermonkey stops automatic updates for it. To keep your changes and still get updates, fork the repository and point `@updateURL` at your fork.

---

## How It Works

**SABER.** The portal is a server-rendered ASP.NET MVC site with no API and no export button. The list shows 7 certificates per page and ignores page-size parameters. The script fetches the list pages with your session cookie, unlocks locked certificates by re-submitting the review page's own form (including its anti-forgery token) to the Skip action, re-reads the list to get every detail link, then parses each detail page into a row. Certificate status comes from the list. For download-only certificates it fetches the certificate PDF and reads it with pdf.js: each value is picked by its position next to its label on page 1.

**EximElite.** The page embeds a publish-to-web Power BI report. The script hooks the report's own `querydata` request to capture its key and filters, then replays it with every detail column selected. It pages through the result, decodes Power BI's compressed response format, and writes the workbook with SheetJS. Table styling, column widths and pivot tables are added by editing the `.xlsx` package with JSZip.

**Combined.** Both scripts are wrapped unchanged in their own functions, and a hostname check decides which one runs. It is generated by `tools/build-combined.js`, so the three files never drift apart.

| Library | Version | Use |
|---|---|---|
| [SheetJS](https://sheetjs.com/) (`xlsx`) | 0.18.5 | Build the workbook |
| [JSZip](https://stuk.github.io/jszip/) | 3.10.1 | Add tables, styles and pivots to the `.xlsx` |
| [pdf.js](https://mozilla.github.io/pdf.js/) | 3.11.174 | Read certificate PDFs (SABER only) |

All are loaded from cdnjs by Tampermonkey (`@require`), pinned to exact versions.

---

## Development

The two single-site scripts are the source of truth. The combined script is generated from them, so never edit it by hand.

```bash
git clone https://github.com/aamer2411/userscripts-saber-eximelite-exporters.git
cd userscripts-saber-eximelite-exporters

# after editing a script in scripts/
npm run build   # regenerate scripts/eximelite-saber-combined.user.js
npm run check   # syntax-check all three scripts
```

Only Node.js is needed, with no dependencies to install. Bump `@version` in a script's header when you change it, so Tampermonkey picks up the update.

---

## Project Structure

```
.
├── scripts/
│   ├── saber-coc-detail-exporter.user.js    SABER exporter (source)
│   ├── eximelite-exporter.user.js           EximElite exporter (source)
│   └── eximelite-saber-combined.user.js     Both tools in one (generated)
├── tools/
│   └── build-combined.js                    Builds the combined script
├── package.json                             build / check commands
└── LICENSE
```

---

## Disclaimer

These are independent tools. They are not affiliated with or endorsed by SABER, SASO, EximElite or Microsoft. They only read data your account can already see. Use them in line with each site's terms of use and your organisation's policies.

---

## License

[MIT](LICENSE)
