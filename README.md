# Shop Sales Comparison Tool — v6.0

Static site (no build step). Publish directory = repository root.

Files: `index.html`, `xls-writer.js`, `vendor/` (ExcelJS 4.4.0, SheetJS 0.18.5), `manifest.json`, `logo-192.png`, `logo-512.png`, `_headers`.

New in 6.0: **Daily Sales** tab. Shops log in, post the day's sales (Total, POS-1, POS-2, Qty, Invoices), see the last 30 days and copy a WhatsApp report image. Data lives in the "Daily Sales" Google Sheet through the Google Apps Script API set in `API_URL` (search `API_URL` in index.html). Shops can bookmark `.../#daily` to open straight on that tab.

Backend code: see `daily-sales-app/Code.gs` (paste into the Sheet's Apps Script, then Deploy > Manage deployments > New version).

5.0: reads and writes Excel 97 `.xls`; start-up notice until 30 Nov 2026 (`ANNOUNCE_UNTIL` in index.html).
