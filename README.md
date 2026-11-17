# Edge Table Cleaner (`edge-table-cleaner`)

> **100% Local Intranet & Financial Web Table Cleaner → Standard Excel (`.xlsx` / `.xls`) & UTF-8 BOM `.csv`**
> Built for finance, procurement, auditing, and operations teams working inside sensitive ERP, OA, and government bidding portals where **uploading internal data to cloud converters is strictly prohibited**.

- **Online Web Cleaner & Extension Download**: [https://a3152557994-ship-it.github.io/edge-table-cleaner/](https://a3152557994-ship-it.github.io/edge-table-cleaner/)
- **Download Manifest V3 Extension (`.zip`)**: [GitHub Release `v1.0.0`](https://github.com/a3152557994-ship-it/edge-table-cleaner/releases/tag/v1.0.0)

---

## Why Edge Table Cleaner?

Copy-pasting complex HTML tables from ERP/OA/bidding portals into Excel usually breaks everything:
1. **Merged Headers & Category Rows (`rowspan` / `colspan`)** shift columns out of alignment and break PivotTables.
2. **Currency Symbols & Thousands Commas (`¥1,280,500.00元`, `$12,450.80`, `(¥45,200.50)`)** turn numbers into un-summable text strings.
3. **Leading-Zero Invoice & Tax Codes (`0038492019`)** get mangled into `3.8492E+09`.
4. **Paginated Tables** repeat multi-row headers on every page.

`edge-table-cleaner` solves all 4 problems in milliseconds — **100% locally in your browser with zero network requests**.

---

## Install Browser Extension (Edge / Chrome Manifest V3)

1. Download `edge-table-cleaner-mv3.zip` from [Releases](https://github.com/a3152557994-ship-it/edge-table-cleaner/releases/latest) and unzip it.
2. Open `edge://extensions` (or `chrome://extensions`) and toggle **Developer mode** ON.
3. Click **Load unpacked** and select the unzipped folder.

---

## Pricing & Offline Ed25519 Licensing

- **Free Tier**: Unlimited single-page table scanning, `rowspan`/`colspan` expansion, financial currency/comma sanitization, and `.xlsx` / `.xls` / `.csv` export.
- **Pro Lifetime Buyout (`¥19 / $9.9`)**: Unlocks **Multi-Page Longitudinal Staging & Automatic Duplicate Header Removal** across paginated ERP/procurement tables. Verified 100% offline via WebCrypto `Ed25519` (`OPC1.<payload>.<sig>`).
- **Order Pro Key**: `3152557994@qq.com` / `a3152557994@gmail.com`

## Development & Tests

```bash
npm test
```
