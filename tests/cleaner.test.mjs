import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  cleanCellHtml,
  parseSingleHtmlTable,
  mergeCleanTables,
  toCsvWithBom,
  toSpreadsheetMlXls,
  toXlsxBytes,
  verifyLicenseUniversal,
  PUBLIC_SPKI_BASE64,
} from "../src/cleaner.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

test("rowspan and colspan complex bidding table expands into aligned rectangular 2D grid", () => {
  const complexHtml = `
    <table>
      <thead>
        <tr>
          <th rowspan="2">分包编号</th>
          <th rowspan="2">采购品目</th>
          <th colspan="2">含税报价明细</th>
        </tr>
        <tr>
          <th>单价</th>
          <th>总价</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td rowspan="2">PKG-01</td>
          <td>核心交换机<sup>[1]</sup></td>
          <td>¥ 12,500.00元</td>
          <td>¥ 25,000.00*</td>
        </tr>
        <tr>
          <td>万兆光模块 [注2]</td>
          <td>$1,200.50</td>
          <td>(¥4,800.25)</td>
        </tr>
      </tbody>
    </table>
  `;

  const parsed = parseSingleHtmlTable(complexHtml);
  assert.equal(parsed.rowCount, 4);
  assert.equal(parsed.colCount, 4);
  assert.equal(parsed.headerRowCount, 2);
  assert.ok(parsed.stats.mergedCellsExpanded >= 4);

  // Header row 1 & 2 properly aligned despite rowspan="2" on cols 0 & 1!
  assert.deepEqual(parsed.rows[0], ["分包编号", "采购品目", "含税报价明细", "含税报价明细"]);
  assert.deepEqual(parsed.rows[1], ["分包编号", "采购品目", "单价", "总价"]);

  // Body row 1 & 2: PKG-01 repeated into row 2 col 0 so Excel pivot works; amounts stripped cleanly!
  assert.deepEqual(parsed.rows[2], ["PKG-01", "核心交换机", "12500.00", "25000.00"]);
  assert.deepEqual(parsed.rows[3], ["PKG-01", "万兆光模块", "1200.50", "-4800.25"]);

  assert.equal(parsed.cells[2][2].isNumeric, true);
  assert.equal(parsed.cells[2][2].numericValue, 12500);
  assert.equal(parsed.cells[3][3].isNumeric, true);
  assert.equal(parsed.cells[3][3].numericValue, -4800.25);
});

test("Leading-zero invoice codes and >15 digit bank numbers stay preserved as text", () => {
  const inv = cleanCellHtml("0038492019");
  assert.equal(inv.isNumeric, false);
  assert.equal(inv.preserveAsText, true);
  assert.equal(inv.text, "0038492019");

  const bank = cleanCellHtml("6222021001092837465");
  assert.equal(bank.isNumeric, false);
  assert.equal(bank.preserveAsText, true);
  assert.equal(bank.text, "6222021001092837465");

  const euro = cleanCellHtml("€ 3,890.75");
  assert.equal(euro.isNumeric, true);
  assert.equal(euro.numericValue, 3890.75);
});

test("Multi-page vertical merge deduplicates repeated headers across pages", () => {
  const page1 = parseSingleHtmlTable(`
    <table>
      <tr><th>发票代码</th><th>供应商</th><th>结算金额</th></tr>
      <tr><td>00192831</td><td>华东机电</td><td>¥108,000.00</td></tr>
      <tr><td>00192832</td><td>博远科技</td><td>¥52,400.50</td></tr>
    </table>
  `);
  const page2 = parseSingleHtmlTable(`
    <table>
      <tr><th>发票代码</th><th>供应商</th><th>结算金额</th></tr>
      <tr><td>00192833</td><td>瑞信网络</td><td>¥31,900.00</td></tr>
    </table>
  `);

  const merged = mergeCleanTables([page1, page2]);
  assert.equal(merged.duplicateHeadersRemoved, 1);
  assert.equal(merged.rowCount, 4); // 1 header + 3 data rows
  assert.equal(merged.rows[3][0], "00192833");
  assert.equal(merged.cells[3][2].numericValue, 31900);
});

test("CSV (UTF-8 BOM), SpreadsheetML (.xls), and genuine .xlsx exports are well-formed", () => {
  const tbl = parseSingleHtmlTable(`
    <table>
      <tr><th>项目</th><th>发票号</th><th>金额</th></tr>
      <tr><td>服务器采购, A区</td><td>00891234</td><td>¥98,500.00元</td></tr>
    </table>
  `);

  const csv = toCsvWithBom(tbl);
  assert.equal(csv.charCodeAt(0), 0xfeff, "CSV must start with UTF-8 BOM");
  assert.match(csv, /"服务器采购, A区",00891234,98500\.00/);

  const xls = toSpreadsheetMlXls(tbl);
  assert.match(xls, /<Data ss:Type="String">00891234<\/Data>/);
  assert.match(xls, /<Data ss:Type="Number">98500<\/Data>/);

  const xlsx = toXlsxBytes(tbl);
  assert.ok(xlsx instanceof Uint8Array);
  // PK\x03\x04 ZIP magic header
  assert.equal(xlsx[0], 0x50);
  assert.equal(xlsx[1], 0x4b);
  assert.equal(xlsx[2], 0x03);
  assert.equal(xlsx[3], 0x04);
  const rawAscii = Buffer.from(xlsx).toString("utf8");
  assert.match(rawAscii, /xl\/worksheets\/sheet1\.xml/);
  assert.match(rawAscii, /<c r="C2" t="n"><v>98500<\/v><\/c>/);
});

test("Universal WebCrypto Ed25519 verifier validates OPC1 tokens and MV3 bundle exists", async () => {
  assert.equal(PUBLIC_SPKI_BASE64, "MCowBQYDK2VwAyEAP5NCVP7CT0tgaGLSGMe4r87SjYRWTM6gN6igBZYSXys=");

  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  const spkiB64 = publicKey.export({ format: "der", type: "spki" }).toString("base64");
  const now = Math.floor(Date.now() / 1000);

  const makeTok = (payload) => {
    const pB64 = Buffer.from(JSON.stringify(payload)).toString("base64url");
    const sB64 = crypto.sign(null, Buffer.from(`OPC1.${pB64}`), privateKey).toString("base64url");
    return `OPC1.${pB64}.${sB64}`;
  };

  const validTok = makeTok({ v: 1, p: "edge-table-cleaner", e: "cfo@corp.com", plan: "pro", iat: now, exp: 0 });
  const okRes = await verifyLicenseUniversal(validTok, "edge-table-cleaner", spkiB64);
  assert.equal(okRes.valid, true);

  const badSig = `${validTok.slice(0, -4)}AAAA`;
  const badRes = await verifyLicenseUniversal(badSig, "edge-table-cleaner", spkiB64);
  assert.equal(badRes.valid, false);

  const wrongProd = makeTok({ v: 1, p: "store-radar-mcp", e: "cfo@corp.com", plan: "pro", iat: now, exp: 0 });
  const mismatchRes = await verifyLicenseUniversal(wrongProd, "edge-table-cleaner", spkiB64);
  assert.equal(mismatchRes.valid, false);
  assert.equal(mismatchRes.reason, "PRODUCT_MISMATCH");

  // Verify MV3 zip & icons exist
  const zipPath = path.join(ROOT, "dist", "edge-table-cleaner-mv3.zip");
  assert.equal(fs.existsSync(zipPath), true, "MV3 zip should be built");
  assert.ok(fs.statSync(zipPath).size > 2000);

  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, "extension", "manifest.json"), "utf8"));
  assert.equal(manifest.manifest_version, 3);
});
