/**
 * Build script for Edge Table Cleaner:
 * 1. Syncs `src/cleaner.mjs` -> `extension/cleaner.js` and `docs/cleaner.mjs`
 * 2. Generates clean RGBA PNG icons (`16x16`, `48x48`, `128x128`) with pure Node `node:zlib`
 * 3. Packages `extension/` into `dist/edge-table-cleaner-mv3.zip` (ready for Microsoft Edge Add-ons)
 */

import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    t[i] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(typeStr, dataBuf) {
  const typeBuf = Buffer.from(typeStr, "ascii");
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(dataBuf.length, 0);
  const crcInput = Buffer.concat([typeBuf, dataBuf]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(crcInput), 0);
  return Buffer.concat([lenBuf, typeBuf, dataBuf, crcBuf]);
}

function generateSpreadsheetIconPng(size) {
  const rawRows = [];
  const radius = Math.max(2, Math.round(size * 0.2));

  for (let y = 0; y < size; y += 1) {
    const row = Buffer.alloc(1 + size * 4);
    row[0] = 0; // PNG filter: None

    for (let x = 0; x < size; x += 1) {
      const idx = 1 + x * 4;

      // Rounded-rect distance check
      let inside = true;
      const dx = x < radius ? radius - x : x >= size - radius ? x - (size - 1 - radius) : 0;
      const dy = y < radius ? radius - y : y >= size - radius ? y - (size - 1 - radius) : 0;
      if (dx * dx + dy * dy > radius * radius) inside = false;

      if (!inside) {
        row[idx] = 0;
        row[idx + 1] = 0;
        row[idx + 2] = 0;
        row[idx + 3] = 0;
        continue;
      }

      // Emerald gradient background
      const ratio = (x + y) / (2 * size);
      let r = Math.round(16 - ratio * 6);
      let g = Math.round(185 - ratio * 45);
      let b = Math.round(129 - ratio * 30);

      // Inner white spreadsheet table grid
      const pad = Math.max(2, Math.round(size * 0.2));
      const lineW = Math.max(1, Math.round(size / 16));
      if (x >= pad && x < size - pad && y >= pad && y < size - pad) {
        const innerW = size - 2 * pad;
        const innerH = size - 2 * pad;
        const rx = x - pad;
        const ry = y - pad;

        const isOuterBorder =
          rx < lineW || rx >= innerW - lineW || ry < lineW || ry >= innerH - lineW;
        const isHeaderFill = ry < Math.round(innerH * 0.32);
        const isRowLine = Math.abs(ry - Math.round(innerH * 0.64)) < lineW;
        const isColLine = Math.abs(rx - Math.round(innerW * 0.42)) < lineW;

        if (isOuterBorder || isHeaderFill || isRowLine || isColLine) {
          r = 248;
          g = 250;
          b = 252;
        }
      }

      row[idx] = r;
      row[idx + 1] = g;
      row[idx + 2] = b;
      row[idx + 3] = 255;
    }
    rawRows.push(row);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const compressed = zlib.deflateSync(Buffer.concat(rawRows));
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  return Buffer.concat([
    sig,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", compressed),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function main() {
  const srcCleaner = path.join(ROOT, "src", "cleaner.mjs");
  const extDir = path.join(ROOT, "extension");
  const iconsDir = path.join(extDir, "icons");
  const docsDir = path.join(ROOT, "docs");
  const distDir = path.join(ROOT, "dist");

  fs.mkdirSync(iconsDir, { recursive: true });
  fs.mkdirSync(docsDir, { recursive: true });
  fs.mkdirSync(distDir, { recursive: true });

  // 1. Copy shared engine into extension/cleaner.js and docs/cleaner.mjs
  fs.copyFileSync(srcCleaner, path.join(extDir, "cleaner.js"));
  fs.copyFileSync(srcCleaner, path.join(docsDir, "cleaner.mjs"));

  // 2. Generate 16/48/128 PNG icons
  for (const size of [16, 48, 128]) {
    const png = generateSpreadsheetIconPng(size);
    fs.writeFileSync(path.join(iconsDir, `icon${size}.png`), png);
  }

  // 3. Create clean MV3 zip in dist/edge-table-cleaner-mv3.zip
  const zipPath = path.join(distDir, "edge-table-cleaner-mv3.zip");
  if (fs.existsSync(zipPath)) {
    fs.unlinkSync(zipPath);
  }
  execFileSync(
    "zip",
    ["-r", "-q", zipPath, "manifest.json", "popup.html", "popup.js", "content.js", "cleaner.js", "icons"],
    { cwd: extDir },
  );

  const stat = fs.statSync(zipPath);
  console.log(`Built MV3 extension package: dist/edge-table-cleaner-mv3.zip (${stat.size} bytes)`);
}

main();
