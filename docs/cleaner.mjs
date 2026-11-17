/**
 * Edge Table Cleaner — Universal Core Engine (Node.js + Browser + MV3 Extension)
 * 100% Local Execution · Zero External Dependencies · Zero Network Uploads
 *
 * Features:
 * - Parses HTML `<table>` strings with complex nested `rowspan` & `colspan` into aligned 2D grids
 * - Strips currency symbols (`¥`, `￥`, `$`, `€`, `£`, `元`), thousands commas, accounting negatives `(1,200.50)`, and footnotes
 * - Preserves leading-zero invoice/tax/account identifiers as text to prevent Excel scientific notation corruption
 * - Vertically merges multi-page tables while deduplicating single-row or multi-row repeated headers
 * - Exports UTF-8 BOM `.csv`, Excel SpreadsheetML `.xls`, and genuine Office Open XML `.xlsx`
 * - Verifies offline `OPC1.<payload>.<sig>` licenses via WebCrypto Ed25519
 */

export const PRODUCT_SLUG = "edge-table-cleaner";
export const PUBLIC_SPKI_BASE64 = "MCowBQYDK2VwAyEAP5NCVP7CT0tgaGLSGMe4r87SjYRWTM6gN6igBZYSXys=";

const HTML_ENTITIES = {
  nbsp: " ",
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  yen: "¥",
  euro: "€",
  pound: "£",
  mdash: "—",
  ndash: "–",
  hellip: "...",
};

export function decodeHtmlEntities(input) {
  return String(input || "")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&([a-zA-Z]+);/g, (full, name) => {
      const lower = name.toLowerCase();
      return lower in HTML_ENTITIES ? HTML_ENTITIES[lower] : full;
    });
}

function parseAttrInt(tagAttrs, attrName, fallback = 1) {
  const re = new RegExp(`\\b${attrName}\\s*=\\s*(?:"(\\d+)"|'(\\d+)'|(\\d+))`, "i");
  const m = String(tagAttrs || "").match(re);
  if (!m) return fallback;
  const val = parseInt(m[1] || m[2] || m[3], 10);
  return Number.isFinite(val) && val >= 1 ? Math.min(val, 200) : fallback;
}

/**
 * Clean a single cell's raw HTML fragment into typed value `{ text, numericValue, isNumeric, stats }`.
 */
export function cleanCellHtml(rawCellHtml, options = {}) {
  const stripCurrency = options.stripCurrency !== false;
  const stripCommas = options.stripCommas !== false;
  const stripFootnotes = options.stripFootnotes !== false;

  const stats = {
    strippedCurrency: false,
    strippedCommas: false,
    strippedFootnotes: false,
  };

  let str = String(rawCellHtml || "");

  // Remove <script> and <style> inside cells
  str = str.replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, "");

  // Strip <sup> footnotes before stripping general HTML tags
  if (stripFootnotes && /<sup\b[^>]*>[\s\S]*?<\/sup>/i.test(str)) {
    str = str.replace(/<sup\b[^>]*>[\s\S]*?<\/sup>/gi, "");
    stats.strippedFootnotes = true;
  }

  // Convert <br> to single space
  str = str.replace(/<br\s*\/?>/gi, " ");
  // Strip remaining HTML tags
  str = str.replace(/<[^>]+>/g, "");
  // Decode HTML entities
  str = decodeHtmlEntities(str);
  // Strip invisible zero-width chars & normalize whitespace
  str = str
    .replace(/[\u200B-\u200D\uFEFF\u00A0]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (stripFootnotes) {
    const beforeFn = str;
    str = str
      .replace(/\s*\[(?:注\s*\d+|Note\s*\d+|\d+)\]\s*$/gi, "")
      .replace(/\s*[①②③④⑤⑥⑦⑧⑨⑩†‡]+\s*$/g, "")
      .replace(/([0-9元%])\s*\*+$/g, "$1")
      .trim();
    if (str !== beforeFn) stats.strippedFootnotes = true;
  }

  if (!str || str === "-" || str === "—" || str === "--" || str === "N/A") {
    return { text: str === "-" || str === "—" || str === "--" ? "" : str, numericValue: null, isNumeric: false, stats };
  }

  // Detect leading-zero codes (e.g. invoice number "00384920", zip "02138") or >15 digit bank/tax IDs:
  // Never coerce these to numbers or Excel will destroy them!
  if (/^0\d+$/.test(str) || /^\d{16,}$/.test(str)) {
    return { text: str, numericValue: null, isNumeric: false, preserveAsText: true, stats };
  }

  // Financial amount candidate normalization
  let candidate = str;
  let isNegativeAccounting = false;

  // Accounting parentheses negative: e.g. "(¥45,200.50)" or "($1,250.00)" or "(1,250.00)"
  const acctMatch = candidate.match(/^\(\s*([^()]+?)\s*\)$/);
  if (acctMatch && /\d/.test(acctMatch[1])) {
    candidate = acctMatch[1].trim();
    isNegativeAccounting = true;
  }

  let currencyRemoved = false;
  if (stripCurrency) {
    const next = candidate
      .replace(/^(?:HK\$|US\$|CA\$|AU\$|NT\$|RMB|CNY|USD|EUR|GBP|JPY)\s*/i, "")
      .replace(/^[+-]?\s*[¥￥$€£]\s*/, (m) => (m.trim().startsWith("-") ? "-" : m.trim().startsWith("+") ? "+" : ""))
      .replace(/^([+-])\s+/, "$1")
      .replace(/\s*(?:元整|人民币|元)$/u, "")
      .trim();
    if (next !== candidate) {
      currencyRemoved = true;
      candidate = next;
    }
  }

  let commasRemoved = false;
  if (stripCommas && /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(candidate)) {
    candidate = candidate.replace(/,/g, "");
    commasRemoved = true;
  }

  // Validate pure finite decimal number (up to 15 significant digits so IEEE-754 float64 is exact)
  const digitsOnly = candidate.replace(/^[+-]/, "").replace(".", "");
  const hasLeadingZeroInt = /^[+-]?0\d+/.test(candidate);
  if (
    !hasLeadingZeroInt &&
    digitsOnly.length > 0 &&
    digitsOnly.length <= 15 &&
    /^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(candidate)
  ) {
    let num = Number(candidate);
    if (isNegativeAccounting && num > 0) {
      num = -num;
    }
    if (Number.isFinite(num)) {
      stats.strippedCurrency = currencyRemoved;
      stats.strippedCommas = commasRemoved;
      // Preserve clean decimal string without scientific notation
      const normalizedText = isNegativeAccounting && !candidate.startsWith("-") ? `-${candidate}` : candidate.replace(/^\+/, "");
      return {
        text: normalizedText,
        numericValue: num,
        isNumeric: true,
        stats,
      };
    }
  }

  return { text: str, numericValue: null, isNumeric: false, stats };
}

/**
 * Extract top-level or sequential `<table>...</table>` blocks from arbitrary HTML.
 */
export function extractHtmlTables(html) {
  const src = String(html || "");
  const tables = [];
  const tagRe = /<\/?table\b[^>]*>/gi;
  let depth = 0;
  let startIdx = -1;
  let match;

  while ((match = tagRe.exec(src)) !== null) {
    const isClose = match[0].startsWith("</");
    if (!isClose) {
      if (depth === 0) startIdx = match.index;
      depth += 1;
    } else if (depth > 0) {
      depth -= 1;
      if (depth === 0 && startIdx !== -1) {
        tables.push(src.slice(startIdx, match.index + match[0].length));
        startIdx = -1;
      }
    }
  }

  // Fallback: if user pasted bare `<tr>...</tr>` rows without outer `<table>` wrapper
  if (tables.length === 0 && /<tr\b/i.test(src)) {
    tables.push(`<table>${src}</table>`);
  }
  return tables;
}

/**
 * Parse a single HTML `<table>` into a rectangular 2D array of cleaned cell objects,
 * accurately expanding `rowspan` and `colspan` attributes.
 */
export function parseSingleHtmlTable(tableHtml, options = {}) {
  const expandMerged = options.expandMerged !== false;
  const repeatRowspanValue = options.repeatRowspanValue !== false;
  const repeatColspanValue = options.repeatColspanValue ?? false;

  const grid = []; // grid[r][c] = cellObj
  let headerRowCount = 0;
  let mergedCellsExpanded = 0;
  let currencyCellsCleaned = 0;
  let commaCellsCleaned = 0;
  let footnoteCellsCleaned = 0;

  // Match <tr> blocks in order
  const trRe = /<tr\b([^>]*)>([\s\S]*?)<\/tr>/gi;
  let trMatch;
  let rowIndex = 0;

  // Detect which rows live inside <thead>
  const theadRanges = [];
  const theadRe = /<thead\b[^>]*>([\s\S]*?)<\/thead>/gi;
  let thm;
  while ((thm = theadRe.exec(tableHtml)) !== null) {
    theadRanges.push([thm.index, thm.index + thm[0].length]);
  }

  while ((trMatch = trRe.exec(tableHtml)) !== null) {
    const trStart = trMatch.index;
    const inThead = theadRanges.some(([s, e]) => trStart >= s && trStart < e);
    const rowInner = trMatch[2];

    if (!grid[rowIndex]) grid[rowIndex] = [];

    const cellRe = /<(th|td)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
    let cellMatch;
    let colCursor = 0;
    let allThInRow = true;
    let cellsSeenInTr = 0;

    while ((cellMatch = cellRe.exec(rowInner)) !== null) {
      cellsSeenInTr += 1;
      const tagName = cellMatch[1].toLowerCase();
      if (tagName !== "th") allThInRow = false;

      const attrs = cellMatch[2];
      const rawInner = cellMatch[3];
      const rowspan = expandMerged ? parseAttrInt(attrs, "rowspan", 1) : 1;
      const colspan = expandMerged ? parseAttrInt(attrs, "colspan", 1) : 1;

      if (rowspan > 1 || colspan > 1) {
        mergedCellsExpanded += 1;
      }

      // Advance colCursor past any slots already occupied by previous rowspans
      while (grid[rowIndex][colCursor] !== undefined) {
        colCursor += 1;
      }

      const cleaned = cleanCellHtml(rawInner, options);
      if (cleaned.stats.strippedCurrency) currencyCellsCleaned += 1;
      if (cleaned.stats.strippedCommas) commaCellsCleaned += 1;
      if (cleaned.stats.strippedFootnotes) footnoteCellsCleaned += 1;

      const isHeaderCell = inThead || tagName === "th";

      for (let dr = 0; dr < rowspan; dr += 1) {
        const targetRow = rowIndex + dr;
        if (!grid[targetRow]) grid[targetRow] = [];
        for (let dc = 0; dc < colspan; dc += 1) {
          const targetCol = colCursor + dc;
          const isOrigin = dr === 0 && dc === 0;
          let shouldFill = isOrigin;
          if (!isOrigin) {
            if (isHeaderCell) {
              shouldFill = true; // repeating header group labels keeps multi-row headers self-describing
            } else if (dc === 0 && dr > 0) {
              shouldFill = repeatRowspanValue;
            } else {
              shouldFill = repeatColspanValue;
            }
          }

          grid[targetRow][targetCol] = shouldFill
            ? {
                text: cleaned.text,
                numericValue: cleaned.numericValue,
                isNumeric: cleaned.isNumeric,
                preserveAsText: Boolean(cleaned.preserveAsText),
                isHeader: isHeaderCell,
              }
            : {
                text: "",
                numericValue: null,
                isNumeric: false,
                preserveAsText: false,
                isHeader: isHeaderCell,
              };
        }
      }

      colCursor += colspan;
    }

    if (cellsSeenInTr > 0 && (inThead || (allThInRow && rowIndex === headerRowCount))) {
      headerRowCount = Math.max(headerRowCount, rowIndex + 1);
    }

    rowIndex += 1;
  }

  // Filter completely blank rows and normalize rectangular column width
  const nonEmptyGrid = grid.filter((row) =>
    Array.isArray(row) && row.some((cell) => cell && String(cell.text).trim() !== ""),
  );

  const maxCols = nonEmptyGrid.reduce((m, r) => Math.max(m, r.length), 0);
  const normalizedCells = nonEmptyGrid.map((r) => {
    const rowOut = [];
    for (let c = 0; c < maxCols; c += 1) {
      rowOut.push(
        r[c] || {
          text: "",
          numericValue: null,
          isNumeric: false,
          preserveAsText: false,
          isHeader: false,
        },
      );
    }
    return rowOut;
  });

  if (headerRowCount === 0 && normalizedCells.length > 1) {
    // Heuristic: if row 0 has zero numeric values and row 1 has >=1 numeric value, treat row 0 as header
    const r0HasNum = normalizedCells[0].some((c) => c.isNumeric);
    const r1HasNum = normalizedCells[1].some((c) => c.isNumeric);
    if (!r0HasNum && r1HasNum) {
      headerRowCount = 1;
    }
  }

  const rows = normalizedCells.map((r) => r.map((c) => c.text));

  return {
    rowCount: rows.length,
    colCount: maxCols,
    headerRowCount: Math.min(headerRowCount, rows.length),
    rows,
    cells: normalizedCells,
    stats: {
      mergedCellsExpanded,
      currencyCellsCleaned,
      commaCellsCleaned,
      footnoteCellsCleaned,
    },
  };
}

/**
 * Parse all `<table>` blocks in an HTML string (or delimited tab/pipe text fallback).
 */
export function cleanHtmlTables(htmlOrText, options = {}) {
  const rawTables = extractHtmlTables(htmlOrText);
  if (rawTables.length > 0) {
    return rawTables
      .map((t) => parseSingleHtmlTable(t, options))
      .filter((t) => t.rowCount > 0 && t.colCount > 0);
  }

  // Fallback for TSV / Markdown pipe tables pasted by user
  const lines = String(htmlOrText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (lines.length === 0) return [];

  const delimiter = lines[0].includes("\t") ? "\t" : lines[0].includes("|") ? "|" : ",";
  const parsedCells = [];
  for (const line of lines) {
    if (delimiter === "|" && /^\|?[\s:-]+\|[\s|:-]*$/.test(line)) continue;
    let parts = line.split(delimiter);
    if (delimiter === "|") {
      if (line.startsWith("|")) parts = parts.slice(1);
      if (line.endsWith("|")) parts = parts.slice(0, -1);
    }
    parsedCells.push(parts.map((p) => cleanCellHtml(p, options)));
  }

  const maxCols = parsedCells.reduce((m, r) => Math.max(m, r.length), 0);
  const cells = parsedCells.map((r) => {
    const out = [];
    for (let c = 0; c < maxCols; c += 1) {
      out.push(r[c] || { text: "", numericValue: null, isNumeric: false, stats: {} });
    }
    return out;
  });

  return [
    {
      rowCount: cells.length,
      colCount: maxCols,
      headerRowCount: cells.length > 1 ? 1 : 0,
      rows: cells.map((r) => r.map((c) => c.text)),
      cells,
      stats: {
        mergedCellsExpanded: 0,
        currencyCellsCleaned: cells.flat().filter((c) => c.stats?.strippedCurrency).length,
        commaCellsCleaned: cells.flat().filter((c) => c.stats?.strippedCommas).length,
        footnoteCellsCleaned: cells.flat().filter((c) => c.stats?.strippedFootnotes).length,
      },
    },
  ];
}

function rowSignature(rowCells) {
  return rowCells.map((c) => String(c?.text ?? c ?? "").trim().toLowerCase()).join("||");
}

/**
 * Vertically merge multiple cleaned tables (e.g., paginated ERP / bidding tables),
 * automatically deduplicating repeated single-row or multi-row headers on Page 2..N.
 */
export function mergeCleanTables(parsedTables, options = {}) {
  const list = Array.isArray(parsedTables) ? parsedTables.filter((t) => t && t.rowCount > 0) : [];
  if (list.length === 0) {
    return {
      rowCount: 0,
      colCount: 0,
      headerRowCount: 0,
      rows: [],
      cells: [],
      duplicateHeadersRemoved: 0,
      stats: { mergedCellsExpanded: 0, currencyCellsCleaned: 0, commaCellsCleaned: 0, footnoteCellsCleaned: 0 },
    };
  }

  const first = list[0];
  const headerDepth = Math.max(1, first.headerRowCount || 1);
  const canonicalHeaderSigs = first.cells.slice(0, headerDepth).map(rowSignature);
  const maxCols = list.reduce((m, t) => Math.max(m, t.colCount), 0);

  const mergedCells = [];
  let duplicateHeadersRemoved = 0;
  const totalStats = {
    mergedCellsExpanded: 0,
    currencyCellsCleaned: 0,
    commaCellsCleaned: 0,
    footnoteCellsCleaned: 0,
  };

  list.forEach((tbl, tblIdx) => {
    totalStats.mergedCellsExpanded += tbl.stats?.mergedCellsExpanded || 0;
    totalStats.currencyCellsCleaned += tbl.stats?.currencyCellsCleaned || 0;
    totalStats.commaCellsCleaned += tbl.stats?.commaCellsCleaned || 0;
    totalStats.footnoteCellsCleaned += tbl.stats?.footnoteCellsCleaned || 0;

    let startRow = 0;
    if (tblIdx > 0 && options.deduplicateHeaders !== false) {
      // Check full multi-row header match first, then single-row header match
      while (startRow < tbl.cells.length && startRow < canonicalHeaderSigs.length) {
        if (rowSignature(tbl.cells[startRow]) === canonicalHeaderSigs[startRow]) {
          startRow += 1;
          duplicateHeadersRemoved += 1;
        } else {
          break;
        }
      }
      // Also strip any mid-table repeated first-row header
    }

    for (let r = startRow; r < tbl.cells.length; r += 1) {
      const row = tbl.cells[r];
      if (tblIdx > 0 && options.deduplicateHeaders !== false && rowSignature(row) === canonicalHeaderSigs[0]) {
        duplicateHeadersRemoved += 1;
        continue;
      }
      const padded = [];
      for (let c = 0; c < maxCols; c += 1) {
        padded.push(
          row[c] || {
            text: "",
            numericValue: null,
            isNumeric: false,
            preserveAsText: false,
            isHeader: false,
          },
        );
      }
      mergedCells.push(padded);
    }
  });

  return {
    rowCount: mergedCells.length,
    colCount: maxCols,
    headerRowCount: Math.min(headerDepth, mergedCells.length),
    rows: mergedCells.map((r) => r.map((c) => c.text)),
    cells: mergedCells,
    duplicateHeadersRemoved,
    stats: totalStats,
  };
}

/**
 * Export cleaned table to standard UTF-8 with BOM `.csv` (Excel double-click safe for Chinese/UTF-8).
 */
export function toCsvWithBom(tableOrRows) {
  const cells = Array.isArray(tableOrRows)
    ? tableOrRows.map((r) => r.map((v) => ({ text: String(v ?? ""), preserveAsText: /^0\d+$/.test(String(v ?? "")) })))
    : tableOrRows.cells || [];

  const csvLines = cells.map((row) =>
    row
      .map((cell) => {
        const val = String(cell?.text ?? "");
        // Escape double quotes per RFC 4180
        if (/[",\r\n]/.test(val)) {
          return `"${val.replace(/"/g, '""')}"`;
        }
        return val;
      })
      .join(","),
  );

  return `\uFEFF${csvLines.join("\r\n")}\r\n`;
}

function escapeXml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Export cleaned table to Excel 2003 XML SpreadsheetML (`.xls`), preserving Number vs String types!
 */
export function toSpreadsheetMlXls(parsedTable, options = {}) {
  const sheetName = escapeXml(options.sheetName || "CleanedTable").slice(0, 31);
  const cells = parsedTable.cells || [];
  const headerRows = parsedTable.headerRowCount || 0;

  const xmlRows = cells
    .map((row, rIdx) => {
      const styleAttr = rIdx < headerRows ? ' ss:StyleID="sHeader"' : "";
      const xmlCells = row
        .map((cell) => {
          const isNum = Boolean(cell.isNumeric && typeof cell.numericValue === "number" && !cell.preserveAsText);
          const type = isNum ? "Number" : "String";
          const val = isNum ? String(cell.numericValue) : escapeXml(cell.text);
          return `<Cell${styleAttr}><Data ss:Type="${type}">${val}</Data></Cell>`;
        })
        .join("");
      return `      <Row>${xmlCells}</Row>`;
    })
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Font ss:FontName="Calibri" ss:Size="11"/>
    </Style>
    <Style ss:ID="sHeader">
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/>
      <Interior ss:Color="#E6F4EA" ss:Pattern="Solid"/>
    </Style>
  </Styles>
  <Worksheet ss:Name="${sheetName}">
    <Table>
${xmlRows}
    </Table>
  </Worksheet>
</Workbook>`;
}

// --- Pure JS CRC32 + Uncompressed ZIP Builder for Genuine `.xlsx` ---
const CRC32_TABLE = (() => {
  const tbl = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    tbl[i] = c >>> 0;
  }
  return tbl;
})();

function crc32Bytes(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function buildStoredZip(files) {
  const enc = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const file of files) {
    const nameBytes = enc.encode(file.name);
    const dataBytes = typeof file.content === "string" ? enc.encode(file.content) : file.content;
    const crc = crc32Bytes(dataBytes);
    const size = dataBytes.length;

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(localHeader.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0x0800, true); // UTF-8 filename flag
    lv.setUint16(8, 0, true); // STORE (no compression)
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0x5621, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);

    localParts.push(localHeader, dataBytes);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(centralHeader.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x5621, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    centralParts.push(centralHeader);
    offset += localHeader.length + size;
  }

  const centralSize = centralParts.reduce((s, p) => s + p.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  const totalLen = offset + centralSize + 22;
  const out = new Uint8Array(totalLen);
  let pos = 0;
  for (const part of [...localParts, ...centralParts, eocd]) {
    out.set(part, pos);
    pos += part.length;
  }
  return out;
}

function colIndexToLetter(colIdx) {
  let n = colIdx + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * Generate a genuine Office Open XML `.xlsx` binary (`Uint8Array`) with zero dependencies!
 */
export function toXlsxBytes(parsedTable, options = {}) {
  const sheetName = escapeXml(options.sheetName || "CleanedTable").slice(0, 31);
  const cells = parsedTable.cells || [];
  const headerRows = parsedTable.headerRowCount || 0;

  const sheetRowsXml = cells
    .map((row, rIdx) => {
      const rowNum = rIdx + 1;
      const styleAttr = rIdx < headerRows ? ' s="1"' : "";
      const rowCellsXml = row
        .map((cell, cIdx) => {
          const ref = `${colIndexToLetter(cIdx)}${rowNum}`;
          if (cell.isNumeric && typeof cell.numericValue === "number" && !cell.preserveAsText) {
            return `<c r="${ref}"${styleAttr} t="n"><v>${cell.numericValue}</v></c>`;
          }
          const txt = escapeXml(cell.text || "");
          return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${txt}</t></is></c>`;
        })
        .join("");
      return `<row r="${rowNum}">${rowCellsXml}</row>`;
    })
    .join("");

  const files = [
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    },
    {
      name: "xl/workbook.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${sheetName}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    },
    {
      name: "xl/styles.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE6F4EA"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>`,
    },
    {
      name: "xl/worksheets/sheet1.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetRowsXml}</sheetData>
</worksheet>`,
    },
  ];

  return buildStoredZip(files);
}

function decodeBase64Standard(b64) {
  if (typeof atob === "function") {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }
  return new Uint8Array(Buffer.from(b64, "base64"));
}

function decodeBase64UrlBytes(str) {
  const normalized = String(str).replace(/-/g, "+").replace(/_/g, "/");
  const pad = "=".repeat((4 - (normalized.length % 4)) % 4);
  return decodeBase64Standard(normalized + pad);
}

/**
 * Universal WebCrypto Ed25519 Verifier (works identically in Node >=20, Chrome/Edge MV3 Extension, and Browser).
 */
export async function verifyLicenseUniversal(
  token,
  expectedProduct = PRODUCT_SLUG,
  pubSpkiBase64 = PUBLIC_SPKI_BASE64,
) {
  if (typeof token !== "string" || !token.trim()) {
    return { valid: false, reason: "INVALID_TYPE" };
  }
  const parts = token.trim().split(".");
  if (parts.length !== 3 || parts[0] !== "OPC1" || !parts[1] || !parts[2]) {
    return { valid: false, reason: "BAD_FORMAT" };
  }

  try {
    const subtle = globalThis.crypto?.subtle;
    if (!subtle) {
      return { valid: false, reason: "WEBCRYPTO_UNAVAILABLE" };
    }
    const spkiBytes = decodeBase64Standard(pubSpkiBase64);
    const publicKey = await subtle.importKey("spki", spkiBytes, { name: "Ed25519" }, false, ["verify"]);
    const sigBytes = decodeBase64UrlBytes(parts[2]);
    const msgBytes = new TextEncoder().encode(`OPC1.${parts[1]}`);
    const verified = await subtle.verify({ name: "Ed25519" }, publicKey, sigBytes, msgBytes);
    if (!verified) {
      return { valid: false, reason: "BAD_SIGNATURE" };
    }

    const payloadJson = new TextDecoder().decode(decodeBase64UrlBytes(parts[1]));
    const payload = JSON.parse(payloadJson);
    if (!payload || typeof payload !== "object" || !("p" in payload)) {
      return { valid: false, reason: "MISSING_FIELDS", payload };
    }
    if (expectedProduct && payload.p !== expectedProduct && payload.p !== "*") {
      return { valid: false, reason: "PRODUCT_MISMATCH", payload };
    }
    if (typeof payload.exp === "number" && payload.exp > 0) {
      const nowSec = Math.floor(Date.now() / 1000);
      if (nowSec > payload.exp) {
        return { valid: false, reason: "EXPIRED", payload };
      }
    }
    return { valid: true, payload };
  } catch (err) {
    return { valid: false, reason: "VERIFY_ERROR", message: err.message };
  }
}
