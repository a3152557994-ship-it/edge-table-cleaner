import {
  parseSingleHtmlTable,
  mergeCleanTables,
  toCsvWithBom,
  toSpreadsheetMlXls,
  toXlsxBytes,
  verifyLicenseUniversal,
} from "./cleaner.js";

let proActive = false;
let stagedHtmlList = [];
let scannedTables = [];

function getCleanOptions() {
  return {
    expandMerged: document.getElementById("optExpand").checked,
    stripCurrency: document.getElementById("optCurrency").checked,
    stripCommas: document.getElementById("optCommas").checked,
    stripFootnotes: document.getElementById("optFootnotes").checked,
  };
}

function triggerDownload(bytesOrText, filename, mimeType) {
  const blob = new Blob([bytesOrText], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

async function loadState() {
  const stored = await chrome.storage.local.get(["opc1_token", "staged_tables_html"]);
  stagedHtmlList = Array.isArray(stored.staged_tables_html) ? stored.staged_tables_html : [];
  document.getElementById("stagedCount").textContent = String(stagedHtmlList.length);

  if (stored.opc1_token) {
    document.getElementById("licenseInput").value = stored.opc1_token;
    const res = await verifyLicenseUniversal(stored.opc1_token);
    if (res.valid) {
      proActive = true;
      const badge = document.getElementById("tierBadge");
      badge.textContent = "Pro 终身版";
      badge.classList.add("pro");
      document.getElementById("licenseMsg").textContent = `✅ 已激活 (${res.payload.p})`;
    }
  }
}

async function scanCurrentTab() {
  const listEl = document.getElementById("tablesList");
  listEl.innerHTML = `<div style="color:#94a3b8;margin-top:8px;">正在本地扫描当前标签页...</div>`;

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !tab.url || /^(chrome|edge|about):/i.test(tab.url)) {
      listEl.innerHTML = `<div style="color:#fca5a5;margin-top:8px;">请在包含网页表格的普通页面（或内网系统）点击插件图标。</div>`;
      return;
    }

    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["content.js"],
    });

    const res = await chrome.tabs.sendMessage(tab.id, { type: "SCAN_TABLES" });
    scannedTables = res?.tables || [];

    if (scannedTables.length === 0) {
      listEl.innerHTML = `<div style="color:#94a3b8;margin-top:8px;">当前页面未检测到 &lt;table&gt; 表格。</div>`;
      return;
    }

    const opts = getCleanOptions();
    listEl.innerHTML = "";

    scannedTables.forEach((item, idx) => {
      const parsed = parseSingleHtmlTable(item.outerHTML, opts);
      const card = document.createElement("div");
      card.className = "table-item";
      card.innerHTML = `
        <div style="font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">#${idx + 1} ${item.title}</div>
        <div class="table-meta">${parsed.rowCount} 行 × ${parsed.colCount} 列 · 展开合并 ${parsed.stats.mergedCellsExpanded} 处 · 清洗金额 ${parsed.stats.currencyCellsCleaned + parsed.stats.commaCellsCleaned} 格</div>
        <div class="actions">
          <button class="btn btn-primary" data-act="xlsx" data-idx="${idx}" type="button">.xlsx</button>
          <button class="btn" data-act="xls" data-idx="${idx}" type="button">.xls</button>
          <button class="btn" data-act="csv" data-idx="${idx}" type="button">.csv</button>
          <button class="btn btn-amber" data-act="stage" data-idx="${idx}" type="button">+ 暂存本页</button>
        </div>
      `;
      card.addEventListener("mouseenter", () => {
        chrome.tabs.sendMessage(tab.id, { type: "HIGHLIGHT_TABLE", index: item.index }).catch(() => {});
      });
      listEl.appendChild(card);
    });

    listEl.querySelectorAll("button[data-act]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const act = btn.getAttribute("data-act");
        const idx = Number(btn.getAttribute("data-idx"));
        const target = scannedTables[idx];
        const parsed = parseSingleHtmlTable(target.outerHTML, getCleanOptions());
        const base = `table-${idx + 1}`;

        if (act === "xlsx") {
          triggerDownload(
            toXlsxBytes(parsed, { sheetName: base }),
            `${base}.xlsx`,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          );
        } else if (act === "xls") {
          triggerDownload(toSpreadsheetMlXls(parsed, { sheetName: base }), `${base}.xls`, "application/vnd.ms-excel");
        } else if (act === "csv") {
          triggerDownload(toCsvWithBom(parsed), `${base}.csv`, "text/csv;charset=utf-8");
        } else if (act === "stage") {
          stagedHtmlList.push(target.outerHTML);
          await chrome.storage.local.set({ staged_tables_html: stagedHtmlList });
          document.getElementById("stagedCount").textContent = String(stagedHtmlList.length);
          document.getElementById("stageMsg").className = "status";
          document.getElementById("stageMsg").textContent = `已暂存第 ${stagedHtmlList.length} 页表格，可翻到下一页继续暂存！`;
        }
      });
    });
  } catch (err) {
    listEl.innerHTML = `<div style="color:#fca5a5;margin-top:8px;">扫描失败：${err.message}</div>`;
  }
}

function exportMerged(format) {
  const msgEl = document.getElementById("stageMsg");
  if (!proActive) {
    msgEl.className = "status err";
    msgEl.textContent = "🔒 跨页去重合并为 Pro 功能，请在下方输入 OPC1. 激活码解锁（¥19 / $9.9 终身买断）。";
    return;
  }
  if (stagedHtmlList.length === 0) {
    msgEl.className = "status err";
    msgEl.textContent = "暂存池为空，请先点击表格下方的「+ 暂存本页」。";
    return;
  }
  const opts = getCleanOptions();
  const parsedList = stagedHtmlList.map((h) => parseSingleHtmlTable(h, opts));
  const merged = mergeCleanTables(parsedList, { deduplicateHeaders: true });

  if (format === "xlsx") {
    triggerDownload(
      toXlsxBytes(merged, { sheetName: "MergedPages" }),
      `merged-${stagedHtmlList.length}pages.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  } else {
    triggerDownload(toCsvWithBom(merged), `merged-${stagedHtmlList.length}pages.csv`, "text/csv;charset=utf-8");
  }
  msgEl.className = "status";
  msgEl.textContent = `✅ 已合并 ${stagedHtmlList.length} 页（共 ${merged.rowCount} 行，自动剔除重复表头 ${merged.duplicateHeadersRemoved} 行）！`;
}

document.getElementById("scanBtn").addEventListener("click", scanCurrentTab);

document.getElementById("clearStagedBtn").addEventListener("click", async () => {
  stagedHtmlList = [];
  await chrome.storage.local.set({ staged_tables_html: [] });
  document.getElementById("stagedCount").textContent = "0";
  document.getElementById("stageMsg").textContent = "已清空跨页暂存池。";
});

document.getElementById("exportMergedXlsxBtn").addEventListener("click", () => exportMerged("xlsx"));
document.getElementById("exportMergedCsvBtn").addEventListener("click", () => exportMerged("csv"));

document.getElementById("verifyBtn").addEventListener("click", async () => {
  const token = document.getElementById("licenseInput").value.trim();
  const msgEl = document.getElementById("licenseMsg");
  const res = await verifyLicenseUniversal(token);
  if (res.valid) {
    proActive = true;
    await chrome.storage.local.set({ opc1_token: token });
    const badge = document.getElementById("tierBadge");
    badge.textContent = "Pro 终身版";
    badge.classList.add("pro");
    msgEl.className = "status";
    msgEl.textContent = `✅ 验签成功！已永久解锁跨页去重合并 (${res.payload.p})`;
  } else {
    msgEl.className = "status err";
    msgEl.textContent = `❌ 激活码无效：${res.reason}`;
  }
});

document.getElementById("buyMailLink").addEventListener("click", (e) => {
  e.preventDefault();
  const subject = encodeURIComponent("[购买激活码] Edge Table Cleaner Pro 终身买断 (¥19 / $9.9)");
  const body = encodeURIComponent(
    "你好！我想购买 Edge Table Cleaner (`edge-table-cleaner`) 终身买断版 OPC1. 激活码 (¥19 / $9.9)。\n备用联系邮箱：a3152557994@gmail.com\n接收激活码邮箱：",
  );
  chrome.tabs.create({ url: `mailto:3152557994@qq.com?cc=a3152557994@gmail.com&subject=${subject}&body=${body}` });
});

loadState().then(scanCurrentTab);
