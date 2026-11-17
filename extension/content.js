/**
 * Content script injected on-demand via `activeTab` + `chrome.scripting.executeScript`.
 * Zero background tracking, zero network calls — 100% safe for sensitive intranet/ERP portals.
 */
(() => {
  if (window.__EDGE_TABLE_CLEANER_INSTALLED__) return;
  window.__EDGE_TABLE_CLEANER_INSTALLED__ = true;

  let highlightedEl = null;
  let prevOutline = "";

  function clearHighlight() {
    if (highlightedEl) {
      highlightedEl.style.outline = prevOutline;
      highlightedEl = null;
    }
  }

  function inferTableTitle(table, idx) {
    const caption = table.querySelector("caption");
    if (caption && caption.textContent.trim()) {
      return caption.textContent.trim().slice(0, 48);
    }
    const aria = table.getAttribute("aria-label") || table.getAttribute("summary");
    if (aria && aria.trim()) return aria.trim().slice(0, 48);

    let prev = table.previousElementSibling;
    for (let step = 0; step < 3 && prev; step += 1) {
      if (/^H[1-6]$/.test(prev.tagName) && prev.textContent.trim()) {
        return prev.textContent.trim().slice(0, 48);
      }
      prev = prev.previousElementSibling;
    }
    const firstRow = table.querySelector("tr");
    if (firstRow && firstRow.textContent.trim()) {
      return firstRow.textContent.replace(/\s+/g, " ").trim().slice(0, 42);
    }
    return `Table #${idx + 1}`;
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (!msg || typeof msg !== "object") return false;

    if (msg.type === "SCAN_TABLES") {
      clearHighlight();
      const nodes = Array.from(document.querySelectorAll("table"));
      const list = nodes
        .map((tbl, idx) => {
          const rows = tbl.querySelectorAll("tr").length;
          const cells = tbl.querySelectorAll("th, td").length;
          if (rows === 0 || cells === 0) return null;
          return {
            index: idx,
            title: inferTableTitle(tbl, idx),
            approxRows: rows,
            approxCells: cells,
            outerHTML: tbl.outerHTML,
          };
        })
        .filter(Boolean);

      sendResponse({ ok: true, pageTitle: document.title || "Page", tables: list });
      return true;
    }

    if (msg.type === "HIGHLIGHT_TABLE") {
      clearHighlight();
      const nodes = document.querySelectorAll("table");
      const target = nodes[msg.index];
      if (target) {
        prevOutline = target.style.outline || "";
        target.style.outline = "3px solid #10b981";
        target.scrollIntoView({ behavior: "smooth", block: "center" });
        highlightedEl = target;
      }
      sendResponse({ ok: true });
      return true;
    }

    if (msg.type === "CLEAR_HIGHLIGHT") {
      clearHighlight();
      sendResponse({ ok: true });
      return true;
    }

    return false;
  });
})();
