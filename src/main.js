import {
  createDocAuthSystem,
  defaultToolbarConfig,
  defaultActions,
} from "@nutrient-sdk/document-authoring";
import {
  initGenerationUI, generationActions, generationToolbarItems,
  TEMPLATES, buildTemplate, activateTemplate,
} from "./generate-ui.js";

let docAuthSystem;
let editor;
let currentDocument;
const UNIT_CM = 0.5;

// ═══════════════════════════════════════════════════════════════════
//  LOADING
// ═══════════════════════════════════════════════════════════════════
function hideLoading() {
  const el = document.getElementById("loading-screen");
  el.classList.add("fade-out");
  setTimeout(() => el.remove(), 300);
}

// ═══════════════════════════════════════════════════════════════════
//  DOCUMENT LOADING — new blank doc + open local DOCX
// ═══════════════════════════════════════════════════════════════════
async function newBlankDocument() {
  const doc = await docAuthSystem.createDocumentFromPlaintext("");
  currentDocument = doc;
  await editor.setCurrentDocument(doc);
}

let openFileInput;
function setupOpenFileInput() {
  openFileInput = document.createElement("input");
  openFileInput.type = "file";
  openFileInput.accept = ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  openFileInput.style.display = "none";
  openFileInput.addEventListener("change", async () => {
    const file = openFileInput.files[0];
    openFileInput.value = "";
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      window.pocToast?.("Only DOCX files can be opened for editing — PDF is the generated output format.", true);
      return;
    }
    try {
      const doc = await docAuthSystem.importDOCX(file);
      currentDocument = doc;
      await editor.setCurrentDocument(doc);
      window.pocToast?.(`Opened ${file.name}`);
    } catch (err) {
      console.error(err);
      window.pocToast?.(`Could not open ${file.name} — see console.`, true);
    }
  });
  document.body.appendChild(openFileInput);
}

// ═══════════════════════════════════════════════════════════════════
//  SAMPLES PANEL
// ═══════════════════════════════════════════════════════════════════
function buildSamplesPanel() {
  const panel = document.getElementById("samples-panel");
  const overlay = document.getElementById("samples-overlay");
  const list = panel.querySelector(".samples-list");
  const TPL_ICON = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="9" x2="9" y2="21"/></svg>`;

  const addSection = (label) => {
    const h = document.createElement("div");
    h.style.cssText = "padding:10px 20px 4px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--app-gray-500);";
    h.textContent = label;
    list.appendChild(h);
  };

  addSection("Sample templates with merge fields");
  for (const tpl of TEMPLATES) {
    const item = document.createElement("div");
    item.className = "sample-item";
    item.innerHTML = `<span class="item-icon">${TPL_ICON}</span><span class="item-label">${tpl.name}<br><span style="font-weight:400;font-size:11px;color:var(--app-gray-500);">${tpl.description}</span></span>`;
    item.addEventListener("click", async () => {
      closeSamplesPanel();
      const blob = await buildTemplate(tpl.id);
      const doc = await docAuthSystem.importDOCX(blob);
      currentDocument = doc;
      await editor.setCurrentDocument(doc);
      activateTemplate(tpl);
    });
    list.appendChild(item);
  }
  overlay.addEventListener("click", closeSamplesPanel);
  panel.querySelector(".panel-close").addEventListener("click", closeSamplesPanel);
}
function toggleSamplesPanel() {
  const p = document.getElementById("samples-panel"), o = document.getElementById("samples-overlay");
  const isOpen = p.classList.contains("open");
  p.classList.toggle("open", !isOpen); o.classList.toggle("open", !isOpen);
}
function closeSamplesPanel() {
  document.getElementById("samples-panel").classList.remove("open");
  document.getElementById("samples-overlay").classList.remove("open");
}

// ═══════════════════════════════════════════════════════════════════
//  TABLE CONFIG MODAL
// ═══════════════════════════════════════════════════════════════════
let tableConfig = {
  columns: [
    { heading: "Column A", rowHeights: [1, 1, 1, 1] },
    { heading: "Column B", rowHeights: [1, 1, 1] },
    { heading: "Column C", rowHeights: [1, 1, 1, 1, 1] },
  ],
  // Horizontal merges: each { microRow, fromCol, toCol }
  hMerges: [],
  // Border overrides: key "col,logRow,side" → true/false
  borders: {},
  // Cell colors: key "col,logRow" → hex color (e.g. "#FF0000")
  cellColors: {},
};

function openTableModal() {
  document.getElementById("table-modal-overlay").classList.add("open");
  renderColumnCards();
  renderPreview();
}
function closeTableModal() {
  document.getElementById("table-modal-overlay").classList.remove("open");
}

function renderColumnCards() {
  const container = document.getElementById("columns-config");
  container.innerHTML = "";
  tableConfig.columns.forEach((col, ci) => {
    const card = document.createElement("div");
    card.className = "col-card";
    let rowsHtml = "";
    col.rowHeights.forEach((h, ri) => {
      rowsHtml += `<div class="row-height-row"><span class="row-height-label">R${ri + 1}</span><input type="number" class="row-height-input" min="0.5" max="20" step="0.5" value="${h}" data-col="${ci}" data-row="${ri}" /><span class="row-height-unit">cm</span></div>`;
    });
    card.innerHTML = `
      <div class="col-title">Col ${ci + 1}</div>
      <input class="col-heading-input" type="text" value="${col.heading}" placeholder="Heading" data-col="${ci}" />
      <div class="row-controls">
        <button type="button" data-col="${ci}" data-dir="-1">-</button>
        <span class="row-count">${col.rowHeights.length}</span>
        <button type="button" data-col="${ci}" data-dir="1">+</button>
      </div>
      <div class="row-label">rows</div>
      <div class="row-heights-list">${rowsHtml}</div>`;
    container.appendChild(card);
  });
  container.querySelectorAll(".row-controls button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const ci = parseInt(btn.dataset.col), dir = parseInt(btn.dataset.dir);
      const col = tableConfig.columns[ci];
      if (dir === 1) col.rowHeights.push(1);
      else if (col.rowHeights.length > 1) col.rowHeights.pop();
      renderColumnCards(); renderPreview();
    });
  });
  container.querySelectorAll(".col-heading-input").forEach((input) => {
    input.addEventListener("input", () => { tableConfig.columns[parseInt(input.dataset.col)].heading = input.value; renderPreview(); });
  });
  container.querySelectorAll(".row-height-input").forEach((input) => {
    input.addEventListener("input", () => {
      const val = parseFloat(input.value);
      if (val > 0) { tableConfig.columns[parseInt(input.dataset.col)].rowHeights[parseInt(input.dataset.row)] = val; renderPreview(); }
    });
  });
}

function renderPreview() {
  const preview = document.getElementById("table-preview");
  const PX_PER_CM = 30;
  const config = tableConfig.columns;
  let html = '<div class="preview-grid">';
  config.forEach((col) => {
    html += '<div class="preview-col">';
    html += `<div class="preview-col-header">${col.heading || "&nbsp;"}</div>`;
    col.rowHeights.forEach((h, ri) => {
      const px = Math.max(20, Math.round(h * PX_PER_CM));
      const isLast = ri === col.rowHeights.length - 1;
      html += `<div class="preview-col-row${isLast ? " last" : ""}" style="height:${px}px">${h}cm</div>`;
    });
    html += "</div>";
  });
  html += "</div>";
  preview.innerHTML = html;
}

function setupTableModal() {
  const overlay = document.getElementById("table-modal-overlay");
  const modal = document.getElementById("table-modal");
  modal.querySelector(".modal-close").addEventListener("click", closeTableModal);
  document.getElementById("table-cancel").addEventListener("click", closeTableModal);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) closeTableModal(); });
  document.getElementById("col-count-value").textContent = tableConfig.columns.length;
  document.getElementById("col-minus").addEventListener("click", () => {
    if (tableConfig.columns.length > 1) {
      tableConfig.columns.pop();
      document.getElementById("col-count-value").textContent = tableConfig.columns.length;
      renderColumnCards(); renderPreview();
    }
  });
  document.getElementById("col-plus").addEventListener("click", () => {
    if (tableConfig.columns.length < 8) {
      tableConfig.columns.push({ heading: `Column ${String.fromCharCode(65 + tableConfig.columns.length)}`, rowHeights: [1, 1, 1] });
      document.getElementById("col-count-value").textContent = tableConfig.columns.length;
      renderColumnCards(); renderPreview();
    }
  });
  document.getElementById("table-insert").addEventListener("click", async () => {
    await insertSpecialTable();
    closeTableModal();
  });
}

// ═══════════════════════════════════════════════════════════════════
//  BUILD & INSERT TABLE (XML micro-rows injected into DOCX)
// ═══════════════════════════════════════════════════════════════════
function escapeXml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildSpecialTableXml(config) {
  const numCols = config.length;
  const totalTwips = 9638;
  const colW = Math.floor(totalTwips / numCols);
  const TWIPS_05CM = 284;
  const B = `w:val="single" w:sz="4" w:color="CCCCCC"`;
  const N = `w:val="none" w:sz="0" w:color="auto"`;
  const borders = tableConfig.borders || {};

  const mergeMap = config.map((col) => {
    const spans = col.rowHeights.map((h) => Math.max(1, Math.round(h / UNIT_CM)));
    const total = spans.reduce((a, b) => a + b, 0);
    const boundaries = new Set();
    // Also track which logical row each micro-row belongs to
    const microToLogical = [];
    let pr = 0;
    for (let li = 0; li < spans.length; li++) {
      for (let s = 0; s < spans[li]; s++) {
        microToLogical.push(li);
      }
      pr += spans[li];
      boundaries.add(pr - 1);
    }
    return { total, boundaries, microToLogical };
  });
  const maxPhys = Math.max(...mergeMap.map((m) => m.total));

  // Helper: check if a border is enabled for a given cell side
  function isBorderOn(col, microRow, side) {
    const info = mergeMap[col];
    if (!info || microRow >= info.microToLogical.length) return false;
    const logRow = info.microToLogical[microRow];
    const key = `${col},${logRow},${side}`;
    return borders[key] !== false;
  }

  // Helper: get cell shading XML
  const cellColors = tableConfig.cellColors || {};
  function getCellColor(col, microRow) {
    const info = mergeMap[col];
    if (!info || microRow >= info.microToLogical.length) return null;
    const logRow = info.microToLogical[microRow];
    return cellColors[`${col},${logRow}`] || null;
  }
  function getShadingXml(col, microRow) {
    const color = getCellColor(col, microRow);
    if (!color) return "";
    return `<w:shd w:val="clear" w:fill="${color.replace("#", "")}"/>`;
  }
  // When a cell has a color, internal borders should match the fill color
  // to avoid white gaps between micro-rows
  function getInvisibleBorder(col, microRow) {
    const color = getCellColor(col, microRow);
    if (color) {
      return `w:val="single" w:sz="4" w:color="${color.replace("#", "")}"`;
    }
    return N;
  }

  let gridXml = "";
  for (let i = 0; i < numCols; i++) gridXml += `<w:gridCol w:w="${colW}"/>`;

  let headerCells = "";
  for (let c = 0; c < numCols; c++) {
    headerCells += `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/><w:tcBorders><w:top ${B}/><w:left ${B}/><w:bottom ${B}/><w:right ${B}/></w:tcBorders></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="1F2937"/></w:rPr><w:t>${escapeXml(config[c].heading)}</w:t></w:r></w:p></w:tc>`;
  }
  let rowsXml = `<w:tr><w:trPr><w:trHeight w:val="${TWIPS_05CM}" w:hRule="exact"/></w:trPr>${headerCells}</w:tr>`;

  // Get horizontal merges
  const hMerges = tableConfig.hMerges || [];

  for (let pr = 0; pr < maxPhys; pr++) {
    let cells = "";
    // Check if this micro-row has a horizontal merge
    const hMerge = hMerges.find((m) => m.microRow === pr);

    let c = 0;
    while (c < numCols) {
      const info = mergeMap[c];
      const isActive = pr < info.total;

      // Check if this cell starts a horizontal span
      if (hMerge && c === hMerge.fromCol && isActive) {
        const spanCount = hMerge.toCol - hMerge.fromCol + 1;
        const spanWidth = colW * spanCount;
        const hid = getInvisibleBorder(c, pr);
        const bBottom = (info.boundaries.has(pr) && isBorderOn(c, pr, "bottom")) ? B : hid;
        const bTop = (pr === 0 && isBorderOn(c, pr, "top")) ? B : hid;
        const bLeft = isBorderOn(c, pr, "left") ? B : hid;
        const bRight = isBorderOn(hMerge.toCol, pr, "right") ? B : hid;
        cells += `<w:tc><w:tcPr><w:tcW w:w="${spanWidth}" w:type="dxa"/><w:gridSpan w:val="${spanCount}"/><w:tcBorders><w:top ${bTop}/><w:left ${bLeft}/><w:bottom ${bBottom}/><w:right ${bRight}/></w:tcBorders>${getShadingXml(c, pr)}</w:tcPr><w:p/></w:tc>`;
        c = hMerge.toCol + 1;
        continue;
      }

      // Skip columns absorbed by horizontal merge
      if (hMerge && c > hMerge.fromCol && c <= hMerge.toCol) {
        c++;
        continue;
      }

      if (!isActive) {
        cells += `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/><w:tcBorders><w:top ${N}/><w:left ${N}/><w:bottom ${N}/><w:right ${N}/></w:tcBorders></w:tcPr><w:p/></w:tc>`;
      } else {
        const hid = getInvisibleBorder(c, pr);
        const bBottom = (info.boundaries.has(pr) && isBorderOn(c, pr, "bottom")) ? B : hid;
        const bTop = (pr === 0 && isBorderOn(c, pr, "top")) ? B : hid;
        const bLeft = isBorderOn(c, pr, "left") ? B : hid;
        const bRight = isBorderOn(c, pr, "right") ? B : hid;
        cells += `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/><w:tcBorders><w:top ${bTop}/><w:left ${bLeft}/><w:bottom ${bBottom}/><w:right ${bRight}/></w:tcBorders>${getShadingXml(c, pr)}</w:tcPr><w:p/></w:tc>`;
      }
      c++;
    }
    rowsXml += `<w:tr><w:trPr><w:trHeight w:val="${TWIPS_05CM}" w:hRule="exact"/></w:trPr>${cells}</w:tr>`;
  }

  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top ${N}/><w:left ${N}/><w:bottom ${N}/><w:right ${N}/><w:insideH ${N}/><w:insideV ${N}/></w:tblBorders></w:tblPr><w:tblGrid>${gridXml}</w:tblGrid>${rowsXml}</w:tbl>`;
}

async function insertSpecialTable() {
  const config = tableConfig.columns.map((c) => ({ heading: c.heading, rowHeights: [...c.rowHeights] }));
  const tableXml = buildSpecialTableXml(config);

  const marker = `__TBL_${Date.now()}__`;
  if (editor.hasActiveCursor()) editor.insertTextAtCursor(marker);
  else {
    await editor.currentDocument().transaction(async ({ draft }) => {
      const sections = draft.body().sections();
      const content = sections[sections.length - 1].content();
      const para = content.addParagraph();
      para.asTextView().setText(marker);
      return { commit: true };
    });
  }

  const exportedBuffer = await editor.currentDocument().exportDOCX();
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(exportedBuffer);
  let docXml = await zip.file("word/document.xml").async("string");

  let replaced = false;
  docXml = docXml.replace(
    /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/g,
    (match, open, inner, close) => {
      if (replaced) return match;
      const plain = inner.replace(/<[^>]+>/g, "");
      if (plain.includes(marker)) { replaced = true; return tableXml; }
      return match;
    }
  );
  if (!replaced) {
    const sectIdx = docXml.lastIndexOf("<w:sectPr");
    if (sectIdx !== -1) docXml = docXml.slice(0, sectIdx) + tableXml + docXml.slice(sectIdx);
    else docXml = docXml.replace("</w:body>", tableXml + "</w:body>");
  }

  zip.file("word/document.xml", docXml);
  const newBlob = await zip.generateAsync({ type: "blob" });
  const newDoc = await docAuthSystem.importDOCX(newBlob);
  currentDocument = newDoc;
  await editor.setCurrentDocument(newDoc);
}

// ═══════════════════════════════════════════════════════════════════
//  CONTEXT MENU INJECTION
// ═══════════════════════════════════════════════════════════════════
function findInShadowDom(selector) {
  const results = [...document.querySelectorAll(selector)];
  function searchNode(node) {
    if (node.shadowRoot) {
      results.push(...node.shadowRoot.querySelectorAll(selector));
      node.shadowRoot.querySelectorAll("*").forEach(searchNode);
    }
    node.querySelectorAll("*").forEach((child) => {
      if (child.shadowRoot) {
        results.push(...child.shadowRoot.querySelectorAll(selector));
        child.shadowRoot.querySelectorAll("*").forEach(searchNode);
      }
    });
  }
  searchNode(document.body);
  return results;
}

function setupContextMenuInjection() {
  document.addEventListener("contextmenu", () => {
    let attempts = 0;
    const tryInject = () => {
      attempts++;
      const menus = findInShadowDom('ul[role="menu"]');
      for (const menu of menus) {
        if (menu.dataset.mergeInjected) continue;
        const hasCut = menu.querySelector('[data-key="cut"], [data-key="copy"]');
        if (!hasCut) continue;
        menu.dataset.mergeInjected = "true";
        injectMergeSplitButtons(menu);
        return;
      }
      if (attempts < 20) setTimeout(tryInject, 100);
    };
    setTimeout(tryInject, 100);
  });
}

function injectMergeSplitButtons(menu) {
  const sep = document.createElement("li");
  sep.setAttribute("role", "separator");
  sep.className = "gukncs0 BaselineUI-Separator";
  sep.dataset.orientation = "horizontal";

  const groupLi = document.createElement("li");
  groupLi.setAttribute("role", "presentation");
  const groupUl = document.createElement("ul");
  groupUl.setAttribute("role", "group");
  groupUl.setAttribute("aria-label", "Table");
  groupUl.className = "xyd2ss0";
  groupUl.style.cssText = "list-style:none;margin:0;padding:4px 0;";

  const refItem = menu.querySelector('li[role="menuitem"]');
  const refClass = refItem ? refItem.className : "";

  groupUl.appendChild(buildCtxItem("Advanced options", refClass, () => { openMergeModal(); }));
  groupLi.appendChild(groupUl);
  menu.appendChild(sep);
  menu.appendChild(groupLi);
}

function buildCtxItem(label, className, onClick) {
  const li = document.createElement("li");
  li.setAttribute("role", "menuitem");
  li.setAttribute("tabindex", "-1");
  if (className) li.className = className;
  li.innerHTML = `<div class="_1yep5z6ih _1yep5z6ii _1yep5z6ij _1yep5z6j8 _1yep5z6j9 _1yep5z6ja _1yep5z69l _1yep5z69m _1yep5z69n BaselineUI-Box"><span>${label}</span></div>`;
  li.addEventListener("click", (e) => {
    e.stopPropagation();
    onClick();
    const wrapper = li.closest("div._1gzftth6");
    if (wrapper) wrapper.remove();
  });
  return li;
}

// ═══════════════════════════════════════════════════════════════════
//  MERGE MODAL (rectangular cell selection)
// ═══════════════════════════════════════════════════════════════════
function openMergeModal() {
  document.getElementById("merge-modal-overlay")?.remove();

  const config = tableConfig.columns;
  const PX_PER_CM = 25;

  const overlay = document.createElement("div");
  overlay.id = "merge-modal-overlay";
  overlay.className = "merge-overlay";

  const modal = document.createElement("div");
  modal.className = "merge-modal";
  modal.innerHTML = `<div class="merge-modal-header"><h3>Merge / Split Cells</h3><p>Click cells to build a rectangular selection, then Merge. Click a large cell to Split.</p></div>`;

  const grid = document.createElement("div");
  grid.className = "merge-grid";

  const cellEls = {};
  const selected = new Set();

  function updateHighlight() {
    const bounds = getSelectionBounds();
    Object.entries(cellEls).forEach(([key, el]) => {
      el.classList.toggle("merge-cell-selected", selected.has(key));
      el.classList.remove("merge-cell-candidate");
    });
    if (selected.size > 0) {
      getNeighborKeys(bounds).forEach((key) => {
        if (!selected.has(key) && cellEls[key]) cellEls[key].classList.add("merge-cell-candidate");
      });
    }
  }

  function getSelectionBounds() {
    if (selected.size === 0) return null;
    let minC = Infinity, maxC = -1, minR = Infinity, maxR = -1;
    for (const key of selected) {
      const [c, r] = key.split(",").map(Number);
      minC = Math.min(minC, c); maxC = Math.max(maxC, c);
      minR = Math.min(minR, r); maxR = Math.max(maxR, r);
    }
    return { minC, maxC, minR, maxR };
  }

  function getNeighborKeys(bounds) {
    if (!bounds) return [];
    const keys = [];
    for (let c = bounds.minC; c <= bounds.maxC; c++) {
      if (bounds.minR > 0) keys.push(`${c},${bounds.minR - 1}`);
      keys.push(`${c},${bounds.maxR + 1}`);
    }
    for (let r = bounds.minR; r <= bounds.maxR; r++) {
      if (bounds.minC > 0) keys.push(`${bounds.minC - 1},${r}`);
      keys.push(`${bounds.maxC + 1},${r}`);
    }
    return keys.filter((k) => cellEls[k]);
  }

  function isValidAdd(col, row) {
    if (selected.size === 0) return true;
    return getNeighborKeys(getSelectionBounds()).includes(`${col},${row}`);
  }

  config.forEach((col, ci) => {
    const colDiv = document.createElement("div");
    colDiv.className = "merge-col";
    const hdr = document.createElement("div");
    hdr.className = "merge-col-header";
    hdr.textContent = col.heading;
    colDiv.appendChild(hdr);

    col.rowHeights.forEach((h, ri) => {
      const px = Math.max(24, Math.round(h * PX_PER_CM));
      const cell = document.createElement("div");
      cell.className = "merge-cell";
      cell.style.height = `${px}px`;
      cell.textContent = `${h}cm`;
      const key = `${ci},${ri}`;
      cellEls[key] = cell;

      cell.addEventListener("click", () => {
        if (selected.has(key)) {
          selected.delete(key);
          const bounds = getSelectionBounds();
          if (bounds) {
            const expected = (bounds.maxC - bounds.minC + 1) * (bounds.maxR - bounds.minR + 1);
            if (selected.size !== expected) selected.clear();
          }
        } else if (isValidAdd(ci, ri)) {
          selected.add(key);
        }
        updateHighlight();
      });

      colDiv.appendChild(cell);
    });
    grid.appendChild(colDiv);
  });

  modal.appendChild(grid);

  const footer = document.createElement("div");
  footer.className = "merge-modal-footer";

  const mergeBtn = document.createElement("button");
  mergeBtn.className = "btn btn-primary";
  mergeBtn.textContent = "Merge selected";
  mergeBtn.addEventListener("click", async () => {
    const bounds = getSelectionBounds();
    if (!bounds || selected.size < 2) return;

    const isVertical = bounds.minC === bounds.maxC && bounds.minR !== bounds.maxR;
    const isHorizontal = bounds.minR === bounds.maxR && bounds.minC !== bounds.maxC;

    if (isVertical) {
      // Vertical merge: combine row heights in the selected column
      const col = tableConfig.columns[bounds.minC];
      if (bounds.minR < col.rowHeights.length) {
        const to = Math.min(bounds.maxR, col.rowHeights.length - 1);
        const merged = col.rowHeights.slice(bounds.minR, to + 1).reduce((a, b) => a + b, 0);
        col.rowHeights.splice(bounds.minR, to - bounds.minR + 1, merged);
      }
    } else if (isHorizontal) {
      // Horizontal merge: store a gridSpan merge for all micro-rows of this logical row
      // Convert logical row to micro-row range (use first selected column as reference)
      const col = tableConfig.columns[bounds.minC];
      const spans = col.rowHeights.map((h) => Math.max(1, Math.round(h / UNIT_CM)));
      let microStart = 0;
      for (let r = 0; r < bounds.minR; r++) microStart += spans[r];
      const microEnd = microStart + spans[bounds.minR] - 1;

      for (let mr = microStart; mr <= microEnd; mr++) {
        tableConfig.hMerges.push({ microRow: mr, fromCol: bounds.minC, toCol: bounds.maxC });
      }

    } else {
      // Mixed: do vertical merge in each column + horizontal merge
      for (let c = bounds.minC; c <= bounds.maxC; c++) {
        const col = tableConfig.columns[c];
        if (bounds.minR < col.rowHeights.length) {
          const to = Math.min(bounds.maxR, col.rowHeights.length - 1);
          const merged = col.rowHeights.slice(bounds.minR, to + 1).reduce((a, b) => a + b, 0);
          col.rowHeights.splice(bounds.minR, to - bounds.minR + 1, merged);
        }
      }
      // Also add horizontal merge for the resulting row
      const col = tableConfig.columns[bounds.minC];
      const spans = col.rowHeights.map((h) => Math.max(1, Math.round(h / UNIT_CM)));
      let microStart = 0;
      for (let r = 0; r < bounds.minR; r++) microStart += spans[r];
      const microEnd = microStart + spans[bounds.minR] - 1;
      for (let mr = microStart; mr <= microEnd; mr++) {
        tableConfig.hMerges.push({ microRow: mr, fromCol: bounds.minC, toCol: bounds.maxC });
      }
    }

    overlay.remove();
    await regenerateTable();
  });

  const splitBtn = document.createElement("button");
  splitBtn.className = "btn btn-secondary";
  splitBtn.textContent = "Split selected";
  splitBtn.addEventListener("click", async () => {
    if (selected.size === 0) return;
    const toSplit = [...selected].map((k) => k.split(",").map(Number));
    toSplit.sort((a, b) => b[1] - a[1]);
    for (const [c, r] of toSplit) {
      const col = tableConfig.columns[c];
      if (col.rowHeights[r] > UNIT_CM) {
        const h = col.rowHeights[r];
        const half = Math.round((h / 2) / UNIT_CM) * UNIT_CM || UNIT_CM;
        const other = Math.round((h - half) / UNIT_CM) * UNIT_CM || UNIT_CM;
        col.rowHeights.splice(r, 1, half, other);
      }
    }
    overlay.remove();
    await regenerateTable();
  });

  const closeBtn = document.createElement("button");
  closeBtn.className = "btn btn-secondary";
  closeBtn.textContent = "Close";
  closeBtn.addEventListener("click", () => overlay.remove());

  const bordersBtn = document.createElement("button");
  bordersBtn.className = "btn btn-secondary";
  bordersBtn.textContent = "Borders...";
  bordersBtn.addEventListener("click", () => {
    overlay.remove();
    openBordersModal();
  });

  const colorsBtn = document.createElement("button");
  colorsBtn.className = "btn btn-secondary";
  colorsBtn.textContent = "Colors...";
  colorsBtn.addEventListener("click", () => {
    overlay.remove();
    openColorsModal();
  });

  footer.appendChild(closeBtn);
  footer.appendChild(colorsBtn);
  footer.appendChild(bordersBtn);
  footer.appendChild(splitBtn);
  footer.appendChild(mergeBtn);
  modal.appendChild(footer);

  overlay.appendChild(modal);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
}

// ═══════════════════════════════════════════════════════════════════
//  BORDERS MODAL
// ═══════════════════════════════════════════════════════════════════
function openBordersModal() {
  document.getElementById("borders-modal-overlay")?.remove();

  const config = tableConfig.columns;
  const PX_PER_CM = 25;
  const borders = tableConfig.borders;

  const overlay = document.createElement("div");
  overlay.id = "borders-modal-overlay";
  overlay.className = "merge-overlay";

  const modal = document.createElement("div");
  modal.className = "merge-modal";
  modal.style.width = "600px";
  modal.innerHTML = `<div class="merge-modal-header"><h3>Edit Borders</h3><p>Click on a border to toggle it on/off.</p></div>`;

  const gridWrap = document.createElement("div");
  gridWrap.style.cssText = "padding:20px 24px;position:relative;";

  // Build a visual grid where each cell shows its 4 borders as clickable edges
  const grid = document.createElement("div");
  grid.className = "borders-grid";

  config.forEach((col, ci) => {
    const colDiv = document.createElement("div");
    colDiv.className = "borders-col";

    // Header
    const hdr = document.createElement("div");
    hdr.className = "borders-cell-header";
    hdr.textContent = col.heading;
    colDiv.appendChild(hdr);

    col.rowHeights.forEach((h, ri) => {
      const px = Math.max(30, Math.round(h * PX_PER_CM));
      const cell = document.createElement("div");
      cell.className = "borders-cell";
      cell.style.height = `${px}px`;

      // Create 4 border edges (top, right, bottom, left)
      ["top", "right", "bottom", "left"].forEach((side) => {
        const key = `${ci},${ri},${side}`;
        const edge = document.createElement("div");
        edge.className = `borders-edge borders-edge-${side}`;
        const isOn = borders[key] !== false; // default = on
        edge.classList.toggle("off", !isOn);
        edge.title = `${side} border (click to toggle)`;

        edge.addEventListener("click", (e) => {
          e.stopPropagation();
          const current = borders[key] !== false;
          borders[key] = !current;
          edge.classList.toggle("off", current);
        });

        cell.appendChild(edge);
      });

      colDiv.appendChild(cell);
    });

    grid.appendChild(colDiv);
  });

  gridWrap.appendChild(grid);
  modal.appendChild(gridWrap);

  // Footer
  const footer = document.createElement("div");
  footer.className = "merge-modal-footer";

  const applyBtn = document.createElement("button");
  applyBtn.className = "btn btn-primary";
  applyBtn.textContent = "Apply";
  applyBtn.addEventListener("click", async () => {
    overlay.remove();
    await regenerateTable();
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-secondary";
  cancelBtn.textContent = "Cancel";
  cancelBtn.addEventListener("click", () => overlay.remove());

  const allOnBtn = document.createElement("button");
  allOnBtn.className = "btn btn-secondary";
  allOnBtn.textContent = "All borders";
  allOnBtn.addEventListener("click", () => {
    config.forEach((col, ci) => {
      col.rowHeights.forEach((_, ri) => {
        ["top", "right", "bottom", "left"].forEach((s) => { borders[`${ci},${ri},${s}`] = true; });
      });
    });
    grid.querySelectorAll(".borders-edge").forEach((e) => e.classList.remove("off"));
  });

  const noneBtn = document.createElement("button");
  noneBtn.className = "btn btn-secondary";
  noneBtn.textContent = "No borders";
  noneBtn.addEventListener("click", () => {
    config.forEach((col, ci) => {
      col.rowHeights.forEach((_, ri) => {
        ["top", "right", "bottom", "left"].forEach((s) => { borders[`${ci},${ri},${s}`] = false; });
      });
    });
    grid.querySelectorAll(".borders-edge").forEach((e) => e.classList.add("off"));
  });

  footer.appendChild(cancelBtn);
  footer.appendChild(noneBtn);
  footer.appendChild(allOnBtn);
  footer.appendChild(applyBtn);
  modal.appendChild(footer);

  overlay.appendChild(modal);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
}

// ═══════════════════════════════════════════════════════════════════
//  COLORS MODAL
// ═══════════════════════════════════════════════════════════════════
const PRESET_COLORS = [
  "#FFFFFF", "#F8F9FA", "#E9ECEF", "#DEE2E6", "#ADB5BD", "#6C757D", "#495057", "#343A40", "#212529", "#000000",
  "#FFF3BF", "#FFEC99", "#FFD43B", "#FCC419", "#FAB005", "#F59F00", "#E67700",
  "#D3F9D8", "#B2F2BB", "#8CE99A", "#69DB7C", "#51CF66", "#40C057", "#2F9E44",
  "#D0EBFF", "#A5D8FF", "#74C0FC", "#4DABF7", "#339AF0", "#228BE6", "#1971C2",
  "#E8F0FE", "#BAD7F2", "#7FB3E0", "#4A90D9", "#2B6CB0", "#1A4F8B", "#0B3D6B",
  "#FFE3E3", "#FFC9C9", "#FFA8A8", "#FF8787", "#FF6B6B", "#FA5252", "#E03131",
  "#F3D9FA", "#EEBEFA", "#D0BFFF", "#B197FC", "#9775FA", "#845EF7", "#7048E8",
];

function openColorsModal() {
  document.getElementById("colors-modal-overlay")?.remove();

  const config = tableConfig.columns;
  const cellColors = tableConfig.cellColors;
  const PX_PER_CM = 25;

  const overlay = document.createElement("div");
  overlay.id = "colors-modal-overlay";
  overlay.className = "merge-overlay";

  const modal = document.createElement("div");
  modal.className = "merge-modal";
  modal.style.width = "650px";
  modal.innerHTML = `<div class="merge-modal-header"><h3>Cell Colors</h3><p>Click a cell, then pick a color.</p></div>`;

  const body = document.createElement("div");
  body.style.cssText = "padding:20px 24px;display:flex;gap:20px;";

  // Left: table grid
  const grid = document.createElement("div");
  grid.className = "merge-grid";
  grid.style.flex = "1";

  let selectedCell = null;
  let selectedKey = null;

  config.forEach((col, ci) => {
    const colDiv = document.createElement("div");
    colDiv.className = "merge-col";
    const hdr = document.createElement("div");
    hdr.className = "merge-col-header";
    hdr.textContent = col.heading;
    colDiv.appendChild(hdr);

    col.rowHeights.forEach((h, ri) => {
      const px = Math.max(24, Math.round(h * PX_PER_CM));
      const cell = document.createElement("div");
      cell.className = "merge-cell colors-cell";
      cell.style.height = `${px}px`;
      cell.style.cursor = "pointer";
      const key = `${ci},${ri}`;
      const color = cellColors[key];
      if (color) cell.style.background = color;

      cell.addEventListener("click", () => {
        if (selectedCell) selectedCell.classList.remove("colors-cell-active");
        selectedCell = cell;
        selectedKey = key;
        cell.classList.add("colors-cell-active");
      });

      colDiv.appendChild(cell);
    });
    grid.appendChild(colDiv);
  });

  // Right: color palette
  const palette = document.createElement("div");
  palette.className = "colors-palette";

  PRESET_COLORS.forEach((color) => {
    const swatch = document.createElement("div");
    swatch.className = "colors-swatch";
    swatch.style.background = color;
    if (color === "#FFFFFF") swatch.style.border = "1px solid #DEE2E6";
    swatch.addEventListener("click", () => {
      if (!selectedKey) return;
      cellColors[selectedKey] = color;
      if (selectedCell) selectedCell.style.background = color;
    });
    palette.appendChild(swatch);
  });

  // Custom color input
  const customRow = document.createElement("div");
  customRow.style.cssText = "margin-top:8px;display:flex;align-items:center;gap:6px;";
  const customInput = document.createElement("input");
  customInput.type = "color";
  customInput.value = "#2563EB";
  customInput.style.cssText = "width:32px;height:32px;border:none;cursor:pointer;padding:0;";
  const customLabel = document.createElement("span");
  customLabel.textContent = "Custom";
  customLabel.style.cssText = "font-size:11px;color:#868E96;";
  customInput.addEventListener("input", () => {
    if (!selectedKey) return;
    cellColors[selectedKey] = customInput.value;
    if (selectedCell) selectedCell.style.background = customInput.value;
  });
  customRow.appendChild(customInput);
  customRow.appendChild(customLabel);
  palette.appendChild(customRow);

  body.appendChild(grid);
  body.appendChild(palette);
  modal.appendChild(body);

  // Footer
  const footer = document.createElement("div");
  footer.className = "merge-modal-footer";

  const applyBtn = document.createElement("button");
  applyBtn.className = "btn btn-primary";
  applyBtn.textContent = "Apply";
  applyBtn.addEventListener("click", async () => {
    overlay.remove();
    await regenerateTable();
  });

  const clearBtn = document.createElement("button");
  clearBtn.className = "btn btn-secondary";
  clearBtn.textContent = "Clear all";
  clearBtn.addEventListener("click", () => {
    Object.keys(cellColors).forEach((k) => delete cellColors[k]);
    grid.querySelectorAll(".colors-cell").forEach((c) => c.style.background = "");
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-secondary";
  cancelBtn.textContent = "Cancel";
  cancelBtn.addEventListener("click", () => overlay.remove());

  footer.appendChild(cancelBtn);
  footer.appendChild(clearBtn);
  footer.appendChild(applyBtn);
  modal.appendChild(footer);

  overlay.appendChild(modal);
  overlay.addEventListener("click", (e) => { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
}

// ═══════════════════════════════════════════════════════════════════
//  REGENERATE TABLE (export → replace XML → reimport)
// ═══════════════════════════════════════════════════════════════════
// Find all top-level <w:tbl> positions (handles nesting)
function findTopLevelTables(xml) {
  const tables = [];
  let i = 0;
  while (i < xml.length) {
    const openIdx = xml.indexOf("<w:tbl>", i);
    if (openIdx === -1) break;

    // Find matching </w:tbl> accounting for nesting
    let depth = 1;
    let j = openIdx + 7; // length of "<w:tbl>"
    while (j < xml.length && depth > 0) {
      const nextOpen = xml.indexOf("<w:tbl>", j);
      const nextClose = xml.indexOf("</w:tbl>", j);
      if (nextClose === -1) break;
      if (nextOpen !== -1 && nextOpen < nextClose) {
        depth++;
        j = nextOpen + 7;
      } else {
        depth--;
        if (depth === 0) {
          tables.push({ start: openIdx, end: nextClose + 8 }); // 8 = "</w:tbl>".length
        }
        j = nextClose + 8;
      }
    }
    i = j;
  }
  return tables;
}

async function regenerateTable() {
  const config = tableConfig.columns.map((c) => ({ heading: c.heading, rowHeights: [...c.rowHeights] }));
  const tableXml = buildSpecialTableXml(config);

  const exportedBuffer = await editor.currentDocument().exportDOCX();
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(exportedBuffer);
  let docXml = await zip.file("word/document.xml").async("string");

  // Find our custom table by its marker
  const tables = findTopLevelTables(docXml);


  // Find our table: it must contain ALL our column headings
  const headings = tableConfig.columns.map(c => c.heading);
  let target = null;
  for (const t of tables) {
    const tblXml = docXml.substring(t.start, t.end);
    const plain = tblXml.replace(/<[^>]+>/g, "");
    if (headings.every(h => plain.includes(h))) {
      target = t;

      break;
    }
  }

  if (!target) {
    target = tables[tables.length - 1];

  }
  const last = target;


  docXml = docXml.substring(0, last.start) + tableXml + docXml.substring(last.end);

  zip.file("word/document.xml", docXml);
  const newBlob = await zip.generateAsync({ type: "blob" });

  // Verify


  // Destroy editor and recreate with new document to force full refresh
  const newDoc = await docAuthSystem.importDOCX(newBlob);
  editor.destroy();
  currentDocument = newDoc;
  editor = await docAuthSystem.createEditor(document.getElementById("editor"), {
    document: newDoc,
    ui: editorUiConfig(),
  });
  editor.on("document.load", () => { currentDocument = editor.currentDocument(); });

}

// ═══════════════════════════════════════════════════════════════════
//  TOOLBAR ICONS
// ═══════════════════════════════════════════════════════════════════
const SAMPLES_ICON = `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>`)}`;
const NEW_ICON = `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>`)}`;
const OPEN_ICON = `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 20h12a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-7L9 5H4a2 2 0 0 0-2 2v11a2 2 0 0 0 2 2z"/><path d="M12 11v6"/><path d="m9 14 3-3 3 3"/></svg>`)}`;
const TABLE_ICON = `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>`)}`;

// ═══════════════════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════════════════
function editorUiConfig() {
  return {
    actions: [
      ...defaultActions,
      { id: "custom.new", label: "New Document", icon: NEW_ICON, handler: () => newBlankDocument() },
      { id: "custom.open", label: "Open File…", icon: OPEN_ICON, handler: () => openFileInput.click() },
      { id: "custom.samples", label: "Samples", icon: SAMPLES_ICON, handler: () => toggleSamplesPanel() },
      { id: "custom.table", label: "Custom Table", icon: TABLE_ICON, handler: () => openTableModal() },
      ...generationActions(),
    ],
    toolbar: {
      ...defaultToolbarConfig,
      items: [
        ...defaultToolbarConfig.items,
        { type: "separator", id: "sep-custom" },
        { type: "action", id: "new-btn", actionId: "custom.new" },
        { type: "action", id: "open-btn", actionId: "custom.open" },
        { type: "action", id: "samples-btn", actionId: "custom.samples" },
        { type: "action", id: "table-btn", actionId: "custom.table" },
        { type: "built-in", id: "table-ops", builtInType: "table-menu" },
        { type: "separator", id: "sep-generation" },
        ...generationToolbarItems(),
      ],
    },
  };
}

async function init() {
  // Empty license key = evaluation mode (watermark on exported PDFs).
  const licenseKey = (import.meta.env.VITE_DOCAUTH_LICENSE_KEY ?? "").trim();
  docAuthSystem = await createDocAuthSystem(licenseKey ? { licenseKey } : {});

  initGenerationUI({
    getEditor: () => editor,
    getSystem: () => docAuthSystem,
  });

  // Open on the flagship demo template: pixel-perfect certificate with merge fields
  const firstTemplate = TEMPLATES[0];
  const firstBlob = await buildTemplate(firstTemplate.id);
  const firstDoc = await docAuthSystem.importDOCX(firstBlob);
  currentDocument = firstDoc;

  editor = await docAuthSystem.createEditor(document.getElementById("editor"), {
    document: firstDoc,
    ui: editorUiConfig(),
  });
  activateTemplate(firstTemplate);

  editor.on("document.load", () => { currentDocument = editor.currentDocument(); });
  setupOpenFileInput();
  buildSamplesPanel();
  setupTableModal();
  setupContextMenuInjection();
  hideLoading();
}

init().catch((err) => { console.error(err); hideLoading(); });
