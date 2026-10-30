// ═══════════════════════════════════════════════════════════════════
//  GENERATION UI — datasets panel, merge-field panel, generate modal
//  (single PDF, simulated email workflow, batch generation)
// ═══════════════════════════════════════════════════════════════════
import { DATASETS, datasetFields, parseCSV } from "./data.js";
import { TEMPLATES, buildTemplate } from "./templates.js";
import { mergeDocx } from "./merge.js";

let ctx; // { getEditor, getSystem, loadBlob }
let currentDataset = DATASETS.taxCertificates;

// ───────────────────────────── helpers ─────────────────────────────
const $ = (sel) => document.querySelector(sel);

function el(tag, className, html) {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (html !== undefined) e.innerHTML = html;
  return e;
}

function toast(msg, isError = false) {
  document.querySelector(".poc-toast")?.remove();
  const t = el("div", "poc-toast" + (isError ? " error" : ""), msg);
  document.body.appendChild(t);
  requestAnimationFrame(() => t.classList.add("show"));
  setTimeout(() => { t.classList.remove("show"); setTimeout(() => t.remove(), 300); }, 3200);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function sanitizeName(s) {
  return String(s).replace(/[^\w-]+/g, "_").slice(0, 40) || "document";
}

function recordLabel(rec) {
  const fields = Object.keys(rec);
  return fields.slice(0, 2).map((f) => rec[f]).join(" — ");
}

// ─────────────────────── generation pipeline ───────────────────────
async function exportCurrentTemplate() {
  return ctx.getEditor().currentDocument().exportDOCX();
}

async function generatePdfFromMerged(mergedBlob) {
  const doc = await ctx.getSystem().importDOCX(mergedBlob);
  const pdfBuffer = await doc.exportPDF();
  doc.destroy?.();
  return new Blob([pdfBuffer], { type: "application/pdf" });
}

async function generateOne(record, format) {
  const templateBuffer = await exportCurrentTemplate();
  const merged = await mergeDocx(templateBuffer, record);
  if (format === "docx") return merged;
  return generatePdfFromMerged(merged);
}

// ───────────────────────── dataset handling ────────────────────────
function setDataset(dataset) {
  currentDataset = dataset;
  const chip = $("#dataset-chip");
  if (chip) chip.textContent = `${dataset.name} · ${dataset.records.length} records`;
  renderFieldsPanel();
  renderDataPanel();
  renderGenerateRecordSelect();
}

export function activateTemplate(tpl) {
  const ds = DATASETS[tpl.datasetId];
  if (ds) setDataset(ds);
}

// ─────────────────────────── slide panels ──────────────────────────
function makeSlidePanel(id, titleText) {
  const overlay = el("div", "panel-overlay");
  overlay.id = `${id}-overlay`;
  const panel = el("div", "slide-panel");
  panel.id = `${id}-panel`;
  panel.innerHTML = `
    <div class="panel-header">
      <h2>${titleText}</h2>
      <button class="panel-close" aria-label="Close">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
    </div>
    <div class="panel-body"></div>`;
  const close = () => { panel.classList.remove("open"); overlay.classList.remove("open"); };
  overlay.addEventListener("click", close);
  panel.querySelector(".panel-close").addEventListener("click", close);
  document.body.appendChild(overlay);
  document.body.appendChild(panel);
  return { panel, overlay, close };
}

function togglePanel(id) {
  const p = $(`#${id}-panel`), o = $(`#${id}-overlay`);
  const isOpen = p.classList.contains("open");
  document.querySelectorAll(".slide-panel.open").forEach((x) => x.classList.remove("open"));
  document.querySelectorAll(".panel-overlay.open").forEach((x) => x.classList.remove("open"));
  if (!isOpen) { p.classList.add("open"); o.classList.add("open"); }
}

// Fields panel
function renderFieldsPanel() {
  const body = $("#fields-panel .panel-body");
  if (!body) return;
  const fields = datasetFields(currentDataset);
  body.innerHTML = `<p class="panel-hint">Click a field to insert it at the cursor position. Placeholders are resolved at generation time from the connected data source.</p>`;
  const list = el("div", "field-list");
  fields.forEach((f) => {
    const item = el("div", "field-item", `<span class="field-tag">{{${f}}}</span><span class="field-sample">${currentDataset.records[0]?.[f] ?? ""}</span>`);
    item.addEventListener("click", () => {
      const editor = ctx.getEditor();
      if (editor.hasActiveCursor()) {
        editor.insertTextAtCursor(`{{${f}}}`);
        toast(`Inserted {{${f}}}`);
      } else {
        toast("Click inside the document first to place the cursor.", true);
      }
    });
    list.appendChild(item);
  });
  body.appendChild(list);
}

// Data panel
function renderDataPanel() {
  const body = $("#data-panel .panel-body");
  if (!body) return;
  body.innerHTML = `<p class="panel-hint">Data connectors live in the host application — this panel simulates the application handing records to the document layer. Pick a sample source or upload a CSV.</p>`;

  const cards = el("div", "data-cards");
  Object.values(DATASETS).forEach((ds) => {
    const card = el("div", "data-card" + (ds.id === currentDataset.id ? " active" : ""),
      `<div class="data-card-name">${ds.name}</div><div class="data-card-desc">${ds.description}</div>`);
    card.addEventListener("click", () => setDataset(ds));
    cards.appendChild(card);
  });

  // CSV upload card
  const csvCard = el("div", "data-card upload", `<div class="data-card-name">Upload CSV…</div><div class="data-card-desc">First row = field names. Becomes the active data source.</div>`);
  const fileInput = el("input");
  fileInput.type = "file";
  fileInput.accept = ".csv,text/csv";
  fileInput.style.display = "none";
  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const parsed = parseCSV(reader.result);
      if (!parsed || !parsed.records.length) { toast("Could not parse CSV (need a header row + data rows).", true); return; }
      setDataset({ id: "csv", name: file.name, description: `${parsed.records.length} uploaded records`, records: parsed.records });
      toast(`Loaded ${parsed.records.length} records from ${file.name}`);
    };
    reader.readAsText(file);
  });
  csvCard.addEventListener("click", () => fileInput.click());
  cards.appendChild(csvCard);
  body.appendChild(cards);
  body.appendChild(fileInput);

  // Records preview
  const fields = datasetFields(currentDataset);
  const previewWrap = el("div", "records-preview");
  previewWrap.innerHTML = `<div class="records-title">Active source: <b>${currentDataset.name}</b> (${currentDataset.records.length} records)</div>`;
  const scroller = el("div", "records-scroll");
  const table = el("table", "records-table");
  table.innerHTML =
    `<thead><tr>${fields.map((f) => `<th>${f}</th>`).join("")}</tr></thead>` +
    `<tbody>${currentDataset.records.slice(0, 6).map((r) =>
      `<tr>${fields.map((f) => `<td>${r[f] ?? ""}</td>`).join("")}</tr>`).join("")}</tbody>`;
  scroller.appendChild(table);
  previewWrap.appendChild(scroller);
  if (currentDataset.records.length > 6) {
    previewWrap.appendChild(el("div", "records-more", `… and ${currentDataset.records.length - 6} more records`));
  }
  body.appendChild(previewWrap);
}

// ─────────────────────────── generate modal ────────────────────────
let previewUrl = null;

function renderGenerateRecordSelect() {
  const select = $("#gen-record-select");
  if (!select) return;
  select.innerHTML = currentDataset.records
    .map((r, i) => `<option value="${i}">${recordLabel(r)}</option>`)
    .join("");
  renderGenerateRecordDetail();
  const batchBtn = $("#gen-batch-btn");
  if (batchBtn) batchBtn.textContent = `Generate all ${currentDataset.records.length} documents`;
}

function renderGenerateRecordDetail() {
  const detail = $("#gen-record-detail");
  const select = $("#gen-record-select");
  if (!detail || !select) return;
  const rec = currentDataset.records[parseInt(select.value || "0", 10)] || {};
  detail.innerHTML = Object.entries(rec)
    .map(([k, v]) => `<div class="gen-kv"><span class="gen-k">${k}</span><span class="gen-v">${v}</span></div>`)
    .join("");
}

function setPreviewBlob(blob) {
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = URL.createObjectURL(blob);
  const frame = $("#gen-preview-frame");
  frame.src = previewUrl;
  $("#gen-preview-empty").style.display = "none";
  frame.style.display = "block";
  $("#gen-workflow").style.display = "none";
}

function selectedRecord() {
  return currentDataset.records[parseInt($("#gen-record-select").value || "0", 10)];
}

function setBusy(busy, label) {
  $("#generate-modal").classList.toggle("busy", busy);
  const status = $("#gen-status");
  status.textContent = busy ? label : "";
}

async function onPreviewPdf() {
  try {
    setBusy(true, "Rendering PDF…");
    const t0 = performance.now();
    const blob = await generateOne(selectedRecord(), "pdf");
    setPreviewBlob(blob);
    setBusy(false);
    toast(`PDF generated in ${((performance.now() - t0) / 1000).toFixed(1)}s`);
  } catch (err) {
    console.error(err); setBusy(false); toast("Generation failed — see console.", true);
  }
}

async function onDownload(format) {
  try {
    setBusy(true, format === "pdf" ? "Rendering PDF…" : "Merging DOCX…");
    const rec = selectedRecord();
    const blob = await generateOne(rec, format);
    downloadBlob(blob, `${sanitizeName(Object.values(rec)[0])}.${format}`);
    setBusy(false);
  } catch (err) {
    console.error(err); setBusy(false); toast("Generation failed — see console.", true);
  }
}

// Simulated transactional workflow: event → merge → render → email
async function onSimulateWorkflow() {
  const rec = selectedRecord();
  const steps = [
    { label: "Business event received from host application", detail: "e.g. approval clicked in a workflow of the host application" },
    { label: "Record merged into template", detail: recordLabel(rec) },
    { label: "PDF rendered by Document Authoring", detail: "pixel-perfect output" },
    { label: `Email dispatched to ${rec.Email ?? "recipient"}`, detail: "SIMULATED — email delivery stays in the host application" },
  ];
  const wf = $("#gen-workflow");
  wf.style.display = "block";
  $("#gen-preview-frame").style.display = "none";
  $("#gen-preview-empty").style.display = "none";
  wf.innerHTML = steps.map((s, i) =>
    `<div class="wf-step" id="wf-step-${i}"><div class="wf-dot"></div><div><div class="wf-label">${s.label}</div><div class="wf-detail">${s.detail}</div></div></div>`).join("");

  const mark = (i, cls) => $(`#wf-step-${i}`).classList.add(cls);
  let blob = null;
  try {
    mark(0, "active"); await new Promise((r) => setTimeout(r, 700)); mark(0, "done");
    mark(1, "active");
    const templateBuffer = await exportCurrentTemplate();
    const merged = await mergeDocx(templateBuffer, rec);
    mark(1, "done");
    mark(2, "active");
    blob = await generatePdfFromMerged(merged);
    mark(2, "done");
    mark(3, "active"); await new Promise((r) => setTimeout(r, 700)); mark(3, "done");
  } catch (err) {
    console.error(err); toast("Workflow simulation failed — see console.", true); return;
  }
  await new Promise((r) => setTimeout(r, 500));
  setPreviewBlob(blob);
  toast("Workflow complete — document generated and 'sent'.");
}

// Batch generation
async function onBatch() {
  const records = currentDataset.records;
  const bar = $("#gen-batch-bar"), fill = $("#gen-batch-fill"), stats = $("#gen-batch-stats");
  bar.style.display = "block";
  fill.style.width = "0%";
  stats.textContent = "Preparing template…";
  const btn = $("#gen-batch-btn");
  btn.disabled = true;

  try {
    const t0 = performance.now();
    const templateBuffer = await exportCurrentTemplate();
    const JSZip = (await import("jszip")).default;
    const outZip = new JSZip();
    const nameField = datasetFields(currentDataset)[0];

    for (let i = 0; i < records.length; i++) {
      const merged = await mergeDocx(templateBuffer, records[i]);
      const pdf = await generatePdfFromMerged(merged);
      outZip.file(`${pad3(i + 1)}_${sanitizeName(records[i][nameField])}.pdf`, pdf);
      const pct = Math.round(((i + 1) / records.length) * 100);
      fill.style.width = `${pct}%`;
      const elapsed = (performance.now() - t0) / 1000;
      stats.textContent = `${i + 1}/${records.length} documents — ${elapsed.toFixed(1)}s elapsed — avg ${(elapsed / (i + 1)).toFixed(2)}s/doc`;
    }

    stats.textContent += " — packaging ZIP…";
    const zipBlob = await outZip.generateAsync({ type: "blob" });
    const total = (performance.now() - t0) / 1000;
    stats.innerHTML = `<b>${records.length} PDFs in ${total.toFixed(1)}s</b> (avg ${(total / records.length).toFixed(2)}s/doc, in-browser WASM — server-side rendering scales linearly with workers)`;
    downloadBlob(zipBlob, `${sanitizeName(currentDataset.name)}_batch.zip`);
  } catch (err) {
    console.error(err); toast("Batch failed — see console.", true);
  } finally {
    btn.disabled = false;
  }
}
function pad3(n) { return String(n).padStart(3, "0"); }

function buildGenerateModal() {
  const overlay = el("div", "merge-overlay");
  overlay.id = "generate-modal-overlay";
  overlay.style.display = "none";
  overlay.innerHTML = `
  <div class="merge-modal" id="generate-modal" style="width:960px;max-width:95vw;max-height:90vh;display:flex;flex-direction:column;">
    <div class="merge-modal-header">
      <h3>Generate Documents</h3>
      <p>Merge the current template with the active data source: <b><span id="gen-dataset-name"></span></b></p>
    </div>
    <div class="gen-body">
      <div class="gen-left">
        <label class="gen-label">Record</label>
        <select id="gen-record-select" class="gen-select"></select>
        <div id="gen-record-detail" class="gen-detail"></div>
        <div class="gen-actions">
          <button class="btn btn-primary" id="gen-preview-btn">Preview PDF</button>
          <button class="btn btn-secondary" id="gen-dl-pdf-btn">Download PDF</button>
          <button class="btn btn-secondary" id="gen-dl-docx-btn">Download DOCX</button>
          <button class="btn btn-workflow" id="gen-workflow-btn">▶ Simulate workflow: generate &amp; email</button>
        </div>
        <div class="gen-batch">
          <div class="gen-batch-title">Batch generation</div>
          <button class="btn btn-secondary" id="gen-batch-btn">Generate all documents</button>
          <div id="gen-batch-bar" class="batch-bar" style="display:none;"><div id="gen-batch-fill" class="batch-fill"></div></div>
          <div id="gen-batch-stats" class="batch-stats"></div>
        </div>
        <div id="gen-status" class="gen-status"></div>
      </div>
      <div class="gen-right">
        <div id="gen-preview-empty" class="gen-preview-empty">Select a record and click <b>Preview PDF</b></div>
        <iframe id="gen-preview-frame" title="PDF preview" style="display:none;"></iframe>
        <div id="gen-workflow" class="gen-workflow" style="display:none;"></div>
      </div>
    </div>
    <div class="merge-modal-footer">
      <button class="btn btn-secondary" id="gen-close-btn">Close</button>
    </div>
  </div>`;
  document.body.appendChild(overlay);

  overlay.addEventListener("click", (e) => { if (e.target === overlay) closeGenerateModal(); });
  $("#gen-close-btn").addEventListener("click", closeGenerateModal);
  $("#gen-record-select").addEventListener("change", renderGenerateRecordDetail);
  $("#gen-preview-btn").addEventListener("click", onPreviewPdf);
  $("#gen-dl-pdf-btn").addEventListener("click", () => onDownload("pdf"));
  $("#gen-dl-docx-btn").addEventListener("click", () => onDownload("docx"));
  $("#gen-workflow-btn").addEventListener("click", onSimulateWorkflow);
  $("#gen-batch-btn").addEventListener("click", onBatch);
}

function openGenerateModal() {
  $("#gen-dataset-name").textContent = `${currentDataset.name} (${currentDataset.records.length} records)`;
  renderGenerateRecordSelect();
  $("#generate-modal-overlay").style.display = "flex";
}
function closeGenerateModal() {
  $("#generate-modal-overlay").style.display = "none";
}

// ─────────────────────────── injected styles ───────────────────────
const STYLES = `
.panel-body { flex:1; overflow-y:auto; padding: 12px 0; }
.panel-hint { font-size:12px; color:var(--app-gray-500); padding: 0 20px 12px; line-height:1.5; }
.field-list { display:flex; flex-direction:column; }
.field-item { padding:10px 20px; cursor:pointer; display:flex; flex-direction:column; gap:2px; transition:background .12s; }
.field-item:hover { background:var(--app-blue-light); }
.field-tag { font-family:ui-monospace,Menlo,monospace; font-size:13px; font-weight:600; color:var(--app-blue); }
.field-sample { font-size:11px; color:var(--app-gray-500); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.data-cards { display:flex; flex-direction:column; gap:8px; padding:0 20px 16px; }
.data-card { border:1px solid var(--app-gray-200); border-radius:8px; padding:12px 14px; cursor:pointer; transition:border-color .12s, background .12s; }
.data-card:hover { border-color:var(--app-blue); }
.data-card.active { border-color:var(--app-blue); background:var(--app-blue-light); }
.data-card.upload { border-style:dashed; }
.data-card-name { font-size:13px; font-weight:600; color:var(--app-gray-900); }
.data-card-desc { font-size:11px; color:var(--app-gray-500); margin-top:2px; }
.records-preview { padding:0 20px; }
.records-title { font-size:12px; color:var(--app-gray-700); margin-bottom:8px; }
.records-scroll { overflow-x:auto; border:1px solid var(--app-gray-200); border-radius:6px; }
.records-table { border-collapse:collapse; font-size:11px; width:100%; }
.records-table th { background:var(--app-gray-100); color:var(--app-navy); text-align:left; padding:6px 8px; white-space:nowrap; position:sticky; top:0; }
.records-table td { padding:5px 8px; border-top:1px solid var(--app-gray-100); white-space:nowrap; color:var(--app-gray-700); }
.records-more { font-size:11px; color:var(--app-gray-500); padding:6px 2px; }
.gen-body { display:flex; gap:0; flex:1; min-height:0; }
.gen-left { width:340px; flex-shrink:0; padding:16px 20px; border-right:1px solid var(--app-gray-200); overflow-y:auto; display:flex; flex-direction:column; gap:10px; }
.gen-right { flex:1; min-width:0; display:flex; flex-direction:column; background:var(--app-gray-100); }
.gen-label { font-size:12px; font-weight:600; color:var(--app-gray-700); }
.gen-select { width:100%; padding:8px 10px; border:1px solid var(--app-gray-300); border-radius:6px; font-family:var(--app-font); font-size:13px; }
.gen-detail { border:1px solid var(--app-gray-200); border-radius:6px; max-height:180px; overflow-y:auto; font-size:11px; }
.gen-kv { display:flex; border-bottom:1px solid var(--app-gray-100); }
.gen-kv:last-child { border-bottom:none; }
.gen-k { width:45%; padding:4px 8px; font-weight:600; color:var(--app-navy); background:var(--app-gray-50); flex-shrink:0; }
.gen-v { padding:4px 8px; color:var(--app-gray-700); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.gen-actions { display:flex; flex-direction:column; gap:8px; }
.gen-actions .btn { width:100%; }
.btn-workflow { background:var(--app-navy); color:#fff; border:none; }
.btn-workflow:hover { background:#374151; }
.gen-batch { border-top:1px solid var(--app-gray-200); padding-top:12px; display:flex; flex-direction:column; gap:8px; }
.gen-batch-title { font-size:12px; font-weight:600; color:var(--app-gray-700); }
.batch-bar { height:8px; background:var(--app-gray-200); border-radius:4px; overflow:hidden; }
.batch-fill { height:100%; width:0; background:var(--app-blue); transition:width .15s; }
.batch-stats { font-size:11px; color:var(--app-gray-700); line-height:1.4; }
.gen-status { font-size:12px; color:var(--app-blue); font-weight:600; min-height:16px; }
#generate-modal.busy { pointer-events:none; opacity:.75; }
.gen-preview-empty { flex:1; display:flex; align-items:center; justify-content:center; color:var(--app-gray-500); font-size:13px; }
#gen-preview-frame { flex:1; width:100%; border:none; min-height:420px; }
.gen-workflow { padding:32px; display:flex; flex-direction:column; gap:0; }
.wf-step { display:flex; gap:14px; padding:0 0 28px; position:relative; opacity:.35; transition:opacity .3s; }
.wf-step:not(:last-child):before { content:""; position:absolute; left:9px; top:22px; bottom:2px; width:2px; background:var(--app-gray-300); }
.wf-step.active, .wf-step.done { opacity:1; }
.wf-dot { width:20px; height:20px; border-radius:50%; background:#fff; border:2px solid var(--app-gray-300); flex-shrink:0; box-sizing:border-box; transition:all .3s; }
.wf-step.active .wf-dot { border-color:var(--app-blue); box-shadow:0 0 0 4px rgba(37,99,235,.2); }
.wf-step.done .wf-dot { background:var(--app-blue); border-color:var(--app-blue); }
.wf-label { font-size:13px; font-weight:600; color:var(--app-gray-900); }
.wf-detail { font-size:11px; color:var(--app-gray-500); margin-top:2px; }
.poc-toast { position:fixed; bottom:24px; left:50%; transform:translate(-50%, 12px); background:var(--app-navy); color:#fff; padding:10px 20px; border-radius:8px; font-size:13px; font-family:var(--app-font); z-index:10005; opacity:0; transition:all .25s; box-shadow:0 6px 20px rgba(0,0,0,.25); }
.poc-toast.show { opacity:1; transform:translate(-50%, 0); }
.poc-toast.error { background:#C92A2A; }
#dataset-chip { margin-left:auto; font-size:12px; color:rgba(255,255,255,.85); background:rgba(255,255,255,.12); padding:4px 12px; border-radius:999px; }
`;

// ─────────────────────────── toolbar icons ─────────────────────────
const svg64 = (s) => `data:image/svg+xml;base64,${btoa(s)}`;
export const FIELD_ICON = svg64(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7V5h16v2"/><path d="M12 5v14"/><path d="M9 19h6"/></svg>`);
export const DATA_ICON = svg64(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/></svg>`);
export const GENERATE_ICON = svg64(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="6 3 20 12 6 21 6 3"/></svg>`);

// ───────────────────────────── public API ──────────────────────────
export { TEMPLATES, buildTemplate, DATASETS };

export function generationActions() {
  return [
    { id: "custom.fields", label: "Insert Field", icon: FIELD_ICON, handler: () => togglePanel("fields") },
    { id: "custom.data", label: "Data", icon: DATA_ICON, handler: () => togglePanel("data") },
    { id: "custom.generate", label: "Generate", icon: GENERATE_ICON, handler: () => openGenerateModal() },
  ];
}

export function generationToolbarItems() {
  return [
    { type: "action", id: "fields-btn", actionId: "custom.fields" },
    { type: "action", id: "data-btn", actionId: "custom.data" },
    { type: "action", id: "generate-btn", actionId: "custom.generate" },
  ];
}

export function initGenerationUI(context) {
  ctx = context;
  window.pocToast = toast;
  const style = document.createElement("style");
  style.textContent = STYLES;
  document.head.appendChild(style);

  // Dataset chip in brand bar
  const chip = el("div");
  chip.id = "dataset-chip";
  document.getElementById("brand-bar").appendChild(chip);

  makeSlidePanel("fields", "Merge Fields");
  makeSlidePanel("data", "Data Source");
  buildGenerateModal();
  setDataset(currentDataset);
}
