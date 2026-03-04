/* global Office, Word */

// ============================================
// BiroA ADMIN – V1.0
// Word taskpane za upravljanje templejt poljima
// ============================================

// ==========================================
// KONFIGURACIJA (defaults – prepisuju se iz XML-a)
// ==========================================
const DEFAULT_CONFIG = {
  baseUrl: "https://raw.githubusercontent.com/baneandreev-byte/BiroA-templates-test/main",
  apiUrl:  "https://api.github.com/repos/baneandreev-byte/BiroA-templates-test/contents",
  token: "",
  branches: [
    { id: "01 IDR", label: "IDR – Idejno rešenje",  files: ["00 IDR Glavna sveska.dotx", "01 IDR Sveska projekta.dotx"] },
    { id: "02 PGD", label: "PGD – Projekat za građevinsku dozvolu", files: ["00 PGD Glavna sveska.dotx", "01 PGD Sveska projekta.dotx"] },
    { id: "03 PZI", label: "PZI – Projekat za izvođenje", files: ["00 PZI Glavna sveska.dotx", "01 PZI Sveska projekta.dotx"] },
    { id: "04 TK",  label: "TK – Tehnička kontrola",  files: ["00 TK Glavna sveska.dotx",  "01 TK Sveska projekta.docx"] },
  ]
};

const ADMIN_XML_NS = "http://biroa.rs/word-addin/admin-config";

const FORMAT_OPTIONS = {
  text:   [
    { value: "text:auto",  label: "Automatski",   hint: "" },
    { value: "text:upper", label: "VELIKA SLOVA",  hint: "Primer: BEOGRAD" },
    { value: "text:lower", label: "mala slova",    hint: "Primer: beograd" },
    { value: "text:title", label: "Naslov",        hint: "Primer: Beograd" },
  ],
  date:   [
    { value: "date:auto",        label: "Kako je uneto",    hint: "" },
    { value: "date:today",       label: "Danas (dd.mm.yyyy)", hint: "Primer: 07.02.2025" },
    { value: "date:dd.mm.yyyy",  label: "dd.mm.yyyy",       hint: "Primer: 07.02.2025" },
    { value: "date:yyyy-mm-dd",  label: "yyyy-mm-dd",       hint: "Primer: 2025-02-07" },
    { value: "date:mmmm.yyyy",   label: "MMMM.yyyy",        hint: "Primer: februar.2025" },
    { value: "date:dd.mmmm.yyyy",label: "dd.MMMM.yyyy",     hint: "Primer: 07.februar.2025" },
  ],
  number: [
    { value: "number:auto", label: "Automatski",  hint: "" },
    { value: "number:int",  label: "Ceo broj",    hint: "Primer: 1.234" },
    { value: "number:2",    label: "2 decimale",  hint: "Primer: 1.234,56" },
    { value: "number:rsd",  label: "RSD",         hint: "Primer: 1.234,56 RSD" },
    { value: "number:eur",  label: "€",           hint: "Primer: 1.234,56 €" },
    { value: "number:usd",  label: "$",           hint: "Primer: 1.234,56 $" },
  ],
};

// ==========================================
// STATE
// ==========================================
let rows = [];               // [{ id, field, type, format, value, description }]
let config = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
let selectedRowIndex = null;

// Drag state
let draggedElement = null;
let draggedId = null;

// Branch modal state
let editingBranchIndex = null;

// Rename modal state
let renamingRowIndex = null;

// Description modal state
let descriptionRowIndex = null;

// ==========================================
// DOM helper
// ==========================================
function el(id) { return document.getElementById(id); }

function setStatus(msg, kind = "info") {
  const s = el("status");
  if (!s) return;
  s.textContent = msg;
  s.className = `status ${kind}`;
  s.classList.remove("hidden");
  clearTimeout(setStatus._t);
  setStatus._t = setTimeout(() => s.classList.add("hidden"), 4000);
}

// ==========================================
// TAG helpers (isti format kao user app)
// ==========================================
function makeTag(key, type, format) {
  return `BA_FIELD|key=${key.trim()}|type=${(type||"text").trim()}|format=${(format||"text:auto").trim()}`;
}

function parseTag(tag) {
  const s = String(tag || "");
  if (!s.startsWith("BA_FIELD|")) return null;
  const out = {};
  for (const p of s.split("|").slice(1)) {
    const [k, ...rest] = p.split("=");
    out[k] = rest.join("=");
  }
  return out.key ? { key: out.key, type: out.type || "text", format: out.format || "text:auto" } : null;
}

function token(key) { return `{${key}}`; }

function xmlEscape(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function escHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ==========================================
// TABS
// ==========================================
function switchTab(tabId) {
  ["btnTabFields","btnTabConfig"].forEach(id => {
    el(id)?.classList.toggle("active", id === tabId);
  });
  el("panelFields")?.classList.toggle("hidden", tabId !== "btnTabFields");
  el("panelConfig")?.classList.toggle("hidden", tabId !== "btnTabConfig");

  if (tabId === "btnTabConfig") {
    loadConfigToForm();
    renderBranches();
  }
}

// ==========================================
// RENDER ROWS
// ==========================================
function renderRows() {
  const container = el("rows");
  if (!container) return;

  if (rows.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📄</div>
        <div>Otvori templejt pa klikni <strong>Skeniraj</strong><br>ili dodaj nova polja ručno.</div>
      </div>`;
    return;
  }

  // Save focus
  const active = document.activeElement;
  const focusedRowEl = active?.closest?.(".row");
  const wasInput = active?.tagName === "INPUT";
  const focusedRowId = focusedRowEl?.dataset?.id;

  container.innerHTML = "";

  rows.forEach((r, idx) => {
    if (!r.id) r.id = crypto.randomUUID();

    const row = document.createElement("div");
    row.className = "row";
    if (idx === selectedRowIndex) row.classList.add("selected");
    row.dataset.id = r.id;
    row.draggable = false;

    row.addEventListener("dragover", handleDragOver);
    row.addEventListener("dragleave", handleDragLeave);
    row.addEventListener("drop", handleDrop);
    row.addEventListener("click", (e) => {
      if (e.target.closest(".drag-handle,.row-btn,.type-badge")) return;
      selectedRowIndex = idx;
      renderRows();
    });

    // Drag handle
    const handle = document.createElement("div");
    handle.className = "drag-handle";
    handle.innerHTML = "⋮⋮";
    handle.title = "Prevuci za premeštanje";
    handle.draggable = true;
    handle.dataset.id = r.id;
    handle.addEventListener("dragstart", handleDragStart);
    handle.addEventListener("dragend",   handleDragEnd);

    // Field input
    const fieldInput = document.createElement("input");
    fieldInput.className = "field-input";
    fieldInput.type = "text";
    fieldInput.placeholder = "Naziv polja";
    fieldInput.value = r.field || "";
    fieldInput.addEventListener("input", (e) => {
      r.field = e.target.value;
    });
    fieldInput.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
    });
    fieldInput.addEventListener("focus", () => {
      selectedRowIndex = idx;
    });

    // Value input (test)
    const valueInput = document.createElement("input");
    valueInput.className = "field-input";
    valueInput.type = "text";
    valueInput.placeholder = "Test vrednost...";
    valueInput.value = r.value || "";
    valueInput.style.color = "#6b7280";
    valueInput.addEventListener("input", (e) => {
      r.value = e.target.value;
    });
    valueInput.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
    });
    valueInput.addEventListener("focus", () => {
      selectedRowIndex = idx;
    });

    // Type badge (klik → otvori modal za izmenu tipa/formata)
    const typeBadge = document.createElement("button");
    typeBadge.className = `type-badge ${r.type || "text"}`;
    const typeLabels = { text: "Tekst", date: "Datum", number: "Broj" };
    typeBadge.textContent = typeLabels[r.type] || "Tekst";
    typeBadge.title = `Tip: ${r.type} | Format: ${r.format}\nKlikni za izmenu`;
    typeBadge.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
      openTypeModal(r);
    });

    // Row actions
    const actions = document.createElement("div");
    actions.className = "row-actions";

    const btnDesc = document.createElement("button");
    btnDesc.className = "row-btn describe";
    btnDesc.innerHTML = "💬";
    btnDesc.title = r.description ? `Objašnjenje: ${r.description}` : "Dodaj objašnjenje za polje";
    if (r.description) btnDesc.classList.add("has-desc");
    btnDesc.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
      openDescriptionModal(idx);
    });

    const btnRename = document.createElement("button");
    btnRename.className = "row-btn rename";
    btnRename.innerHTML = "✎";
    btnRename.title = "Preimenuj polje u dokumentu";
    btnRename.addEventListener("click", (e) => {
      e.stopPropagation();
      selectedRowIndex = idx;
      openRenameModal(idx);
    });

    const btnDel = document.createElement("button");
    btnDel.className = "row-btn delete";
    btnDel.innerHTML = "×";
    btnDel.title = "Obriši red";
    btnDel.addEventListener("click", (e) => {
      e.stopPropagation();
      rows.splice(idx, 1);
      if (selectedRowIndex === idx) selectedRowIndex = null;
      else if (selectedRowIndex > idx) selectedRowIndex--;
      renderRows();
    });

    actions.appendChild(btnDesc);
    actions.appendChild(btnRename);
    actions.appendChild(btnDel);

    row.appendChild(handle);
    row.appendChild(fieldInput);
    row.appendChild(valueInput);
    row.appendChild(typeBadge);
    row.appendChild(actions);
    container.appendChild(row);
  });

  // Restore focus
  if (wasInput && focusedRowId) {
    const targetRow = container.querySelector(`[data-id="${focusedRowId}"]`);
    const inp = targetRow?.querySelector("input");
    if (inp) setTimeout(() => inp.focus(), 0);
  }
}

// ==========================================
// DRAG & DROP
// ==========================================
function handleDragStart(e) {
  draggedElement = e.currentTarget.closest(".row");
  draggedId = e.currentTarget.dataset.id;
  draggedElement.classList.add("dragging");
  e.dataTransfer.effectAllowed = "move";
  e.dataTransfer.setData("text/plain", draggedId);
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
  const target = e.target.closest(".row");
  if (target && target !== draggedElement) {
    document.querySelectorAll(".row").forEach(r => { if (r !== target) r.classList.remove("drag-over"); });
    target.classList.add("drag-over");
  }
  return false;
}

function handleDragLeave(e) {
  e.target.closest(".row")?.classList.remove("drag-over");
}

function handleDrop(e) {
  e.stopPropagation();
  const target = e.target.closest(".row");
  if (!target || target === draggedElement) return false;

  const fromIdx = rows.findIndex(r => r.id === draggedId);
  const toIdx   = rows.findIndex(r => r.id === target.dataset.id);
  if (fromIdx === -1 || toIdx === -1) return false;

  const [moved] = rows.splice(fromIdx, 1);
  rows.splice(toIdx, 0, moved);

  if (selectedRowIndex === fromIdx) selectedRowIndex = toIdx;
  else if (fromIdx < selectedRowIndex && toIdx >= selectedRowIndex) selectedRowIndex--;
  else if (fromIdx > selectedRowIndex && toIdx <= selectedRowIndex) selectedRowIndex++;

  renderRows();
  return false;
}

function handleDragEnd() {
  draggedElement?.classList.remove("dragging");
  document.querySelectorAll(".row").forEach(r => r.classList.remove("drag-over"));
  draggedElement = null;
  draggedId = null;
}

// ==========================================
// TYPE/FORMAT MODAL (klik na badge)
// ==========================================
function openTypeModal(row) {
  const modal = el("modal");
  const backdrop = el("modalBackdrop");
  if (!modal || !backdrop) return;

  el("modalFieldName").textContent = row.field || "(bez naziva)";

  // Set radio
  document.querySelectorAll('input[name="ftype"]').forEach(r => {
    r.checked = r.value === (row.type || "text");
  });

  updateFormatOptions(row.type || "text", row.format || "text:auto");

  // Listen type change
  document.querySelectorAll('input[name="ftype"]').forEach(r => {
    r.onchange = () => updateFormatOptions(r.value, null);
  });

  // Listen format hint
  const fs = el("formatSelect");
  if (fs) {
    fs.onchange = () => {
      const opt = fs.options[fs.selectedIndex];
      el("formatHint").textContent = opt?.getAttribute("data-hint") || "";
    };
  }

  modal.classList.remove("hidden");
  backdrop.classList.remove("hidden");
}

function closeTypeModal() {
  el("modal")?.classList.add("hidden");
  el("modalBackdrop")?.classList.add("hidden");
}

function saveTypeModal() {
  if (selectedRowIndex === null || selectedRowIndex >= rows.length) return;
  const r = rows[selectedRowIndex];

  const checked = document.querySelector('input[name="ftype"]:checked');
  if (checked) r.type = checked.value;

  const fs = el("formatSelect");
  if (fs) r.format = fs.value;

  closeTypeModal();
  renderRows();
  setStatus(`Ažurirano: ${r.field} (${r.type})`, "info");
}

function updateFormatOptions(type, currentFormat) {
  const fs = el("formatSelect");
  const hint = el("formatHint");
  if (!fs) return;

  fs.innerHTML = "";
  const opts = FORMAT_OPTIONS[type] || FORMAT_OPTIONS.text;
  opts.forEach(opt => {
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.label;
    o.setAttribute("data-hint", opt.hint);
    if (currentFormat && opt.value === currentFormat) o.selected = true;
    fs.appendChild(o);
  });

  const sel = fs.options[fs.selectedIndex];
  if (hint) hint.textContent = sel?.getAttribute("data-hint") || "";
}

// ==========================================
// RENAME MODAL
// ==========================================
function openRenameModal(idx) {
  renamingRowIndex = idx;
  const r = rows[idx];
  el("renameOldName").textContent = r.field || "(prazno)";
  const inp = el("renameNewName");
  if (inp) inp.value = r.field || "";

  el("modalRename")?.classList.remove("hidden");
  el("modalRenameBackdrop")?.classList.remove("hidden");
  setTimeout(() => inp?.focus(), 50);
}

function closeRenameModal() {
  el("modalRename")?.classList.add("hidden");
  el("modalRenameBackdrop")?.classList.add("hidden");
  renamingRowIndex = null;
}

// ==========================================
// DESCRIPTION MODAL
// ==========================================
function openDescriptionModal(idx) {
  descriptionRowIndex = idx;
  const r = rows[idx];
  el("descFieldName").textContent = r.field || "(prazno)";
  const ta = el("descText");
  if (ta) ta.value = r.description || "";

  el("modalDesc")?.classList.remove("hidden");
  el("modalDescBackdrop")?.classList.remove("hidden");
  setTimeout(() => ta?.focus(), 50);
}

function closeDescriptionModal() {
  el("modalDesc")?.classList.add("hidden");
  el("modalDescBackdrop")?.classList.add("hidden");
  descriptionRowIndex = null;
}

function saveDescriptionModal() {
  if (descriptionRowIndex === null) return;
  const r = rows[descriptionRowIndex];
  r.description = (el("descText")?.value || "").trim();
  closeDescriptionModal();
  renderRows();
  setStatus(`Objašnjenje ${r.description ? "sačuvano" : "obrisano"}: ${r.field}`, "success");
}

async function doRename() {
  if (renamingRowIndex === null) return;
  const r = rows[renamingRowIndex];
  const oldKey = r.field;
  const newName = el("renameNewName")?.value.trim();

  if (!newName) { setStatus("Unesi novo ime.", "warn"); return; }
  if (newName === oldKey) { closeRenameModal(); return; }

  // Update in-memory
  r.field = newName;

  // Update content controls in document
  let updated = 0;
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag,title");
      await context.sync();

      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta || meta.key !== oldKey) continue;
        const newTag = makeTag(newName, meta.type, meta.format);
        cc.tag = newTag;
        cc.title = newName;
        cc.insertText(token(newName), Word.InsertLocation.replace);
        updated++;
      }

      await context.sync();
    });

    setStatus(`Preimenovano "${oldKey}" → "${newName}" (${updated} polje${updated !== 1 ? "a" : ""} ažurirano).`, "success");
  } catch (err) {
    console.error("Rename greška:", err);
    setStatus("Greška pri preimenovanju u dokumentu.", "error");
  }

  closeRenameModal();
  renderRows();
}

// ==========================================
// INSERT FIELD (ubaci u dokument)
// ==========================================
async function insertFieldAtSelection() {
  if (selectedRowIndex === null) {
    setStatus("Selektuj red u tabeli prvo.", "warn");
    return;
  }

  const r = rows[selectedRowIndex];
  const key = (r.field || "").trim();
  if (!key) {
    setStatus("Unesi naziv polja.", "warn");
    return;
  }

  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const cc = sel.insertContentControl();
      cc.tag = makeTag(key, r.type, r.format);
      cc.title = key;
      cc.appearance = "BoundingBox";
      try { cc.cannotDelete = true; cc.cannotEdit = false; } catch {}
      cc.insertText(token(key), Word.InsertLocation.replace);
      await context.sync();
    });

    setStatus(`Polje "${key}" ubačeno u dokument.`, "success");
  } catch (err) {
    console.error("Insert greška:", err);
    setStatus("Greška pri ubacivanju polja.", "error");
  }
}

// ==========================================
// SCAN DOCUMENT – učitaj polja iz dokumenta
// ==========================================
async function scanDocument() {
  setStatus("Skeniram dokument...", "info");

  // Namespace koji klijentska aplikacija koristi za čuvanje vrednosti
  const CLIENT_XML_NS = "http://biroa.rs/word-addin/state";

  try {
    const found = [];
    const seenKeys = new Set();

    await Word.run(async (context) => {
      // 1) Učitaj content controls (za tip/format)
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta || seenKeys.has(meta.key)) continue;
        seenKeys.add(meta.key);
        found.push({
          id: crypto.randomUUID(),
          field: meta.key,
          type: meta.type,
          format: meta.format,
          value: "",
          description: "",
        });
      }

      // 2) Učitaj sačuvane vrednosti iz klijentskog XML state-a
      const parts = context.document.customXmlParts;
      parts.load("items");
      await context.sync();
      for (const p of parts.items) p.load("namespaceUri");
      await context.sync();

      const clientPart = parts.items.find(p => p.namespaceUri === CLIENT_XML_NS);
      if (clientPart) {
        const xmlResult = clientPart.getXml();
        await context.sync();
        const str = xmlResult.value || "";
        // Parsiraj <item field="..." value="..."/> atribute
        const re = /<item\s+([^/>]+?)\s*\/>/g;
        let m;
        const savedValues = new Map();
        while ((m = re.exec(str))) {
          const attrs = m[1];
          const getAttr = (name) => {
            const rm = new RegExp(`${name}="([^"]*)"`);
            const mm = rm.exec(attrs);
            if (!mm) return "";
            return mm[1]
              .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
              .replace(/&gt;/g, ">").replace(/&lt;/g, "<").replace(/&amp;/g, "&");
          };
          const f = getAttr("field");
          const v = getAttr("value");
          if (f) savedValues.set(f, v);
        }
        // Ubaci sačuvane vrednosti u pronađena polja
        for (const row of found) {
          if (savedValues.has(row.field)) {
            row.value = savedValues.get(row.field);
          }
        }
      }
    });

    if (found.length === 0) {
      setStatus("Nije pronađeno nijedno BA_FIELD polje u dokumentu.", "warn");
      return;
    }

    rows = found;
    selectedRowIndex = null;
    renderRows();
    setStatus(`Skeniranje gotovo: ${found.length} polje${found.length !== 1 ? "a" : ""} pronađeno.`, "success");
  } catch (err) {
    console.error("Scan greška:", err);
    setStatus("Greška pri skeniranju dokumenta.", "error");
  }
}

// ==========================================
// DELETE CONTENT CONTROL (ukloni wrapper, zadrži tekst)
// Koristi se kad se klikne × na redu koji IMA CC u dokumentu
// ==========================================
async function removeFieldFromDocument(key) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag,cannotDelete");
      await context.sync();

      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta || meta.key !== key) continue;
        try { cc.cannotDelete = false; } catch {}
        cc.delete(true); // keepContent = true
      }

      await context.sync();
    });
  } catch (err) {
    console.error("Remove CC greška:", err);
  }
}

// ==========================================
// GITHUB: SNIMI TEMPLEJT
// Tok: izaberi granu → izaberi fajl → unesi commit poruku → commit
// ==========================================

function openGitHubSaveModal() {
  el("modalGHTitle").textContent = "Snimi templejt na GitHub";
  el("modalGHBody").innerHTML = "";
  el("modalGHFooter").innerHTML = `<button id="btnModalGHCancel" class="btn-secondary">Otkaži</button>`;
  el("btnModalGHCancel")?.addEventListener("click", closeGHModal);

  el("modalGH")?.classList.remove("hidden");
  el("modalGHBackdrop")?.classList.remove("hidden");

  renderGHStep1();
}

function openGitHubOpenModal() {
  el("modalGHTitle").textContent = "Otvori templejt sa GitHub-a";
  el("modalGHBody").innerHTML = "";
  el("modalGHFooter").innerHTML = `<button id="btnModalGHCancel" class="btn-secondary">Otkaži</button>`;
  el("btnModalGHCancel")?.addEventListener("click", closeGHModal);

  el("modalGH")?.classList.remove("hidden");
  el("modalGHBackdrop")?.classList.remove("hidden");

  renderGHOpenStep1();
}

function closeGHModal() {
  el("modalGH")?.classList.add("hidden");
  el("modalGHBackdrop")?.classList.add("hidden");
}

// === SNIMI: Korak 1 – izbor grane ===
function renderGHStep1() {
  const body = el("modalGHBody");
  body.innerHTML = `<p class="gh-step-title">Izaberi vrstu projekta:</p>`;

  if (config.branches.length === 0) {
    body.innerHTML += `<div class="gh-error">Nema grana u konfiguraciji. Dodaj grane u tabu KONFIGURACIJA.</div>`;
    return;
  }

  config.branches.forEach(branch => {
    const btn = document.createElement("button");
    btn.className = "gh-btn";
    btn.innerHTML = `<div class="gh-btn-id">${escHtml(branch.id)}</div><div class="gh-btn-label">${escHtml(branch.label)}</div>`;
    btn.addEventListener("click", () => renderGHStep2(branch));
    body.appendChild(btn);
  });
}

// === SNIMI: Korak 2 – izbor fajla ===
function renderGHStep2(branch) {
  const body = el("modalGHBody");
  body.innerHTML = `
    <span class="gh-back" id="ghBack1">← Nazad</span>
    <p class="gh-step-title">${escHtml(branch.id)} – izaberi svesku:</p>
  `;
  el("ghBack1")?.addEventListener("click", renderGHStep1);

  if (branch.files.length === 0) {
    body.innerHTML += `<div class="gh-error">Nema fajlova u ovoj grani.</div>`;
    return;
  }

  branch.files.forEach(fileName => {
    const btn = document.createElement("button");
    btn.className = "gh-btn";
    btn.innerHTML = `📄 ${escHtml(fileName)}`;
    btn.addEventListener("click", () => renderGHStep3(branch, fileName));
    body.appendChild(btn);
  });
}

// === SNIMI: Korak 3 – commit poruka ===
function renderGHStep3(branch, fileName) {
  const body = el("modalGHBody");
  body.innerHTML = `
    <span class="gh-back" id="ghBack2">← Nazad</span>
    <p class="gh-step-title">Snimi: <strong>${escHtml(branch.id)}/${escHtml(fileName)}</strong></p>
    <div class="gh-commit-form">
      <div class="form-group">
        <label class="form-label">Commit poruka:</label>
        <input type="text" id="ghCommitMsg" class="form-input"
          placeholder="npr. Dodato polje INVESTITOR"
          value="Ažuriran templejt: ${escHtml(fileName)}" />
      </div>
      <div class="form-hint">
        Dokument mora biti sačuvan (.dotx) pre nego što klikneš Snimi.
      </div>
    </div>
  `;
  el("ghBack2")?.addEventListener("click", () => renderGHStep2(branch));

  // Update footer
  const footer = el("modalGHFooter");
  footer.innerHTML = `
    <button id="btnGHCancelFinal" class="btn-secondary">Otkaži</button>
    <button id="btnGHCommit" class="btn-primary">↑ Commit na GitHub</button>
  `;
  el("btnGHCancelFinal")?.addEventListener("click", closeGHModal);
  el("btnGHCommit")?.addEventListener("click", () => doGitHubCommit(branch, fileName));
}

// === SNIMI: Izvrši commit ===
async function doGitHubCommit(branch, fileName) {
  const commitMsg = el("ghCommitMsg")?.value.trim() || `Ažuriran templejt: ${fileName}`;

  if (!config.token) {
    setStatus("GitHub token nije podešen. Podesi u tab KONFIGURACIJA.", "error");
    closeGHModal();
    return;
  }

  const body = el("modalGHBody");
  body.innerHTML = `
    <div class="gh-loading">
      <div class="gh-loading-icon">⏳</div>
      <div>Priprema dokumenta...</div>
    </div>`;
  el("modalGHFooter").innerHTML = "";

  try {
    // 1. Uzmi sadržaj trenutnog dokumenta kao Base64 .docx
    let base64Content = null;

    await Word.run(async (context) => {
      const doc = context.document;
      const body2 = doc.body;
      body2.load("text"); // just to trigger sync
      await context.sync();

      // Snimi dokument i uzmi kao base64
      // Word Online / Desktop: getDocumentAsBase64
      const result = doc.getContentControls();
      result.load("tag");
      await context.sync();

      // Pravi način: Office.js – doc.body.getOoxml() ne daje ceo docx
      // Koristimo: context.document.save() pa čitamo fajl
      // Ali to nije dostupno u taskpane-u bez special permissions.
      // Alternativa: koristimo Word.run i serializujemo sve CC tagove u XML
      // pa modifikujemo .dotx fajl koji je skinut sa GitHuba.
    });

    // Prava strategija:
    // 1. Skinuti originalni .dotx sa GitHub-a
    // 2. U njemu ažurirati BA_FIELD tagove prema rows[] stanju
    // 3. Uploadovati modifikovani .dotx

    body.innerHTML = `<div class="gh-loading"><div class="gh-loading-icon">⬇️</div><div>Skidam original sa GitHub-a...</div></div>`;

    const rawUrl = buildRawUrl(branch.id, fileName);
    const dlResp = await fetch(rawUrl);

    if (!dlResp.ok) throw new Error(`Download greška: HTTP ${dlResp.status}`);
    const originalBuffer = await dlResp.arrayBuffer();

    body.innerHTML = `<div class="gh-loading"><div class="gh-loading-icon">🔧</div><div>Ažuriram polja u templejtu...</div></div>`;

    // Modifikuj .dotx – ažuriraj tagove BA_FIELD po rows[]
    const modifiedBuffer = await updateDocxTags(originalBuffer);

    body.innerHTML = `<div class="gh-loading"><div class="gh-loading-icon">⬆️</div><div>Uplodujem na GitHub...</div></div>`;

    // Dobavi SHA za PUT (GitHub API zahteva SHA za update)
    const apiPath = `${config.apiUrl}/${encodeURIComponent(branch.id)}/${encodeURIComponent(fileName)}`;
    const shaResp = await fetch(apiPath, {
      headers: {
        "Accept": "application/vnd.github.v3+json",
        "Authorization": `token ${config.token}`
      }
    });

    let sha = null;
    if (shaResp.ok) {
      const shaData = await shaResp.json();
      sha = shaData.sha;
    }

    // Konvertuj buffer u base64
    const uint8 = new Uint8Array(modifiedBuffer);
    let bin = "";
    for (let i = 0; i < uint8.length; i++) bin += String.fromCharCode(uint8[i]);
    const b64 = btoa(bin);

    // PUT request na GitHub API
    const putBody = {
      message: commitMsg,
      content: b64,
    };
    if (sha) putBody.sha = sha;

    const putResp = await fetch(apiPath, {
      method: "PUT",
      headers: {
        "Accept": "application/vnd.github.v3+json",
        "Authorization": `token ${config.token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(putBody)
    });

    if (!putResp.ok) {
      const errData = await putResp.json().catch(() => ({}));
      throw new Error(`GitHub PUT greška: HTTP ${putResp.status} – ${errData.message || ""}`);
    }

    body.innerHTML = `
      <div class="gh-loading">
        <div class="gh-loading-icon">✅</div>
        <div style="color:#14532d; font-weight:600;">Uspešno sačuvano!</div>
        <div style="font-size:12px; color:#6b7280; margin-top:6px;">${escHtml(branch.id)}/${escHtml(fileName)}</div>
      </div>`;
    el("modalGHFooter").innerHTML = `<button id="btnGHDone" class="btn-primary">Zatvori</button>`;
    el("btnGHDone")?.addEventListener("click", closeGHModal);

    setStatus(`✅ Templejt sačuvan na GitHub: ${fileName}`, "success");

  } catch (err) {
    console.error("GitHub commit greška:", err);
    body.innerHTML = `
      <div class="gh-loading">
        <div class="gh-loading-icon" style="font-size:28px;">❌</div>
        <div class="gh-error" style="margin-top:8px;">${escHtml(err.message)}</div>
      </div>`;
    el("modalGHFooter").innerHTML = `<button id="btnGHRetry" class="btn-secondary">Nazad</button>`;
    el("btnGHRetry")?.addEventListener("click", () => renderGHStep3(branch, fileName));
    setStatus(`Greška: ${err.message}`, "error");
  }
}

// === OTVORI: Korak 1 – izbor grane ===
function renderGHOpenStep1() {
  const body = el("modalGHBody");
  body.innerHTML = `<p class="gh-step-title">Izaberi vrstu projekta:</p>`;

  config.branches.forEach(branch => {
    const btn = document.createElement("button");
    btn.className = "gh-btn";
    btn.innerHTML = `<div class="gh-btn-id">${escHtml(branch.id)}</div><div class="gh-btn-label">${escHtml(branch.label)}</div>`;
    btn.addEventListener("click", () => renderGHOpenStep2(branch));
    body.appendChild(btn);
  });
}

// === OTVORI: Korak 2 – izbor fajla ===
function renderGHOpenStep2(branch) {
  const body = el("modalGHBody");
  body.innerHTML = `
    <span class="gh-back" id="ghOpenBack1">← Nazad</span>
    <p class="gh-step-title">${escHtml(branch.id)} – izaberi svesku:</p>
  `;
  el("ghOpenBack1")?.addEventListener("click", renderGHOpenStep1);

  branch.files.forEach(fileName => {
    const btn = document.createElement("button");
    btn.className = "gh-btn";
    btn.innerHTML = `📄 ${escHtml(fileName)}`;
    btn.addEventListener("click", () => doOpenFromGitHub(branch, fileName));
    body.appendChild(btn);
  });
}

// === OTVORI: Skini i otvori dokument ===
async function doOpenFromGitHub(branch, fileName) {
  const body = el("modalGHBody");
  body.innerHTML = `<div class="gh-loading"><div class="gh-loading-icon">⏳</div><div>Skinam ${escHtml(fileName)}...</div></div>`;
  el("modalGHFooter").innerHTML = "";

  try {
    const url = buildRawUrl(branch.id, fileName);
    const resp = await fetch(url);

    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buffer = await resp.arrayBuffer();

    // Konvertuj u Base64 i otvori u Wordu
    const uint8 = new Uint8Array(buffer);
    let bin = "";
    for (let i = 0; i < uint8.length; i++) bin += String.fromCharCode(uint8[i]);
    const b64 = btoa(bin);

    await Word.run(async (context) => {
      const doc = context.application.createDocument(b64);
      doc.open();
      await context.sync();
    });

    closeGHModal();
    setStatus(`Otvoren: ${fileName}. Klikni Skeniraj da učitaš polja.`, "success");

  } catch (err) {
    console.error("Open greška:", err);
    body.innerHTML = `<div class="gh-loading"><div class="gh-error">Greška: ${escHtml(err.message)}</div></div>`;
    el("modalGHFooter").innerHTML = `<button id="btnGHOpenBack" class="btn-secondary">Nazad</button>`;
    el("btnGHOpenBack")?.addEventListener("click", () => renderGHOpenStep2(branch));
  }
}

// ==========================================
// DOCX TAG UPDATE (JSZip)
// ==========================================

// Ažurira BA_FIELD tagove u .docx/dotx fajlu prema rows[]
async function updateDocxTags(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file("word/document.xml").async("string");

  // Izgradi map: oldKey -> row (za brzu pretragu)
  const rowMap = new Map();
  for (const r of rows) {
    if (r.field) rowMap.set(r.field.trim(), r);
  }

  // Zameni w:val atribute u w:tag elementima
  // <w:tag w:val="BA_FIELD|key=...|type=...|format=..."/>
  const updated = docXml.replace(
    /(<w:tag\s+w:val=")([^"]*?)(")/g,
    (match, pre, tagVal, post) => {
      const meta = parseTag(tagVal);
      if (!meta) return match; // nije BA_FIELD tag, ne diraj

      const row = rowMap.get(meta.key);
      if (!row) return match; // ne postoji u rows[], ostavi kakav je

      const newTag = makeTag(row.field, row.type, row.format);
      return `${pre}${newTag}${post}`;
    }
  );

  // Takođe ažuriraj w:alias (title) – prikazuje se u Word UI
  const updatedWithAlias = updated.replace(
    /(<w:alias\s+w:val=")([^"]*?)(")/g,
    (match, pre, oldTitle, post) => {
      const row = rows.find(r => r.field && r.field.trim() === oldTitle.trim());
      if (!row) return match;
      return `${pre}${xmlEscape(row.field)}${post}`;
    }
  );

  zip.file("word/document.xml", updatedWithAlias);

  // Ako postoji .dotx (template), ažuriraj i word/settings.xml ako postoji
  // (ne menjamo, samo repackujemo)

  const result = await zip.generateAsync({ type: "arraybuffer" });
  return result;
}

// ==========================================
// GITHUB URL builder
// ==========================================
function buildRawUrl(branchId, fileName) {
  const base = (config.baseUrl || "").replace(/\/$/, "");
  return `${base}/${branchId.split("/").map(encodeURIComponent).join("/")}/${encodeURIComponent(fileName)}`;
}

// ==========================================
// CONFIG TAB
// ==========================================
function loadConfigToForm() {
  el("cfgBaseUrl").value = config.baseUrl || "";
  el("cfgApiUrl").value  = config.apiUrl  || "";
  el("cfgToken").value   = config.token   || "";
}

function saveConfigFromForm() {
  config.baseUrl = (el("cfgBaseUrl")?.value || "").trim();
  config.apiUrl  = (el("cfgApiUrl")?.value  || "").trim();
  config.token   = (el("cfgToken")?.value   || "").trim();
  saveConfigToDocument();
  setStatus("Podešavanja sačuvana.", "success");

  // Update header sub
  updateHeaderSub();
}

async function testConnection() {
  const apiUrl = (el("cfgApiUrl")?.value || "").trim();
  if (!apiUrl) { setStatus("Unesi API URL.", "warn"); return; }

  setStatus("Testiram konekciju...", "info");
  try {
    const token = (el("cfgToken")?.value || "").trim();
    const resp = await fetch(apiUrl, {
      headers: {
        "Accept": "application/vnd.github.v3+json",
        ...(token ? { "Authorization": `token ${token}` } : {})
      }
    });
    if (resp.ok) setStatus(`✅ Konekcija uspešna! (HTTP ${resp.status})`, "success");
    else if (resp.status === 401) setStatus("❌ Neautorizovan. Proveri token.", "error");
    else if (resp.status === 404) setStatus("❌ Repo nije pronađen. Proveri URL.", "error");
    else setStatus(`⚠️ HTTP ${resp.status}`, "warn");
  } catch (err) {
    setStatus(`❌ Greška: ${err.message}`, "error");
  }
}

// ==========================================
// BRANCHES (Config tab)
// ==========================================
function renderBranches() {
  const list = el("branchesList");
  if (!list) return;

  list.innerHTML = "";

  if (config.branches.length === 0) {
    list.innerHTML = `<div style="color:#9ca3af; font-size:13px; font-style:italic; padding:8px 0;">Nema grana.</div>`;
    return;
  }

  config.branches.forEach((branch, idx) => {
    const card = document.createElement("div");
    card.className = "branch-card";

    const header = document.createElement("div");
    header.className = "branch-card-header";
    header.addEventListener("click", () => card.classList.toggle("open"));

    header.innerHTML = `
      <span class="branch-toggle">▶</span>
      <div class="branch-card-info">
        <div class="branch-card-id">${escHtml(branch.id)}</div>
        <div class="branch-card-label">${escHtml(branch.label)}</div>
      </div>
      <span class="branch-badge">${branch.files.length} fajla</span>
    `;

    const actions = document.createElement("div");
    actions.className = "branch-actions";

    const btnEdit = document.createElement("button");
    btnEdit.className = "branch-btn";
    btnEdit.innerHTML = "✏️";
    btnEdit.title = "Izmeni";
    btnEdit.addEventListener("click", (e) => { e.stopPropagation(); openBranchModal(idx); });

    const btnDel = document.createElement("button");
    btnDel.className = "branch-btn del";
    btnDel.innerHTML = "🗑️";
    btnDel.title = "Obriši";
    btnDel.addEventListener("click", (e) => {
      e.stopPropagation();
      if (confirm(`Obrisati granu "${branch.id}"?`)) {
        config.branches.splice(idx, 1);
        saveConfigToDocument();
        renderBranches();
        setStatus(`Grana "${branch.id}" obrisana.`, "success");
      }
    });

    actions.appendChild(btnEdit);
    actions.appendChild(btnDel);
    header.appendChild(actions);

    // Files list
    const filesDiv = document.createElement("div");
    filesDiv.className = "branch-files";

    branch.files.forEach((f, fidx) => {
      const row = document.createElement("div");
      row.className = "branch-file-row";
      row.innerHTML = `<span class="branch-file-name">📄 ${escHtml(f)}</span>`;

      const btnDelFile = document.createElement("button");
      btnDelFile.className = "branch-file-del";
      btnDelFile.innerHTML = "×";
      btnDelFile.title = "Ukloni fajl";
      btnDelFile.addEventListener("click", (e) => {
        e.stopPropagation();
        if (confirm(`Ukloniti "${f}" iz grane "${branch.id}"?`)) {
          config.branches[idx].files.splice(fidx, 1);
          saveConfigToDocument();
          renderBranches();
        }
      });

      row.appendChild(btnDelFile);
      filesDiv.appendChild(row);
    });

    card.appendChild(header);
    card.appendChild(filesDiv);
    list.appendChild(card);
  });
}

function openBranchModal(branchIndex = null) {
  editingBranchIndex = branchIndex;

  const title   = el("modalBranchTitle");
  const idInp   = el("branchIdInput");
  const labInp  = el("branchLabelInput");
  const filesEd = el("branchFilesEditor");

  let branch = { id: "", label: "", files: [] };
  if (branchIndex !== null) {
    branch = JSON.parse(JSON.stringify(config.branches[branchIndex]));
    title.textContent = "Izmeni granu";
    idInp.value = branch.id;
    idInp.disabled = true;
    idInp.style.background = "#f3f4f6";
  } else {
    title.textContent = "Nova grana";
    idInp.value = "";
    idInp.disabled = false;
    idInp.style.background = "#fff";
  }

  labInp.value = branch.label;
  filesEd.innerHTML = "";
  if (branch.files.length === 0) addBranchFileRow("");
  else branch.files.forEach(f => addBranchFileRow(f));

  el("modalBranch")?.classList.remove("hidden");
  el("modalBranchBackdrop")?.classList.remove("hidden");
  setTimeout(() => (branchIndex === null ? idInp : labInp).focus(), 50);
}

function closeBranchModal() {
  el("modalBranch")?.classList.add("hidden");
  el("modalBranchBackdrop")?.classList.add("hidden");
  editingBranchIndex = null;
}

function addBranchFileRow(value = "") {
  const filesEd = el("branchFilesEditor");
  const row = document.createElement("div");
  row.className = "file-editor-row";

  const inp = document.createElement("input");
  inp.type = "text";
  inp.placeholder = "npr. 00 PGD Glavna sveska.dotx";
  inp.value = value;

  const btnDel = document.createElement("button");
  btnDel.className = "btn-del-file";
  btnDel.innerHTML = "×";
  btnDel.addEventListener("click", () => {
    const rows = filesEd.querySelectorAll(".file-editor-row");
    if (rows.length > 1) row.remove();
    else inp.value = "";
  });

  row.appendChild(inp);
  row.appendChild(btnDel);
  filesEd.appendChild(row);
}

function saveBranchModal() {
  const id    = el("branchIdInput")?.value.trim();
  const label = el("branchLabelInput")?.value.trim();

  if (!label) { setStatus("Unesi naziv grane.", "error"); return; }

  if (editingBranchIndex === null && !id) {
    setStatus("Unesi ID grane.", "error");
    return;
  }

  if (editingBranchIndex === null && config.branches.some(b => b.id === id)) {
    setStatus(`Grana "${id}" već postoji.`, "error");
    return;
  }

  const files = Array.from(el("branchFilesEditor")?.querySelectorAll("input") || [])
    .map(i => i.value.trim())
    .filter(Boolean);

  if (editingBranchIndex !== null) {
    config.branches[editingBranchIndex].label = label;
    config.branches[editingBranchIndex].files = files;
    setStatus(`Grana ažurirana.`, "success");
  } else {
    config.branches.push({ id, label, files });
    setStatus(`Grana "${id}" dodata.`, "success");
  }

  saveConfigToDocument();
  renderBranches();
  closeBranchModal();
}

// ==========================================
// XML PERSISTENCE (config)
// ==========================================
function buildConfigXml() {
  let xml = `<admin-config xmlns="${ADMIN_XML_NS}">`;
  xml += `<repo baseUrl="${xmlEscape(config.baseUrl)}" apiUrl="${xmlEscape(config.apiUrl)}" token="${xmlEscape(config.token)}"/>`;
  xml += `<branches>`;
  config.branches.forEach(b => {
    xml += `<branch id="${xmlEscape(b.id)}" label="${xmlEscape(b.label)}">`;
    b.files.forEach(f => { xml += `<file name="${xmlEscape(f)}"/>`; });
    xml += `</branch>`;
  });
  xml += `</branches></admin-config>`;
  return xml;
}

async function saveConfigToDocument() {
  try {
    const xml = buildConfigXml();
    await Word.run(async (ctx) => {
      const parts = ctx.document.customXmlParts;
      parts.load("items");
      await ctx.sync();
      for (const p of parts.items) p.load("namespaceUri");
      await ctx.sync();
      for (const p of parts.items) { if (p.namespaceUri === ADMIN_XML_NS) p.delete(); }
      await ctx.sync();
      parts.add(xml);
      await ctx.sync();
    });
  } catch (err) {
    console.error("Save config greška:", err);
  }
}

async function loadConfigFromDocument() {
  try {
    let loaded = false;
    await Word.run(async (ctx) => {
      const parts = ctx.document.customXmlParts;
      parts.load("items");
      await ctx.sync();
      for (const p of parts.items) p.load("namespaceUri");
      await ctx.sync();

      const mine = parts.items.find(p => p.namespaceUri === ADMIN_XML_NS);
      if (!mine) return;

      const xml = mine.getXml();
      await ctx.sync();

      const parsed = parseConfigXml(xml.value || "");
      if (parsed) { config = parsed; loaded = true; }
    });

    if (!loaded) config = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  } catch (err) {
    console.error("Load config greška:", err);
    config = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  }
}

function parseConfigXml(str) {
  try {
    const doc = new DOMParser().parseFromString(str, "text/xml");
    if (doc.querySelector("parsererror")) return null;

    const repoEl = doc.querySelector("repo");
    const result = {
      baseUrl: repoEl?.getAttribute("baseUrl") || DEFAULT_CONFIG.baseUrl,
      apiUrl:  repoEl?.getAttribute("apiUrl")  || DEFAULT_CONFIG.apiUrl,
      token:   repoEl?.getAttribute("token")   || "",
      branches: []
    };

    doc.querySelectorAll("branch").forEach(b => {
      const id    = b.getAttribute("id")    || "";
      const label = b.getAttribute("label") || "";
      const files = Array.from(b.querySelectorAll("file"))
        .map(f => f.getAttribute("name") || "").filter(Boolean);
      if (id) result.branches.push({ id, label, files });
    });

    return result;
  } catch { return null; }
}

// ==========================================
// HEADER SUB
// ==========================================
function updateHeaderSub() {
  const sub = el("headerSub");
  if (!sub) return;
  try {
    const url = new URL(config.baseUrl);
    const parts = url.pathname.split("/").filter(Boolean);
    sub.textContent = parts.slice(0,2).join("/") || "GitHub templejti";
  } catch {
    sub.textContent = "GitHub templejti";
  }
}

// ==========================================
// TEST FILL / CLEAR
// ==========================================

function applyFormat(type, format, rawValue) {
  const v = String(rawValue ?? "");
  if (!v) return "";

  if (type === "number") {
    const n = Number(v.replace(/\./g, "").replace(",", ".").replace(/[^\d.-]/g, ""));
    if (isNaN(n)) return v;
    const fmt = (num, dec) => {
      const fixed = num.toFixed(dec);
      const parts = fixed.split(".");
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      return dec > 0 ? parts[0] + "," + parts[1] : parts[0];
    };
    if (format === "number:int") return fmt(Math.round(n), 0);
    if (format === "number:2")   return fmt(n, 2);
    if (format === "number:rsd") return fmt(n, 2) + " RSD";
    if (format === "number:eur") return fmt(n, 2) + " €";
    if (format === "number:usd") return fmt(n, 2) + " $";
    return String(n);
  }

  if (type === "date") {
    const months = ["januar","februar","mart","april","maj","jun","jul","avgust","septembar","oktobar","novembar","decembar"];
    if (format === "date:today") {
      const d = new Date();
      return `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`;
    }
    if (format === "date:mmmm.yyyy" || format === "date:dd.mmmm.yyyy") {
      let d;
      if (v.includes(".")) { const p = v.split("."); d = new Date(p[2], p[1]-1, p[0]); }
      else if (v.includes("-")) { d = new Date(v); }
      if (!d || isNaN(d.getTime())) return v;
      const mn = months[d.getMonth()];
      if (format === "date:mmmm.yyyy") return `${mn}.${d.getFullYear()}`;
      return `${String(d.getDate()).padStart(2,"0")}.${mn}.${d.getFullYear()}`;
    }
    return v;
  }

  if (type === "text") {
    if (format === "text:upper") return v.toUpperCase();
    if (format === "text:lower") return v.toLowerCase();
    if (format === "text:title") return v.replace(/\b\w/g, l => l.toUpperCase());
  }

  return v;
}

async function fillTest() {
  const map = new Map();
  for (const r of rows) {
    const key = (r.field || "").trim();
    if (!key) continue;
    const raw = (r.value || "").trim();
    map.set(key, raw ? applyFormat(r.type, r.format, raw) : token(key));
  }

  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      let filled = 0;
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;
        const val = map.get(meta.key) ?? token(meta.key);
        cc.insertText(val, Word.InsertLocation.replace);
        filled++;
      }

      await context.sync();
      setStatus(`Popunjeno ${filled} polja (test vrednosti).`, "success");
    });
  } catch (err) {
    console.error("Fill test greška:", err);
    setStatus("Greška pri popunjavanju.", "error");
  }
}

async function clearTest() {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items/tag");
      await context.sync();

      let cleared = 0;
      for (const cc of ccs.items) {
        const meta = parseTag(cc.tag || "");
        if (!meta) continue;
        cc.insertText(token(meta.key), Word.InsertLocation.replace);
        cleared++;
      }

      await context.sync();
      setStatus(`Očišćeno ${cleared} polja → {POLJE} vraćen.`, "info");
    });
  } catch (err) {
    console.error("Clear test greška:", err);
    setStatus("Greška pri čišćenju.", "error");
  }
}

// ==========================================
// BIND UI
// ==========================================
function bindUi() {
  // Tabs
  el("btnTabFields")?.addEventListener("click", () => switchTab("btnTabFields"));
  el("btnTabConfig")?.addEventListener("click", () => switchTab("btnTabConfig"));

  // Header GitHub btn → picker: otvori ili snimi
  el("btnGitHub")?.addEventListener("click", () => {
    // Prikaži mini menu
    showGHMenu();
  });

  // Fields tab
  el("btnInsertField")?.addEventListener("click", insertFieldAtSelection);
  el("btnScanDoc")?.addEventListener("click", scanDocument);
  el("btnFillTest")?.addEventListener("click", fillTest);
  el("btnClearTest")?.addEventListener("click", clearTest);
  el("btnSyncGitHub")?.addEventListener("click", openGitHubSaveModal);
  el("btnAddRow")?.addEventListener("click", () => {
    rows.push({ id: crypto.randomUUID(), field: "", type: "text", format: "text:auto", value: "", description: "" });
    renderRows();
    // Focus last row input
    setTimeout(() => {
      const inputs = el("rows")?.querySelectorAll("input");
      inputs?.[inputs.length - 1]?.focus();
    }, 50);
  });

  // Type modal
  el("btnModalClose")?.addEventListener("click", closeTypeModal);
  el("btnModalCancel")?.addEventListener("click", closeTypeModal);
  el("btnModalOk")?.addEventListener("click", saveTypeModal);
  el("modalBackdrop")?.addEventListener("click", (e) => { if (e.target === el("modalBackdrop")) closeTypeModal(); });
  el("modal")?.addEventListener("click", e => e.stopPropagation());

  // Rename modal
  el("btnRenameClose")?.addEventListener("click", closeRenameModal);
  el("btnRenameCancel")?.addEventListener("click", closeRenameModal);
  el("btnRenameOk")?.addEventListener("click", doRename);
  el("renameNewName")?.addEventListener("keydown", (e) => { if (e.key === "Enter") doRename(); });
  el("modalRenameBackdrop")?.addEventListener("click", (e) => { if (e.target === el("modalRenameBackdrop")) closeRenameModal(); });
  el("modalRename")?.addEventListener("click", e => e.stopPropagation());

  // Description modal
  el("btnDescClose")?.addEventListener("click", closeDescriptionModal);
  el("btnDescCancel")?.addEventListener("click", closeDescriptionModal);
  el("btnDescSave")?.addEventListener("click", saveDescriptionModal);
  el("btnDescClear")?.addEventListener("click", () => {
    if (el("descText")) el("descText").value = "";
    saveDescriptionModal();
  });
  el("modalDescBackdrop")?.addEventListener("click", (e) => { if (e.target === el("modalDescBackdrop")) closeDescriptionModal(); });
  el("modalDesc")?.addEventListener("click", e => e.stopPropagation());

  // GitHub modal
  el("modalGHBackdrop")?.addEventListener("click", (e) => { if (e.target === el("modalGHBackdrop")) closeGHModal(); });
  el("modalGH")?.addEventListener("click", e => e.stopPropagation());

  // Config tab
  el("btnSaveConfig")?.addEventListener("click", () => {
    config.baseUrl = (el("cfgBaseUrl")?.value || "").trim();
    config.apiUrl  = (el("cfgApiUrl")?.value  || "").trim();
    config.token   = (el("cfgToken")?.value   || "").trim();
    saveConfigToDocument();
    updateHeaderSub();
    setStatus("Podešavanja sačuvana.", "success");
  });
  el("btnTestConn")?.addEventListener("click", testConnection);
  el("btnAddBranch")?.addEventListener("click", () => openBranchModal(null));

  // Branch modal
  el("btnBranchClose")?.addEventListener("click", closeBranchModal);
  el("btnBranchCancel")?.addEventListener("click", closeBranchModal);
  el("btnBranchSave")?.addEventListener("click", saveBranchModal);
  el("btnAddBranchFile")?.addEventListener("click", () => addBranchFileRow(""));
  el("modalBranchBackdrop")?.addEventListener("click", (e) => { if (e.target === el("modalBranchBackdrop")) closeBranchModal(); });
  el("modalBranch")?.addEventListener("click", e => e.stopPropagation());
}

// ==========================================
// GITHUB MENU (header button)
// ==========================================
function showGHMenu() {
  // Kreiraj mini dropdown
  const existing = document.getElementById("ghMenu");
  if (existing) { existing.remove(); return; }

  const menu = document.createElement("div");
  menu.id = "ghMenu";
  menu.style.cssText = `
    position: fixed; top: 52px; right: 14px;
    background: #fff; border: 1px solid #e5e7eb;
    border-radius: 8px; box-shadow: 0 8px 20px rgba(0,0,0,0.15);
    z-index: 2000; min-width: 200px; overflow: hidden;
  `;

  const items = [
    { icon: "📂", label: "Otvori templejt sa GitHub-a", fn: () => { menu.remove(); openGitHubOpenModal(); } },
    { icon: "↑",  label: "Snimi templejt na GitHub",    fn: () => { menu.remove(); openGitHubSaveModal(); } },
  ];

  items.forEach(item => {
    const btn = document.createElement("button");
    btn.style.cssText = `
      display: flex; align-items: center; gap: 10px;
      width: 100%; padding: 11px 14px; border: none;
      background: #fff; text-align: left; cursor: pointer;
      font-size: 13px; color: #1f2937; font-family: inherit;
      border-bottom: 1px solid #f3f4f6; transition: background 0.1s;
    `;
    btn.innerHTML = `<span>${item.icon}</span><span>${item.label}</span>`;
    btn.addEventListener("mouseenter", () => btn.style.background = "#f3f4f6");
    btn.addEventListener("mouseleave", () => btn.style.background = "#fff");
    btn.addEventListener("click", item.fn);
    menu.appendChild(btn);
  });

  document.body.appendChild(menu);

  // Zatvori na klik van menija
  setTimeout(() => {
    document.addEventListener("click", function handler() {
      menu.remove();
      document.removeEventListener("click", handler);
    });
  }, 10);
}

// ==========================================
// INIT
// ==========================================
Office.onReady(async () => {
  console.log("⚙️ BiroA Admin – Office.onReady");

  try {
    await loadConfigFromDocument();
    console.log("✅ Config učitan:", config.branches.length, "grana");
  } catch (err) {
    console.error("Load config greška:", err);
  }

  updateHeaderSub();
  renderRows();
  bindUi();

  console.log("✅ BiroA Admin spreman");
});
