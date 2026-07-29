/* ============================================================
   Pearl Fashions — Nightsuit Costing
   Storage: Google Sheets (data rows) + Google Drive (images)
   Nothing here talks to any server other than Google's own APIs.
   ============================================================ */

const SHEET_NAME = "CostingSheets";
const HEADERS = [
  "ID", "Date Added", "Fabric Image URL", "Design Image URL",
  "Supplier Name", "Bill Number", "Bill Date",
  "Total Fabric Bill Amount (₹)",
  "Fabric Rate (per kg/meter)", "Total Fabric Amount (kg/meter)",
  "Cutting Master", "Style Code",
  "Fabric Length (in)", "Fabric Width (in)", "GSM", "Number Of Pieces",
  "Cloth Used Per Piece (kg/meter)", "Fabric Cost Per Piece", "Total Fabric Cost",
  "Job Work Description", "Job Work Cost", "Fixed Cost", "Extra Accessories Cost", "MRP Percent",
  "Total Cost Per Piece", "Final MRP", "Notes",
  "Manual Cloth Used (kg/meter)", "Use Manual Cloth Used", "Cutting Remarks (Sizes & Ratios)"
];
// column indices (0-based) for readability when reading rows back
const COL = {
  ID:0, DATE:1, FABRIC_IMG:2, DESIGN_IMG:3, SUPPLIER:4, BILL_NO:5, BILL_DATE:6,
  TOTAL_FABRIC_BILL_AMOUNT:7,
  FABRIC_RATE:8, TOTAL_FABRIC_AMOUNT:9,
  CUTTING_MASTER:10, STYLE:11,
  FAB_LENGTH:12, FAB_WIDTH:13, GSM:14, NUM_PIECES:15,
  CLOTH_USED:16, FABRIC_COST_PIECE:17, TOTAL_FABRIC_COST:18,
  JOB_DESC:19, JOB_COST:20, FIXED_COST:21, ACCESSORIES_COST:22, MRP_PCT:23,
  VARIABLE_COST:24, FINAL_MRP:25, NOTES:26,
  MANUAL_CLOTH_USED:27, USE_MANUAL_CLOTH:28, CUTTING_REMARKS:29
};

let CONFIG = {
  clientId: localStorage.getItem("ncs_clientId") || "",
  apiKey: localStorage.getItem("ncs_apiKey") || "",
  sheetId: localStorage.getItem("ncs_sheetId") || "",
  folderId: localStorage.getItem("ncs_folderId") || ""
};

let tokenClient = null;
let gapiReady = false;
let sheetGid = null; // numeric sheetId of the CostingSheets tab, needed for row deletion / formatting
let fabricImage = { url: "", fileId: "" };
let designImage = { url: "", fileId: "" };
let allRows = []; // cache of loaded rows, each {rowIndex, values:[...]}

let editingRowIndex = null; // set when editing an existing entry instead of creating a new one
let editingId = null;
let editingDate = null;

/* ---------------- boot ---------------- */
window.addEventListener("DOMContentLoaded", () => {
  document.getElementById("cfgClientId").value = CONFIG.clientId;
  document.getElementById("cfgApiKey").value = CONFIG.apiKey;
  document.getElementById("cfgSheetId").value = CONFIG.sheetId;
  document.getElementById("cfgFolderId").value = CONFIG.folderId;

  wireTabs();
  wireSettings();
  wireUploads();
  wireForm();

  document.getElementById("signInBtn").addEventListener("click", handleSignIn);
  document.getElementById("refreshBtn").addEventListener("click", loadEntries);
  document.getElementById("cancelEditBtn").addEventListener("click", cancelEdit);

  if (isConfigured()) {
    document.getElementById("setupBanner").style.display = "none";
    initGoogle();
  }
});

function isConfigured() {
  return CONFIG.clientId && CONFIG.apiKey && CONFIG.sheetId && CONFIG.folderId;
}

/* ---------------- tabs ---------------- */
function wireTabs() {
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      const tab = btn.dataset.tab;
      document.getElementById("tab-new").style.display = tab === "new" ? "block" : "none";
      document.getElementById("tab-list").style.display = tab === "list" ? "block" : "none";
      if (tab === "list") loadEntries();
    });
  });
}

/* ---------------- settings modal ---------------- */
function openSettings() { document.getElementById("settingsModal").classList.add("open"); }
function closeSettings() { document.getElementById("settingsModal").classList.remove("open"); }

function wireSettings() {
  document.getElementById("settingsBtn").addEventListener("click", openSettings);
  document.getElementById("saveSettingsBtn").addEventListener("click", () => {
    CONFIG.clientId = document.getElementById("cfgClientId").value.trim();
    CONFIG.apiKey = document.getElementById("cfgApiKey").value.trim();
    CONFIG.sheetId = document.getElementById("cfgSheetId").value.trim();
    CONFIG.folderId = document.getElementById("cfgFolderId").value.trim();
    localStorage.setItem("ncs_clientId", CONFIG.clientId);
    localStorage.setItem("ncs_apiKey", CONFIG.apiKey);
    localStorage.setItem("ncs_sheetId", CONFIG.sheetId);
    localStorage.setItem("ncs_folderId", CONFIG.folderId);
    closeSettings();
    if (isConfigured()) {
      document.getElementById("setupBanner").style.display = "none";
      initGoogle();
    } else {
      alert("Please fill in all four fields.");
    }
  });
}

/* ---------------- Google init ---------------- */
function initGoogle() {
  gapi.load("client", async () => {
    await gapi.client.init({ apiKey: CONFIG.apiKey });
    await gapi.client.load("https://sheets.googleapis.com/$discovery/rest?version=v4");
    await gapi.client.load("https://www.googleapis.com/discovery/v1/apis/drive/v3/rest");
    gapiReady = true;
  });

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.clientId,
    scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file",
    callback: async (resp) => {
      if (resp.error) {
        setStatus(false, "Sign-in failed");
        return;
      }
      gapi.client.setToken(resp);
      setStatus(true, "Connected to Google");
      document.getElementById("signInBtn").textContent = "Signed in ✓";
      await getSheetGid();
      await ensureHeaders();
    }
  });
}

function handleSignIn() {
  if (!isConfigured()) { openSettings(); return; }
  if (!tokenClient) { alert("Still initializing, try again in a second."); return; }
  tokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
}

function setStatus(on, text) {
  document.getElementById("statusDot").classList.toggle("on", on);
  document.getElementById("statusText").textContent = text;
}

/* ---------------- sheet bootstrap ---------------- */
async function ensureHeaders() {
  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.sheetId,
      range: `${SHEET_NAME}!A1:Z1`
    });
    if (!res.result.values || res.result.values.length === 0) {
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.sheetId,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: "RAW",
        resource: { values: [HEADERS] }
      });
    }
    await formatHeaderRow();
  } catch (e) {
    console.error("ensureHeaders error", e);
    if (e.status === 400 || (e.result && e.result.error && e.result.error.status === "INVALID_ARGUMENT")) {
      alert(`Could not find a tab named "${SHEET_NAME}" in your sheet. Please create/rename a tab to exactly "${SHEET_NAME}" and reload.`);
    }
  }
}

// bold white text on a dark-blue band, frozen so it always stays visible while scrolling
async function formatHeaderRow() {
  if (sheetGid === null) return;
  try {
    await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.sheetId,
      resource: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: sheetGid, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  backgroundColor: { red: 0.078, green: 0.129, blue: 0.239 },
                  textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 10 },
                  horizontalAlignment: "LEFT"
                }
              },
              fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
          },
          {
            updateSheetProperties: {
              properties: { sheetId: sheetGid, gridProperties: { frozenRowCount: 1 } },
              fields: "gridProperties.frozenRowCount"
            }
          }
        ]
      }
    });
  } catch (e) { console.error("formatHeaderRow error", e); }
}

async function getSheetGid() {
  try {
    const meta = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: CONFIG.sheetId });
    const sheet = meta.result.sheets.find(s => s.properties.title === SHEET_NAME);
    if (sheet) sheetGid = sheet.properties.sheetId;
  } catch (e) { console.error(e); }
}

/* ---------------- image upload ---------------- */
function wireUploads() {
  setupUploadBox("fabricImageInput", "fabricUploadBox", "fabricUploadStatus", fabricImage);
  setupUploadBox("designImageInput", "designUploadBox", "designUploadStatus", designImage);
}

function setupUploadBox(inputId, boxId, statusId, targetObj) {
  const input = document.getElementById(inputId);
  const box = document.getElementById(boxId);
  const status = document.getElementById(statusId);

  input.addEventListener("change", async () => {
    const file = input.files[0];
    if (!file) return;
    if (!gapi.client.getToken()) {
      alert("Please sign in with Google first.");
      input.value = "";
      return;
    }
    status.textContent = "Uploading…";
    try {
      const { fileId, viewUrl } = await uploadImageToDrive(file);
      targetObj.fileId = fileId;
      targetObj.url = viewUrl;
      box.querySelector(".placeholder")?.remove();
      let img = box.querySelector("img");
      if (!img) { img = document.createElement("img"); box.prepend(img); }
      img.src = URL.createObjectURL(file);
      status.textContent = "Uploaded ✓";
    } catch (e) {
      console.error(e);
      status.textContent = "Upload failed";
      alert("Image upload failed. Check your Drive folder ID and permissions in Settings.");
    }
  });
}

// used when editing an entry: shows the already-uploaded image without needing a new file
function setExistingImagePreview(boxId, url) {
  const box = document.getElementById(boxId);
  if (!box) return;
  box.querySelector(".placeholder")?.remove();
  let img = box.querySelector("img");
  if (!url) {
    img?.remove();
    if (!box.querySelector(".placeholder")) {
      const span = document.createElement("span");
      span.className = "placeholder";
      span.textContent = "Click to upload photo";
      box.prepend(span);
    }
    return;
  }
  if (!img) { img = document.createElement("img"); box.prepend(img); }
  img.src = thumbUrl(url);
}

async function uploadImageToDrive(file) {
  const accessToken = gapi.client.getToken().access_token;
  const metadata = { name: `${Date.now()}_${file.name}`, parents: [CONFIG.folderId] };

  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", file);

  const uploadRes = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
    { method: "POST", headers: { Authorization: `Bearer ${accessToken}` }, body: form }
  );
  const uploadJson = await uploadRes.json();
  if (!uploadJson.id) throw new Error("No file id returned from Drive upload");
  const fileId = uploadJson.id;

  // make it viewable by anyone with the link, so it can be shown in the app / sheet
  await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}/permissions`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ role: "reader", type: "anyone" })
  });

  const viewUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
  return { fileId, viewUrl };
}

function thumbUrl(driveUrl) {
  // convert stored view url into a Drive thumbnail url for fast, reliable <img> rendering
  const match = driveUrl && driveUrl.match(/id=([^&]+)/);
  if (!match) return driveUrl || "";
  return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w400`;
}

/* ---------------- live calculations ---------------- */
const calcFieldIds = [
  "fabricRate", "totalFabricAmount", "numPieces", "jobWorkCost", "fixedCost", "accessoriesCost", "mrpPercent",
  "manualClothUsed"
];

function wireForm() {
  calcFieldIds.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", recalc);
    else console.warn(`wireForm(): no element with id "${id}"`);
  });
  const useManualEl = document.getElementById("useManualCloth");
  if (useManualEl) useManualEl.addEventListener("change", recalc);
  document.getElementById("entryForm").addEventListener("submit", handleSave);
  document.getElementById("resetFormBtn").addEventListener("click", () => { cancelEdit(); });
  recalc();
}

function num(id) {
  const el = document.getElementById(id);
  if (!el) { console.warn(`num(): no element with id "${id}"`); return 0; }
  const v = parseFloat(el.value);
  return Number.isFinite(v) ? v : 0;
}

// sets textContent on an element if it exists, without letting a missing
// element (or bad value) stop the other calc fields from updating
function setCalcText(id, text) {
  const el = document.getElementById(id);
  if (!el) { console.warn(`setCalcText(): no element with id "${id}"`); return; }
  el.textContent = text;
}

function recalc() {
  const fabricRate = num("fabricRate");
  const totalFabricAmount = num("totalFabricAmount");
  const numPieces = num("numPieces");
  const jobWorkCost = num("jobWorkCost");
  const fixedCost = num("fixedCost");
  const accessoriesCost = num("accessoriesCost");
  const mrpPercent = num("mrpPercent");
  const manualClothUsed = num("manualClothUsed");
  const useManualCloth = document.getElementById("useManualCloth")?.checked || false;

  // Total fabric bill amount is auto-calculated: fabric rate × total fabric amount
  const totalFabricBillAmount = fabricRate * totalFabricAmount;
  const billAmountEl = document.getElementById("totalFabricBillAmount");
  if (billAmountEl) billAmountEl.value = totalFabricBillAmount ? totalFabricBillAmount.toFixed(2) : "";

  // Cloth used per piece: use the manual override figure if the tick is checked, else fall back to the formula
  const formulaClothUsed = numPieces > 0 ? totalFabricAmount / numPieces : 0;
  const clothUsed = (useManualCloth && manualClothUsed > 0) ? manualClothUsed : formulaClothUsed;

  const fabricCostPerPiece = fabricRate * clothUsed;
  const totalFabricCost = fabricRate * totalFabricAmount;
  // Total cost / piece = fabric cost/piece + job work cost + fixed cost + accessories cost
  const totalCostPerPiece = fabricCostPerPiece + jobWorkCost + fixedCost + accessoriesCost;
  // Final MRP = total cost/piece multiplied directly by the MRP number (not treated as a %)
  const finalMRP = totalCostPerPiece * mrpPercent;

  setCalcText("calcClothUsed", clothUsed.toFixed(3));
  setCalcText("calcFabricPerPiece", "₹" + fabricCostPerPiece.toFixed(2));
  setCalcText("calcTotalCost", "₹" + totalCostPerPiece.toFixed(2));
  setCalcText("calcMRP", "₹" + finalMRP.toFixed(2));

  return { clothUsed, fabricCostPerPiece, totalFabricCost, totalFabricBillAmount, totalCostPerPiece, finalMRP };
}

/* ---------------- save / update entry ---------------- */
async function handleSave(e) {
  e.preventDefault();
  if (!gapi.client.getToken()) { alert("Please sign in with Google first."); return; }

  const c = recalc();
  const row = [
    editingId || ("NS-" + Date.now()),
    editingDate || new Date().toISOString().slice(0, 10),
    fabricImage.url,
    designImage.url,
    document.getElementById("supplierName").value,
    document.getElementById("billNumber").value,
    document.getElementById("billDate").value,
    c.totalFabricBillAmount.toFixed(2),
    num("fabricRate"),
    num("totalFabricAmount"),
    document.getElementById("cuttingMaster").value,
    document.getElementById("styleCode").value,
    num("fabLength"),
    num("fabWidth"),
    num("fabGSM"),
    num("numPieces"),
    c.clothUsed.toFixed(3),
    c.fabricCostPerPiece.toFixed(2),
    c.totalFabricCost.toFixed(2),
    document.getElementById("jobWorkDesc").value,
    num("jobWorkCost"),
    num("fixedCost"),
    num("accessoriesCost"),
    num("mrpPercent"),
    c.totalCostPerPiece.toFixed(2),
    c.finalMRP.toFixed(2),
    document.getElementById("notes").value,
    num("manualClothUsed"),
    document.getElementById("useManualCloth").checked ? "TRUE" : "FALSE",
    document.getElementById("cuttingRemarks").value
  ];

  const saveBtn = document.getElementById("saveBtn");
  const isEdit = editingRowIndex !== null;
  saveBtn.disabled = true; saveBtn.textContent = isEdit ? "Updating…" : "Saving…";
  try {
    if (isEdit) {
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.sheetId,
        range: `${SHEET_NAME}!A${editingRowIndex}:Z${editingRowIndex}`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [row] }
      });
    } else {
      await gapi.client.sheets.spreadsheets.values.append({
        spreadsheetId: CONFIG.sheetId,
        range: `${SHEET_NAME}!A:A`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [row] }
      });
    }
    saveBtn.textContent = isEdit ? "Updated ✓" : "Saved ✓";
    setTimeout(() => { saveBtn.disabled = false; }, 1200);
    cancelEdit(); // clears edit state + resets form
    document.querySelector('.tab-btn[data-tab="list"]').click();
  } catch (err) {
    console.error(err);
    alert("Could not save — check your sign-in and sheet ID in Settings.");
    saveBtn.textContent = isEdit ? "Update costing sheet" : "Save costing sheet";
    saveBtn.disabled = false;
  }
}

function resetForm() {
  document.getElementById("entryForm").reset();
  fabricImage = { url: "", fileId: "" };
  designImage = { url: "", fileId: "" };
  ["fabricUploadBox", "designUploadBox"].forEach(id => {
    const box = document.getElementById(id);
    box.querySelector("img")?.remove();
    if (!box.querySelector(".placeholder")) {
      const span = document.createElement("span");
      span.className = "placeholder";
      span.textContent = "Click to upload " + (id === "fabricUploadBox" ? "fabric photo" : "design photo");
      box.prepend(span);
    }
  });
  document.getElementById("fabricUploadStatus").textContent = "";
  document.getElementById("designUploadStatus").textContent = "";
  recalc();
}

/* ---------------- edit mode ---------------- */
function editEntry(rowIndex) {
  const row = allRows.find(r => r.rowIndex === rowIndex);
  if (!row) return;
  const v = row.values;
  const get = (idx) => v[idx] || "";

  editingRowIndex = rowIndex;
  editingId = get(COL.ID);
  editingDate = get(COL.DATE);

  document.getElementById("supplierName").value = get(COL.SUPPLIER);
  document.getElementById("billNumber").value = get(COL.BILL_NO);
  document.getElementById("billDate").value = get(COL.BILL_DATE);
  document.getElementById("totalFabricBillAmount").value = get(COL.TOTAL_FABRIC_BILL_AMOUNT);
  document.getElementById("fabricRate").value = get(COL.FABRIC_RATE);
  document.getElementById("totalFabricAmount").value = get(COL.TOTAL_FABRIC_AMOUNT);
  document.getElementById("cuttingMaster").value = get(COL.CUTTING_MASTER);
  document.getElementById("styleCode").value = get(COL.STYLE);
  document.getElementById("fabLength").value = get(COL.FAB_LENGTH);
  document.getElementById("fabWidth").value = get(COL.FAB_WIDTH);
  document.getElementById("fabGSM").value = get(COL.GSM);
  document.getElementById("numPieces").value = get(COL.NUM_PIECES);
  document.getElementById("jobWorkDesc").value = get(COL.JOB_DESC);
  document.getElementById("jobWorkCost").value = get(COL.JOB_COST);
  document.getElementById("fixedCost").value = get(COL.FIXED_COST);
  document.getElementById("accessoriesCost").value = get(COL.ACCESSORIES_COST);
  document.getElementById("mrpPercent").value = get(COL.MRP_PCT);
  document.getElementById("notes").value = get(COL.NOTES);
  document.getElementById("manualClothUsed").value = get(COL.MANUAL_CLOTH_USED);
  document.getElementById("useManualCloth").checked = get(COL.USE_MANUAL_CLOTH) === "TRUE";
  document.getElementById("cuttingRemarks").value = get(COL.CUTTING_REMARKS);

  fabricImage = { url: v[COL.FABRIC_IMG] || "", fileId: "" };
  designImage = { url: v[COL.DESIGN_IMG] || "", fileId: "" };
  setExistingImagePreview("fabricUploadBox", fabricImage.url);
  setExistingImagePreview("designUploadBox", designImage.url);

  recalc();
  document.getElementById("saveBtn").textContent = "Update costing sheet";
  document.getElementById("editBanner").style.display = "flex";
  document.querySelector('.tab-btn[data-tab="new"]').click();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function cancelEdit() {
  editingRowIndex = null;
  editingId = null;
  editingDate = null;
  document.getElementById("saveBtn").textContent = "Save costing sheet";
  document.getElementById("editBanner").style.display = "none";
  resetForm();
}

/* ---------------- list / view entries ---------------- */
async function loadEntries() {
  const wrap = document.getElementById("listWrap");
  if (!gapi.client.getToken()) {
    wrap.innerHTML = `<div class="empty-state"><h3>Not signed in</h3><p>Sign in with Google to load your styles.</p></div>`;
    return;
  }
  wrap.innerHTML = `<div class="empty-state">Loading…</div>`;
  try {
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.sheetId,
      range: `${SHEET_NAME}!A2:Z`
    });
    const values = res.result.values || [];
    allRows = values.map((v, i) => ({ rowIndex: i + 2, values: v }));
    renderList();
  } catch (err) {
    console.error(err);
    wrap.innerHTML = `<div class="empty-state"><h3>Couldn't load</h3><p>Check your Sheet ID and that you're signed in.</p></div>`;
  }
}

function renderList() {
  const wrap = document.getElementById("listWrap");
  if (allRows.length === 0) {
    wrap.innerHTML = `<div class="empty-state"><h3>No styles yet</h3><p>Add your first costing sheet from the "New Style" tab.</p></div>`;
    return;
  }
  const rows = allRows.map(r => {
    const v = r.values;
    const get = (idx) => v[idx] || "";
    return `<tr>
      <td><img class="thumb" src="${thumbUrl(get(COL.FABRIC_IMG))}" onerror="this.style.visibility='hidden'"></td>
      <td>${get(COL.STYLE) || "—"}</td>
      <td>${get(COL.SUPPLIER) || "—"}</td>
      <td>${get(COL.CUTTING_MASTER) || "—"}</td>
      <td>${get(COL.NUM_PIECES) || "0"}</td>
      <td><span class="mrp-tag">₹${get(COL.FINAL_MRP) || "0.00"}</span></td>
      <td>
        <button class="btn btn-dark btn-small" onclick="viewDetail(${r.rowIndex})">View</button>
        <button class="btn btn-ghost btn-small" style="color:var(--blue);border-color:#DCD5C8;" onclick="editEntry(${r.rowIndex})">Edit</button>
        <button class="btn btn-danger btn-small" onclick="deleteEntry(${r.rowIndex})">Delete</button>
      </td>
    </tr>`;
  }).join("");

  wrap.innerHTML = `<table>
    <thead><tr>
      <th></th><th>Style</th><th>Supplier</th><th>Cutting Master</th><th>Pieces</th><th>Final MRP</th><th></th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function viewDetail(rowIndex) {
  const row = allRows.find(r => r.rowIndex === rowIndex);
  if (!row) return;
  const v = row.values;
  const get = (idx) => v[idx] || "—";
  const html = `
    <h2>${get(COL.STYLE)}</h2>
    <p class="settings-help">${get(COL.ID)} · Added ${get(COL.DATE)}</p>
    <div class="detail-imgs">
      ${v[COL.FABRIC_IMG] ? `<img src="${thumbUrl(v[COL.FABRIC_IMG])}">` : ""}
      ${v[COL.DESIGN_IMG] ? `<img src="${thumbUrl(v[COL.DESIGN_IMG])}">` : ""}
    </div>
    <hr class="stitch">
    <div class="detail-grid">
      <div class="detail-item"><div class="k">Supplier</div><div class="v">${get(COL.SUPPLIER)}</div></div>
      <div class="detail-item"><div class="k">Bill No / Date</div><div class="v">${get(COL.BILL_NO)} · ${get(COL.BILL_DATE)}</div></div>
      <div class="detail-item"><div class="k">Total Fabric Bill Amount</div><div class="v">₹${get(COL.TOTAL_FABRIC_BILL_AMOUNT)}</div></div>
      <div class="detail-item"><div class="k">Fabric Rate</div><div class="v">₹${get(COL.FABRIC_RATE)} per kg/m</div></div>
      <div class="detail-item"><div class="k">Total Fabric Amount</div><div class="v">${get(COL.TOTAL_FABRIC_AMOUNT)} kg/m</div></div>
      <div class="detail-item"><div class="k">Cutting Master</div><div class="v">${get(COL.CUTTING_MASTER)}</div></div>
      <div class="detail-item"><div class="k">Fabric L × W · GSM</div><div class="v">${get(COL.FAB_LENGTH)}in × ${get(COL.FAB_WIDTH)}in · ${get(COL.GSM)}</div></div>
      <div class="detail-item"><div class="k">Pieces</div><div class="v">${get(COL.NUM_PIECES)}</div></div>
      <div class="detail-item"><div class="k">Cloth Used / Piece</div><div class="v">${get(COL.CLOTH_USED)} kg/m${get(COL.USE_MANUAL_CLOTH) === "TRUE" ? " (manual override)" : ""}</div></div>
      <div class="detail-item"><div class="k">Fabric Cost / Piece</div><div class="v">₹${get(COL.FABRIC_COST_PIECE)}</div></div>
      <div class="detail-item"><div class="k">Total Fabric Cost</div><div class="v">₹${get(COL.TOTAL_FABRIC_COST)}</div></div>
      <div class="detail-item"><div class="k">Job Work</div><div class="v">${get(COL.JOB_DESC)} — ₹${get(COL.JOB_COST)}</div></div>
      <div class="detail-item"><div class="k">Fixed Cost</div><div class="v">₹${get(COL.FIXED_COST)}</div></div>
      <div class="detail-item"><div class="k">Accessories Cost</div><div class="v">₹${get(COL.ACCESSORIES_COST)}</div></div>
      <div class="detail-item"><div class="k">MRP %</div><div class="v">${get(COL.MRP_PCT)}%</div></div>
    </div>
    <hr class="stitch">
    <div class="calc-panel">
      <div class="calc-row"><span>Total cost / piece</span><span class="val">₹${get(COL.VARIABLE_COST)}</span></div>
      <div class="calc-row final"><span>Final MRP</span><span class="val">₹${get(COL.FINAL_MRP)}</span></div>
    </div>
    ${v[COL.CUTTING_REMARKS] ? `<p style="margin-top:14px;font-size:13.5px;color:var(--ink-dim);"><strong>Sizes &amp; Ratios:</strong> ${get(COL.CUTTING_REMARKS)}</p>` : ""}
    ${v[COL.NOTES] ? `<p style="margin-top:8px;font-size:13.5px;color:var(--ink-dim);"><strong>Notes:</strong> ${get(COL.NOTES)}</p>` : ""}
  `;
  document.getElementById("detailContent").innerHTML = html;
  document.getElementById("detailModal").classList.add("open");
}
function closeDetail() { document.getElementById("detailModal").classList.remove("open"); }

async function deleteEntry(rowIndex) {
  if (!confirm("Delete this costing sheet? This can't be undone.")) return;
  if (sheetGid === null) await getSheetGid();
  try {
    await gapi.client.sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.sheetId,
      resource: {
        requests: [{
          deleteDimension: {
            range: { sheetId: sheetGid, dimension: "ROWS", startIndex: rowIndex - 1, endIndex: rowIndex }
          }
        }]
      }
    });
    await loadEntries();
  } catch (e) {
    console.error(e);
    alert("Could not delete this row.");
  }
}
