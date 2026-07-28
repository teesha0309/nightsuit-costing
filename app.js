/* ============================================================
   Nightsuit Costing Studio — app logic
   Storage: Google Sheets (data rows) + Google Drive (images)
   Nothing here talks to any server other than Google's own APIs.
   ============================================================ */

const SHEET_NAME = "CostingSheets";
const HEADERS = [
  "ID", "Date Added", "Fabric Image URL", "Design Image URL",
  "Supplier Name", "Bill Number", "Bill Date",
  "Total Fabric Purchased (kg)", "Total Bill Amount", "Fabric Price Per Kg",
  "Cutting Master", "Style Code",
  "Fabric Length (in)", "Fabric Width (in)", "GSM", "Avg Fabric Per Piece (kg)", "Number Of Pieces",
  "Total Avg Cloth Used (kg)", "Fabric Cost Per Piece", "Total Fabric Cost",
  "Job Work Description", "Job Work Cost", "Fixed Cost", "Profit Percent",
  "Variable Cost Per Piece", "Final MRP", "Notes"
];
// column indices (0-based) for readability when reading rows back
const COL = {
  ID:0, DATE:1, FABRIC_IMG:2, DESIGN_IMG:3, SUPPLIER:4, BILL_NO:5, BILL_DATE:6,
  TOTAL_KG:7, TOTAL_BILL:8, PRICE_PER_KG:9, CUTTING_MASTER:10, STYLE:11,
  FAB_LENGTH:12, FAB_WIDTH:13, GSM:14, AVG_PER_PIECE:15, NUM_PIECES:16,
  TOTAL_CLOTH:17, FABRIC_COST_PIECE:18, TOTAL_FABRIC_COST:19,
  JOB_DESC:20, JOB_COST:21, FIXED_COST:22, PROFIT_PCT:23,
  VARIABLE_COST:24, FINAL_MRP:25, NOTES:26
};

let CONFIG = {
  clientId: localStorage.getItem("ncs_clientId") || "",
  apiKey: localStorage.getItem("ncs_apiKey") || "",
  sheetId: localStorage.getItem("ncs_sheetId") || "",
  folderId: localStorage.getItem("ncs_folderId") || ""
};

let tokenClient = null;
let gapiReady = false;
let sheetGid = null; // numeric sheetId of the CostingSheets tab, needed for row deletion
let fabricImage = { url: "", fileId: "" };
let designImage = { url: "", fileId: "" };
let allRows = []; // cache of loaded rows, each {rowIndex, values:[...]}

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
      await ensureHeaders();
      await getSheetGid();
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
      range: `${SHEET_NAME}!A1:AA1`
    });
    if (!res.result.values || res.result.values.length === 0) {
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.sheetId,
        range: `${SHEET_NAME}!A1`,
        valueInputOption: "RAW",
        resource: { values: [HEADERS] }
      });
    }
  } catch (e) {
    console.error("ensureHeaders error", e);
    if (e.status === 400 || (e.result && e.result.error && e.result.error.status === "INVALID_ARGUMENT")) {
      alert(`Could not find a tab named "${SHEET_NAME}" in your sheet. Please create/rename a tab to exactly "${SHEET_NAME}" and reload.`);
    }
  }
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
  "totalFabricKg", "totalBillAmount", "avgPerPiece", "numPieces", "jobWorkCost", "fixedCost", "profitPercent"
];

function wireForm() {
  calcFieldIds.forEach(id => document.getElementById(id).addEventListener("input", recalc));
  document.getElementById("entryForm").addEventListener("submit", handleSave);
  document.getElementById("resetFormBtn").addEventListener("click", resetForm);
  recalc();
}

function num(id) { return parseFloat(document.getElementById(id).value) || 0; }

function recalc() {
  const totalFabricKg = num("totalFabricKg");
  const totalBillAmount = num("totalBillAmount");
  const avgPerPiece = num("avgPerPiece");
  const numPieces = num("numPieces");
  const jobWorkCost = num("jobWorkCost");
  const fixedCost = num("fixedCost");
  const profitPercent = num("profitPercent");

  const pricePerKg = totalFabricKg > 0 ? totalBillAmount / totalFabricKg : 0;
  const totalCloth = avgPerPiece * numPieces;
  const fabricCostPerPiece = pricePerKg * avgPerPiece;
  const totalFabricCost = pricePerKg * totalCloth;
  const variableCostPerPiece = fabricCostPerPiece + jobWorkCost;
  const finalMRP = (variableCostPerPiece + fixedCost) * (1 + profitPercent / 100);

  document.getElementById("calcPricePerKg").textContent = "₹" + pricePerKg.toFixed(2);
  document.getElementById("calcTotalCloth").textContent = totalCloth.toFixed(3) + " kg";
  document.getElementById("calcFabricPerPiece").textContent = "₹" + fabricCostPerPiece.toFixed(2);
  document.getElementById("calcTotalFabric").textContent = "₹" + totalFabricCost.toFixed(2);
  document.getElementById("calcVariable").textContent = "₹" + variableCostPerPiece.toFixed(2);
  document.getElementById("calcMRP").textContent = "₹" + finalMRP.toFixed(2);

  return { pricePerKg, totalCloth, fabricCostPerPiece, totalFabricCost, variableCostPerPiece, finalMRP };
}

/* ---------------- save entry ---------------- */
async function handleSave(e) {
  e.preventDefault();
  if (!gapi.client.getToken()) { alert("Please sign in with Google first."); return; }

  const c = recalc();
  const row = [
    "NS-" + Date.now(),
    new Date().toISOString().slice(0, 10),
    fabricImage.url,
    designImage.url,
    document.getElementById("supplierName").value,
    document.getElementById("billNumber").value,
    document.getElementById("billDate").value,
    num("totalFabricKg"),
    num("totalBillAmount"),
    c.pricePerKg.toFixed(2),
    document.getElementById("cuttingMaster").value,
    document.getElementById("styleCode").value,
    num("fabLength"),
    num("fabWidth"),
    num("fabGSM"),
    num("avgPerPiece"),
    num("numPieces"),
    c.totalCloth.toFixed(3),
    c.fabricCostPerPiece.toFixed(2),
    c.totalFabricCost.toFixed(2),
    document.getElementById("jobWorkDesc").value,
    num("jobWorkCost"),
    num("fixedCost"),
    num("profitPercent"),
    c.variableCostPerPiece.toFixed(2),
    c.finalMRP.toFixed(2),
    document.getElementById("notes").value
  ];

  const saveBtn = document.getElementById("saveBtn");
  saveBtn.disabled = true; saveBtn.textContent = "Saving…";
  try {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: CONFIG.sheetId,
      range: `${SHEET_NAME}!A:A`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [row] }
    });
    saveBtn.textContent = "Saved ✓";
    setTimeout(() => { saveBtn.textContent = "Save costing sheet"; saveBtn.disabled = false; }, 1200);
    resetForm();
    document.querySelector('.tab-btn[data-tab="list"]').click();
  } catch (err) {
    console.error(err);
    alert("Could not save — check your sign-in and sheet ID in Settings.");
    saveBtn.textContent = "Save costing sheet"; saveBtn.disabled = false;
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
      range: `${SHEET_NAME}!A2:AA`
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
      <div class="detail-item"><div class="k">Total Purchased / Bill Amount</div><div class="v">${get(COL.TOTAL_KG)} kg · ₹${get(COL.TOTAL_BILL)}</div></div>
      <div class="detail-item"><div class="k">Fabric Price / kg</div><div class="v">₹${get(COL.PRICE_PER_KG)}</div></div>
      <div class="detail-item"><div class="k">Cutting Master</div><div class="v">${get(COL.CUTTING_MASTER)}</div></div>
      <div class="detail-item"><div class="k">Fabric L × W · GSM</div><div class="v">${get(COL.FAB_LENGTH)}in × ${get(COL.FAB_WIDTH)}in · ${get(COL.GSM)}</div></div>
      <div class="detail-item"><div class="k">Avg / Piece · Pieces</div><div class="v">${get(COL.AVG_PER_PIECE)}kg × ${get(COL.NUM_PIECES)}</div></div>
      <div class="detail-item"><div class="k">Total Cloth Used</div><div class="v">${get(COL.TOTAL_CLOTH)} kg</div></div>
      <div class="detail-item"><div class="k">Fabric Cost / Piece</div><div class="v">₹${get(COL.FABRIC_COST_PIECE)}</div></div>
      <div class="detail-item"><div class="k">Total Fabric Cost</div><div class="v">₹${get(COL.TOTAL_FABRIC_COST)}</div></div>
      <div class="detail-item"><div class="k">Job Work</div><div class="v">${get(COL.JOB_DESC)} — ₹${get(COL.JOB_COST)}</div></div>
      <div class="detail-item"><div class="k">Fixed Cost</div><div class="v">₹${get(COL.FIXED_COST)}</div></div>
      <div class="detail-item"><div class="k">Profit %</div><div class="v">${get(COL.PROFIT_PCT)}%</div></div>
    </div>
    <hr class="stitch">
    <div class="calc-panel">
      <div class="calc-row"><span>Variable cost / piece</span><span class="val">₹${get(COL.VARIABLE_COST)}</span></div>
      <div class="calc-row final"><span>Final MRP</span><span class="val">₹${get(COL.FINAL_MRP)}</span></div>
    </div>
    ${v[COL.NOTES] ? `<p style="margin-top:14px;font-size:13.5px;color:var(--ink-dim);"><strong>Notes:</strong> ${get(COL.NOTES)}</p>` : ""}
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
