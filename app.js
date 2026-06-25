/* ============================================================
   Distribution Solutions — Reorder & Sales Dashboard
   Pure client-side. Integrates two reports:
     1. Apprise ERP "Inventory On Order and Sales" (.xlsx)  -> PO / reorder
     2. Amazon FBA restock / inventory health (.csv)        -> ship-in to FBA
   Matched by stripping the leading "FBA" from FBA SKUs.
   ============================================================ */
"use strict";

const DAYS_PER_MONTH = 30.4;

/* ---------- Global state ---------- */
const STATE = {
  erp: null,             // { products, months, completeMonths, partialMonth, meta }
  fba: null,             // { rows, byKey, meta }
  products: [],          // merged, computed product objects (used by UI)
  months: [], completeMonths: [], partialMonth: null, meta: {},
  hasERP: false, hasFBA: false,
  view: "po",            // po | fba
  overrides: {},         // productCode -> override PO order qty
  fbaOverrides: {},      // productCode -> override ship-to-FBA qty
  expanded: new Set(),   // productCodes with an open detail row
  sort: { key: "urgency", dir: "desc" },
  bound: false,          // persistent control listeners attached once
  // ----- focus list (ERP noise filter) -----
  whitelist: new Set(typeof window !== "undefined" && window.MOVE_FORWARD_CODES ? window.MOVE_FORWARD_CODES : []),
  newItemMin: (typeof window !== "undefined" && window.NEW_ITEM_MIN) ? window.NEW_ITEM_MIN : Infinity,
  whitelistSource: "bundled",   // bundled | uploaded
  showAll: false,        // when true, ignore the focus filter and show every ERP SKU
};

/* A SKU is on the ERP "focus list" if it's a move-forward item, a new item
   (code >= newItemMin), or an FBA item (always kept). */
function isFocus(p) {
  if (p.fba) return true;
  if (STATE.whitelist.has(p.code)) return true;
  const n = parseInt(p.code, 10);
  return isFinite(n) && n >= STATE.newItemMin;
}
function poVisible(p) { return p.hasERP && (STATE.showAll || p.included); }

const ASSUMPTIONS = {
  window: 6, weighting: "even", growth: 0,
  leadTime: 45, coverage: 60, safety: 21, pack: 1,
  fbaCover: 56,          // FBA coverage target days (used when Amazon gives no rec.)
};

/* ---------- Helpers ---------- */
const $ = (sel) => document.querySelector(sel);
const num = (v) => (typeof v === "number" && isFinite(v) ? v : 0);
const tn = (v) => { if (v == null || v === "") return 0; const n = parseFloat(String(v).replace(/,/g, "")); return isFinite(n) ? n : 0; };
const str = (v) => (v == null ? "" : String(v));
const fmt = (n) => (n == null || !isFinite(n) ? "—" : Math.round(n).toLocaleString());
const fmt1 = (n) => (n == null || !isFinite(n) ? "—" : (Math.round(n * 10) / 10).toLocaleString());
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const avg = (arr) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);
const round1 = (n) => Math.round(num(n) * 10) / 10;

/* ============================================================
   FILE LOADING  (accepts ERP .xlsx and FBA .csv, keeps both)
   ============================================================ */
/* Read one file, detect its type, parse it into STATE. Returns a Promise that
   resolves with { kind, count, label, error } — does NOT build the dashboard. */
function ingestFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const probe = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null })[0] || [];
        const H = new Set(probe.map((h) => String(h == null ? "" : h).trim()));
        const isMoveFwd = H.has("Continue Relationship Y/N") || H.has("Closeout Y/N");
        const isFBA = H.has("fnsku") || H.has("fba-inventory-level-health-status") ||
          (H.has("sku") && H.has("asin") && H.has("available"));
        const isERP = !!wb.Sheets["report-data"] || (H.has("Product Code") && H.has("On Hand"));

        if (isMoveFwd) { const n = parseWhitelist(wb); resolve({ kind: "whitelist", count: n, label: "focus list" }); }
        else if (isFBA) { STATE.fba = parseFBA(wb); resolve({ kind: "fba", count: STATE.fba.rows.length, label: "Amazon FBA report" }); }
        else if (isERP) { STATE.erp = parseERP(wb); resolve({ kind: "erp", count: STATE.erp.products.length, label: "Inventory On Order report" }); }
        else resolve({ kind: "unknown", error: "This file wasn't recognized as the inventory or FBA report." });
      } catch (err) { console.error(err); resolve({ kind: "unknown", error: "Could not read this file (" + err.message + ")." }); }
    };
    reader.readAsArrayBuffer(file);
  });
}

/* ----- Upload screen wiring (two slots: FBA + Inventory) ----- */
function setupUploadScreen() {
  const slots = [
    { key: "fba", input: "#fileFba", drop: "#dropFba" },
    { key: "erp", input: "#fileErp", drop: "#dropErp" },
  ];
  slots.forEach((s) => {
    const input = $(s.input), drop = $(s.drop);
    input.addEventListener("change", (e) => { if (e.target.files[0]) handleSlotFile(s.key, e.target.files[0]); input.value = ""; });
    ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("dragover"); }));
    ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("dragover"); }));
    drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) handleSlotFile(s.key, f); });
  });
  $("#submitBtn").addEventListener("click", onSubmitUploads);
  $("#reuploadBtn").addEventListener("click", showUploadScreen);
  updateUploadUI();
}

function handleSlotFile(slotKey, file) {
  const statusEl = slotKey === "fba" ? $("#statusFba") : $("#statusErp");
  statusEl.className = "uc-status reading";
  statusEl.textContent = "Reading “" + file.name + "” …";
  ingestFile(file).then((res) => {
    if (res.error) { statusEl.className = "uc-status err"; statusEl.textContent = "⚠ " + res.error; return; }
    if (res.kind === "whitelist") {
      statusEl.className = "uc-status ok";
      statusEl.innerHTML = "✓ Focus list updated (" + res.count.toLocaleString() + " items). You can still add the report above.";
      updateUploadUI(); return;
    }
    // if the dropped file was actually the other report, gently note it landed in the right place
    if (res.kind !== slotKey) {
      const other = res.kind === "fba" ? $("#statusFba") : $("#statusErp");
      other.dataset.note = "↗ Detected as the " + res.label + " — added to the correct box.";
    }
    updateUploadUI();
  });
}

function fileMeta() {
  const m = {};
  if (STATE.erp) m.erp = STATE.erp.products.length;
  if (STATE.fba) m.fba = STATE.fba.rows.length;
  return m;
}

function updateUploadUI() {
  const setCard = (key, loaded, count, dateStr) => {
    const card = key === "fba" ? $("#cardFba") : $("#cardErp");
    const st = key === "fba" ? $("#statusFba") : $("#statusErp");
    card.classList.toggle("loaded", loaded);
    if (loaded) {
      st.className = "uc-status ok";
      st.innerHTML = "✓ Loaded — <b>" + count.toLocaleString() + "</b> items" + (dateStr ? " <span class='muted'>(" + esc(dateStr) + ")</span>" : "") +
        (st.dataset.note ? "<br><span class='muted'>" + esc(st.dataset.note) + "</span>" : "");
    } else if (st.className.indexOf("reading") === -1 && st.className.indexOf("err") === -1) {
      st.className = "uc-status"; st.textContent = "Not uploaded yet";
    }
  };
  const erpDate = STATE.erp && STATE.erp.meta["printed-date"] ? String(STATE.erp.meta["printed-date"]).split(" ")[0] : "";
  const fbaDate = STATE.fba && STATE.fba.meta.snapshotDate ? String(STATE.fba.meta.snapshotDate) : "";
  setCard("fba", !!STATE.fba, STATE.fba ? STATE.fba.rows.length : 0, fbaDate);
  setCard("erp", !!STATE.erp, STATE.erp ? STATE.erp.products.length : 0, erpDate);

  const n = (STATE.erp ? 1 : 0) + (STATE.fba ? 1 : 0);
  const summary = $("#uploadSummary");
  if (n === 0) summary.innerHTML = "<span class='muted'>Add at least one report to continue. Both is best.</span>";
  else if (n === 2) summary.innerHTML = "<b class='sum-ok'>✓ Both reports ready.</b> Click Submit to view the dashboard.";
  else summary.innerHTML = "<b class='sum-warn'>Only the " + (STATE.erp ? "Inventory" : "FBA") + " report is added.</b> We recommend adding both before you continue.";
  $("#submitBtn").disabled = n === 0;
  $("#uploadConfirm").classList.add("hidden");
}

function onSubmitUploads() {
  const both = STATE.erp && STATE.fba;
  if (!STATE.erp && !STATE.fba) return;
  if (both) return goToDashboard();
  // only one uploaded — ask before continuing
  const have = STATE.erp ? "Inventory On Order" : "Amazon FBA";
  const missing = STATE.erp ? "Amazon FBA" : "Inventory On Order";
  const box = $("#uploadConfirm");
  box.classList.remove("hidden");
  box.innerHTML =
    `<div class="confirm-card">
       <div class="confirm-q">You've only added the <b>${have}</b> report.</div>
       <div class="confirm-sub">The dashboard works best with <b>both</b>. Do you want to add the ${missing} report first?</div>
       <div class="confirm-btns">
         <button class="btn" id="confirmAddOther">Go back &amp; add ${missing}</button>
         <button class="btn btn-accent" id="confirmContinue">Continue with just ${have}</button>
       </div>
     </div>`;
  $("#confirmAddOther").addEventListener("click", () => box.classList.add("hidden"));
  $("#confirmContinue").addEventListener("click", goToDashboard);
}

function goToDashboard() {
  mergeData();
  if (!STATE.products.length) { alert("No product rows found. Please check the file(s)."); return; }
  STATE.view = STATE.hasERP ? "po" : "fba";
  buildUI();
  $("#uploadScreen").classList.add("hidden");
  $("#app").classList.remove("hidden");
}

function showUploadScreen() {
  $("#app").classList.add("hidden");
  $("#uploadScreen").classList.remove("hidden");
  updateUploadUI();
}

// initialize the upload screen as soon as the page is ready
setupUploadScreen();

/* ============================================================
   PARSING — ERP
   ============================================================ */
function parseERP(wb) {
  const meta = {};
  if (wb.Sheets["Extra Information"]) {
    XLSX.utils.sheet_to_json(wb.Sheets["Extra Information"], { header: 1 })
      .forEach((r) => { if (r[0]) meta[r[0]] = r[1]; });
  }
  const dataSheet = wb.Sheets["report-data"] || wb.Sheets[wb.SheetNames[wb.SheetNames.length - 1]];
  const rows = XLSX.utils.sheet_to_json(dataSheet, { header: 1, raw: true, defval: null });
  const header = rows[0].map((h) => (h == null ? "" : String(h).trim()));
  const colOf = (name) => header.indexOf(name);
  const C = {
    code: colOf("Product Code"), name: colOf("Product Name"), desc: colOf("Description"),
    category: colOf("Category Code"), subcat: colOf("Subcategory Code"),
    supCode: colOf("Primary Supplier Code"), supplier: colOf("Supplier Name"),
    onHand: colOf("On Hand"), inTransit: colOf("In Transit"), backorder: colOf("Backordered"),
    onOrder: colOf("On Order"), available: colOf("Available"), discontinued: colOf("Discontinued"),
  };

  const monthCols = [];
  header.forEach((h, i) => { if (/^\d{1,2}\/\d{4}$/.test(h)) monthCols.push(i); });
  const ytdIdx = [];
  header.forEach((h, i) => { if (/Y-?T-?D/i.test(h)) ytdIdx.push(i); });
  const firstYTD = ytdIdx.length ? ytdIdx[0] : Infinity;
  const salesMonthCols = monthCols.filter((i) => i < firstYTD);

  let partialLabel = null;
  const pd = parseDate(meta["printed-date"]);
  if (pd) partialLabel = (pd.getMonth() + 1) + "/" + pd.getFullYear();

  const months = salesMonthCols.map((idx) => ({ label: header[idx], idx }));
  const completeMonths = months.filter((m) => m.label !== partialLabel);

  const products = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const code = row[C.code];
    if (code == null || code === "") continue;
    products.push({
      code: String(code).trim(),
      name: str(row[C.name]), desc: str(row[C.desc]),
      category: row[C.category] != null ? String(row[C.category]) : "Uncat",
      subcat: str(row[C.subcat]),
      supplier: row[C.supplier] != null ? String(row[C.supplier]) : "(no supplier)",
      supCode: str(row[C.supCode]),
      onHand: num(row[C.onHand]), inTransit: num(row[C.inTransit]), backorder: num(row[C.backorder]),
      onOrder: num(row[C.onOrder]), available: num(row[C.available]),
      discontinued: row[C.discontinued] === true || row[C.discontinued] === "True",
      salesByMonth: months.map((m) => num(row[m.idx])),
      completeSales: completeMonths.map((m) => num(row[m.idx])),
      hasERP: true, fba: null,
    });
  }
  return {
    products, months, completeMonths,
    partialMonth: months.some((m) => m.label === partialLabel) ? partialLabel : null,
    meta,
  };
}

/* ============================================================
   PARSING — FBA restock CSV (parsed by SheetJS, mapped by name)
   ============================================================ */
function parseFBA(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const header = rows[0].map((h) => (h == null ? "" : String(h).trim()));
  const idx = (name) => header.indexOf(name);
  const c = {
    sku: idx("sku"), name: idx("product-name"), asin: idx("asin"), fnsku: idx("fnsku"),
    available: idx("available"), inboundWorking: idx("inbound-working"), inboundShipped: idx("inbound-shipped"),
    inboundReceived: idx("inbound-received"), inboundQty: idx("inbound-quantity"), reserved: idx("Total Reserved Quantity"),
    t7: idx("units-shipped-t7"), t30: idx("units-shipped-t30"), t60: idx("units-shipped-t60"), t90: idx("units-shipped-t90"),
    daysSupply: idx("days-of-supply"),
    totalDaysSupply: idx("Total Days of Supply (including units from open shipments)"),
    health: idx("fba-inventory-level-health-status"), recAction: idx("recommended-action"),
    recShipQty: idx("Recommended ship-in quantity"), recShipDate: idx("Recommended ship-in date"),
    price: idx("your-price"), salesPrice: idx("sales-price"),
    excess: idx("estimated-excess-quantity"), minLevel: idx("fba-minimum-inventory-level"),
    sellThrough: idx("sell-through"), salesRank: idx("sales-rank"),
    supplier: idx("supplier"), snapshot: idx("snapshot-date"), weeksCover: idx("weeks-of-cover-t30"),
  };
  const list = [], byKey = new Map();
  let snapshot = null;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const rawSku = row[c.sku];
    if (rawSku == null) continue;
    const key = String(rawSku).replace(/^FBA/i, "").trim();
    if (!key) continue; // skip junk "FBA" / blank rows
    if (!snapshot && c.snapshot >= 0) snapshot = row[c.snapshot];
    const f = {
      key, sku: String(rawSku), productName: str(row[c.name]), asin: str(row[c.asin]), fnsku: str(row[c.fnsku]),
      available: tn(row[c.available]),
      inboundWorking: tn(row[c.inboundWorking]), inboundShipped: tn(row[c.inboundShipped]), inboundReceived: tn(row[c.inboundReceived]),
      inboundQty: tn(row[c.inboundQty]), reserved: tn(row[c.reserved]),
      t7: tn(row[c.t7]), t30: tn(row[c.t30]), t60: tn(row[c.t60]), t90: tn(row[c.t90]),
      daysSupply: tn(row[c.daysSupply]), totalDaysSupply: tn(row[c.totalDaysSupply]),
      health: str(row[c.health]) || "Unknown", recAction: str(row[c.recAction]),
      recShipQty: tn(row[c.recShipQty]), recShipDate: str(row[c.recShipDate]),
      price: tn(row[c.price]) || tn(row[c.salesPrice]),
      excess: tn(row[c.excess]), minLevel: tn(row[c.minLevel]),
      sellThrough: tn(row[c.sellThrough]), salesRank: tn(row[c.salesRank]),
      supplier: str(row[c.supplier]), weeksCover: tn(row[c.weeksCover]),
    };
    f.inboundTotal = f.inboundQty || (f.inboundWorking + f.inboundShipped + f.inboundReceived);
    if (byKey.has(key)) mergeFbaRow(byKey.get(key), f); // same SKU listed twice (e.g. 2 ASINs/conditions)
    else { list.push(f); byKey.set(key, f); }
  }
  return { rows: list, byKey, meta: { snapshotDate: excelOrStr(snapshot), source: "fba" } };
}

/* SheetJS may parse a date cell as an Excel serial number; render it readably. */
function excelOrStr(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d)) return (d.getUTCMonth() + 1) + "/" + d.getUTCDate() + "/" + d.getUTCFullYear();
  }
  return String(v);
}

/* ============================================================
   PARSING — move-forward focus list (updates the ERP whitelist)
   ============================================================ */
function parseWhitelist(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const header = rows[0].map((h) => (h == null ? "" : String(h).trim()));
  const codeCol = header.indexOf("Product Code");
  const set = new Set();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const code = row[codeCol];
    if (code != null && code !== "") set.add(String(code).trim());
  }
  STATE.whitelist = set;
  STATE.whitelistSource = "uploaded";
  return set.size;
}

/* Merge a second FBA listing for the same SKU into the first.
   Sum physical/velocity fields; for Amazon's recommendation take the max (don't
   double-order one ASIN); keep the most-severe health; force days-of-supply to
   recompute from the aggregated available + velocity. */
const HEALTH_RANK = { "Out of stock": 4, "Low stock": 3, "Excess": 2, "Healthy": 1, "Unknown": 0, "": 0 };
function mergeFbaRow(a, b) {
  ["available", "inboundWorking", "inboundShipped", "inboundReceived", "inboundQty", "inboundTotal",
   "reserved", "t7", "t30", "t60", "t90", "excess"].forEach((k) => { a[k] += b[k]; });
  a.recShipQty = Math.max(a.recShipQty, b.recShipQty);
  a.minLevel = Math.max(a.minLevel, b.minLevel);
  a.price = Math.max(a.price, b.price);
  if ((HEALTH_RANK[b.health] || 0) > (HEALTH_RANK[a.health] || 0)) a.health = b.health;
  if (!a.recShipDate && b.recShipDate) a.recShipDate = b.recShipDate;
  a.daysSupply = 0; // recompute from aggregated figures
  a.merged = (a.merged || 1) + 1;
}

function parseDate(s) { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; }

/* ============================================================
   MERGE  — attach FBA to ERP products by matched key; keep FBA-only too
   ============================================================ */
function mergeData() {
  STATE.hasERP = !!STATE.erp;
  STATE.hasFBA = !!STATE.fba;
  const products = [];
  const erpKeys = new Set();

  if (STATE.erp) {
    STATE.erp.products.forEach((p) => {
      p.hasERP = true;
      p.fba = STATE.fba ? (STATE.fba.byKey.get(p.code) || null) : null;
      erpKeys.add(p.code);
      products.push(p);
    });
  }
  if (STATE.fba) {
    STATE.fba.rows.forEach((f) => {
      if (erpKeys.has(f.key)) return; // already attached
      products.push({           // FBA-only pseudo product
        code: f.key, name: f.productName, desc: "", category: "Amazon FBA", subcat: "",
        supplier: f.supplier || "(Amazon FBA only)", supCode: "",
        onHand: 0, inTransit: 0, backorder: 0, onOrder: 0, available: 0, discontinued: false,
        salesByMonth: STATE.erp ? STATE.erp.months.map(() => 0) : [],
        completeSales: STATE.erp ? STATE.erp.completeMonths.map(() => 0) : [],
        hasERP: false, fba: f,
      });
    });
  }
  // flag each product against the ERP focus list (FBA always kept)
  products.forEach((p) => { p.included = isFocus(p); });

  STATE.products = products;
  STATE.months = STATE.erp ? STATE.erp.months : [];
  STATE.completeMonths = STATE.erp ? STATE.erp.completeMonths : [];
  STATE.partialMonth = STATE.erp ? STATE.erp.partialMonth : null;
  STATE.meta = STATE.erp ? STATE.erp.meta : (STATE.fba ? STATE.fba.meta : {});
}

/* ============================================================
   COMPUTATION
   ============================================================ */
function computeAll() {
  const A = ASSUMPTIONS;
  const win = Math.min(A.window, STATE.completeMonths.length || A.window);

  STATE.products.forEach((p) => {
    // ---------- Warehouse / PO metrics (ERP) ----------
    if (p.hasERP) {
      const cs = p.completeSales, n = cs.length;
      const recent = cs.slice(Math.max(0, n - win)).map((v) => Math.max(0, v));
      let velMonthly;
      if (A.weighting === "recent" && recent.length > 1) {
        let wsum = 0, vsum = 0;
        recent.forEach((v, i) => { const w = i + 1; wsum += w; vsum += v * w; });
        velMonthly = vsum / wsum;
      } else velMonthly = avg(recent);

      const effDaily = (velMonthly / DAYS_PER_MONTH) * (1 + A.growth / 100);
      const effMonthly = velMonthly * (1 + A.growth / 100);
      const last3 = avg(cs.slice(Math.max(0, n - 3)));
      const prior3 = avg(cs.slice(Math.max(0, n - 6), Math.max(0, n - 3)));
      let trendPct = null;
      if (prior3 > 0) trendPct = (last3 - prior3) / prior3;
      else if (last3 > 0) trendPct = 5;
      if (trendPct != null) trendPct = Math.max(-1, Math.min(5, trendPct));

      const incoming = p.onOrder + p.inTransit;
      const position = p.available + incoming;
      const coverDays = effDaily > 0 ? p.available / effDaily : Infinity;
      const coverDaysPos = effDaily > 0 ? position / effDaily : Infinity;
      const targetUnits = effDaily * (A.leadTime + A.coverage + A.safety);
      let suggested = 0;
      if (!p.discontinued && effDaily > 0) {
        suggested = Math.max(0, Math.ceil(targetUnits - position));
        if (A.pack > 1 && suggested > 0) suggested = Math.ceil(suggested / A.pack) * A.pack;
      }
      let status;
      if (p.discontinued) status = "Discontinued";
      else if (velMonthly <= 0 && p.onHand > 0) status = "Dead stock";
      else if (velMonthly <= 0) status = "No sales";
      else if (p.available <= 0) status = "Stockout";
      else if (coverDays < A.leadTime) status = "Critical";
      else if (coverDays < A.leadTime + A.coverage) status = "Reorder soon";
      else if (coverDays > 365) status = "Overstock";
      else status = "Healthy";
      let urgency = 0;
      if (effDaily > 0) {
        urgency = Math.max(0, (A.leadTime + A.safety) - coverDaysPos) * effDaily;
        if (p.available <= 0) urgency += 1e6 + effMonthly;
      }
      Object.assign(p, {
        velMonthly, effMonthly, effDaily, incoming, position, coverDays, coverDaysPos,
        proj30: effDaily * 30, proj60: effDaily * 60, proj90: effDaily * 90,
        suggested, status, urgency, trendPct,
        ytdNet: cs.reduce((a, b) => a + b, 0),
        last12: cs.slice(Math.max(0, n - 12)).reduce((a, b) => a + b, 0),
      });
    } else {
      Object.assign(p, {
        velMonthly: 0, effMonthly: 0, effDaily: 0, incoming: 0, position: 0,
        coverDays: Infinity, coverDaysPos: Infinity, proj30: 0, proj60: 0, proj90: 0,
        suggested: 0, status: "Amazon only", urgency: 0, trendPct: null, ytdNet: 0, last12: 0,
      });
    }

    // ---------- FBA ship-in metrics ----------
    if (p.fba) {
      const f = p.fba;
      const velM = f.t30 > 0 ? f.t30 : (f.t90 > 0 ? f.t90 / 3 : 0);
      const velD = velM / DAYS_PER_MONTH;
      const fbaPosition = f.available + f.inboundTotal;
      const coverDays = f.daysSupply > 0 ? f.daysSupply : (velD > 0 ? f.available / velD : (f.available > 0 ? Infinity : 0));
      let shipSug = f.recShipQty > 0 ? f.recShipQty
        : (velD > 0 ? Math.max(0, Math.ceil(velD * A.fbaCover - fbaPosition)) : 0);
      const whseAvail = p.hasERP ? Math.max(0, p.available) : null;
      const shipNow = whseAvail == null ? shipSug : Math.min(shipSug, whseAvail);
      const shortfall = whseAvail == null ? 0 : Math.max(0, shipSug - whseAvail);

      let action;
      if (f.available <= 0 && (velM > 0 || f.recShipQty > 0)) action = "FBA out";
      else if (shipSug > 0 && shortfall > 0) action = "Ship + reorder";
      else if (shipSug > 0) action = "Ship to FBA";
      else if (f.health === "Excess" || f.excess > 0) action = "FBA excess";
      else action = "FBA OK";

      let fbaUrgency = shipSug;
      if (f.available <= 0 && velM > 0) fbaUrgency += 1e6 + velM;

      p.fbaCalc = { velM, velD, fbaPosition, coverDays, shipSug, whseAvail, shipNow, shortfall, action, fbaUrgency };
    } else {
      p.fbaCalc = null;
    }
  });
}

function finalQty(p) {
  return Object.prototype.hasOwnProperty.call(STATE.overrides, p.code) ? STATE.overrides[p.code] : p.suggested;
}
function finalShip(p) {
  if (!p.fbaCalc) return 0;
  return Object.prototype.hasOwnProperty.call(STATE.fbaOverrides, p.code) ? STATE.fbaOverrides[p.code] : p.fbaCalc.shipNow;
}

/* ============================================================
   UI BUILD
   ============================================================ */
function buildUI() {
  const m = STATE.meta, bits = [];
  if (m["company-name"]) bits.push(esc(m["company-name"]));
  if (STATE.hasERP) {
    const focus = STATE.products.filter((p) => p.hasERP && p.included).length;
    bits.push(`ERP ✓ ${focus.toLocaleString()} focus SKUs <span class="muted">of ${STATE.erp.products.length.toLocaleString()}</span>`);
  }
  if (STATE.hasFBA) bits.push(`FBA ✓ ${STATE.fba.rows.length.toLocaleString()} SKUs`);
  bits.push(`Focus list: ${STATE.whitelistSource === "uploaded" ? "uploaded" : "bundled"} (${STATE.whitelist.size.toLocaleString()})`);
  if (m["printed-date"]) bits.push("ERP " + esc(String(m["printed-date"]).split(" ")[0]));
  if (STATE.hasFBA && STATE.fba.meta.snapshotDate) bits.push("FBA " + esc(String(STATE.fba.meta.snapshotDate)));
  $("#reportMeta").innerHTML = bits.join(" &nbsp;•&nbsp; ");

  buildViewToggle();
  buildAssumptions();
  buildFilters();
  bindControls();
  refresh();
}

function buildViewToggle() {
  const wrap = $("#viewToggle");
  if (!(STATE.hasERP && STATE.hasFBA)) { wrap.innerHTML = ""; wrap.classList.add("hidden"); }
  else {
    wrap.classList.remove("hidden");
    wrap.innerHTML =
      `<button class="seg-btn ${STATE.view === "po" ? "active" : ""}" data-view="po">📦 Reorder / PO &nbsp;<small>supplier → warehouse</small></button>` +
      `<button class="seg-btn ${STATE.view === "fba" ? "active" : ""}" data-view="fba">🚚 FBA restock &nbsp;<small>warehouse → Amazon</small></button>`;
    wrap.querySelectorAll(".seg-btn").forEach((b) => b.addEventListener("click", () => {
      STATE.view = b.dataset.view;
      STATE.sort = STATE.view === "po" ? { key: "urgency", dir: "desc" } : { key: "fbaUrgency", dir: "desc" };
      buildViewToggle(); buildFilters(); $("#search").value = ""; refresh();
    }));
  }
}

function buildAssumptions() {
  const A = ASSUMPTIONS;
  const winNote = STATE.partialMonth ? `Current month (${STATE.partialMonth}) is partial and excluded from velocity.` : "";
  const fbaField = STATE.hasFBA ? `
    <div class="assume">
      <label>FBA coverage target (days)</label>
      <input id="aFbaCover" type="number" min="0" step="7" value="${A.fbaCover}" />
      <span class="hint">Used when Amazon gives no ship-in qty</span>
    </div>` : "";
  $("#assumptions").innerHTML = `
    <div class="assume"><label>Velocity window</label>
      <select id="aWindow">
        <option value="3"${A.window==3?" selected":""}>Last 3 months</option>
        <option value="6"${A.window==6?" selected":""}>Last 6 months</option>
        <option value="12"${A.window==12?" selected":""}>Last 12 months</option>
      </select><span class="hint">${winNote}</span></div>
    <div class="assume"><label>Weighting</label>
      <select id="aWeight">
        <option value="even"${A.weighting==="even"?" selected":""}>Even average</option>
        <option value="recent"${A.weighting==="recent"?" selected":""}>Weight recent higher</option>
      </select></div>
    <div class="assume"><label>Growth adjustment (%)</label>
      <input id="aGrowth" type="number" step="5" value="${A.growth}" /><span class="hint">+/- applied to projections</span></div>
    <div class="assume"><label>Supplier lead time (days)</label>
      <input id="aLead" type="number" min="0" step="1" value="${A.leadTime}" /></div>
    <div class="assume"><label>Coverage target (days)</label>
      <input id="aCover" type="number" min="0" step="1" value="${A.coverage}" /><span class="hint">Stock to hold beyond lead time</span></div>
    <div class="assume"><label>Safety stock (days)</label>
      <input id="aSafety" type="number" min="0" step="1" value="${A.safety}" /></div>
    <div class="assume"><label>Round order up to multiple of</label>
      <input id="aPack" type="number" min="1" step="1" value="${A.pack}" /><span class="hint">Case / carton pack</span></div>
    ${fbaField}`;

  const rewire = () => {
    A.window = parseInt($("#aWindow").value, 10);
    A.weighting = $("#aWeight").value;
    A.growth = parseFloat($("#aGrowth").value) || 0;
    A.leadTime = Math.max(0, parseFloat($("#aLead").value) || 0);
    A.coverage = Math.max(0, parseFloat($("#aCover").value) || 0);
    A.safety = Math.max(0, parseFloat($("#aSafety").value) || 0);
    A.pack = Math.max(1, parseInt($("#aPack").value, 10) || 1);
    if ($("#aFbaCover")) A.fbaCover = Math.max(0, parseFloat($("#aFbaCover").value) || 0);
    refresh();
  };
  ["aWindow","aWeight","aGrowth","aLead","aCover","aSafety","aPack","aFbaCover"].forEach((id) => {
    const el = $("#" + id); if (el) el.addEventListener("change", rewire);
  });
}

function buildFilters() {
  const pool = STATE.products.filter((p) => STATE.view === "fba" ? p.fbaCalc : poVisible(p));
  const suppliers = [...new Set(pool.map((p) => p.supplier))].sort();
  const cats = [...new Set(pool.map((p) => p.category))].sort();
  $("#supplierFilter").innerHTML = `<option value="">All suppliers (${suppliers.length})</option>` +
    suppliers.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
  $("#categoryFilter").innerHTML = `<option value="">All categories</option>` +
    cats.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join("");
  const statuses = STATE.view === "fba"
    ? ["FBA out", "Ship + reorder", "Ship to FBA", "FBA excess", "FBA OK"]
    : ["Stockout", "Critical", "Reorder soon", "Healthy", "Overstock", "Dead stock", "No sales", "Discontinued"];
  $("#statusFilter").innerHTML = `<option value="">All statuses</option>` +
    statuses.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");

  // export buttons + legend per view
  const fbaView = STATE.view === "fba";
  $("#exportPO").classList.toggle("hidden", fbaView || !STATE.hasERP);
  $("#exportFBA").classList.toggle("hidden", !fbaView || !STATE.hasFBA);
  $("#showAllWrap").classList.toggle("hidden", fbaView); // focus filter only applies to ERP/PO view
  $("#showAllSkus").checked = STATE.showAll;
  $("#legend").innerHTML = fbaView
    ? `<b>FBA restock:</b> ship from your warehouse into Amazon.
       <span class="badge b-stockout">FBA out</span> out of stock at Amazon ·
       <span class="badge b-reorder">Ship + reorder</span> ship now, warehouse can't fully cover → also cut a PO ·
       <span class="badge b-healthy">Ship to FBA</span> warehouse covers it ·
       <span class="badge b-over">FBA excess</span> too much at Amazon.`
    : `<b>Reorder / PO:</b> order from your supplier into your warehouse. Showing your <b>focus list</b> (move-forward items, new items ≥ ${STATE.newItemMin === Infinity ? "n/a" : STATE.newItemMin.toLocaleString()}, and all FBA items); noise is hidden — tick “Show all SKUs” to reveal everything. Edit any Order Qty before exporting.`;
  $("#footNote").textContent = fbaView
    ? "Ship Qty = min(recommended ship-in, warehouse available). Override before exporting; \"Ship + reorder\" means also raise a PO."
    : "Suggested order = (lead-time + coverage + safety) demand − (available + on order). Override any quantity before exporting.";
}

function bindControls() {
  if (STATE.bound) return; // persistent elements — bind once, even across multiple uploads
  STATE.bound = true;
  ["#search","#supplierFilter","#categoryFilter","#statusFilter","#groupBySupplier","#onlyActionable"].forEach((s) =>
    $(s).addEventListener(s === "#search" ? "input" : "change", renderTable));
  $("#exportPO").addEventListener("click", exportPO);
  $("#exportFBA").addEventListener("click", exportFBA);
  $("#exportAll").addEventListener("click", exportAll);
  $("#showAllSkus").addEventListener("change", (e) => { STATE.showAll = e.target.checked; buildFilters(); refresh(); });

  $("#gridBody").addEventListener("input", (e) => {
    if (!e.target.classList.contains("qty-input")) return;
    const code = e.target.dataset.code, kind = e.target.dataset.kind, v = e.target.value.trim();
    const map = kind === "fba" ? STATE.fbaOverrides : STATE.overrides;
    if (v === "") delete map[code]; else map[code] = Math.max(0, parseInt(v, 10) || 0);
    e.target.classList.toggle("overridden", v !== "");
  });
  $("#gridBody").addEventListener("click", (e) => {
    if (e.target.closest("input, a, button")) return;
    const tr = e.target.closest("tr.data-row");
    if (!tr) return;
    const code = tr.dataset.code;
    STATE.expanded.has(code) ? STATE.expanded.delete(code) : STATE.expanded.add(code);
    renderTable();
  });
}

function refresh() { computeAll(); renderKPIs(); renderInsights(); renderTable(); }

/* ============================================================
   KPIs & INSIGHTS  (view-aware)
   ============================================================ */
function renderKPIs() {
  const P = STATE.products;
  let cards;
  if (STATE.view === "fba") {
    const F = P.filter((p) => p.fbaCalc);
    const out = F.filter((p) => p.fbaCalc.action === "FBA out");
    const ship = F.filter((p) => finalShip(p) > 0);
    const shortfall = F.filter((p) => p.fbaCalc.action === "Ship + reorder");
    const excess = F.filter((p) => p.fbaCalc.action === "FBA excess");
    const units = ship.reduce((a, p) => a + finalShip(p), 0);
    const sold30 = F.reduce((a, p) => a + (p.fba.t30 || 0), 0);
    cards = [
      { v: F.length, l: "FBA SKUs", s: "tracked at Amazon", cls: "" },
      { v: fmt(sold30), l: "FBA units sold (30d)", s: "across all listings", cls: "good" },
      { v: out.length, l: "FBA stockouts", s: "out of stock at Amazon", cls: out.length ? "alert" : "good" },
      { v: ship.length, l: "SKUs to ship in", s: "below FBA coverage", cls: ship.length ? "warn" : "good" },
      { v: fmt(units), l: "Units to ship to FBA", s: "from warehouse", cls: "warn" },
      { v: shortfall.length, l: "Ship + reorder", s: "warehouse can't cover", cls: shortfall.length ? "alert" : "good" },
    ];
  } else {
    const E = P.filter(poVisible);
    const sellable = E.filter((p) => p.velMonthly > 0 && !p.discontinued);
    const stockouts = E.filter((p) => p.status === "Stockout");
    const reorder = E.filter((p) => finalQty(p) > 0);
    const dead = E.filter((p) => p.status === "Dead stock");
    const over = E.filter((p) => p.status === "Overstock");
    cards = [
      { v: E.length.toLocaleString(), l: "Active SKUs", s: sellable.length + " selling", cls: "" },
      { v: fmt(E.reduce((a, p) => a + Math.max(0, p.last12), 0)), l: "Units sold (12 mo)", s: "net of returns", cls: "good" },
      { v: stockouts.length, l: "Stockouts w/ demand", s: "out of stock, still selling", cls: stockouts.length ? "alert" : "good" },
      { v: reorder.length, l: "SKUs to reorder", s: "below target coverage", cls: reorder.length ? "warn" : "good" },
      { v: fmt(reorder.reduce((a, p) => a + finalQty(p), 0)), l: "Units to order", s: "across all suppliers", cls: "warn" },
      { v: over.length + dead.length, l: "Over / dead stock", s: over.length + " over • " + dead.length + " dead", cls: (over.length + dead.length) ? "warn" : "good" },
    ];
  }
  $("#kpis").innerHTML = cards.map((c) => `
    <div class="kpi ${c.cls}"><div class="kpi-value">${c.v}</div>
      <div class="kpi-label">${c.l}</div><div class="kpi-sub">${c.s}</div></div>`).join("");
}

function renderInsights() {
  const P = STATE.products;
  const li = (name, val) => `<li><span class="li-name">${esc(name)}</span><span class="li-val">${val}</span></li>`;
  const card = (title, icon, items, render) => `
    <div class="insight${items.length ? "" : " empty-i"}"><h3>${icon} ${title}</h3>
      ${items.length ? "<ul>" + items.map(render).join("") + "</ul>" : "<p class='muted'>Nothing here — good news.</p>"}</div>`;

  if (STATE.view === "fba") {
    const F = P.filter((p) => p.fbaCalc);
    const out = F.filter((p) => p.fbaCalc.action === "FBA out").sort((a, b) => b.fbaCalc.velM - a.fbaCalc.velM).slice(0, 6);
    const ship = F.filter((p) => p.fbaCalc.action === "Ship to FBA").sort((a, b) => finalShip(b) - finalShip(a)).slice(0, 6);
    const short = F.filter((p) => p.fbaCalc.action === "Ship + reorder").sort((a, b) => b.fbaCalc.shortfall - a.fbaCalc.shortfall).slice(0, 6);
    const excess = F.filter((p) => p.fbaCalc.action === "FBA excess").sort((a, b) => (b.fba.excess || 0) - (a.fba.excess || 0)).slice(0, 6);
    $("#insights").innerHTML =
      card("FBA stockouts (losing sales)", "🚨", out, (p) => li(`${p.code} · ${p.name}`, `${fmt(p.fbaCalc.velM)}/mo`)) +
      card("Ship to FBA now", "🚚", ship, (p) => li(`${p.code} · ${p.name}`, `${fmt(finalShip(p))} u`)) +
      card("Ship + reorder (warehouse short)", "⚠️", short, (p) => li(`${p.code} · ${p.name}`, `short ${fmt(p.fbaCalc.shortfall)}`)) +
      card("FBA excess (consider removal)", "🧊", excess, (p) => li(`${p.code} · ${p.name}`, `${fmt(p.fba.excess)} excess`));
  } else {
    const E = P.filter(poVisible);
    const urgent = E.filter((p) => p.status === "Stockout" || p.status === "Critical").sort((a, b) => b.urgency - a.urgency).slice(0, 6);
    const growing = E.filter((p) => p.velMonthly >= 1 && p.trendPct != null && p.trendPct > 0.15).sort((a, b) => b.trendPct - a.trendPct).slice(0, 6);
    const declining = E.filter((p) => p.last12 > 5 && p.trendPct != null && p.trendPct < -0.15).sort((a, b) => a.trendPct - b.trendPct).slice(0, 6);
    const dead = E.filter((p) => p.status === "Dead stock" && p.onHand > 0).sort((a, b) => b.onHand - a.onHand).slice(0, 6);
    $("#insights").innerHTML =
      card("Most urgent to reorder", "🚨", urgent, (p) => li(`${p.code} · ${p.name}`, `${fmt(finalQty(p))} u`)) +
      card("Fastest growing", "📈", growing, (p) => li(`${p.code} · ${p.name}`, `<span class="trend-up">+${Math.round(p.trendPct*100)}%</span>`)) +
      card("Slowing down", "📉", declining, (p) => li(`${p.code} · ${p.name}`, `<span class="trend-down">${Math.round(p.trendPct*100)}%</span>`)) +
      card("Dead stock (cash tied up)", "🧊", dead, (p) => li(`${p.code} · ${p.name}`, `${fmt(p.onHand)} on hand`));
  }
}

/* ============================================================
   TABLE  (column sets per view; fits screen via fixed layout)
   ============================================================ */
const COLUMNS_PO = [
  { key: "status", label: "Action", left: true, w: 8 },
  { key: "code", label: "Code", left: true, w: 6 },
  { key: "name", label: "Product Name", left: true, w: 20, ellip: true },
  { key: "supplier", label: "Supplier", left: true, w: 12, ellip: true },
  { key: "trend", label: "12-mo trend", left: true, nosort: true, w: 9 },
  { key: "available", label: "Avail", w: 5 },
  { key: "incoming", label: "On Order", w: 6 },
  { key: "velMonthly", label: "Vel/mo", w: 5 },
  { key: "coverDays", label: "Cover", w: 5 },
  { key: "proj30", label: "Proj 30", w: 5 },
  { key: "proj60", label: "Proj 60", w: 5 },
  { key: "proj90", label: "Proj 90", w: 5 },
  { key: "suggested", label: "Suggested", w: 6 },
  { key: "order", label: "Order Qty", nosort: true, w: 8 },
];
const COLUMNS_FBA = [
  { key: "fbaAction", label: "FBA Action", left: true, w: 11 },
  { key: "code", label: "Code", left: true, w: 7 },
  { key: "name", label: "Product Name", left: true, w: 24, ellip: true },
  { key: "fbaAvail", label: "FBA Avail", w: 8 },
  { key: "fbaInbound", label: "Inbound", w: 7 },
  { key: "fbaVel", label: "FBA/mo", w: 7 },
  { key: "fbaCover", label: "Days Supply", w: 9 },
  { key: "fbaHealth", label: "Health", left: true, w: 9, ellip: true },
  { key: "whseAvail", label: "Whse Avail", w: 8 },
  { key: "shipSug", label: "Rec Ship-in", w: 9 },
  { key: "ship", label: "Ship Qty", nosort: true, w: 9 },
];
const activeColumns = () => (STATE.view === "fba" ? COLUMNS_FBA : COLUMNS_PO);

const STATUS_CLASS = {
  "Stockout": "b-stockout", "Critical": "b-critical", "Reorder soon": "b-reorder",
  "Healthy": "b-healthy", "Overstock": "b-over", "Dead stock": "b-dead",
  "No sales": "b-nosales", "Discontinued": "b-disc", "Amazon only": "b-disc",
  "FBA out": "b-stockout", "Ship + reorder": "b-reorder", "Ship to FBA": "b-healthy",
  "FBA excess": "b-over", "FBA OK": "b-nosales",
};
const HEALTH_CLASS = { "Out of stock": "b-stockout", "Low stock": "b-reorder", "Healthy": "b-healthy", "Excess": "b-over" };

function sortVal(p, key) {
  switch (key) {
    case "order": return finalQty(p);
    case "ship": return finalShip(p);
    case "trend": return p.trendPct == null ? -Infinity : p.trendPct;
    case "status": return p.urgency;
    case "fbaAction": return p.fbaCalc ? p.fbaCalc.fbaUrgency : -Infinity;
    case "fbaAvail": return p.fba ? p.fba.available : -Infinity;
    case "fbaInbound": return p.fba ? p.fba.inboundTotal : -Infinity;
    case "fbaVel": return p.fbaCalc ? p.fbaCalc.velM : -Infinity;
    case "fbaCover": return p.fbaCalc ? p.fbaCalc.coverDays : -Infinity;
    case "fbaHealth": return p.fba ? p.fba.health : "";
    case "whseAvail": return p.fbaCalc && p.fbaCalc.whseAvail != null ? p.fbaCalc.whseAvail : -Infinity;
    case "shipSug": return p.fbaCalc ? p.fbaCalc.shipSug : -Infinity;
    case "fbaUrgency": return p.fbaCalc ? p.fbaCalc.fbaUrgency : -Infinity;
    default: return p[key];
  }
}

function getFilteredSorted() {
  const q = $("#search").value.trim().toLowerCase();
  const sup = $("#supplierFilter").value, cat = $("#categoryFilter").value, st = $("#statusFilter").value;
  const onlyAction = $("#onlyActionable").checked;
  const fba = STATE.view === "fba";

  let rows = STATE.products.filter((p) => {
    if (fba ? !p.fbaCalc : !poVisible(p)) return false;
    if (sup && p.supplier !== sup) return false;
    if (cat && p.category !== cat) return false;
    if (st) { if (fba ? p.fbaCalc.action !== st : p.status !== st) return false; }
    if (onlyAction) {
      if (fba) { if (!(finalShip(p) > 0 || p.fbaCalc.action === "FBA out")) return false; }
      else { if (!(finalQty(p) > 0 || p.status === "Stockout" || p.status === "Critical")) return false; }
    }
    if (q) {
      const hay = (p.code + " " + p.name + " " + p.supplier + " " + (p.fba ? p.fba.asin + " " + p.fba.sku : "")).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  const { key, dir } = STATE.sort, mul = dir === "asc" ? 1 : -1;
  rows.sort((a, b) => {
    let av = sortVal(a, key), bv = sortVal(b, key);
    if (typeof av === "string" || typeof bv === "string") return String(av).localeCompare(String(bv)) * mul;
    av = av == null || !isFinite(av) ? -Infinity : av;
    bv = bv == null || !isFinite(bv) ? -Infinity : bv;
    return (av - bv) * mul;
  });
  return rows;
}

function renderTable() {
  const cols = activeColumns();
  $("#gridCols").innerHTML = cols.map((c) => `<col style="width:${c.w}%" />`).join("");
  const arrow = (k) => STATE.sort.key === k ? `<span class="arrow">${STATE.sort.dir === "asc" ? "▲" : "▼"}</span>` : "";
  $("#gridHead").innerHTML = "<tr>" + cols.map((c) =>
    `<th class="${c.left ? "left" : ""}" data-key="${c.key}" ${c.nosort ? "" : 'data-sortable="1"'}>${c.label} ${c.nosort ? "" : arrow(c.key)}</th>`
  ).join("") + "</tr>";
  $("#gridHead").querySelectorAll("th[data-sortable]").forEach((th) => th.addEventListener("click", () => {
    const k = th.dataset.key;
    if (STATE.sort.key === k) STATE.sort.dir = STATE.sort.dir === "asc" ? "desc" : "asc";
    else { STATE.sort.key = k; STATE.sort.dir = (k === "code" || k === "name" || k === "supplier" || k === "fbaHealth") ? "asc" : "desc"; }
    renderTable();
  }));

  const rows = getFilteredSorted();
  const groupBy = $("#groupBySupplier").checked;
  const body = $("#gridBody"), empty = $("#emptyState");
  if (!rows.length) { body.innerHTML = ""; empty.classList.remove("hidden"); }
  else empty.classList.add("hidden");
  const rowOf = STATE.view === "fba" ? rowHTML_FBA : rowHTML_PO;

  let html = "";
  if (groupBy) {
    const groups = {};
    rows.forEach((p) => { (groups[p.supplier] = groups[p.supplier] || []).push(p); });
    Object.keys(groups).sort().forEach((sup) => {
      const g = groups[sup];
      const units = g.reduce((a, p) => a + (STATE.view === "fba" ? finalShip(p) : finalQty(p)), 0);
      const word = STATE.view === "fba" ? "units to ship" : "units to order";
      html += `<tr class="group-row"><td colspan="${cols.length}">${esc(sup)}
        <span class="grp-sub">${g.length} SKUs • ${fmt(units)} ${word}</span></td></tr>`;
      g.forEach((p) => { html += rowOf(p) + detailHTML(p); });
    });
  } else {
    rows.forEach((p) => { html += rowOf(p) + detailHTML(p); });
  }
  body.innerHTML = html;

  const total = STATE.products.filter((p) => STATE.view === "fba" ? p.fbaCalc : poVisible(p)).length;
  const noiseNote = (STATE.view !== "fba" && !STATE.showAll && STATE.hasERP)
    ? ` <span class="muted">(focus list — ${(STATE.products.filter((p) => p.hasERP).length - total).toLocaleString()} noise SKUs hidden)</span>` : "";
  $("#rowCount").innerHTML = `${rows.length.toLocaleString()} of ${total.toLocaleString()} SKUs shown${noiseNote}`;
}

function trCommon(p) {
  return `class="data-row${STATE.expanded.has(p.code) ? " open" : ""}" data-code="${esc(p.code)}"`;
}
function caret(p) { return `<span class="caret">${STATE.expanded.has(p.code) ? "▾" : "▸"}</span>`; }

function rowHTML_PO(p) {
  const cover = p.coverDays === Infinity ? "∞" : fmt(p.coverDays);
  const ov = finalQty(p), overridden = Object.prototype.hasOwnProperty.call(STATE.overrides, p.code);
  const fbaTag = p.fba ? ` <span class="mini-tag" title="Also sold on Amazon FBA">FBA</span>` : "";
  return `<tr ${trCommon(p)}>
    <td class="left">${caret(p)}<span class="badge ${STATUS_CLASS[p.status]}">${p.status}</span></td>
    <td class="left pc">${esc(p.code)}</td>
    <td class="left ellip" title="${esc(p.name)}">${esc(p.name)}${fbaTag}</td>
    <td class="left ellip" title="${esc(p.supplier)}">${esc(p.supplier)}</td>
    <td class="left">${sparkline(p.salesByMonth)} ${trendBadge(p.trendPct)}</td>
    <td class="num">${fmt(p.available)}</td>
    <td class="num">${p.incoming ? fmt(p.incoming) : '<span class="muted">0</span>'}</td>
    <td class="num">${fmt1(p.velMonthly)}</td>
    <td class="num">${cover}</td>
    <td class="num">${fmt(p.proj30)}</td>
    <td class="num">${fmt(p.proj60)}</td>
    <td class="num">${fmt(p.proj90)}</td>
    <td class="num">${p.suggested ? "<b>" + fmt(p.suggested) + "</b>" : '<span class="muted">0</span>'}</td>
    <td class="num"><input class="qty-input ${overridden ? "overridden" : ""}" type="number" min="0" data-code="${esc(p.code)}" data-kind="po" value="${ov}" /></td>
  </tr>`;
}

function rowHTML_FBA(p) {
  const f = p.fba, fc = p.fbaCalc;
  const cover = fc.coverDays === Infinity ? "∞" : fmt(fc.coverDays);
  const ov = finalShip(p), overridden = Object.prototype.hasOwnProperty.call(STATE.fbaOverrides, p.code);
  const whse = fc.whseAvail == null ? '<span class="muted">n/a</span>'
    : (fc.shortfall > 0 ? `<span class="trend-down">${fmt(fc.whseAvail)}</span>` : fmt(fc.whseAvail));
  const hCls = HEALTH_CLASS[f.health] || "b-nosales";
  return `<tr ${trCommon(p)}>
    <td class="left">${caret(p)}<span class="badge ${STATUS_CLASS[fc.action]}">${fc.action}</span></td>
    <td class="left pc">${esc(p.code)}</td>
    <td class="left ellip" title="${esc(p.name)}">${esc(p.name)}</td>
    <td class="num">${f.available <= 0 ? '<span class="trend-down">0</span>' : fmt(f.available)}</td>
    <td class="num">${f.inboundTotal ? fmt(f.inboundTotal) : '<span class="muted">0</span>'}</td>
    <td class="num">${fmt(fc.velM)}</td>
    <td class="num">${cover}</td>
    <td class="left"><span class="badge ${hCls}">${esc(f.health)}</span></td>
    <td class="num">${whse}</td>
    <td class="num">${fc.shipSug ? "<b>" + fmt(fc.shipSug) + "</b>" : '<span class="muted">0</span>'}</td>
    <td class="num"><input class="qty-input ${overridden ? "overridden" : ""}" type="number" min="0" data-code="${esc(p.code)}" data-kind="fba" value="${ov}" /></td>
  </tr>`;
}

function detailHTML(p) {
  if (!STATE.expanded.has(p.code)) return "";
  const cols = activeColumns().length;
  const sections = [];

  // ----- warehouse monthly chart (ERP only) -----
  if (p.hasERP && STATE.months.length) {
    const vals = p.salesByMonth, max = Math.max(1, ...vals.map((v) => Math.abs(v)));
    const peak = Math.max(...vals);
    const bars = STATE.months.map((m, i) => {
      const v = vals[i], isPartial = m.label === STATE.partialMonth;
      const cls = v < 0 ? "neg" : (isPartial ? "partial" : (v === peak && peak > 0 ? "peak" : ""));
      return `<div class="mbar-row"><span class="mbar-label">${esc(m.label)}${isPartial ? " <span class='muted'>·part</span>" : ""}</span>
        <span class="mbar-track"><span class="mbar-fill ${cls}" style="width:${(Math.max(0, Math.abs(v) / max) * 100).toFixed(1)}%"></span></span>
        <span class="mbar-val">${v.toLocaleString()}</span></div>`;
    }).join("");
    sections.push(`<div class="detail-section chart-section">
      <h4>📊 Warehouse monthly units sold</h4><div class="mchart">${bars}</div></div>`);
  }

  const stat = (l, v, hl) => `<div class="dstat${hl ? " hl" : ""}"><span class="dstat-l">${l}</span><span class="dstat-v">${v}</span></div>`;

  // ----- warehouse / PO stats -----
  if (p.hasERP) {
    sections.push(`<div class="detail-section"><h4>📦 Warehouse / PO</h4><div class="dlist">${[
      stat("On hand", fmt(p.onHand)), stat("Available", fmt(p.available)),
      stat("On order", fmt(p.onOrder)), stat("In transit", fmt(p.inTransit)),
      stat("Velocity / mo", fmt1(p.velMonthly)), stat("Days of cover", p.coverDays === Infinity ? "∞" : fmt(p.coverDays)),
      stat("12-mo net units", fmt(p.ytdNet)), stat("Trend (3 vs prior 3)", trendBadge(p.trendPct)),
      stat("Supplier", esc(p.supplier)),
      stat("Suggested order", fmt(p.suggested), true),
    ].join("")}</div></div>`);
  }

  // ----- amazon fba stats -----
  if (p.fba) {
    const f = p.fba, fc = p.fbaCalc;
    sections.push(`<div class="detail-section"><h4>🚚 Amazon FBA${p.hasERP ? "" : " · Amazon-only SKU"}</h4><div class="dlist">${[
      stat("FBA SKU", esc(f.sku)), stat("ASIN", esc(f.asin || "—")),
      stat("FBA available", fmt(f.available)), stat("Inbound", fmt(f.inboundTotal)),
      stat("Reserved", fmt(f.reserved)),
      stat("Health", `<span class="badge ${HEALTH_CLASS[f.health] || "b-nosales"}">${esc(f.health)}</span>`),
      stat("Sold 7 / 30 / 90d", `${fmt(f.t7)} / ${fmt(f.t30)} / ${fmt(f.t90)}`),
      stat("Days of supply", fc.coverDays === Infinity ? "∞" : fmt(fc.coverDays)),
      stat("Amazon rec. ship-in", fmt(f.recShipQty) + (f.recShipDate ? " <span class='muted'>by " + esc(f.recShipDate) + "</span>" : "")),
      stat("Warehouse can cover?", fc.whseAvail == null ? "n/a" : (fc.shortfall > 0 ? `<span class="trend-down">No — short ${fmt(fc.shortfall)}</span>` : `<span class="trend-up">Yes</span>`)),
      stat("Amazon price", f.price ? "$" + f.price.toFixed(2) : "—"),
      stat("Excess units", fmt(f.excess)), stat("Sales rank", f.salesRank ? "#" + fmt(f.salesRank) : "—"),
      stat("Suggested ship qty", fmt(fc.shipSug), true),
    ].join("")}</div></div>`);
  }

  return `<tr class="detail-row"><td colspan="${cols}">
    <div class="detail-title">${esc(p.code)} · ${esc(p.name)}</div>
    <div class="detail-grid">${sections.join("")}</div>
  </td></tr>`;
}

function trendBadge(pct) {
  if (pct == null) return '<span class="trend-flat">·</span>';
  const v = Math.round(pct * 100);
  if (v > 15) return `<span class="trend-up">▲ ${v}%</span>`;
  if (v < -15) return `<span class="trend-down">▼ ${v}%</span>`;
  return `<span class="trend-flat">≈</span>`;
}
function sparkline(vals) {
  if (!vals || !vals.length) return "";
  const v = vals.map((x) => Math.max(0, x)), w = 64, h = 16, max = Math.max(1, ...v);
  const step = v.length > 1 ? w / (v.length - 1) : w;
  const pts = v.map((x, i) => `${(i * step).toFixed(1)},${(h - (x / max) * (h - 2) - 1).toFixed(1)}`).join(" ");
  return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><polyline fill="none" stroke="var(--blue)" stroke-width="1.5" points="${pts}" /></svg>`;
}

/* ============================================================
   EXPORTS
   ============================================================ */
function timestamp() {
  const d = STATE.meta["printed-date"] ? new Date(STATE.meta["printed-date"]) : new Date();
  const p = (x) => String(x).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function exportPO() {
  const rows = STATE.products.filter((p) => poVisible(p) && finalQty(p) > 0)
    .sort((a, b) => a.supplier.localeCompare(b.supplier) || b.urgency - a.urgency);
  if (!rows.length) { alert("No items currently have an order quantity greater than 0."); return; }
  const header = ["Supplier", "Supplier Code", "Product Code", "Product Name", "Category", "On Hand", "Available",
    "On Order", "Vel/mo", "Cover Days", "Proj 90d", "Also on FBA?", "Suggested Qty", "Order Qty"];
  const aoa = [header];
  rows.forEach((p) => aoa.push([
    p.supplier, p.supCode, p.code, p.name, p.category, Math.round(p.onHand), Math.round(p.available),
    Math.round(p.incoming), round1(p.velMonthly), p.coverDays === Infinity ? "" : Math.round(p.coverDays),
    Math.round(p.proj90), p.fba ? "Yes" : "", p.suggested, finalQty(p),
  ]));
  const sumMap = {};
  rows.forEach((p) => { const s = (sumMap[p.supplier] = sumMap[p.supplier] || { skus: 0, units: 0, code: p.supCode }); s.skus++; s.units += finalQty(p); });
  const sumAoa = [["Supplier", "Supplier Code", "SKUs to Order", "Total Units"]];
  Object.keys(sumMap).sort().forEach((s) => sumAoa.push([s, sumMap[s].code, sumMap[s].skus, sumMap[s].units]));
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet(aoa);
  ws1["!cols"] = [{wch:22},{wch:12},{wch:14},{wch:38},{wch:12},{wch:8},{wch:9},{wch:9},{wch:7},{wch:9},{wch:8},{wch:9},{wch:11},{wch:9}];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sumAoa), "Supplier Summary");
  XLSX.utils.book_append_sheet(wb, ws1, "Reorder POs");
  XLSX.writeFile(wb, `Reorder_PO_${timestamp()}.xlsx`);
}

function exportFBA() {
  const rows = STATE.products.filter((p) => p.fbaCalc && finalShip(p) > 0)
    .sort((a, b) => b.fbaCalc.fbaUrgency - a.fbaCalc.fbaUrgency);
  if (!rows.length) { alert("No items currently have a ship-to-FBA quantity greater than 0."); return; }
  const header = ["Product Code", "FBA SKU", "ASIN", "Product Name", "FBA Available", "Inbound", "FBA Sold 30d",
    "Days of Supply", "Health", "Warehouse Available", "Amazon Rec Ship-in", "Suggested Ship", "Ship Qty",
    "Warehouse Short?", "Amazon Rec Date", "Action"];
  const aoa = [header];
  rows.forEach((p) => { const f = p.fba, fc = p.fbaCalc; aoa.push([
    p.code, f.sku, f.asin, p.name, Math.round(f.available), Math.round(f.inboundTotal), Math.round(f.t30),
    fc.coverDays === Infinity ? "" : Math.round(fc.coverDays), f.health,
    fc.whseAvail == null ? "" : Math.round(fc.whseAvail), Math.round(f.recShipQty), fc.shipSug, finalShip(p),
    fc.shortfall > 0 ? "SHORT " + Math.round(fc.shortfall) : "", f.recShipDate, fc.action,
  ]); });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{wch:14},{wch:14},{wch:12},{wch:40},{wch:11},{wch:8},{wch:10},{wch:11},{wch:11},{wch:14},{wch:14},{wch:11},{wch:9},{wch:14},{wch:14},{wch:14}];
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: header.length - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws, "FBA Ship-in");
  XLSX.writeFile(wb, `Ship_to_FBA_${timestamp()}.xlsx`);
}

function exportAll() {
  const monthLabels = STATE.months.map((m) => m.label);
  const head = ["Status", "Product Code", "Product Name", "Description", "Supplier", "Supplier Code", "Category",
    "On Hand", "Available", "On Order", "In Transit", "Vel/mo", "Trend %", "Cover Days", "Proj 30d", "Proj 60d",
    "Proj 90d", "Suggested Qty", "Order Qty", "12-mo Net Units"];
  const fbaHead = STATE.hasFBA ? ["On FBA?", "FBA Available", "FBA Inbound", "FBA Sold 30d", "FBA Days Supply",
    "FBA Health", "FBA Action", "Suggested Ship", "Ship Qty"] : [];
  const header = [...head, ...fbaHead, ...monthLabels];
  const aoa = [header];
  const rows = STATE.products.filter((p) => poVisible(p) || p.fba)
    .sort((a, b) => (b.urgency || 0) + (b.fbaCalc ? b.fbaCalc.fbaUrgency : 0) - ((a.urgency || 0) + (a.fbaCalc ? a.fbaCalc.fbaUrgency : 0)));
  rows.forEach((p) => {
    const base = [p.status, p.code, p.name, p.desc, p.supplier, p.supCode, p.category,
      Math.round(p.onHand), Math.round(p.available), Math.round(p.onOrder), Math.round(p.inTransit),
      round1(p.velMonthly), p.trendPct == null ? "" : Math.round(p.trendPct * 100),
      p.coverDays === Infinity ? "" : Math.round(p.coverDays), Math.round(p.proj30), Math.round(p.proj60),
      Math.round(p.proj90), p.suggested, finalQty(p), Math.round(p.ytdNet)];
    const fbaCells = STATE.hasFBA ? (p.fba ? [
      "Yes", Math.round(p.fba.available), Math.round(p.fba.inboundTotal), Math.round(p.fba.t30),
      p.fbaCalc.coverDays === Infinity ? "" : Math.round(p.fbaCalc.coverDays), p.fba.health, p.fbaCalc.action,
      p.fbaCalc.shipSug, finalShip(p),
    ] : ["", "", "", "", "", "", "", "", ""]) : [];
    aoa.push([...base, ...fbaCells, ...p.salesByMonth]);
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: header.length - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws, "Analysis");
  XLSX.writeFile(wb, `Inventory_Analysis_${timestamp()}.xlsx`);
}
