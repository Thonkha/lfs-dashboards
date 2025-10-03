// -------------------- CONFIG --------------------
// If you want auto Google Sheets loading, set these two variables.
// Otherwise leave empty and use Excel upload as backup.
const SHEET_ID = "1j2sFV5nT-Ij0rwDHa4k2pX34V0dQQ0-iY_2NVagpXac";    // e.g. "1AbC..."; leave empty to skip sheet fetch
const API_KEY  = "AIzaSyDjgYbFOH0uMwniKeiM6lZ5BJbOcupeM74";    // Google API key (optional)

// Column names
const COL = {
  BRANCH: "BRANCH",
  DATE_OUT: "DATE OUT",
  CORPSE: "CORPSE NAME",
  GENDER: "GENDER",
  DATE_IN: "DATE IN",
  DEST: "DESTINATION",
  COFFIN: "COFFIN CODE",
  SERVICE: "SERVICE TYPE",
  TIME_OUT: "TIME OUT",
  SERVICE_TIME: "SERVICE TIME",
  MODE_PAY: "MODE OF PAYMENT",
  PAY_CAT: "PAYMENT CATEGORY",
  TOTAL: "TOTAL AMOUNT",
  PRENEED: "PRENEED",
  CASH: "CASH",
  SOCIETY: "SOCIETY"
};

// -------------------- GLOBALS --------------------
let mainData = [];        // original dataset
let filters = {           // UI-driven filters
  dateStart: null,
  dateEnd: null,
  branch: "All",
  payment: "All",
  service: "All"
};

// DOM references
const errorDiv = document.getElementById("errorMessage");
const branchSel = document.getElementById("branchFilter");
const paySel = document.getElementById("paymentFilter");
const serviceSel = document.getElementById("serviceFilter");
const dateStartEl = document.getElementById("dateStart");
const dateEndEl = document.getElementById("dateEnd");
const resetBtn = document.getElementById("resetFilter");

// -------------------- UTIL HELPERS --------------------
function parseDateFlexible(s) {
  if (!s && s !== 0) return null;
  let str = String(s).trim();
  // if already Date
  if (str instanceof Date) return str;
  // try ISO / mm/dd/yyyy first
  let d = new Date(str);
  if (!isNaN(d)) return d;
  // try dd/mm/yyyy
  const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2],10), c = parseInt(m[3],10);
    // if first >12, assume dd/mm/yyyy
    if (a > 12) return new Date(`${b}/${a}/${c}`);
    // otherwise treat as mm/dd/yyyy
    return new Date(`${a}/${b}/${c}`);
  }
  return null;
}

function parseTimeToHourMin(t) {
  if (!t) return null;
  const s = String(t).trim();
  // Accept hh:mm or hh:mm AM/PM
  const m = s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
  if (!m) return null;
  let hh = parseInt(m[1],10);
  const mm = parseInt(m[2],10);
  const ampm = (m[3]||"").toUpperCase();
  if (ampm === "PM" && hh < 12) hh += 12;
  if (ampm === "AM" && hh === 12) hh = 0;
  return { hour: hh, min: mm };
}

function safeNum(v) {
  if (v == null || v === "") return 0;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g,'')); // remove commas, currency
  return isNaN(n) ? 0 : n;
}

function formatMoney(v) {
  return "M " + (v || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// -------------------- DATA LOADING --------------------
async function loadFromGoogleSheet() {
  if (!SHEET_ID || !API_KEY) return null;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1?key=${API_KEY}`;
  try {
    const res = await fetch(url);
    const json = await res.json();
    if (!json.values) throw new Error("No values returned");
    const headers = json.values[0].map(h => h.trim().toUpperCase());
    const rows = json.values.slice(1);
    const data = rows.map(r => {
      const obj = {};
      headers.forEach((h,i) => obj[h] = r[i] ?? "");
      return obj;
    });
    document.getElementById("dataSource").textContent = "Google Sheets (auto)";
    return data;
  } catch (err) {
    console.warn("Google Sheets load failed:", err);
    return null;
  }
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    const arr = new Uint8Array(e.target.result);
    const wb = XLSX.read(arr, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "" });
    // normalize keys
    json = json.map(row => {
      const o = {};
      Object.keys(row).forEach(k => o[k.trim().toUpperCase()] = row[k]);
      return o;
    });
    document.getElementById("dataSource").textContent = "Excel (uploaded)";
    initializeDashboard(json);
  };
  reader.readAsArrayBuffer(file);
}

// -------------------- INITIALIZE --------------------
async function init() {
  // attach handlers
  document.getElementById("fileInput").addEventListener("change", handleExcelUpload);
  dateStartEl.addEventListener("change", () => { filters.dateStart = dateStartEl.value ? new Date(dateStartEl.value) : null; refreshFromFilters(); });
  dateEndEl.addEventListener("change", () => { filters.dateEnd = dateEndEl.value ? new Date(dateEndEl.value) : null; refreshFromFilters(); });
  branchSel.addEventListener("change", () => { filters.branch = branchSel.value; refreshFromFilters(); });
  paySel.addEventListener("change", () => { filters.payment = paySel.value; refreshFromFilters(); });
  serviceSel.addEventListener("change", () => { filters.service = serviceSel.value; refreshFromFilters(); });
  resetBtn.addEventListener("click", () => {
    filters = { dateStart: null, dateEnd: null, branch: "All", payment: "All", service: "All" };
    dateStartEl.value = ""; dateEndEl.value = "";
    branchSel.value = "All"; paySel.value = "All"; serviceSel.value = "All";
    refreshFromFilters();
  });

  // try Google Sheets first
  const sheetData = await loadFromGoogleSheet();
  if (sheetData && sheetData.length) {
    initializeDashboard(sheetData);
  } else {
    // wait for manual Excel upload
    errorDiv.style.display = "block";
    errorDiv.textContent = "Google Sheets not loaded. Please upload Excel file as backup.";
  }
}

// -------------------- NORMALIZE & INDEX --------------------
function normalizeAndIndex(raw) {
  // map to canonical columns (case-insensitive)
  const data = raw.map(r => {
    const obj = {};
    // copy keys as uppercase (already may be uppercase)
    Object.keys(r).forEach(k => obj[k.trim().toUpperCase()] = r[k]);
    // helpers: parse dates/times/numbers
    obj._DATE_OUT = parseDateFlexible(obj[COL.DATE_OUT]);
    obj._DATE_IN = parseDateFlexible(obj[COL.DATE_IN]);
    obj._STAY_DAYS = (obj._DATE_OUT && obj._DATE_IN) ? Math.max(0, Math.round((obj._DATE_OUT - obj._DATE_IN) / (1000*60*60*24))) : null;
    const toTime = parseTimeToHourMin(obj[COL.TIME_OUT]);
    obj._TIME_OUT_H = toTime ? toTime.hour : null;
    obj._SERVICE_TIME = parseTimeToHourMin(obj[COL.SERVICE_TIME]);
    if (obj._SERVICE_TIME && obj._TIME_OUT_H != null) {
      obj._SERVICE_DURATION_H = (obj._SERVICE_TIME.hour + obj._SERVICE_TIME.min/60) - obj._TIME_OUT_H;
    } else obj._SERVICE_DURATION_H = null;
    obj._TOTAL = safeNum(obj[COL.TOTAL]);
    obj._PRENEED = safeNum(obj[COL.PRENEED]);
    obj._CASH = safeNum(obj[COL.CASH]);
    obj._PAYMENT_MODE = (obj[COL.MODE_PAY] || "").toString().trim().toUpperCase();
    obj._PAY_CAT = (obj[COL.PAY_CAT] || "").toString().trim().toUpperCase();
    obj._BRANCH = (obj[COL.BRANCH] || "").toString().trim();
    obj._SERVICE = (obj[COL.SERVICE] || "").toString().trim();
    obj._COFFIN = (obj[COL.COFFIN] || "").toString().trim();
    obj._DEST = (obj[COL.DEST] || "").toString().trim();
    obj._GENDER = (obj[COL.GENDER] || "").toString().trim().toUpperCase();
    return obj;
  });
  return data;
}

// -------------------- KPI CALCULATIONS --------------------
function computeKPIs(data) {
  const k = {};
  k.totalDispatches = data.length;
  k.avgStay = data.reduce((s,d)=> s + (d._STAY_DAYS || 0),0) / (data.length || 1);
  k.avgStay = Math.round(k.avgStay*10)/10;
  k.dispatchByService = {};
  data.forEach(d => { const s = d._SERVICE || "Unknown"; k.dispatchByService[s] = (k.dispatchByService[s]||0)+1; });
  k.coffinCounts = {};
  data.forEach(d => { const c = d._COFFIN || "Unknown"; k.coffinCounts[c] = (k.coffinCounts[c]||0)+1; });
  k.paymentMix = {};
  data.forEach(d => { const p = d._PAYMENT_MODE || "UNKNOWN"; k.paymentMix[p] = (k.paymentMix[p]||0)+1; });
  k.totalRevenue = data.reduce((s,d)=> s + (d._TOTAL||0), 0);
  k.totalPreneed = data.reduce((s,d)=> s + (d._PRENEED||0), 0);
  k.totalCash = data.reduce((s,d)=> s + (d._CASH||0), 0);
  k.avgRevenuePerCase = data.length ? k.totalRevenue / data.length : 0;
  k.missingTotals = data.filter(d => (!d[COL.TOTAL] || safeNum(d[COL.TOTAL])===0)).length;
  k.afCount = data.filter(d => (d._PAYMENT_MODE || "").includes("A/F") || (d._PAYMENT_MODE || "").includes("AF")).length;
  k.turnaroundWithin7 = data.filter(d => (d._STAY_DAYS != null && d._STAY_DAYS <= 7)).length;
  k.turnaroundWithin7Pct = data.length ? Math.round(1000 * k.turnaroundWithin7 / data.length) / 10 : 0;
  // peak hours
  k.hourCounts = {};
  data.forEach(d => { const h = d._TIME_OUT_H; if (h!=null) k.hourCounts[h] = (k.hourCounts[h]||0)+1; });
  // branch stats
  k.branchStats = {};
  data.forEach(d => {
    const b = d._BRANCH || "Unknown";
    if (!k.branchStats[b]) k.branchStats[b] = { count:0, staySum:0 };
    k.branchStats[b].count++;
    k.branchStats[b].staySum += (d._STAY_DAYS || 0);
  });
  // compute average stay per branch
  Object.keys(k.branchStats).forEach(b => {
    const o = k.branchStats[b];
    o.avgStay = o.count ? Math.round( (o.staySum / o.count) *10)/10 : 0;
  });
  return k;
}

// -------------------- RENDERING --------------------
function renderKPIs(k) {
  const container = document.getElementById("kpis");
  container.innerHTML = `
    <div class="card"><h3>Total Dispatches</h3><div class="metric-value">${k.totalDispatches}</div></div>
    <div class="card"><h3>Average Stay (days)</h3><div class="metric-value">${k.avgStay}</div></div>
    <div class="card"><h3>Total Revenue</h3><div class="metric-value">${formatMoney(k.totalRevenue)}</div></div>
    <div class="card"><h3>Avg Revenue / Case</h3><div class="metric-value">${formatMoney(k.avgRevenuePerCase)}</div></div>
    <div class="card"><h3>Preneed Total</h3><div class="metric-value">${formatMoney(k.totalPreneed)}</div></div>
    <div class="card"><h3>Missing Totals</h3><div class="metric-value">${k.missingTotals}</div></div>
    <div class="card"><h3>AF / After-Funeral Cases</h3><div class="metric-value">${k.afCount}</div></div>
    <div class="card"><h3>Turnaround ≤ 7 days</h3><div class="metric-value">${k.turnaroundWithin7Pct}%</div></div>
    
  `;
}

function renderTimeSeries(data) {
  // group by date out (YYYY-MM-DD)
  const map = {};
  data.forEach(d => {
    if (!d._DATE_OUT) return;
    const key = d._DATE_OUT.toISOString().slice(0,10);
    map[key] = (map[key]||0) + 1;
  });
  const x = Object.keys(map).sort();
  const y = x.map(k => map[k]);
  Plotly.react("dispatchTimeSeries", [{ x, y, type: "scatter", mode:"lines+markers", line:{color:"#800000"} }], { title:"Dispatches Over Time", margin:{t:30} });
}

function renderAvgStayByBranch(k) {
  const labels = Object.keys(k.branchStats);
  const values = labels.map(l => k.branchStats[l].avgStay);
  Plotly.react("avgStayByBranch", [{ x: labels, y: values, type:"bar", marker:{color:"#FFD700"} }], { title:"Average Stay (days) by Branch", margin:{t:30} });
}

function renderServiceTypePie(k) {
  const labels = Object.keys(k.dispatchByService);
  const values = labels.map(l => k.dispatchByService[l]);
  Plotly.react("serviceTypePie", [{ labels, values, type:"pie", hole:0.35 }], { title:"Dispatches by Service Type", margin:{t:30} });
}

function renderCoffinTop(k) {
  const items = Object.entries(k.coffinCounts).sort((a,b)=>b[1]-a[1]).slice(0,12);
  const labels = items.map(i=>i[0]);
  const values = items.map(i=>i[1]);
  Plotly.react("coffinTop", [{ x: values, y: labels, orientation:"h", type:"bar", marker:{color:"#800000"} }], { title:"Top Coffin Types", margin:{t:30} });
}

function renderPaymentMix(k) {
  const labels = Object.keys(k.paymentMix);
  const values = labels.map(l=>k.paymentMix[l]);
  Plotly.react("paymentMix", [{ labels, values, type:"pie" }], { title:"Payment Mode Mix", margin:{t:30} });
}

function renderRevenueTrend(data) {
  const map = {};
  data.forEach(d => {
    if (!d._DATE_OUT) return;
    const key = d._DATE_OUT.toISOString().slice(0,7); // month
    map[key] = (map[key]||0) + (d._TOTAL||0);
  });
  const x = Object.keys(map).sort();
  const y = x.map(k => map[k]);
  Plotly.react("revenueTrend", [{ x, y, type:"bar", marker:{color:"#d4af37"} }], { title:"Revenue by Month", margin:{t:30} });
}

function renderHourlyHeat(k) {
  // hourCounts -> convert to array 0..23
  const hours = Array.from({length:24}, (_,i) => i);
  const counts = hours.map(h => k.hourCounts[h] || 0);
  // plot as bar (heat-like)
  Plotly.react("hourlyHeat", [{
    x: hours.map(h => `${h}:00`),
    y: counts,
    type: "bar",
    marker: {color: counts.map(v=> v>0 ? "#800000" : "#eee")}
  }], { title: "Dispatches by Hour (Time Out)", margin:{t:30} });
}

function renderDispatchTable(data) {
  const wrap = document.createElement("div");
  wrap.className = "table-wrap";
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headers = ["BRANCH","DATE OUT","CORPSE NAME","GENDER","DATE IN","DESTINATION","COFFIN","SERVICE","TIME OUT","PAYMENT","TOTAL"];
  thead.innerHTML = "<tr>" + headers.map(h=>`<th>${h}</th>`).join("") + "</tr>";
  const tbody = document.createElement("tbody");
  data.slice(0,500).forEach(r => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${r._BRANCH||''}</td>
      <td>${r._DATE_OUT ? r._DATE_OUT.toISOString().slice(0,10) : ''}</td>
      <td>${r[COL.CORPSE] || ''}</td>
      <td>${r._GENDER||''}</td>
      <td>${r._DATE_IN ? r._DATE_IN.toISOString().slice(0,10) : ''}</td>
      <td>${r._DEST||''}</td>
      <td>${r._COFFIN||''}</td>
      <td>${r._SERVICE||''}</td>
      <td>${r._TIME_OUT_H!=null? r._TIME_OUT_H + ":00" : ''}</td>
      <td>${r._PAYMENT_MODE || ''}</td>
      <td>${r._TOTAL? formatMoney(r._TOTAL): ''}</td>
    `;
    tbody.appendChild(row);
  });
  table.appendChild(thead); table.appendChild(tbody); wrap.appendChild(table);
  const container = document.getElementById("dispatchTable");
  container.innerHTML = "";
  container.appendChild(wrap);
}

// -------------------- FILTERS & REFRESH --------------------
function refreshFromFilters() {
  // reapply filters to mainData and re-render
  const filtered = mainData.filter(r => {
    // dateOut filter
    if (filters.dateStart) {
      if (!r._DATE_OUT || r._DATE_OUT < filters.dateStart) return false;
    }
    if (filters.dateEnd) {
      // include whole day of end
      const endOfDay = new Date(filters.dateEnd); endOfDay.setHours(23,59,59,999);
      if (!r._DATE_OUT || r._DATE_OUT > endOfDay) return false;
    }
    // branch
    if (filters.branch && filters.branch !== "All") {
      if ((r._BRANCH || "") !== filters.branch) return false;
    }
    // payment
    if (filters.payment && filters.payment !== "All") {
      if ((r._PAYMENT_MODE || "") !== filters.payment) return false;
    }
    // service
    if (filters.service && filters.service !== "All") {
      if ((r._SERVICE || "") !== filters.service) return false;
    }
    return true;
  });
  // re-render from filtered
  renderAllFromData(filtered);
}

function renderAllFromData(raw) {
  const data = normalizeAndIndex(raw);
  filteredData = data; // currently displayed
  const k = computeKPIs(data);
  renderKPIs(k);
  renderTimeSeries(data);
  renderAvgStayByBranch(k);
  renderServiceTypePie(k);
  renderCoffinTop(k);
  renderPaymentMix(k);
  renderRevenueTrend(data);
  renderHourlyHeat(k);
  renderDispatchTable(data);
  // update filter selects dynamically
  populateFilterSelects();
}

function populateFilterSelects() {
  // populate branch, payment, service selects from mainData
  const branches = new Set(); const pays = new Set(); const services = new Set();
  mainData.forEach(r => {
    const br = (r[COL.BRANCH]||"").toString().trim();
    if (br) branches.add(br);
    const pm = (r[COL.MODE_PAY]||"").toString().trim().toUpperCase();
    if (pm) pays.add(pm);
    const sv = (r[COL.SERVICE]||"").toString().trim();
    if (sv) services.add(sv);
  });
  // helper to populate
  function setOptions(el, items) {
    const current = el.value;
    el.innerHTML = `<option value="All">All</option>` + Array.from(items).sort().map(x=>`<option value="${x}">${x}</option>`).join("");
    if ([...items].includes(current)) el.value = current;
  }
  setOptions(branchSel, branches);
  setOptions(paySel, pays);
  setOptions(serviceSel, services);
}

// -------------------- ENTRY POINT --------------------
function initializeDashboard(rawData) {
  errorDiv.style.display = "none";
  if (!rawData || !rawData.length) {
    errorDiv.textContent = "No data provided.";
    errorDiv.style.display = "block";
    return;
  }
  mainData = rawData.map(r => {
    // ensure keys uppercase
    const obj = {};
    Object.keys(r).forEach(k => obj[k.trim().toUpperCase()] = r[k]);
    return obj;
  });

  // initialize filters empty
  filters = { dateStart: null, dateEnd: null, branch: "All", payment: "All", service: "All" };
  dateStartEl.value = ""; dateEndEl.value = "";
  branchSel.innerHTML = "<option>All</option>";
  paySel.innerHTML = "<option>All</option>";
  serviceSel.innerHTML = "<option>All</option>";

  // run first render
  refreshFromFilters();
}

// try to init
init();
