/* Updated script.js — robust date handling and properly-sorted time-series for Plotly */

const fileInput = document.getElementById('fileInput');
fileInput.addEventListener('change', handleFile, false);

/* ----------------- Helpers ----------------- */

function parseCurrency(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let s = String(val).trim();
  s = s.replace(/[^\d.\-]/g, ''); // remove currency letters, commas, spaces
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function toDate(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel serial date -> convert to JS date
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  if (typeof val === 'string') {
    const s = val.trim();
    // mm/dd/yyyy
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(parseInt(m[3],10), parseInt(m[1],10)-1, parseInt(m[2],10));
    // yyyy-mm-dd or yyyy/mm/dd
    m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m) return new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10));
    // fallback
    const d = new Date(s);
    if (!isNaN(d)) return d;
  }
  return null;
}

/* ISO week helpers */
function getISOWeekYear(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return { year: date.getUTCFullYear(), week: weekNo };
}

function isoWeekMonday(year, week) {
  // returns Date (UTC midnight) of Monday of ISO week
  const simple = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
  const dow = simple.getUTCDay();
  const mon = new Date(simple);
  mon.setUTCDate(simple.getUTCDate() - ((dow + 6) % 7));
  // convert to local Date object for display (so Plotly shows local date)
  return new Date(mon.getUTCFullYear(), mon.getUTCMonth(), mon.getUTCDate());
}

/* Currency formatter for display */
const currencyFormatter = new Intl.NumberFormat('en-ZA', { style: 'currency', currency: 'ZAR' });

/* ----------------- File handling ----------------- */

function handleFile(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const name = file.name.toLowerCase();

  if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      processRows(rows);
    };
    reader.readAsArrayBuffer(file);
  } else {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: function(results) {
        processRows(results.data);
      },
      error: function(err){ alert('CSV parse error: ' + err.message); }
    });
  }
}

/* ----------------- Main processing ----------------- */

function processRows(rows) {
  if (!rows || rows.length === 0) {
    alert('No data found');
    return;
  }

  // normalize header map (lowercase trimmed)
  const headerMap = {};
  Object.keys(rows[0]).forEach(h => headerMap[h.trim().toLowerCase()] = h);

  // map common synonyms
  const pick = (candidates) => {
    for (const c of candidates) {
      if (headerMap[c]) return headerMap[c];
    }
    return null;
  };

  const keyMap = {
    claimNumber: pick(['claim number','claim_no','claimno','claimnumber','claim #']),
    claimDate:   pick(['claim date','date of death','date','date_of_death']),
    coverAmount: pick(['cover amount','cover','coveramount']),
    paidAmount:  pick(['paid amount','paid','paidamount','amount paid']),
    status:      pick(['claim status','status']),
    phase:       pick(['claim phase status','claim phase','phase']),
    cause:       pick(['cause of death','causeofdeath','cause'])
  };

  // normalize & parse rows
  const processed = rows.map(r => {
    const rawDateVal = keyMap.claimDate ? r[keyMap.claimDate] : (r['Claim Date'] || r['Date']);
    const dt = toDate(rawDateVal);
    return {
      claimNumber: keyMap.claimNumber ? r[keyMap.claimNumber] : (r['Claim Number'] || ''),
      dateRaw: rawDateVal,
      date: dt,
      cover: parseCurrency(keyMap.coverAmount ? r[keyMap.coverAmount] : (r['Cover Amount'] || r['Cover'] || 0)),
      paid:  parseCurrency(keyMap.paidAmount ? r[keyMap.paidAmount] : (r['Paid Amount'] || r['Paid'] || 0)),
      status: (keyMap.status ? r[keyMap.status] : (r['Claim Status'] || r['Status'] || 'Unknown')) || 'Unknown',
      phase:  (keyMap.phase ? r[keyMap.phase] : (r['Claim Phase Status'] || r['Phase'] || 'Unknown')) || 'Unknown',
      cause:  (keyMap.cause ? r[keyMap.cause] : (r['Cause of Death'] || r['Cause'] || 'Unknown')) || 'Unknown'
    };
  });

  renderFromProcessed(processed);
}

/* ----------------- Rendering ----------------- */

function renderFromProcessed(data) {
  const rows = data.filter(r => r.claimNumber || r.date);

  // metrics
  const totalClaims = rows.length;
  const totalPaid = rows.reduce((s,r)=> s + (r.paid||0), 0);
  const totalCover = rows.reduce((s,r)=> s + (r.cover||0), 0);
  const paidPct = totalCover > 0 ? (totalPaid/totalCover*100) : 0;

  document.querySelector('#metricTotalClaims .metric-value').innerText = totalClaims;
  document.querySelector('#metricTotalPaid .metric-value').innerText = currencyFormatter.format(totalPaid);
  document.querySelector('#metricTotalCover .metric-value').innerText = currencyFormatter.format(totalCover);
  document.querySelector('#metricPaidPct .metric-value').innerText = paidPct.toFixed(1) + '%';

  // aggregates (use maps keyed by time millis to ensure correct sort)
  const statusMap = new Map();
  const phaseMap = new Map();
  const causeMap  = new Map();
  const monthMap  = new Map(); // key = millis of first day of month
  const weekMap   = new Map(); // key = millis of monday

  rows.forEach(r => {
    statusMap.set(r.status, (statusMap.get(r.status)||0) + 1);
    phaseMap.set(r.phase, (phaseMap.get(r.phase)||0) + 1);
    causeMap.set(r.cause, (causeMap.get(r.cause)||0) + 1);

    if (r.date instanceof Date && !isNaN(r.date)) {
      // month start
      const monthStart = new Date(r.date.getFullYear(), r.date.getMonth(), 1);
      const monthKey = monthStart.getTime();
      if (!monthMap.has(monthKey)) monthMap.set(monthKey, 0);
      monthMap.set(monthKey, monthMap.get(monthKey) + 1);

      // ISO week Monday
      const wy = getISOWeekYear(r.date);
      const mon = isoWeekMonday(wy.year, wy.week);
      const weekKey = mon.getTime();
      if (!weekMap.has(weekKey)) weekMap.set(weekKey, 0);
      weekMap.set(weekKey, weekMap.get(weekKey) + 1);
    }
  });

  // convert maps to sorted arrays
  const statusLabels = Array.from(statusMap.keys());
  const statusValues = statusLabels.map(k => statusMap.get(k));

  const phaseLabels = Array.from(phaseMap.keys());
  const phaseValues = phaseLabels.map(k => phaseMap.get(k));

  const topCausesArr = Array.from(causeMap.entries()).sort((a,b)=> b[1]-a[1]).slice(0,10);
  const causeLabels = topCausesArr.map(e=> e[0]);
  const causeValues = topCausesArr.map(e=> e[1]);

  const monthsArr = Array.from(monthMap.entries()).map(([k,v]) => ({ date: new Date(Number(k)), count: v }));
  monthsArr.sort((a,b)=> a.date - b.date);
  const monthsX = monthsArr.map(m => m.date);
  const monthsY = monthsArr.map(m => m.count);

  const weeksArr = Array.from(weekMap.entries()).map(([k,v]) => ({ date: new Date(Number(k)), count: v }));
  weeksArr.sort((a,b)=> a.date - b.date);
  const weeksX = weeksArr.map(w => w.date);
  const weeksY = weeksArr.map(w => w.count);

  // Main monthly line (large)
  Plotly.react('mainLineChart', [{
    x: monthsX,
    y: monthsY,
    mode: 'lines+markers',
    line: { color: '#800000' },
    marker: { size: 6 }
  }], {
    title: 'Monthly Claims (by Month Start)',
    xaxis: { type: 'date', tickformat: '%b %Y', tickangle: -45 },
    margin: { t: 40 },
    height: 380
  });

  // monthlyTrend (full width)
  Plotly.react('monthlyTrend', [{
    x: monthsX,
    y: monthsY,
    mode: 'lines+markers',
    line: { color: '#FFD700' },
    marker: { size: 6 }
  }], {
    title: 'Monthly Claims Trend',
    xaxis: { type: 'date', tickformat: '%b %Y', tickangle: -45 },
    margin: { t: 40 },
    height: 420
  });

  // weeklyTrend (full width) — show Monday date on axis
  Plotly.react('weeklyTrend', [{
    x: weeksX,
    y: weeksY,
    mode: 'lines+markers',
    line: { color: '#3f000f' },
    marker: { size: 6 }
  }], {
    title: 'Weekly Claims Trend (week start = Monday)',
    xaxis: { type: 'date', tickformat: '%Y-%m-%d', tickangle: -45 },
    margin: { t: 40 },
    height: 420
  });

  // small charts on the right
  Plotly.react('claimStatusPie', [{
    type: 'pie',
    labels: statusLabels,
    values: statusValues,
    hole: 0.36
  }], { title: 'Claims by Status', height: 260 });

  Plotly.react('claimPhaseBar', [{
    type: 'bar',
    x: phaseLabels,
    y: phaseValues,
    marker: { color: '#000000' }
  }], { title: 'Claim Phase Status', height: 260 });

  Plotly.react('topCauses', [{
    type: 'bar',
    x: causeValues,
    y: causeLabels,
    orientation: 'h',
    marker: { color: '#FFD700' }
  }], { title: 'Top Causes (Top 10)', height: 300, margin: { l: 140 } });

  Plotly.react('coverVsPaid', [{
    type: 'bar',
    x: ['Cover Amount', 'Paid Amount'],
    y: [totalCover, totalPaid],
    marker: { color: ['#6c757d', '#800000'] }
  }], { title: 'Cover vs Paid', height: 240 });

  // ensure responsiveness
  setTimeout(()=> window.dispatchEvent(new Event('resize')), 150);
}
