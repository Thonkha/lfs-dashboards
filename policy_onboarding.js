// --------------------------
// Global Data
// --------------------------
let mainData = [];   // full dataset
let filteredData = []; // currently displayed data
let activeMonth = null; // track selected month
let activeWeek = null;  // track selected week

// --------------------------
// Load Google Sheet
// --------------------------
window.onload = loadGoogleSheetData;

async function loadGoogleSheetData() {
    const sheetId = "1NLBYEfipHHvdpyLPrVfhqybRgFLlWEqcUkYwqWe8bUI"; 
    const apiKey = "AIzaSyDjgYbFOH0uMwniKeiM6lZ5BJbOcupeM74";         
    const range = "Sheet1";                
    const errorDiv = document.getElementById("errorMessage");

    try {
        const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?key=${apiKey}`);
        const result = await response.json();

        if (!result.values || result.values.length === 0) {
            console.warn("No data found in Google Sheet.");
            errorDiv.textContent = "⚠️ No data found in Google Sheet. Please upload Excel file as backup.";
            errorDiv.style.display = "block";
            return;
        }

        errorDiv.textContent = "";
        errorDiv.style.display = "none";

        const headers = result.values[0];
        const rows = result.values.slice(1);

        const jsonData = rows.map(row => {
            let obj = {};
            headers.forEach((h, i) => {
                obj[h.trim().toUpperCase()] = row[i] || "";
            });
            return obj;
        });

        console.log("✅ Google Sheet data loaded successfully.");
        updateDashboard(jsonData, true);

    } catch (err) {
        console.error("Failed to load Google Sheet.", err);
        errorDiv.textContent = "⚠️ Failed to load Google Sheet. Please upload Excel file as backup.";
        errorDiv.style.display = "block";
    }
}

// --------------------------
// Excel Upload
// --------------------------
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet);

        jsonData = jsonData.map(row => {
            const newRow = {};
            for (let key in row) {
                newRow[key.trim().toUpperCase()] = row[key];
            }
            return newRow;
        });

        const errorDiv = document.getElementById("errorMessage");
        errorDiv.textContent = "";
        errorDiv.style.display = "none";

        console.log("✅ Excel data loaded successfully.");
        updateDashboard(jsonData, true);
    };
    reader.readAsArrayBuffer(file);
}
document.getElementById("fileInput").addEventListener("change", handleExcelUpload);

// --------------------------
// Reset Filter Button
// --------------------------
document.getElementById('resetFilter')?.addEventListener('click', () => {
    activeMonth = null;
    activeWeek = null;
    updateDashboard(mainData, false);
});

// --------------------------
// Apply Filters Together
// --------------------------
function applyFilters() {
    let filtered = [...mainData];

    if (activeMonth) {
        filtered = filtered.filter(d => d.MONTH === parseInt(activeMonth));
    }
    if (activeWeek) {
        filtered = filtered.filter(d => d.WEEK === parseInt(activeWeek));
    }

    return filtered;
}

// --------------------------
// Dashboard Updater
// --------------------------
function updateDashboard(data, saveMain = true) {
    if (saveMain) mainData = data;
    filteredData = applyFilters();

    // If no filters, use original data
    if (filteredData.length === 0) filteredData = data;

    // Normalize
    filteredData.forEach(d => {
        d.MONTH = parseInt(d.MONTH) || 0;
        d.WEEK = parseInt(d.WEEK) || 0;
        d.DAY = parseInt(d.DAY) || 0;
        d.ACTION = (d.ACTION || "").trim();
        d.STATUS = (d.STATUS || "").trim();
        d["PLAN TYPE"] = (d["PLAN TYPE"] || "").trim();
        d.REGION = (d.REGION || "").trim();
        d.BRANCH = (d.BRANCH || "").trim();
        d["CAPTURED BY"] = (d["CAPTURED BY"] || "").trim();
    });

    // === KPIs ===
    const totalPolicies = filteredData.length;
    const NewlyCapturedPolicies = filteredData.filter(d => d.ACTION.toUpperCase() === "NEW").length;
    const NewPolicies = filteredData.filter(d => 
        d.ACTION?.toUpperCase() === "NEW" && d.STATUS?.toUpperCase() === "ON TRIAL"
    ).length;
    const existedPolicies = filteredData.filter(d => 
        d.ACTION?.toUpperCase() === "NEW" && d.STATUS?.toUpperCase() === "ACTIVE"
    ).length;
    const upgrades = filteredData.filter(d => d.ACTION.toUpperCase() === "UPGRADE").length;
    const downgrades = filteredData.filter(d => d.ACTION.toUpperCase() === "DOWNGRADE").length;
    const activeCount = filteredData.filter(d => d.STATUS.toUpperCase() === "ACTIVE").length;
    const ontrial = filteredData.filter(d => d.STATUS.toUpperCase() === "ON TRIAL").length;
    const conversionRate = totalPolicies ? ((activeCount / ontrial) * 100).toFixed(1) : 0;
    const cancelled = filteredData.filter(d => d.STATUS.toUpperCase() === "CANCELLED").length;
  
    document.getElementById("totalNewPolicies").textContent = NewlyCapturedPolicies;
    document.getElementById("newCount").textContent = NewPolicies;
    document.getElementById("Upgrades").textContent = upgrades;
    document.getElementById("Downgrades").textContent = downgrades;
    document.getElementById("activeCount").textContent = existedPolicies;
    document.getElementById("conversionRate").textContent = conversionRate + "%";
    document.getElementById("cancelledCount").textContent = cancelled;

    // === Weekly Trend ===
    const weekActionMap = {};
    filteredData.forEach(d => {
        const week = d.WEEK;
        const action = d.ACTION;
        if (!weekActionMap[action]) {
            weekActionMap[action] = {};
        }
        weekActionMap[action][week] = (weekActionMap[action][week] || 0) + 1;
    });

    const traces = Object.keys(weekActionMap).map(action => ({
        x: Object.keys(weekActionMap[action]),
        y: Object.values(weekActionMap[action]),
        type: "scatter",
        mode: "lines+markers",
        name: action
    }));

    Plotly.react("weeklyTrend", traces, {
        title: "Weekly Trend by ACTION",
        margin: { t: 30 },
        xaxis: { title: "Week" },
        yaxis: { title: "Count" }
    });

    document.getElementById("weeklyTrend").on('plotly_click', function(event){
        activeWeek = event.points[0].x;
        filteredData = applyFilters();
        updateDashboard(mainData, false);
    });

    // === Monthly Trend ===
    const monthMap = {};
    filteredData.forEach(d => monthMap[d.MONTH] = (monthMap[d.MONTH] || 0) + 1);
    Plotly.react("monthlyTrend", [{
        x: Object.keys(monthMap),
        y: Object.values(monthMap),
        type: "bar",
        marker: { color: "#800000" }
    }], { title: "Monthly New Policies", margin: { t: 30 } });

    document.getElementById("monthlyTrend").on('plotly_click', function(event){
        activeMonth = event.points[0].x;
        filteredData = applyFilters();
        updateDashboard(mainData, false);
    });
  

    // === Plan Type Distribution ===
    const planMap = {};
    data.forEach(d => planMap[d["PLAN TYPE"]] = (planMap[d["PLAN TYPE"]] || 0) + 1);
    Plotly.react("planTypeDist", [{
        x: Object.keys(planMap),
        y: Object.values(planMap),
        type: "bar",
        marker: { color: "#FFD700" }
    }], { title: "Plan Type Distribution", margin: { t: 30 } });

    document.getElementById("planTypeDist").on('plotly_click', function(event){
        const clickedPlan = event.points[0].x;
        const filtered = mainData.filter(d => d["PLAN TYPE"] === clickedPlan);
        updateDashboard(filtered, false);
    });

    // === Status Split ===
    const statusMap = {};
    data.forEach(d => statusMap[d.STATUS] = (statusMap[d.STATUS] || 0) + 1);
    Plotly.react("statusSplit", [{
        labels: Object.keys(statusMap),
        values: Object.values(statusMap),
        type: "pie"
    }], { title: "Status Split", margin: { t: 30 } });

    // === Regional Uptake ===
    const regionMap = {};
    data.forEach(d => regionMap[d.REGION] = (regionMap[d.REGION] || 0) + 1);
    Plotly.react("regionUptake", [{
        x: Object.values(regionMap),
        y: Object.keys(regionMap),
        type: "bar",
        orientation: "h",
        marker: { color: "#800000" }
    }], { title: "Regional Uptake", margin: { t: 30 } });

    document.getElementById("regionUptake").on('plotly_click', function(event){
        const clickedRegion = event.points[0].y;
        const filtered = mainData.filter(d => d.REGION === clickedRegion);
        updateDashboard(filtered, false);
    });

    // === Branch Leaderboard ===
    const branchMap = {};
    data.forEach(d => branchMap[d.BRANCH] = (branchMap[d.BRANCH] || 0) + 1);
    const branchSorted = Object.entries(branchMap).sort((a,b)=>b[1]-a[1]);
    Plotly.react("branchLeaderboard", [{
        x: branchSorted.map(d=>d[0]),
        y: branchSorted.map(d=>d[1]),
        type: "bar",
        marker: { color: "#FFD700" }
    }], { title: "Branch Leaderboard", margin: { t: 30 } });

    document.getElementById("branchLeaderboard").on('plotly_click', function(event){
        const clickedBranch = event.points[0].x;
        const filtered = mainData.filter(d => d.BRANCH === clickedBranch);
        updateDashboard(filtered, false);
    });

    // === Captured By Ranking ===
    const capMap = {};
    data.forEach(d => capMap[d["CAPTURED BY"]] = (capMap[d["CAPTURED BY"]] || 0) + 1);
    const capSorted = Object.entries(capMap).sort((a,b)=>b[1]-a[1]);
    Plotly.react("capturedByRank", [{
        x: capSorted.map(d=>d[0]),
        y: capSorted.map(d=>d[1]),
        type: "bar",
        marker: { color: "#800000" }
    }], { title: "Captured By Ranking", margin: { t: 30 } });

    document.getElementById("capturedByRank").on('plotly_click', function(event){
        const clickedUser = event.points[0].x;
        const filtered = mainData.filter(d => d["CAPTURED BY"] === clickedUser);
        updateDashboard(filtered, false);
    });
}
