// --------------------------
// Global Data
// --------------------------
let mainData = [];  // stores full dataset
let filteredData = []; // currently displayed data

// --------------------------
// Load Google Sheet on page load
// --------------------------
window.onload = loadGoogleSheetData;

async function loadGoogleSheetData() {
    const sheetId = "1NLBYEfipHHvdpyLPrVfhqybRgFLlWEqcUkYwqWe8bUI"; // replace with actual ID
    const apiKey = "AIzaSyDjgYbFOH0uMwniKeiM6lZ5BJbOcupeM74";         // replace with Google Cloud API key
    const range = "Sheet1";                // sheet tab name
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
// Excel Backup Loader
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
    updateDashboard(mainData, false);
});

// --------------------------
// Dashboard Updater
// --------------------------
function updateDashboard(data, saveMain = true) {
    if (saveMain) mainData = data;
    filteredData = data;

    // Normalize and parse
    data.forEach(d => {
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
    const totalPolicies = data.length;
    const NewPolicies = data.filter(d => d.ACTION.toUpperCase()).length;
    const trialCount = data.filter(d => d.STATUS.toUpperCase() === "ON TRIAL").length;
    const upgrades = data.filter(d => d.ACTION.toUpperCase() === "UPGRADE").length;
    const downgrades = data.filter(d => d.ACTION.toUpperCase() === "DOWNGRADE").length;
    const activeCount = data.filter(d => d.STATUS.toUpperCase() === "ACTIVE").length;
    const conversionRate = totalPolicies ? ((activeCount / totalPolicies) * 100).toFixed(1) : 0;

    document.getElementById("totalNewPolicies").textContent = NewPolicies;
    document.getElementById("trialCount").textContent = trialCount;
    document.getElementById("Upgrades").textContent = upgrades;
    document.getElementById("Downgrades").textContent = downgrades;
    document.getElementById("activeCount").textContent = activeCount;
    document.getElementById("conversionRate").textContent = conversionRate + "%";

// === Weekly Trend by ACTION ===
    const weekActionMap = {};

    // Group data by week and action
    data.forEach(d => {
        const week = d.WEEK;
        const action = d.ACTION;
        if (!weekActionMap[action]) {
            weekActionMap[action] = {};
        }
        weekActionMap[action][week] = (weekActionMap[action][week] || 0) + 1;
    });

    // Prepare traces for Plotly
    const traces = Object.keys(weekActionMap).map(action => {
        return {
            x: Object.keys(weekActionMap[action]),
            y: Object.values(weekActionMap[action]),
            type: "scatter",
            mode: "lines+markers",
            name: action   // label line by ACTION value
        };
    });

    // Plot the chart
    Plotly.react("weeklyTrend", traces, {
        title: "Weekly Trend by ACTION",
        margin: { t: 30 },
        xaxis: { title: "Week" },
        yaxis: { title: "Count" }
    });

    // Weekly interactivity
    document.getElementById("weeklyTrend").on('plotly_click', function(event){
        const clickedWeek = event.points[0].x;
        const filtered = mainData.filter(d => d.WEEK === parseInt(clickedWeek));
        updateDashboard(filtered, false);
    });

    // === Monthly Trend ===
    const monthMap = {};
    data.forEach(d => monthMap[d.MONTH] = (monthMap[d.MONTH] || 0) + 1);
    Plotly.react("monthlyTrend", [{
        x: Object.keys(monthMap),
        y: Object.values(monthMap),
        type: "bar",
        marker: { color: "#800000" }
    }], { title: "Monthly New Policies", margin: { t: 30 } });

    document.getElementById("monthlyTrend")?.on('plotly_click', function(event){
        const clickedMonth = event.points[0].x;
        const filtered = mainData.filter(d => d.MONTH === parseInt(clickedMonth));
        updateDashboard(filtered, false);
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
