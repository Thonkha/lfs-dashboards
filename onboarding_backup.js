window.onload = loadGoogleSheetData;


async function loadGoogleSheetData() {
    const sheetId = "1NLBYEfipHHvdpyLPrVfhqybRgFLlWEqcUkYwqWe8bUI";
    const apiKey = "AIzaSyDjgYbFOH0uMwniKeiM6lZ5BJbOcupeM74"; // create in Google Cloud console
    const range = "Sheet1"; // tab name

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}?key=${apiKey}`;

    const response = await fetch(url);
    const result = await response.json();

    // First row = headers
    const headers = result.values[0];
    const rows = result.values.slice(1);

    const jsonData = rows.map(row => {
        let obj = {};
        headers.forEach((h, i) => {
            obj[h.trim().toUpperCase()] = row[i] || "";
        });
        return obj;
    });

    updateDashboard(jsonData);
}

function updateDashboard(data) {
    data.forEach(d => {
        d.MONTH = parseInt(d.MONTH);
        d.WEEK = parseInt(d.WEEK);
        d.DAY = parseInt(d.DAY);
        d.ACTION = (d.ACTION || "").trim();
        d.STATUS = (d.STATUS || "").trim();
        d["PLAN TYPE"] = (d["PLAN TYPE"] || "").trim();
        d.REGION = (d.REGION || "").trim();
        d.BRANCH = (d.BRANCH || "").trim();
        d["CAPTURED BY"] = (d["CAPTURED BY"] || "").trim();
    });

    // KPIs
    const totalPolicies = data.length;
    const trialCount = data.filter(d => d.STATUS.toUpperCase() === "ON TRIAL").length;
    const upgrades = data.filter( d => d.ACTION.toUpperCase() === "UPGRADE").length;
    const downgrades = data.filter( d => d.ACTION.toUpperCase() === "DOWNGRADE").length;
    const activeCount = data.filter(d => d.STATUS.toUpperCase() === "ACTIVE").length;
    const conversionRate = totalPolicies ? ((activeCount / totalPolicies) * 100).toFixed(1) : 0;

    document.getElementById("totalPolicies").textContent = totalPolicies;
    document.getElementById("trialCount").textContent = trialCount;
    document.getElementById("Upgrades").textContent = upgrades;
    document.getElementById("Downgrades").textContent = downgrades;
    document.getElementById("activeCount").textContent = activeCount;
    document.getElementById("conversionRate").textContent = conversionRate + "%";

    // Weekly Trend
    const weekMap = {};
    data.forEach(d => {
        weekMap[d.WEEK] = (weekMap[d.WEEK] || 0) + 1;
    });
    Plotly.newPlot("weeklyTrend", [{
        x: Object.keys(weekMap),
        y: Object.values(weekMap),
        type: "scatter",
        mode: "lines+markers",
        line: { color: "#800000" }
    }], { title: "Weekly New Policies", margin: { t: 30 } });

    // Plan Type Distribution
    const planMap = {};
    data.forEach(d => {
        planMap[d["PLAN TYPE"]] = (planMap[d["PLAN TYPE"]] || 0) + 1;
    });
    Plotly.newPlot("planTypeDist", [{
        x: Object.keys(planMap),
        y: Object.values(planMap),
        type: "bar",
        marker: { color: "#FFD700" }
    }], { title: "Plan Type Distribution", margin: { t: 30 } });

    // Status Split
    const statusMap = {};
    data.forEach(d => {
        statusMap[d.STATUS] = (statusMap[d.STATUS] || 0) + 1;
    });
    Plotly.newPlot("statusSplit", [{
        labels: Object.keys(statusMap),
        values: Object.values(statusMap),
        type: "pie"
    }], { title: "Status Split", margin: { t: 30 } });

    // Regional Uptake
    const regionMap = {};
    data.forEach(d => {
        regionMap[d.REGION] = (regionMap[d.REGION] || 0) + 1;
    });
    Plotly.newPlot("regionUptake", [{
        x: Object.values(regionMap),
        y: Object.keys(regionMap),
        type: "bar",
        orientation: "h",
        marker: { color: "#800000" }
    }], { title: "Regional Uptake", margin: { t: 30 } });

    // Branch Leaderboard
    const branchMap = {};
    data.forEach(d => {
        branchMap[d.BRANCH] = (branchMap[d.BRANCH] || 0) + 1;
    });
    const branchSorted = Object.entries(branchMap).sort((a, b) => b[1] - a[1]);
    Plotly.newPlot("branchLeaderboard", [{
        x: branchSorted.map(d => d[0]),
        y: branchSorted.map(d => d[1]),
        type: "bar",
        marker: { color: "#FFD700" }
    }], { title: "Branch Leaderboard", margin: { t: 30 } });

    // Captured By Ranking
    const capMap = {};
    data.forEach(d => {
        capMap[d["CAPTURED BY"]] = (capMap[d["CAPTURED BY"]] || 0) + 1;
    });
    const capSorted = Object.entries(capMap).sort((a, b) => b[1] - a[1]);
    Plotly.newPlot("capturedByRank", [{
        x: capSorted.map(d => d[0]),
        y: capSorted.map(d => d[1]),
        type: "bar",
        marker: { color: "#800000" }
    }], { title: "Captured By Ranking", margin: { t: 30 } });
    let currentFilter = null;

function drawCharts(filteredData = data) {
    // Example: Draw Region Bar Chart
    let regionCounts = {};
    filteredData.forEach(row => {
        regionCounts[row.REGION] = (regionCounts[row.REGION] || 0) + 1;
    });

    Plotly.newPlot('regionChart', [{
        x: Object.keys(regionCounts),
        y: Object.values(regionCounts),
        type: 'bar',
        marker: { color: '#800000' }
    }], {
        title: 'Policies by Region'
    });

    // Example: Draw Plan Type Pie Chart
    let planCounts = {};
    filteredData.forEach(row => {
        planCounts[row['PLAN TYPE']] = (planCounts[row['PLAN TYPE']] || 0) + 1;
    });

    Plotly.newPlot('planChart', [{
        labels: Object.keys(planCounts),
        values: Object.values(planCounts),
        type: 'pie',
        marker: { colors: ['#800000', '#FFD700', '#000000', '#FFFFFF'] }
    }], {
        title: 'Plan Type Distribution'
    });

    // Add other charts here...
}

// Initial draw
drawCharts();

// Click event for Region Chart
document.getElementById('regionChart').on('plotly_click', function(dataClick) {
    let clickedRegion = dataClick.points[0].x;
    currentFilter = { column: 'REGION', value: clickedRegion };
    let filteredData = data.filter(row => row[currentFilter.column] === currentFilter.value);
    drawCharts(filteredData);
});

// Reset button
document.getElementById('resetBtn').addEventListener('click', function() {
    currentFilter = null;
    drawCharts(data);
});

}

