document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('dateStart').addEventListener('change', updateDashboard);
document.getElementById('dateEnd').addEventListener('change', updateDashboard);

let claimsData = [];
let filteredData = [];

function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = e => {
        const rows = e.target.result.split("\n").map(r => r.split(","));
        const headers = rows[0].map(h => h.trim());
        claimsData = rows.slice(1).map(row => {
            let obj = {};
            headers.forEach((h, i) => obj[h] = row[i]?.trim());
            return obj;
        });
        updateDashboard();
    };
    reader.readAsText(file);
}

function updateDashboard() {
    const startDate = document.getElementById('dateStart').value;
    const endDate = document.getElementById('dateEnd').value;

    filteredData = claimsData.filter(row => {
        if (!row["CLAIM DATE"]) return false;
        let claimDate = new Date(row["CLAIM DATE"]);
        if (startDate && claimDate < new Date(startDate)) return false;
        if (endDate && claimDate > new Date(endDate)) return false;
        return true;
    });

    renderMetrics();
    renderCharts();
}

function renderMetrics() {
    let totalClaims = filteredData.length;
    let totalCover = filteredData.reduce((sum, r) => sum + (parseFloat(r["COVER AMOUNT"]) || 0), 0);
    let avgCover = totalClaims ? (totalCover / totalClaims).toFixed(2) : 0;

    document.getElementById('metrics').innerHTML = `
        <div class="card"><h3>Total Claims</h3><div class="metric-value">${totalClaims}</div></div>
        <div class="card"><h3>Total Cover Amount</h3><div class="metric-value">${totalCover}</div></div>
        <div class="card"><h3>Average Cover</h3><div class="metric-value">${avgCover}</div></div>
    `;
}

function renderCharts() {
    // Helper: group counts
    const groupCount = (arr, key) => arr.reduce((acc, r) => {
        acc[r[key]] = (acc[r[key]] || 0) + 1;
        return acc;
    }, {});

    // Claims by Branch
    let branchCounts = groupCount(filteredData, "BRANCH");
    Plotly.newPlot('claimsByBranch', [{
        x: Object.keys(branchCounts),
        y: Object.values(branchCounts),
        type: 'bar',
        marker: { color: '#800000' }
    }], { title: 'Claims by Branch' });

    // Claims by Society
    let societyCounts = groupCount(filteredData, "SOCIETY NAME");
    Plotly.newPlot('claimsBySociety', [{
        x: Object.keys(societyCounts),
        y: Object.values(societyCounts),
        type: 'bar',
        marker: { color: 'gold' }
    }], { title: 'Claims by Society' });

    // Claims by Cover Type
    let coverCounts = groupCount(filteredData, "COVER TYPE");
    Plotly.newPlot('claimsByCoverType', [{
        labels: Object.keys(coverCounts),
        values: Object.values(coverCounts),
        type: 'pie'
    }], { title: 'Claims by Cover Type' });

    // Claims by Gender
    let genderCounts = groupCount(filteredData, "GENDER");
    Plotly.newPlot('claimsByGender', [{
        labels: Object.keys(genderCounts),
        values: Object.values(genderCounts),
        type: 'pie',
        marker: { colors: ['#ff9999', '#9999ff'] }
    }], { title: 'Claims by Gender' });

    // Claims by Coffin Used
    let coffinCounts = groupCount(filteredData, "COFFIN USED");
    Plotly.newPlot('claimsByCoffin', [{
        x: Object.keys(coffinCounts),
        y: Object.values(coffinCounts),
        type: 'bar',
        marker: { color: '#964B00' }
    }], { title: 'Claims by Coffin Used' });

    // Average Cover Amount by Society
    let societyTotals = {};
    let societyNums = {};
    filteredData.forEach(r => {
        let s = r["SOCIETY NAME"];
        societyTotals[s] = (societyTotals[s] || 0) + (parseFloat(r["COVER AMOUNT"]) || 0);
        societyNums[s] = (societyNums[s] || 0) + 1;
    });
    let avgSociety = Object.keys(societyTotals).map(s => ({
        society: s,
        avg: societyTotals[s] / societyNums[s]
    }));
    avgSociety.sort((a, b) => b.avg - a.avg);
    Plotly.newPlot('avgCoverBySociety', [{
        x: avgSociety.map(a => a.society),
        y: avgSociety.map(a => a.avg),
        type: 'bar',
        marker: { color: '#228B22' }
    }], { title: 'Average Cover Amount by Society' });

    // Claims by Month (seasonality)
    let monthCounts = {};
    filteredData.forEach(r => {
        let d = new Date(r["CLAIM DATE"]);
        let month = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        monthCounts[month] = (monthCounts[month] || 0) + 1;
    });
    Plotly.newPlot('claimsByMonth', [{
        x: Object.keys(monthCounts),
        y: Object.values(monthCounts),
        type: 'scatter',
        mode: 'lines+markers',
        marker: { color: '#0000FF' }
    }], { title: 'Claims by Month' });

    // Age Distribution
    let ages = filteredData.map(r => {
        if (!r["DATE OF BIRTH"]) return null;
        let dob = new Date(r["DATE OF BIRTH"]);
        let age = new Date().getFullYear() - dob.getFullYear();
        return age;
    }).filter(a => a !== null && !isNaN(a));
    Plotly.newPlot('ageDistribution', [{
        x: ages,
        type: 'histogram',
        marker: { color: '#FFA500' }
    }], { title: 'Age Distribution of Deceased' });

    // Drill-downs
    document.getElementById('claimsByBranch').on('plotly_click', data => {
        let branch = data.points[0].x;
        filteredData = claimsData.filter(r => r["BRANCH"] === branch);
        renderMetrics();
        renderCharts();
    });

    document.getElementById('claimsBySociety').on('plotly_click', data => {
        let society = data.points[0].x;
        filteredData = claimsData.filter(r => r["SOCIETY NAME"] === society);
        renderMetrics();
        renderCharts();
    });
}
