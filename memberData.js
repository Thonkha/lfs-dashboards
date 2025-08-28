document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("fileInput");
  const resetButton = document.getElementById("resetFilter");
  const errorMessage = document.getElementById("errorMessage");

  let dataset = [];

  // Upload Excel
  fileInput.addEventListener("change", handleFile);

  async function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      dataset = XLSX.utils.sheet_to_json(sheet);

      if (dataset.length === 0) {
        showError("⚠️ No data found in Excel file.");
        return;
      }

      clearError();
      updateDashboard(dataset);
    } catch (err) {
      showError("Error reading file: " + err.message);
    }
  }

  function showError(msg) {
    errorMessage.textContent = msg;
  }

  function clearError() {
    errorMessage.textContent = "";
  }

  // Reset filter
  resetButton.addEventListener("click", () => {
    if (dataset.length > 0) {
      updateDashboard(dataset);
    }
  });

  // Update Dashboard KPIs + Charts
  function updateDashboard(data) {
    // KPIs
    const total = data.length;
    const active = data.filter(d => d.Status === "Active").length;
    const inactive = data.filter(d => d.Status === "Inactive").length;
    const conversion = total > 0 ? ((active / total) * 100).toFixed(1) : 0;

    document.getElementById("totalRecords").textContent = total;
    document.getElementById("activeCount").textContent = active;
    document.getElementById("inactiveCount").textContent = inactive;
    document.getElementById("conversionRate").textContent = conversion + "%";

    // Weekly Trend (dummy grouping by date column if available)
    const weekly = {};
    data.forEach(d => {
      if (d.Date) {
        const week = new Date(d.Date).toISOString().slice(0, 10);
        weekly[week] = (weekly[week] || 0) + 1;
      }
    });

    Plotly.newPlot("weeklyTrend", [{
      x: Object.keys(weekly),
      y: Object.values(weekly),
      type: "bar"
    }], { title: "Weekly Trend" });

    // Monthly Trend
    const monthly = {};
    data.forEach(d => {
      if (d.Date) {
        const month = new Date(d.Date).toISOString().slice(0, 7);
        monthly[month] = (monthly[month] || 0) + 1;
      }
    });

    Plotly.newPlot("monthlyTrend", [{
      x: Object.keys(monthly),
      y: Object.values(monthly),
      type: "scatter",
      mode: "lines+markers"
    }], { title: "Monthly Trend" });

    // Status Split
    const statusCounts = {};
    data.forEach(d => {
      const status = d.Status || "Unknown";
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    });

    Plotly.newPlot("statusSplit", [{
      labels: Object.keys(statusCounts),
      values: Object.values(statusCounts),
      type: "pie"
    }], { title: "Status Split" });

    // Region Distribution
    const regionCounts = {};
    data.forEach(d => {
      const region = d.Region || "Unknown";
      regionCounts[region] = (regionCounts[region] || 0) + 1;
    });

    Plotly.newPlot("regionDist", [{
      x: Object.keys(regionCounts),
      y: Object.values(regionCounts),
      type: "bar"
    }], { title: "Region Distribution" });
  }
});
