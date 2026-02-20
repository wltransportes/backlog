let originalData = [];
let statusChart;
let stationChart;

const uploadInput = document.getElementById("upload");
const filterStation = document.getElementById("filterStation");
const filterStatus = document.getElementById("filterStatus");
const filterDriver = document.getElementById("filterDriver");
const searchInput = document.getElementById("searchInput");

uploadInput.addEventListener("change", handleFile);
filterStation.addEventListener("change", applyFilters);
filterStatus.addEventListener("change", applyFilters);
filterDriver.addEventListener("change", applyFilters);
searchInput.addEventListener("keyup", handleSearch);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    originalData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    populateFilters();
    updateDashboard(originalData);
  };

  reader.readAsArrayBuffer(file);
}

function populateFilters() {
  const stations = [...new Set(originalData.map(d => d["Station Name"]).filter(Boolean))];
  const status = [...new Set(originalData.map(d => d["Latest Status"]).filter(Boolean))];
  const drivers = [...new Set(originalData.map(d => d["Latest User Name"]).filter(Boolean))];

  filterStation.innerHTML = '<option value="">Filtrar por Station</option>';
  filterStatus.innerHTML = '<option value="">Filtrar por Status</option>';
  filterDriver.innerHTML = '<option value="">Filtrar por Entregador</option>';

  stations.sort().forEach(s => {
    filterStation.innerHTML += `<option value="${s}">${s}</option>`;
  });

  status.sort().forEach(s => {
    filterStatus.innerHTML += `<option value="${s}">${s}</option>`;
  });

  drivers.sort().forEach(d => {
    filterDriver.innerHTML += `<option value="${d}">${d}</option>`;
  });
}

function applyFilters() {
  let filtered = [...originalData];

  const station = filterStation.value;
  const status = filterStatus.value;
  const driver = filterDriver.value;

  if (station) {
    filtered = filtered.filter(d => d["Station Name"] === station);
  }

  if (status) {
    filtered = filtered.filter(d => d["Latest Status"] === status);
  }

  if (driver) {
    filtered = filtered.filter(d => d["Latest User Name"] === driver);
  }

  updateDashboard(filtered);
}

function updateDashboard(data) {
  document.getElementById("totalPendencias").innerText = data.length;

  document.getElementById("totalStations").innerText =
    new Set(data.map(d => d["Station Name"]).filter(Boolean)).size;

  document.getElementById("totalStatus").innerText =
    new Set(data.map(d => d["Latest Status"]).filter(Boolean)).size;

  document.getElementById("totalDrivers").innerText =
    new Set(data.map(d => d["Latest User Name"]).filter(Boolean)).size;

  updateCharts(data);
  updateTable(data);
}

function updateCharts(data) {
  const statusCount = {};
  const stationCount = {};

  data.forEach(d => {
    const status = d["Latest Status"] || "Sem Status";
    const station = d["Station Name"] || "Sem Station";

    statusCount[status] = (statusCount[status] || 0) + 1;
    stationCount[station] = (stationCount[station] || 0) + 1;
  });

  if (statusChart) statusChart.destroy();
  if (stationChart) stationChart.destroy();

  statusChart = new Chart(document.getElementById("statusChart"), {
    type: "bar",
    data: {
      labels: Object.keys(statusCount),
      datasets: [{
        label: "Pendências por Status",
        data: Object.values(statusCount),
        backgroundColor: "#ff2e2e"
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          labels: { color: "#fff" }
        }
      },
      scales: {
        x: { ticks: { color: "#fff" } },
        y: { ticks: { color: "#fff" } }
      }
    }
  });

  stationChart = new Chart(document.getElementById("stationChart"), {
    type: "bar",
    data: {
      labels: Object.keys(stationCount),
      datasets: [{
        label: "Pendências por Station",
        data: Object.values(stationCount),
        backgroundColor: "#ff2e2e"
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          labels: { color: "#fff" }
        }
      },
      scales: {
        x: { ticks: { color: "#fff" } },
        y: { ticks: { color: "#fff" } }
      }
    }
  });
}

function updateTable(data) {
  const tbody = document.getElementById("tableBody");
  tbody.innerHTML = "";

  data.forEach(d => {
    const row = document.createElement("tr");

    row.innerHTML = `
      <td>${d["Shipment ID"] || "-"}</td>
      <td>${d["Station Name"] || "-"}</td>
      <td>${d["Latest Status"] || "-"}</td>
      <td>${d["Latest User Name"] || "-"}</td>
      <td>${d["LM Hub Aging"] || "-"}</td>
      <td>${d["Delivery Attempts"] || "-"}</td>
    `;

    tbody.appendChild(row);
  });
}

function handleSearch() {
  const value = searchInput.value.toLowerCase();
  const rows = document.querySelectorAll("#tableBody tr");

  rows.forEach(row => {
    row.style.display = row.innerText.toLowerCase().includes(value)
      ? ""
      : "none";
  });
}
