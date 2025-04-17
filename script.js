let salesData = [];
let barChart, pieChart;

document.addEventListener("DOMContentLoaded", () => {
  fetch("data.json")
    .then((res) => res.json())
    .then((data) => {
      salesData = data;
      populateTable(data);
      updateCategoryFilter(data);
      updateTotalSales(data);
      updateCharts(data);
    });

  document.getElementById("category").addEventListener("change", (e) => {
    const filtered = filterData(salesData, e.target.value);
    applyFilters(filtered);
    updateTotalSales(data);
  });

  document.getElementById("searchBox").addEventListener("input", (e) => {
    const text = e.target.value.toLowerCase();
    const filtered = filterData(salesData, document.getElementById("category").value);
    const searched = filtered.filter((item) => item.product.toLowerCase().includes(text));
    applyFilters(searched);
  });

  document.getElementById("toggleTheme").addEventListener("click", () => {
    document.body.classList.toggle("dark");
  });
});

function populateTable(data) {
  const tbody = document.querySelector("#salesTable tbody");
  tbody.innerHTML = "";
  data.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${item.product}</td>
      <td>${item.category}</td>
      <td>${item.region}</td>
      <td>$${item.sales}</td>
    `;
    tbody.appendChild(row);
  });
}

function filterData(data, category) {
  return category === "all" ? data : data.filter((item) => item.category === category);
}

function updateCategoryFilter(data) {
  const categories = [...new Set(data.map((d) => d.category))];
  const select = document.getElementById("category");
  categories.forEach((cat) => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    select.appendChild(opt);
  });
}
function updateTotalSales(data) {
  const total = data.reduce((sum, item) => sum + Number(item.sales), 0); // Asegúrate de que sales sea un número
  document.getElementById("totalSales").textContent = `${total}`;

  const topProduct = data.reduce((a, b) => (a.sales > b.sales ? a : b), { sales: 0 });
  document.getElementById("topProduct").textContent = topProduct.product;

  const regionSales = {};
  data.forEach(item => {
    regionSales[item.region] = (regionSales[item.region] || 0) + Number(item.sales); // Convierte sales a número
  });
  const topRegion = Object.entries(regionSales).sort((a, b) => b[1] - a[1])[0];
  document.getElementById("topRegion").textContent = topRegion[0];
}
function updateCharts(data) {
  // Datos para el gráfico de barras
  const salesByCategory = {};
  data.forEach(item => {
    salesByCategory[item.category] = (salesByCategory[item.category] || 0) + item.sales;
  });

  const categories = Object.keys(salesByCategory);
  const sales = Object.values(salesByCategory);

  // Gráfico de barras
  const barCtx = document.getElementById("barChart").getContext("2d");
  new Chart(barCtx, {
    type: "bar",
    data: {
      labels: categories,
      datasets: [{
        label: "Sales by Category",
        data: sales,
        backgroundColor: "#4e79a7"
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: {
          display: true,
          text: "Sales by Category",
          font: { size: 16 }
        }
      },
      scales: {
        y: { beginAtZero: true }
      }
    }
  });

  // Datos para el gráfico de líneas
  const salesByRegion = {};
  data.forEach(item => {
    salesByRegion[item.region] = (salesByRegion[item.region] || 0) + item.sales;
  });

  const regions = Object.keys(salesByRegion);
  const regionSales = Object.values(salesByRegion);

  // Gráfico de líneas
  const lineCtx = document.getElementById("lineChart").getContext("2d");
  new Chart(lineCtx, {
    type: "line",
    data: {
      labels: regions,
      datasets: [{
        label: "Sales by Region",
        data: regionSales,
        borderColor: "#f28e2c",
        backgroundColor: "rgba(242, 142, 44, 0.2)",
        fill: true
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: {
          display: true,
          text: "Sales by Region",
          font: { size: 16 }
        }
      },
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
}

function applyFilters(filteredData) {
  populateTable(filteredData);
  updateTotalSales(filteredData);
  updateCharts(filteredData);
}
function exportToExcel() {
  if (!salesData || salesData.length === 0) {
    console.error("No data available to export.");
    alert("No data available to export.");
    return;
  }

  // Crear una hoja de cálculo
  const worksheetData = [
    ["Product", "Category", "Region", "Sales"], // Encabezados
    ...salesData.map(item => [item.product, item.category, item.region, item.sales]) // Filas de datos
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

  // Aplicar estilos a los encabezados
  const headerRange = XLSX.utils.decode_range(worksheet['!ref']);
  for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
    if (!worksheet[cellAddress]) continue;
    worksheet[cellAddress].s = {
      font: { bold: true, color: { rgb: "FFFFFF" } }, // Texto blanco
      fill: { fgColor: { rgb: "4F81BD" } }, // Fondo azul
      alignment: { horizontal: "center", vertical: "center" }
    };
  }

  // Agregar filtro automático
  worksheet['!autofilter'] = { ref: worksheet['!ref'] };

  // Crear un libro de trabajo
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Data");

  // Exportar el archivo Excel
  XLSX.writeFile(workbook, "sales_data.xlsx");
}

document.getElementById("exportButton").addEventListener("click", exportToExcel);
