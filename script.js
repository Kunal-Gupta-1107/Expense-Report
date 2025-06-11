function addRow() {
  const tbody = document.getElementById("table-body");
  const row = document.createElement("tr");

  row.innerHTML = `
    <td><input type="date" onchange="calculateTotal()"></td>
    <td><input type="text" onchange="calculateTotal()"></td>
    <td><input type="text" onchange="calculateTotal()"></td>
    <td><input type="text" onchange="calculateTotal()"></td>
    <td><input type="number" class="km" onchange="calculateTotal()"></td>
    <td class="amount">0</td>
  `;

  tbody.appendChild(row);
}

function calculateTotal() {
  const rate = parseFloat(document.getElementById("rate").value) || 0;
  const kms = document.querySelectorAll(".km");
  let total = 0;

  kms.forEach(kmInput => {
    const km = parseFloat(kmInput.value) || 0;
    const amount = km * rate;
    kmInput.closest("tr").querySelector(".amount").innerText = amount.toFixed(2);
    total += amount;
  });

  document.getElementById("total").innerText = total.toFixed(2);
}

// function downloadPDF() {
//   html2canvas(document.querySelector("#form-container")).then(canvas => {
//     const imgData = canvas.toDataURL("image/png");
//     const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
//     // const imgProps = pdf.getImageProperties(imgData);
//     const pdfWidth = pdf.internal.pageSize.getWidth();
//     const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
//     pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
//     pdf.save("Mileage_Report.pdf");
//   });
// }

// function downloadExcel() {
//   const table = document.getElementById("mileage-table");
//   const wb = XLSX.utils.book_new();
//   const ws = XLSX.utils.table_to_sheet(table);
//   XLSX.utils.book_append_sheet(wb, ws, "Report");
//   XLSX.writeFile(wb, "Mileage_Report.xlsx");
// }
// function exportToExcel() {
//     // Get all form inputs using correct IDs
//     const empName = document.getElementById("empName").value;
//     const empId = document.getElementById("empId").value;
//     const vehicle = document.getElementById("vehicle").value;
//     const payFrom = document.getElementById("payFrom").value;
//     const payTo = document.getElementById("payTo").value;
//     const rate = document.getElementById("rate").value;
//     const total = document.getElementById("total").innerText;

//     // Prepare data array
//     const data = [
//         ["Employee Name", empName],
//         ["Employee ID", empId],
//         ["Vehicle Description", vehicle],
//         ["Pay Period From", payFrom],
//         ["Pay Period To", payTo],
//         ["Mileage Rate (per km)", rate],
//         [],
//         ["Date", "Description", "Starting Location", "Ending Location", "KMs", "Amount"]
//     ];

//     // Loop through each row
//     const rows = document.querySelectorAll("#table-body tr");
//     rows.forEach(row => {
//         const cols = row.querySelectorAll("input");
//         const rowData = Array.from(cols).map(input => input.value);
//         data.push(rowData);
//     });

//     data.push([]);
//     data.push(["Total Reimbursement", total]);

//     // Create worksheet and workbook
//     const worksheet = XLSX.utils.aoa_to_sheet(data);
//     const workbook = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(workbook, worksheet, "Mileage Report");

//     // Download as Excel
//     XLSX.writeFile(workbook, "Employee_Mileage_Report.xlsx");
// }


async function downloadPDF() {
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF();

  doc.setFontSize(16);
doc.text("Employee Mileage Expense Report", 14, 15);

doc.setFontSize(12);
const empName = document.getElementById("empName").value;
const empId = document.getElementById("empId").value;
const vehicle = document.getElementById("vehicle").value;
const payFrom = document.getElementById("payFrom").value;
const payTo = document.getElementById("payTo").value;
const rate = document.getElementById("rate").value;
const total = document.getElementById("total").innerText;

const labels = [
  "Employee Name",
  "Employee ID",
  "Vehicle Description",
  "Pay Period From",
  "Pay Period To",
  "Mileage Rate (per km)"
];

const values = [
  empName,
  empId,
  vehicle,
  payFrom,
  payTo,
  rate
];

let startY = 25;
for (let i = 0; i < labels.length; i++) {
  doc.text(`${labels[i]}:`, 14, startY + i * 7);   // Label
  doc.text(values[i], 80, startY + i * 7);         // Value aligned separately
}
  // Collects Table Row
  const tableBody = [];
  const rows = document.querySelectorAll("#table-body tr");

  rows.forEach(row => {
    const inputs = row.querySelectorAll("input");
    const rowData = Array.from(inputs).map(input => input.value);
    tableBody.push(rowData);
  });

  // Add Table
  doc.autoTable({
    startY: 70,
    head: [["Date", "Description", "Starting Location", "Ending Location", "KMs", "Amount"]],
    body: tableBody,
    theme: 'grid',
    styles: { fontSize: 10 },
  });

  // Add Total
  doc.setFontSize(12);
  doc.text(`Total :  ${total}`, 14, doc.lastAutoTable.finalY + 10);

  const today = new Date();
  const dateStr = today.toLocaleDateString('en-GB').replace(/\//g, '-');
  doc.save(`Employee_Mileage_Report_${dateStr}.pdf`);
}












async function exportStyledExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Mileage Report");

  // Set column structure
  sheet.columns = [
    { header: "", width: 15 },
    { header: "", width: 25 },
    { header: "", width: 20 },
    { header: "", width: 20 },
    { header: "", width: 10 },
    { header: "", width: 15 }
  ];

  // Add Employee Info
  const empInfo = [
    ["Employee Mileage Expense Report"],
    [],
    ["Employee Name:", document.getElementById("empName").value],
    ["Employee ID:", document.getElementById("empId").value],
    ["Vehicle Description:", document.getElementById("vehicle").value],
    ["Pay Period From:", document.getElementById("payFrom").value],
    ["Pay Period To:", document.getElementById("payTo").value],
    ["Mileage Rate (per km):", document.getElementById("rate").value],
    []
  ];

  empInfo.forEach(row => sheet.addRow(row));
  sheet.addRow(); // Blank row table se pahle

  // Add Table Headers with styling
  const headerRow = sheet.addRow(["Date", "Description", "Starting Location", "Ending Location", "KMs", "Amount"]);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF007ACC" }
    };
    cell.alignment = { horizontal: "center" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  });

  // Add Table Rows
  const rows = document.querySelectorAll("#table-body tr");
  rows.forEach(row => {
    const inputs = row.querySelectorAll("input");
    const rowData = Array.from(inputs).map(input => input.value);
    const dataRow = sheet.addRow(rowData);

    dataRow.eachCell((cell) => {
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    });
  });

  // Add Total
  sheet.addRow([]);
  const total = document.getElementById("total").innerText;
  const totalRow = sheet.addRow(["Total", "", "", "", "", `₹ ${total}`]);
  totalRow.eachCell((cell, colNumber) => {
    if (colNumber === 6) {
      cell.font = { bold: true };
    }
    cell.border = {
      top: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" }
    };
  });

  // Save File
  const today = new Date();
  const dateStr = today.toLocaleDateString('en-GB').replace(/\//g, '-');
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `Employee_Mileage_Report_${dateStr}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
