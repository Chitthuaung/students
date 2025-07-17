let globalData = [];

document.getElementById("upload").addEventListener("change", function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(evt) {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    globalData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    renderTable(globalData);
  };

  reader.readAsArrayBuffer(file);
});

function renderTable(data) {
  let html = '<table contenteditable="true">';
  data.forEach(row => {
    html += "<tr>";
    row.forEach(cell => {
      html += `<td>${cell !== undefined ? cell : ''}</td>`;
    });
    html += "</tr>";
  });
  html += "</table>";
  document.getElementById("table-container").innerHTML = html;
}

function saveExcel() {
  const table = document.querySelector("table");
  const newData = [];

  for (let i = 0; i < table.rows.length; i++) {
    const row = [];
    for (let j = 0; j < table.rows[i].cells.length; j++) {
      row.push(table.rows[i].cells[j].innerText);
    }
    newData.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(newData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  XLSX.writeFile(wb, "edited_file.xlsx");
}
