document.getElementById("upload").addEventListener("change", handleFile, false);
document.getElementById("clear").addEventListener("click", () => {
  document.getElementById("table-container").innerHTML = "";
  document.getElementById("json-output").textContent = "";
});

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    displayTable(json);
    const keys = json[0];
    const dataJson = json.slice(1).map(row => {
      let obj = {};
      keys.forEach((key, i) => obj[key] = row[i]);
      return obj;
    });
    document.getElementById("json-output").textContent = JSON.stringify(dataJson, null, 2);
  };

  reader.readAsArrayBuffer(file);
}

function displayTable(data) {
  const tableContainer = document.getElementById("table-container");
  const table = document.createElement("table");
  data.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  tableContainer.innerHTML = "";
  tableContainer.appendChild(table);
}
