document.getElementById("upload").addEventListener("change", handleFile, false);
document.getElementById("clear").addEventListener("click", () => {
  document.getElementById("table-container").innerHTML = "";
  document.getElementById("json-output").textContent = "";
  document.getElementById("paste-input").value = "";
});

document.querySelectorAll(".tab").forEach(tab => {
  tab.addEventListener("click", () => {
    const tabName = tab.dataset.tab;

    document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(tc => tc.classList.remove("active"));

    tab.classList.add("active");
    
    const tabContent = document.getElementById(tabName + "-view");
    if (tabContent) {
      tabContent.classList.add("active");

      // Restore table content if switching back to table view
      if (tabName === "table" && document.getElementById("table-container").innerHTML.trim() === "") {
        if (window.lastTableData) {
          displayTable(window.lastTableData);
        }
      }
    } else {
      console.warn(`No content found for tab '${tabName}'`);
    }
    

    // Restore table if it was hidden
    if (tabName === "table" && document.getElementById("table-container").innerHTML.trim() === "") {
      if (window.lastTableData) {
        displayTable(window.lastTableData);
      }
    }
  });
});

document.getElementById("copy-btn").addEventListener("click", () => {
  const jsonText = document.getElementById("json-output").textContent;
  navigator.clipboard.writeText(jsonText).then(() => {
    alert("JSON copied to clipboard!");
  });
});

document.getElementById("parse-paste").addEventListener("click", () => {
  const text = document.getElementById("paste-input").value.trim();
  if (!text) return;

  let delimiterChoice = document.getElementById("delimiter").value;
  let delimiter = delimiterChoice === "auto"
    ? (text.includes("\t") ? "\t" : ",")
    : delimiterChoice;
  const rows = text.split(/\r?\n/).map(row => row.split(delimiter));
  displayTable(rows);
  const headers = rows[0];
  const dataJson = rows.slice(1).map(row => {
    let obj = {};
    headers.forEach((key, i) => obj[key] = row[i]);
    return obj;
  });
  document.getElementById("json-output").textContent = JSON.stringify(dataJson, null, 2);
});

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
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

  if (data.length > 0) {
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    data[0].forEach(header => {
      const th = document.createElement("th");
      th.textContent = header;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    data.slice(1).forEach(row => {
      const tr = document.createElement("tr");
      row.forEach(cell => {
        const td = document.createElement("td");
        td.textContent = cell;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  tableContainer.innerHTML = "";
  tableContainer.appendChild(table);
}