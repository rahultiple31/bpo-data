(function () {
  const STORAGE_KEY = "bpo-data-entry-records";

  const fieldLabels = [
    ["imageName", "Image Name"],
    ["recordNumber", "Record Number"],
    ["accountType", "Type of Account"],
    ["registeredAddress", "Regd Address"],
    ["fullName", "Full Name"],
    ["fatherHusbandName", "Father/Husband Name"],
    ["dateOfBirth", "Date of Birth"],
    ["gender", "Gender"],
    ["nationality", "Nationality"],
    ["address", "Address"],
    ["city", "City/District/Town"],
    ["state", "State"],
    ["pincode", "Pincode"],
    ["mobileNumber", "Mobile Number"],
    ["email", "Email"],
    ["contact1", "Contact 1"],
    ["contact2", "Contact 2"],
    ["addressVerification", "Address Verification"],
    ["idVerification", "ID Verification"],
    ["centerName", "Center Name"],
    ["centerCode", "Center Code"],
    ["remarks", "Remarks"]
  ];

  const form = document.getElementById("entryForm");
  const imageInput = document.getElementById("imageInput");
  const imagePreview = document.getElementById("imagePreview");
  const imageName = document.getElementById("imageName");
  const previewViewport = document.getElementById("previewViewport");
  const messageDialog = document.getElementById("messageDialog");
  const dataDialog = document.getElementById("dataDialog");
  const resultTable = document.getElementById("resultTable");
  const rowCount = document.getElementById("rowCount");
  const remainingTime = document.getElementById("remainingTime");

  let records = loadRecords();
  let selectedImageUrls = [];
  let countdownSeconds = (7 * 60 * 60) + (54 * 60);

  function loadRecords() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    } catch (error) {
      return [];
    }
  }

  function saveRecords() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
  }

  function collectFormData() {
    const data = {};
    const values = new FormData(form);

    fieldLabels.forEach(([key]) => {
      data[key] = (values.get(key) || "").toString().trim();
    });

    data.updatedAt = new Date().toLocaleString();
    return data;
  }

  function fillForm(record) {
    fieldLabels.forEach(([key]) => {
      const control = form.elements[key];
      if (control) {
        control.value = record[key] || "";
      }
    });
  }

  function clearEntryFields() {
    fieldLabels.forEach(([key]) => {
      if (key !== "imageName" && form.elements[key]) {
        form.elements[key].value = "";
      }
    });
    form.elements.remarks.value = "N.A.";
  }

  function showMessage(message) {
    document.getElementById("dialogMessage").textContent = message;

    if (messageDialog.showModal) {
      messageDialog.showModal();
    } else {
      alert(message);
    }
  }

  function closeDialog(dialog) {
    if (dialog.open) {
      dialog.close();
    }
  }

  function updatePreview(file) {
    selectedImageUrls.forEach((url) => URL.revokeObjectURL(url));
    selectedImageUrls = [];

    if (!file) {
      imagePreview.removeAttribute("src");
      previewViewport.classList.remove("has-image");
      imageName.value = "";
      return;
    }

    const url = URL.createObjectURL(file);
    selectedImageUrls.push(url);
    imagePreview.src = url;
    imageName.value = file.name;
    previewViewport.classList.add("has-image");
  }

  function handleImageSelection(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) {
      return;
    }

    updatePreview(files[0]);
    clearEntryFields();
  }

  function renderDataTable() {
    const thead = resultTable.querySelector("thead");
    const tbody = resultTable.querySelector("tbody");
    const headers = fieldLabels.concat([["updatedAt", "Updated At"]]);

    thead.innerHTML = "";
    tbody.innerHTML = "";

    const headerRow = document.createElement("tr");
    headers.forEach(([, label]) => {
      const th = document.createElement("th");
      th.textContent = label;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    records.forEach((record) => {
      const row = document.createElement("tr");
      headers.forEach(([key]) => {
        const cell = document.createElement("td");
        cell.textContent = record[key] || "";
        row.appendChild(cell);
      });
      tbody.appendChild(row);
    });

    rowCount.textContent = `${records.length} ${records.length === 1 ? "record" : "records"}`;
  }

  function openDataDialog() {
    renderDataTable();

    if (dataDialog.showModal) {
      dataDialog.showModal();
    } else {
      exportExcel();
    }
  }

  function escapeXml(value) {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function buildHtmlWorkbook() {
    const headers = fieldLabels.concat([["updatedAt", "Updated At"]]);
    const rows = records.map((record) => (
      `<tr>${headers.map(([key]) => `<td>${escapeXml(record[key] || "")}</td>`).join("")}</tr>`
    )).join("");

    return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }
    th, td { border: 1px solid #999; padding: 5px 8px; white-space: nowrap; }
    th { background: #e8e8e8; font-weight: 700; }
  </style>
</head>
<body>
  <table>
    <thead><tr>${headers.map(([, label]) => `<th>${escapeXml(label)}</th>`).join("")}</tr></thead>
    <tbody>${rows}</tbody>
  </table>
</body>
</html>`;
  }

  function downloadFile(filename, type, content) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function exportExcel() {
    if (!records.length) {
      showMessage("No data available to export");
      return;
    }

    downloadFile(
      "data-entry-results.xls",
      "application/vnd.ms-excel;charset=utf-8",
      buildHtmlWorkbook()
    );
  }

  function backupXml() {
    if (!records.length) {
      showMessage("No data available for XML backup");
      return;
    }

    const xmlRows = records.map((record) => {
      const fields = fieldLabels
        .concat([["updatedAt", "Updated At"]])
        .map(([key]) => `    <${key}>${escapeXml(record[key] || "")}</${key}>`)
        .join("\n");
      return `  <record>\n${fields}\n  </record>`;
    }).join("\n");

    downloadFile(
      "data-entry-backup.xml",
      "application/xml;charset=utf-8",
      `<?xml version="1.0" encoding="UTF-8"?>\n<records>\n${xmlRows}\n</records>`
    );
  }

  function searchCity() {
    const city = prompt("Enter city to search");
    if (!city) {
      return;
    }

    const match = records.find((record) => (
      (record.city || "").toLowerCase() === city.trim().toLowerCase()
    ));

    if (match) {
      fillForm(match);
      showMessage("City record loaded");
    } else {
      showMessage("City not found");
    }
  }

  function showNotes() {
    const current = form.elements.remarks.value || "";
    const notes = prompt("Notes", current);
    if (notes !== null) {
      form.elements.remarks.value = notes;
    }
  }

  function updateCounter() {
    document.getElementById("totalFormsButton").textContent = `Total Forms: ${records.length}`;
  }

  function tickTimer() {
    countdownSeconds = Math.max(0, countdownSeconds - 1);
    const hours = String(Math.floor(countdownSeconds / 3600)).padStart(2, "0");
    const minutes = String(Math.floor((countdownSeconds % 3600) / 60)).padStart(2, "0");
    const seconds = String(countdownSeconds % 60).padStart(2, "0");
    remainingTime.textContent = `${hours}:${minutes}:${seconds}`;
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();

    const data = collectFormData();
    const existingIndex = records.findIndex((record) => (
      record.imageName === data.imageName && record.recordNumber === data.recordNumber
    ));

    if (existingIndex >= 0) {
      records[existingIndex] = data;
    } else {
      records.push(data);
    }

    saveRecords();
    updateCounter();
    showMessage("Data Updated successfully");
  });

  imageInput.addEventListener("change", handleImageSelection);
  document.getElementById("selectImageButton").addEventListener("click", () => imageInput.click());
  document.getElementById("showDataButton").addEventListener("click", openDataDialog);
  document.getElementById("backupXmlButton").addEventListener("click", backupXml);
  document.getElementById("searchCityButton").addEventListener("click", searchCity);
  document.getElementById("notesButton").addEventListener("click", showNotes);
  document.getElementById("exportExcelButton").addEventListener("click", exportExcel);
  document.getElementById("exitButton").addEventListener("click", () => window.close());
  document.getElementById("totalFormsButton").addEventListener("click", () => {
    showMessage(`Total Forms: ${records.length}`);
  });

  document.getElementById("clearDataButton").addEventListener("click", () => {
    if (confirm("Clear all saved records?")) {
      records = [];
      saveRecords();
      updateCounter();
      renderDataTable();
    }
  });

  document.getElementById("closeDialogButton").addEventListener("click", () => closeDialog(messageDialog));
  document.getElementById("closeDialogX").addEventListener("click", () => closeDialog(messageDialog));
  document.getElementById("closeDataX").addEventListener("click", () => closeDialog(dataDialog));

  updateCounter();
  setInterval(tickTimer, 1000);
})();
