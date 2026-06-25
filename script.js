let exportData = [];

const fileInput = document.getElementById("fileInput");
const processBtn = document.getElementById("processBtn");
const downloadBtn = document.getElementById("downloadBtn");
const downloadExcelBtn = document.getElementById("downloadExcelBtn");
const dropArea = document.getElementById("dropArea");
const browseBtn = document.getElementById("browseBtn");
const fileName = document.getElementById("fileName");
const resultBody = document.getElementById("resultBody");
const statusBox = document.getElementById("status");
const totalContacts = document.getElementById("totalContacts");
const totalSchools = document.getElementById("totalSchools");
const totalDistricts = document.getElementById("totalDistricts");

function text(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function splitName(fullName) {
  const name = text(fullName);

  if (name.includes(",")) {
    const [lastName, ...firstNameParts] = name.split(",");
    return {
      firstName: firstNameParts.join(",").trim(),
      lastName: lastName.trim()
    };
  }

  const parts = name.split(/\s+/).filter(Boolean);
  const lastName = parts.pop() || "";

  return {
    firstName: parts.join(" "),
    lastName
  };
}

function escapeHtml(value) {
  return text(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function csvCell(value) {
  return `"${text(value).replace(/"/g, '""')}"`;
}

function setStatus(message, type = "info") {
  statusBox.className = type;
  statusBox.textContent = message;
}

function updateFileName() {
  fileName.textContent = fileInput.files[0]?.name || "No file selected";
}

function updateDashboard() {
  const schools = new Set(exportData.map((item) => item.school).filter(Boolean));
  const districts = new Set(exportData.map((item) => item.district).filter(Boolean));

  totalContacts.textContent = exportData.length;
  totalSchools.textContent = schools.size;
  totalDistricts.textContent = districts.size;
}

function renderPreview() {
  resultBody.innerHTML = exportData
    .slice(0, 10)
    .map((item) => `
      <tr>
        <td>${escapeHtml(item.firstName)}</td>
        <td>${escapeHtml(item.lastName)}</td>
        <td>${escapeHtml(item.email)}</td>
        <td>${escapeHtml(item.school)}</td>
        <td>${escapeHtml(item.district)}</td>
        <td>${escapeHtml(item.state)}</td>
      </tr>
    `)
    .join("");
}

function cleanRows(rows) {
  let currentState = "";

  return rows.reduce((contacts, row, index) => {
    if (index === 0 || !row || !text(row[0])) {
      return contacts;
    }

    const { firstName, lastName } = splitName(row[0]);
    const rowState = text(row[5]);

    if (rowState) {
      currentState = rowState;
    }

    contacts.push({
      firstName,
      lastName,
      position: text(row[1]),
      email: text(row[2]),
      school: text(row[3]),
      district: text(row[4]),
      state: currentState
    });

    return contacts;
  }, []);
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        if (!worksheet) {
          reject(new Error("No worksheet found in this file."));
          return;
        }

        resolve(XLSX.utils.sheet_to_json(worksheet, { header: 1 }));
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error("Could not read the selected file."));
    reader.readAsArrayBuffer(file);
  });
}

function downloadBlob(blob, fileNameForDownload) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileNameForDownload;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

processBtn.addEventListener("click", async () => {
  const file = fileInput.files[0];

  if (!file) {
    setStatus("Please select an Excel or CSV file first.", "error");
    return;
  }

  try {
    setStatus("Cleaning file...");

    const rows = await readWorkbook(file);
    exportData = cleanRows(rows);

    renderPreview();
    updateDashboard();

    if (!exportData.length) {
      setStatus("No contacts found. Check that your first column contains names.", "error");
      return;
    }

    setStatus(`Complete. ${exportData.length} contact${exportData.length === 1 ? "" : "s"} ready for download.`, "success");
  } catch (error) {
    exportData = [];
    renderPreview();
    updateDashboard();
    setStatus(error.message || "Something went wrong while cleaning the file.", "error");
  }
});

downloadBtn.addEventListener("click", () => {
  if (!exportData.length) {
    setStatus("Clean a file before downloading CSV.", "error");
    return;
  }

  const header = ["First Name", "Last Name", "Position", "Email", "School", "District", "State"];
  const rows = exportData.map((item) => [
    item.firstName,
    item.lastName,
    item.position,
    item.email,
    item.school,
    item.district,
    item.state
  ]);

  const csv = [header, ...rows]
    .map((row) => row.map(csvCell).join(","))
    .join("\n");

  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8;" }), "EBE_Cleaned_Data.csv");
});

downloadExcelBtn.addEventListener("click", () => {
  if (!exportData.length) {
    setStatus("Clean a file before downloading Excel.", "error");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Cleaned Data");
  XLSX.writeFile(workbook, "EBE_Cleaned_Data.xlsx");
});

browseBtn.addEventListener("click", () => {
  fileInput.click();
});

fileInput.addEventListener("change", updateFileName);

dropArea.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropArea.classList.add("dragover");
});

dropArea.addEventListener("dragleave", () => {
  dropArea.classList.remove("dragover");
});

dropArea.addEventListener("drop", (event) => {
  event.preventDefault();
  dropArea.classList.remove("dragover");

  if (event.dataTransfer.files.length) {
    fileInput.files = event.dataTransfer.files;
    updateFileName();
  }
});
