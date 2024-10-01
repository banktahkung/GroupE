// + Excel File Creation
// > Test function to download Excel
async function downloadExcel(filename = "data.xlsx") {
  if (Object.keys(OutputTempTable).length == 0) {
    await FileConversion(FileData);
  }

  const { ExportHeader, ExportData } = transformData(OutputTempTable);

  // * Create a new workbook
  const wb = XLSX.utils.book_new();

  // * Convert JSON to a worksheet with specified headers
  const ws = XLSX.utils.json_to_sheet(ExportData, { header: ExportHeader });

  // * Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // * Generate a file name and save it
  XLSX.writeFile(wb, filename);
}

// > Transform JSON data into headers and rows
function transformData(jsonData) {
  let headers = [];
  let rows = [];

  // * Get the header of json data
  const HeaderKey = Object.keys(jsonData);

  // * Check if jsonData has at least one row to extract headers
  if (HeaderKey.length > 0) {
    // * Extract headers from the first row
    const firstRow = jsonData[HeaderKey[0]];
    headers = Object.values(firstRow);

    // * Process remaining rows
    HeaderKey.slice(1).forEach((rowKey) => {
      const row = jsonData[rowKey];

      let rowData = {};
      Object.keys(row).forEach((colKey, index) => {
        rowData[headers[index]] = row[colKey];
      });
      rows.push(rowData);
    });
  }

  return { ExportHeader: headers, ExportData: rows };
}

// + Export Window Attribute
// > Export window attribute and functions
function TriggerExportWindow() {
  // * Get the export window
  const ExportWindow = document.getElementById("fileOutputWindow");

  // * Check whether the file is already convert or not
  if (CurrentProfile && CurrentFile) {
    // * Displaying and Closing the window trigger function
    if (WindowExport) {
      WindowExport = false;
      ExportWindow.style.display = "none";
    } else {
      WindowExport = true;
      ExportWindow.style.display = "block";
    }
  } else if (!CurrentFile) {
    WarningMessage.style.fontSize = `30px`;
    WarningMessage.textContent = "File content not found";

    ClosingWarningWindow();
  } else {
    WarningMessage.style.fontSize = `40px`;
    WarningMessage.textContent = "Invalid Profile";

    ClosingWarningWindow();
  }
}
