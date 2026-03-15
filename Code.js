// Main function runs when the spreadsheet is opened, creating a custom menu
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Get the UI object for the Spreadsheet

  // Create the custom menu and add items for both functionalities
  ui.createMenu('GetReady Actions')
    .addSubMenu(
      ui.createMenu('Data Transfer') // Submenu for Data Transfer tasks
        .addItem('Move Weekly Records to Raw Data', 'moveWeeklyRecordToRawData')
    )
    .addSubMenu(
      ui.createMenu('Export') // Submenu for exporting data
        .addItem('Download couple_finances.csv', 'downloadCoupleFinancesCSV')
    )
    .addToUi(); // Add the custom menu to the UI
}

/**
 * Data Transfer Functionality
 */

// Function to move data from Cash Flow to Raw Data
function moveWeeklyRecordToRawData() {
  moveRecordsToRawData('cash_flow');
}

// Core function to handle moving data and clearing the source sheet
function moveRecordsToRawData(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = 'raw_data';

  // Fetch source and target sheets by name
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName);

  // Handle missing sheets
  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert(`Error: One or both sheets (${sourceSheetName}, ${targetSheetName}) not found.`);
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();

  // Check for empty source sheet
  if (lastRow < 2 || lastColumn < 1) {
    SpreadsheetApp.getUi().alert(`No data available to move from ${sourceSheetName}.`);
    return;
  }

  // Get data from the source sheet
  const sourceData = sourceSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const filteredData = sourceData.filter(row => row.some(cell => cell !== ""));

  // Handle no data after filtering
  if (filteredData.length === 0) {
    SpreadsheetApp.getUi().alert(`No data found to move from ${sourceSheetName}.`);
    return;
  }

  // Determine the last row in the target sheet
  const lastRowWithData = targetSheet.getRange("A1:A").getValues().filter(String).length;
  const insertRow = lastRowWithData === 0 ? 1 : lastRowWithData + 1;

  // Insert data into the target sheet
  targetSheet.getRange(insertRow, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // Clear columns A to F in the source sheet
  sourceSheet.getRange(6, 1, lastRow - 1, 6).clearContent();
  SpreadsheetApp.getUi().alert(`Records have been moved from ${sourceSheetName} to ${targetSheetName}, and the source sheet has been cleaned.`);
}

/**
 * Export Functionality
 */

// Downloads the full raw_data sheet as couple_finances.csv
// Dates are formatted as dd/MM/yyyy (dayfirst) so couple_report.py can read them
function downloadCoupleFinancesCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('raw_data');

  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('Error: raw_data sheet not found.');
    return;
  }

  const data = rawSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data found in raw_data.');
    return;
  }

  const tz = Session.getScriptTimeZone();

  // Export all columns as-is, only formatting Date cells to dd/MM/yyyy
  const csvRows = data.map((row, rowIndex) =>
    row.map(cell => {
      if (rowIndex > 0 && cell instanceof Date) {
        return Utilities.formatDate(cell, tz, 'dd/MM/yyyy');
      }
      return cell;
    })
  ).filter((row, rowIndex) =>
    rowIndex === 0 || row.some(cell => cell !== '') // keep header + non-empty rows
  );

  // Escape cells that contain commas, quotes, or newlines
  const csvContent = csvRows.map(row =>
    row.map(cell => {
      const s = String(cell === null || cell === undefined ? '' : cell);
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')
  ).join('\n');

  // Trigger a browser download via an HTML dialog
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <body>
        <p style="font-family:Arial,sans-serif;font-size:13px;">Preparing download...</p>
        <script>
          const csv = ${JSON.stringify(csvContent)};
          const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
          const url  = URL.createObjectURL(blob);
          const a    = document.createElement('a');
          a.href     = url;
          a.download = 'couple_finances.csv';
          document.body.appendChild(a);
          a.click();
          URL.revokeObjectURL(url);
          setTimeout(() => google.script.host.close(), 1500);
        </script>
      </body>
    </html>
  `).setWidth(250).setHeight(80);

  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading couple_finances.csv');
}
