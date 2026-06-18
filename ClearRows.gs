function showClearRowsSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('ClearRowsUI')
      .setTitle('Row Selection')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function clearSelectedRows(rowNumbersString) {
  if (!rowNumbersString) {
    SpreadsheetApp.getUi().alert('No rows were selected.');
    return;
  }
  const rowNumbers = JSON.parse(rowNumbersString);
  if (!rowNumbers || rowNumbers.length === 0) {
    SpreadsheetApp.getUi().alert('No rows were selected.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  let clearedCount = 0;

  try {
    rowNumbers.forEach(row => {
      if (typeof row !== 'number' || row < 1) {
        console.warn(`Invalid row number skipped: ${row}`);
        return; // Skip invalid entries
      }
      // Clear columns B-H
      sheet.getRange(`B${row}:H${row}`).clearContent();
      // Clear cells C4. H4
      sheet.getRange('C4').clearContent();
      sheet.getRange('H4').clearContent();
      clearedCount++;
    });

    if (clearedCount > 0) {
      ui.alert(`Successfully cleared data from ${clearedCount} row(s).`);
    } else {
      ui.alert('No valid rows were processed.');
    }

  } catch (e) {
    ui.alert(`Error: ${e.toString()}`);
  }
}