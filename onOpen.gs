/**
 * @OnlyCurrentDoc
 * This script creates a custom menu to trigger an AI-powered cost analysis
 * for construction projects based on data in the "SPECIFICATIONS" sheet.
 * and a file picker menu for uploading specs.
 */

/**
 * Creates a custom menu in the spreadsheet when the file is opened.
 */

/**
 * Creates a custom menu in the spreadsheet UI when the file is opened.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
    ui.createMenu('⬆️ File Uploader ⬆️')
    .addItem('Upload Plans', 'showPlanPicker')
    .addItem('Upload Specifications', 'showSpecPicker')
    .addToUi();

   ui.createMenu('💲 A.I. COST ANALYZER 💲')
      .addItem('💲 Generate Cost Analysis', 'runCostAnalysis')
      .addToUi();

   ui.createMenu('❌Clear Rowz❌')
    .addItem('Open Row Clear Tool', 'showClearRowsSidebar') // This will call your existing function
    .addToUi();   


}