/**
 * Main function to run the cost analysis. It reads data from the "SPECIFICATIONS" sheet,
 * calls the Gemini A.I. model with project details and files, and then creates a
 * new, formatted spreadsheet with the analysis results.
 */
function runCostAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Use a non-blocking "toast" message to show processing status.
    // A toast allows the script to continue running in the background.
    // The "-1" duration means it will show until manually removed.
    spreadsheet.toast('Communicating with Gem A.I. to analyze your project...', 'Analysis in Progress...', -1);

    const specSheet = spreadsheet.getSheetByName('SPECIFICATIONS');
    if (!specSheet) {
      throw new Error('Sheet named "SPECIFICATIONS" was not found. Please ensure it exists.');
    }

    // --- 1. GATHER DATA FROM THE SHEET ---
    const projectName = specSheet.getRange('ProjectName').getValue();
    if (!projectName || projectName.trim() === '') {
      throw new Error('Project name is missing. Please enter a name in cell C7 (named "ProjectName").');
    }

    const planUrl = specSheet.getRange('PlanFiles').getValue();
    const specUrl = specSheet.getRange('SpecFiles').getValue();
    const detailsText = specSheet.getRange('Details').getValues()
      .map(row => row.join(' ').trim()) // Join all cells in a row
      .filter(rowText => rowText !== '') // Filter out empty rows
      .join('\n'); // Join all rows into a single text block

    // Check that at least some information has been provided.
    if (detailsText.trim() === '' && !planUrl && !specUrl) {
      throw new Error('No data found. Please provide project details in the "Details" range (B7:H101) or file links in cells C4 and/or H4.');
    }

    // --- 2. PREPARE FILES AND PROMPT FOR THE A.I. ---
    const fileParts = [];
    const planBlob = getDriveFileBlobFromUrl(planUrl);
    if (planBlob) {
      planBlob.setName("construction_plan_file");
      fileParts.push(planBlob);
    }

    const specBlob = getDriveFileBlobFromUrl(specUrl);
    if (specBlob) {
      specBlob.setName("project_specifications_file");
      fileParts.push(specBlob);
    }

    const prompt = `You are "Gem", a specialized A.I. COST ANALYZER for construction projects. Your task is to provide a detailed cost analysis based on the attached files (construction plans, specifications) and the following project details.

    Project Details from Spreadsheet:
    ---
    ${detailsText}
    ---

    Please analyze all provided information and generate a comprehensive cost analysis. The output must be formatted in Markdown for a spreadsheet.
    1.  Use single hash (#) for main section titles (e.g., '# Materials Cost').
    2.  Use double hash (##) for subsection titles (e.g., '## Framing').
    3.  Use standard Markdown tables for any itemized lists. Make sure tables have a header row.
    4.  Provide clear narrative explanations for your reasoning.

    Begin the analysis.`;

    // --- 3. CALL THE GEMINI A.I. MODEL ---
    // The gemini-1.5-pro-latest model is used for its multimodal capabilities (text and files).
    const generativeModel = GenerativeAi.getGenerativeModel({ model: 'gemini-1.5-pro-latest' });
    const result = generativeModel.generateContent([prompt, ...fileParts]);
    const responseText = result.text();

    // Check for an empty response from the model before proceeding.
    if (!responseText || responseText.trim() === '') {
        throw new Error('The A.I. model returned an empty response. This could be due to the content policy or an issue with the input files.');
    }

    // --- 4. CREATE AND FORMAT THE NEW WORKBOOK ---
    const newWorkbook = SpreadsheetApp.create(projectName);
    const analysisSheet = newWorkbook.getSheets()[0];
    analysisSheet.setName('AI Cost Analysis');
    analysisSheet.setHiddenGridlines(true);

    const lines = responseText.split('\n');
    let currentRow = 1;

    // Use a standard 'for' loop to reliably check the previous line (for table headers).
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmedLine = line.trim();

      if (trimmedLine.length === 0) {
        currentRow++; // Add a space for blank lines.
        continue;
      }

      const range = analysisSheet.getRange(currentRow, 1);

      if (trimmedLine.startsWith('##')) { // Subsection Title
        range.setValue(trimmedLine.replace(/##/g, '').trim());
        range.setFontFamily('Special Elite').setFontSize(14).setFontWeight('normal');
      } else if (trimmedLine.startsWith('#')) { // Main Section Title
        range.setValue(trimmedLine.replace(/#/g, '').trim());
        range.setFontFamily('Special Elite').setFontSize(14).setFontWeight('bold');
        analysisSheet.getRange(currentRow, 1, 1, 5).merge();
      } else if (trimmedLine.startsWith('|') && trimmedLine.endsWith('|')) { // Table Row
        const columns = trimmedLine.split('|').slice(1, -1).map(c => c.trim());
        const tableRowRange = analysisSheet.getRange(currentRow, 1, 1, columns.length);
        tableRowRange.setValues([columns]);
        tableRowRange.setFontFamily('McLaren').setFontSize(11).setFontWeight('normal');

        // Bold table headers. This check is more robust using the loop index 'i'.
        const prevLine = lines[i - 1]?.trim();
        if (prevLine && prevLine.includes('|') && prevLine.includes('---')) {
          tableRowRange.setFontWeight('bold').setBackground('#eeeeee');
        }
      } else { // Regular Paragraph Text
        // Merging cells allows long text to wrap nicely within a defined area.
        const textRange = analysisSheet.getRange(currentRow, 1, 1, 5);
        textRange.merge();
        textRange.setValue(trimmedLine);
        textRange.setFontFamily('McLaren').setFontSize(11).setWrap(true);
      }
      currentRow++;
    }

    // Auto-resize columns for readability
    for (let i = 1; i <= analysisSheet.getLastColumn(); i++) {
      analysisSheet.autoResizeColumn(i);
    }

    // --- 5. SHOW SUCCESS MESSAGE ---
    // Remove the "in progress" toast.
    spreadsheet.toast('Analysis complete.', '✅ Success!', 5);

    const finalUrl = newWorkbook.getUrl();
    const htmlOutput = HtmlService.createHtmlOutput(
      `<body><h3>✅ Analysis Complete!</h3><p>Your new workbook, "${projectName}", is ready.</p><p><a href="${finalUrl}" target="_blank">Click here to open the new spreadsheet.</a></p></body>`
    ).setWidth(400).setHeight(150);
    ui.showModalDialog(htmlOutput, 'Success');

  } catch (error) {
    // Remove the "in progress" toast before showing the error.
    spreadsheet.toast('An error occurred.', 'Script Failed', 5);
    // Show a detailed error message if something goes wrong.
    ui.alert('An Error Occurred', `Script failed: ${error.message}\n\nPlease check your sheet setup and file links, then try again.`, ui.ButtonSet.OK);
  }
}

/**
 * A helper function to get a file blob from a Google Drive URL.
 * It extracts the file ID from the URL and fetches the file.
 * @param {string} url The Google Drive URL.
 * @return {GoogleAppsScript.Base.Blob|null} The file blob or null if the URL is invalid or file is inaccessible.
 */
function getDriveFileBlobFromUrl(url) {
  if (!url || typeof url !== 'string' || url.trim() === '') {
    return null;
  }
  try {
    // This regex extracts the long alphanumeric ID from a Google Drive sharing URL.
    const fileId = url.match(/[-\w]{25,}/)[0];
    return DriveApp.getFileById(fileId).getBlob();
  } catch (error) {
    throw new Error(`Could not access the file at URL: ${url}. Please ensure it is a valid Google Drive link and that you have permission to view the file.`);
  }
}