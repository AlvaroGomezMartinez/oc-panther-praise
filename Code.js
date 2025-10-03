/**
 * @fileoverview oc-panther-praise
 *
 * This Google Apps Script project automates the generation of personalized praise slides
 * for teachers based on form responses collected in a Google Sheets spreadsheet.
 *
 * The script processes new form submissions, uses a template slide from a Google Slides
 * presentation, replaces placeholder text with actual praise details, and appends the
 * customized slides to a target presentation. It tracks processed timestamps using
 * Script Properties to prevent duplicate processing of submissions.
 *
 * Key features:
 * - Integrates with Google Sheets, Google Slides, Google Forms, and Google Sites.
 * - Handles data validation and error logging.
 * - Maintains state to avoid reprocessing existing entries. In other words, it only creates slides for new form submissions.
 *
 * Required APIs:
 * - Google Slides API (enabled as advanced service in appsscript.json)
 *
 * Setup Instructions:
 * 1. Create a Google Form with the following fields:
 *    - Short answer (open-ended): "YOUR first and last name:"
 *    - Short answer (open-ended): "Name of the OC Staff member you are praising:"
 *    - Dropdown: "What is the email of the person you are rewarding? Dropdown or type the name you are looking for." (populate with every teacher's email)
 *    - Short answer (open-ended): "Why are they so awesome? This will appear in the email to the recipient."
 *    Ensure the form collects submitter emails.
 * 2. Link the form to a Google Sheets spreadsheet to store the responses.
 * 3. Create two Google Slides presentations:
 *    - One for the template slide (containing placeholders like "<<Name of the OC Staff member you are praising:>>", "<<Why are they so awesome? This will appear in the email to the recipient.>>", "<<YOUR first and last name:>>").
 *    - Another for the output where the merged slides will be appended.
 * 4. Create a bound Google Apps Script project in the Google Sheets spreadsheet from step 2 above. Insert this script into the Apps Script IDE.
 * 5. In the spreadsheet, create a new sheet named "Setup". In cell A2, enter "Template Google Slide ID:", and in A3, enter "Target Google Slide ID:". Enter the respective presentation IDs in B2 and B3.
 * 6. Set up an "On form submit" trigger for the textMerging() function to automate slide creation upon new submissions.
 * 7. Publish the target Google Slides presentation to the web and embed it in a Google Site for easy access and visibility.
 *
 * Additional Setup:
 * If you want to send the Panther Praise certificate to the recipient, you can use Autocrat to automate that process.
 * 
 * Implementation with People who will Submit the Panther Praise:
 * 1. Create a QR-code or link to the Google Form and share it with staff to encourage submissions.
 *
 * Troubleshooting:
 * - Ensure all IDs (template presentation, target presentation) are correctly entered in the Setup sheet.
 * - Check that the form fields match the expected structure.
 * - Review logs for any errors during execution.
 * - Clear out Script Properties in the Google Apps Script project's Properties Service if necessary to reprocess entries.
 * - Maintain current emails in the Google Form dropdown.
 * - Use the testTextMerging() function to manually test the script without a form submission.
 * 
 * 
 * @summary Automates praise slide creation from form responses to celebrate teacher achievements.
 * @author Academic Technology Coaches - Meredith Jones & Alvaro Gomez
 * @since 2025-10-03
*/

/**
 * Retrieves unprocessed rows from the form responses sheet.
 * @param {Sheet} sheet - The 'Form Responses 1' sheet.
 * @param {Array<string>} processedIds - Array of already processed timestamps.
 * @returns {Array<Object>} Array of unprocessed row objects with timestamp, teacherName, praise, and fromName.
 */
function getUnprocessedRows(sheet, processedIds) {
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  const unprocessed = [];
  for (let i = 0; i < values.length; ++i) {
    const row = values[i];
    const rawTimestamp = row[0];
    if (!rawTimestamp) continue;
    const timestamp = new Date(rawTimestamp).toISOString();
    const teacherName = row[3].toString().trim();
    const praise = row[5].toString().trim().replace(/\n/g, ' ');
    const fromName = row[2].toString().trim();
    if (processedIds.includes(timestamp) || !teacherName || !praise || !fromName) continue;
    unprocessed.push({ timestamp, teacherName, praise, fromName });
  }
  return unprocessed;
}

/**
 * Creates a new praise slide by copying the template and replacing placeholders.
 * @param {Slide} templateSlide - The template slide to copy.
 * @param {Presentation} targetPresentation - The presentation to append the new slide to.
 * @param {Object} row - The row data with teacherName, praise, and fromName.
 */
function createPraiseSlide(templateSlide, targetPresentation, row) {
  const newSlide = targetPresentation.appendSlide(templateSlide);
  newSlide.replaceAllText('<<Name of the OC Staff member you are praising:>>', row.teacherName);
  newSlide.replaceAllText('<<Why are they so awesome? This will appear in the email to the recipient.>>', row.praise);
  newSlide.replaceAllText('<<YOUR first and last name:>>', row.fromName);
  console.log(`✅ Added slide for ${row.teacherName}`);
}

/**
 * Updates the stored processed timestamps in script properties.
 * @param {Properties} scriptProps - The script properties service.
 * @param {Array<string>} processedIds - The updated array of processed timestamps.
 */
function updateProcessedIds(scriptProps, processedIds) {
  scriptProps.setProperty('processedTimestamps', JSON.stringify(processedIds));
}

/**
 * Merges text from a Google Sheets spreadsheet into a Google Slides presentation.
 * This function processes form responses to create personalized praise slides for teachers.
 * It uses a template slide and appends new slides to the target presentation based on new data.
 * Processed timestamps are tracked to avoid duplicates.
 * Presentation IDs are read from the 'Setup' sheet in the bound spreadsheet.
 * @returns {string} A success message indicating the number of new slides added, or an error message if something fails.
 */
function textMerging() {
  // Get configuration from Setup sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = spreadsheet.getSheetByName('Setup');
  if (!setupSheet) {
    throw new Error('Setup sheet not found. Please create a "Setup" sheet with Template Google Slide ID in B2 and Target Google Slide ID in B3.');
  }
  const templatePresentationId = setupSheet.getRange('B2').getValue();
  const targetPresentationId = setupSheet.getRange('B3').getValue();
  const dataSpreadsheetId = spreadsheet.getId();

  if (!templatePresentationId || !targetPresentationId) {
    throw new Error('Template or Target Presentation ID is missing in the Setup sheet.');
  }

  try {
    // Log file IDs
    console.log('Template ID:', templatePresentationId);
    console.log('Target ID:', targetPresentationId);
    console.log('Spreadsheet ID:', dataSpreadsheetId);

    // Open sheet and get data
    let sheet, templatePresentation, templateSlide, targetPresentation;
    try {
      sheet = SpreadsheetApp.openById(dataSpreadsheetId).getSheetByName('Form Responses 1');
    } catch (err) {
      throw new Error(`Failed to open spreadsheet or retrieve data: ${err.message}`);
    }

    try {
      // Open template and target decks
      templatePresentation = SlidesApp.openById(templatePresentationId);
      templateSlide = templatePresentation.getSlides()[0];
      targetPresentation = SlidesApp.openById(targetPresentationId);
    } catch (err) {
      throw new Error(`Failed to open presentations: ${err.message}`);
    }

    // Use ScriptProperties to store processed timestamps
    const scriptProps = PropertiesService.getScriptProperties();
    let processedIds = JSON.parse(scriptProps.getProperty('processedTimestamps') || '[]');

    // Get unprocessed rows
    const unprocessedRows = getUnprocessedRows(sheet, processedIds);

    // Process each unprocessed row
    const newProcessed = [];
    for (const row of unprocessedRows) {
      createPraiseSlide(templateSlide, targetPresentation, row);
      newProcessed.push(row.timestamp);
    }

    // Update stored processed IDs
    updateProcessedIds(scriptProps, processedIds.concat(newProcessed));

    console.log(`✅ Completed: ${newProcessed.length} new slides added.`);
    return `✅ Successfully added ${newProcessed.length} new slides.`;

  } catch (err) {
    console.error('❌ Error object:', err);
    const errorMsg = `❌ Failed with error: ${err.message || err.toString()}`;
    console.log(errorMsg);
    return errorMsg;
  }
}

/**
 * Test function to simulate textMerging without a real trigger.
 * Useful for debugging and manual runs. Run this from the Apps Script editor.
 */
function testTextMerging() {
  console.log('Testing textMerging...');
  const result = textMerging();
  console.log('Test result:', result);
}
