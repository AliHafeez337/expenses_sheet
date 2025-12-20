/**
 * Migration Script: Add Category Summary Rows to Existing Sheet
 * 
 * This migration is broken into separate steps to avoid timeout:
 * Step 1: Add summary rows (can be run in batches)
 * Step 2: Apply formulas to summary rows (can be run in batches)
 * Step 3: Update control panel formulas (quick)
 * 
 * Run each step separately from the menu.
 */

/**
 * STEP 1: Add Summary Rows (Structure Only)
 * Adds summary rows after each category total without formulas
 * Can be run in batches if you have many categories
 */
function migrateStep1_AddSummaryRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var lastRow = sheet.getLastRow();
  var categoryTotalRows = [];
  
  // Find all category total rows
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    ui.alert('No category totals found. Make sure you have categories in your sheet.');
    return;
  }
  
  // Ask how many to process
  var batchResponse = ui.prompt(
    'Migration Step 1: Add Summary Rows',
    'Found ' + categoryTotalRows.length + ' categories.\n\n' +
    'This will add summary row structure after each category total.\n' +
    'Formulas will be added in Step 2.\n\n' +
    'How many categories to process in this run?\n' +
    '(Enter a number like 5 or 10, or "all" for all):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (batchResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var batchText = batchResponse.getResponseText().trim().toLowerCase();
  var batchSize = batchText === 'all' ? categoryTotalRows.length : parseInt(batchText);
  
  if (isNaN(batchSize) || batchSize < 1) {
    ui.alert('Invalid number. Please enter a number or "all".');
    return;
  }
  
  // Check which ones already have summary rows
  var rowsToProcess = [];
  for (var i = 0; i < categoryTotalRows.length && rowsToProcess.length < batchSize; i++) {
    var categoryTotalRow = categoryTotalRows[i];
    var nextRow = categoryTotalRow + 1;
    
    // Check if summary row already exists
    if (nextRow <= sheet.getLastRow()) {
      var nextNote = sheet.getRange(nextRow, 1).getNote();
      if (nextNote === '[CategorySummary]') {
        continue; // Skip, already exists
      }
    }
    
    rowsToProcess.push({
      originalRow: categoryTotalRow,
      index: i
    });
  }
  
  if (rowsToProcess.length === 0) {
    ui.alert('All categories already have summary rows!');
    return;
  }
  
  // Process from bottom to top to maintain row numbers
  rowsToProcess.sort(function(a, b) { return b.originalRow - a.originalRow; });
  
  var addedCount = 0;
  for (var i = 0; i < rowsToProcess.length; i++) {
    var categoryTotalRow = rowsToProcess[i].originalRow;
    
    // Insert summary row
    sheet.insertRowsAfter(categoryTotalRow, 1);
    var summaryRow = categoryTotalRow + 1;
    
    // Set labels and note
    sheet.getRange(summaryRow, 1).setNote('[CategorySummary]');
    sheet.getRange(summaryRow, 2).setValue('My total for this category');
    sheet.getRange(summaryRow, 4).setValue('Wife\'s total for this category');
    sheet.getRange(summaryRow, 6).setValue('My donations for this category');
    sheet.getRange(summaryRow, 8).setValue('Wife\'s donations for this category');
    
    // Format summary row
    sheet.getRange(summaryRow, 1, 1, 9)
      .setBackground('#e8f4f8')
      .setFontWeight('normal')
      .setFontSize(9);
    
    addedCount++;
    
    // Flush every 3 rows
    if (addedCount % 3 === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(200);
    }
  }
  
  SpreadsheetApp.flush();
  
  var remaining = categoryTotalRows.length - (categoryTotalRows.length - rowsToProcess.length) - addedCount;
  
  ui.alert('Step 1 Progress', 
    'Added ' + addedCount + ' summary row(s).\n\n' +
    (remaining > 0 ? 'Remaining: ' + remaining + ' categories.\n\nRun Step 1 again to continue.' : 'All summary rows added!\n\nProceed to Step 2.'),
    ui.ButtonSet.OK);
}

/**
 * STEP 2: Apply Formulas to Summary Rows
 * Adds formulas to existing summary rows
 * Can be run in batches
 */
function migrateStep2_ApplyFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var lastRow = sheet.getLastRow();
  var summaryRows = [];
  
  // Find all summary rows
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategorySummary]') {
      summaryRows.push(row);
    }
  }
  
  if (summaryRows.length === 0) {
    ui.alert('No summary rows found. Please run Step 1 first.');
    return;
  }
  
  // Check which ones need formulas
  var rowsNeedingFormulas = [];
  for (var i = 0; i < summaryRows.length; i++) {
    var summaryRow = summaryRows[i];
    var formulaC = sheet.getRange(summaryRow, 3).getFormula();
    if (!formulaC || formulaC === '') {
      rowsNeedingFormulas.push(summaryRow);
    }
  }
  
  if (rowsNeedingFormulas.length === 0) {
    ui.alert('All summary rows already have formulas!');
    return;
  }
  
  var batchResponse = ui.prompt(
    'Migration Step 2: Apply Formulas',
    'Found ' + rowsNeedingFormulas.length + ' summary rows needing formulas.\n\n' +
    'This will add formulas to the summary rows created in Step 1.\n\n' +
    'How many to process in this run?\n' +
    '(Enter a number like 5 or 10, or "all" for all):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (batchResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var batchText = batchResponse.getResponseText().trim().toLowerCase();
  var batchSize = batchText === 'all' ? rowsNeedingFormulas.length : parseInt(batchText);
  
  if (isNaN(batchSize) || batchSize < 1) {
    ui.alert('Invalid number. Please enter a number or "all".');
    return;
  }
  
  var processedCount = 0;
  for (var i = 0; i < rowsNeedingFormulas.length && processedCount < batchSize; i++) {
    var summaryRow = rowsNeedingFormulas[i];
    applyCategorySummaryRowFormulas(summaryRow, sheet);
    processedCount++;
    
    // Flush every 3 rows
    if (processedCount % 3 === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(200);
    }
  }
  
  SpreadsheetApp.flush();
  
  var remaining = rowsNeedingFormulas.length - processedCount;
  
  ui.alert('Step 2 Progress', 
    'Applied formulas to ' + processedCount + ' summary row(s).\n\n' +
    (remaining > 0 ? 'Remaining: ' + remaining + ' rows.\n\nRun Step 2 again to continue.' : 'All formulas applied!\n\nProceed to Step 3.'),
    ui.ButtonSet.OK);
}

/**
 * STEP 3: Update Control Panel Formulas
 * Quick step - updates B5, B6, B12, B13 to use summary rows
 */
function migrateStep3_UpdateControlPanel() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    'Migration Step 3: Update Control Panel',
    'This will update control panel formulas (B5, B6, B12, B13) to use category summary rows.\n\n' +
    'This is a quick operation.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  updateControlPanelSummaries();
  SpreadsheetApp.flush();
  
  ui.alert('Step 3 Complete!', 
    'Control panel formulas have been updated to use category summary rows.\n\n' +
    'Migration complete!',
    ui.ButtonSet.OK);
}

/**
 * Helper: Check Migration Status
 * Shows which steps are complete
 */
function checkMigrationStatus() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();
  
  var categoryTotalRows = [];
  var summaryRows = [];
  var summaryRowsWithFormulas = 0;
  
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    } else if (note === '[CategorySummary]') {
      summaryRows.push(row);
      var formulaB = sheet.getRange(row, 2).getFormula();
      if (formulaB && formulaB !== '') {
        summaryRowsWithFormulas++;
      }
    }
  }
  
  var status = 'MIGRATION STATUS\n';
  status += '================\n\n';
  status += 'Total Categories: ' + categoryTotalRows.length + '\n';
  status += 'Summary Rows Added: ' + summaryRows.length + '\n';
  status += 'Summary Rows with Formulas: ' + summaryRowsWithFormulas + '\n\n';
  
  if (summaryRows.length === 0) {
    status += '❌ Step 1: Not started\n';
    status += '❌ Step 2: Not started\n';
    status += '❌ Step 3: Not started\n';
  } else if (summaryRows.length < categoryTotalRows.length) {
    status += '⏳ Step 1: In progress (' + summaryRows.length + '/' + categoryTotalRows.length + ')\n';
    status += '⏳ Step 2: Waiting\n';
    status += '⏳ Step 3: Waiting\n';
  } else if (summaryRowsWithFormulas < summaryRows.length) {
    status += '✅ Step 1: Complete\n';
    status += '⏳ Step 2: In progress (' + summaryRowsWithFormulas + '/' + summaryRows.length + ')\n';
    status += '⏳ Step 3: Waiting\n';
  } else {
    // Check if control panel uses summary rows
    var b5Formula = sheet.getRange('B5').getFormula();
    var usesSummaryRows = (b5Formula.indexOf('C') >= 0 || b5Formula.indexOf('E') >= 0) && b5Formula.split('+').length <= categoryTotalRows.length + 5;
    
    status += '✅ Step 1: Complete\n';
    status += '✅ Step 2: Complete\n';
    if (usesSummaryRows) {
      status += '✅ Step 3: Complete\n';
      status += '\n🎉 Migration Complete!';
    } else {
      status += '⏳ Step 3: Needs update\n';
    }
  }
  
  ui.alert('Migration Status', status, ui.ButtonSet.OK);
}

/**
 * Apply formulas to category summary row
 * (Same function as in other scripts)
 */
function applyCategorySummaryRowFormulas(row, sheet) {
  // Find the category total row (should be the row before this summary row)
  var categoryTotalRow = row - 1;
  
  // Find all [Me] and [Wife] rows in this category
  var meRows = [];
  var wifeRows = [];
  
  // Search backwards from category total row to find category header
  var categoryHeaderRow = -1;
  for (var r = categoryTotalRow - 1; r >= 27; r--) {
    var cellValue = sheet.getRange(r, 1).getValue();
    var note = sheet.getRange(r, 1).getNote();
    
    // If we hit a category header (has value but no note), stop
    if (cellValue !== '' && !note) {
      categoryHeaderRow = r;
      break;
    }
    
    // Collect [Me] and [Wife] rows
    if (note === '[Me]') {
      meRows.push(r);
    } else if (note === '[Wife]') {
      wifeRows.push(r);
    }
  }
  
  if (categoryHeaderRow === -1) {
    return; // Couldn't find category header
  }
  
  // Column C: Sum of all [Me] rows' Column B
  if (meRows.length > 0) {
    var meTotalTerms = [];
    for (var i = 0; i < meRows.length; i++) {
      meTotalTerms.push('B' + meRows[i]);
    }
    sheet.getRange(row, 3).setFormula('=' + meTotalTerms.join('+'));
  }
  
  // Column E: Sum of all [Wife] rows' Column B
  if (wifeRows.length > 0) {
    var wifeTotalTerms = [];
    for (var i = 0; i < wifeRows.length; i++) {
      wifeTotalTerms.push('B' + wifeRows[i]);
    }
    sheet.getRange(row, 5).setFormula('=' + wifeTotalTerms.join('+'));
  }
  
  // Column G: Sum of all [Me] rows' donation columns (all 31 days)
  if (meRows.length > 0) {
    var myDonationTerms = [];
    for (var i = 0; i < meRows.length; i++) {
      var meRow = meRows[i];
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        myDonationTerms.push(getColumnLetter(baseCol + 3) + meRow);
      }
    }
    if (myDonationTerms.length > 0) {
      sheet.getRange(row, 7).setFormula('=' + myDonationTerms.join('+'));
    }
  }
  
  // Column I: Sum of all [Wife] rows' donation columns (all 31 days)
  if (wifeRows.length > 0) {
    var wifeDonationTerms = [];
    for (var i = 0; i < wifeRows.length; i++) {
      var wifeRow = wifeRows[i];
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        wifeDonationTerms.push(getColumnLetter(baseCol + 3) + wifeRow);
      }
    }
    if (wifeDonationTerms.length > 0) {
      sheet.getRange(row, 9).setFormula('=' + wifeDonationTerms.join('+'));
    }
  }
}

/**
 * Update control panel summaries to use category summary rows
 * (Same function as in other scripts)
 */
function updateControlPanelSummaries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Find all category summary rows
  var summaryRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategorySummary]') {
      summaryRows.push(row);
    }
  }
  
  if (summaryRows.length === 0) {
    // Fallback: if no summary rows exist, use old method
    var meRows = [];
    var wifeRows = [];
    var myDonationTerms = [];
    var wifeDonationTerms = [];
    
    for (var row = 28; row <= lastRow; row++) {
      var note = sheet.getRange(row, 1).getNote();
      if (note === '[Me]') {
        meRows.push('B' + row);
        for (var day = 1; day <= 31; day++) {
          var baseCol = 3 + (day - 1) * 4;
          myDonationTerms.push(getColumnLetter(baseCol + 3) + row);
        }
      } else if (note === '[Wife]') {
        wifeRows.push('B' + row);
        for (var day = 1; day <= 31; day++) {
          var baseCol = 3 + (day - 1) * 4;
          wifeDonationTerms.push(getColumnLetter(baseCol + 3) + row);
        }
      }
    }
    
    if (meRows.length > 0) {
      sheet.getRange('B5').setFormula('=' + meRows.join('+'));
    }
    if (wifeRows.length > 0) {
      sheet.getRange('B6').setFormula('=' + wifeRows.join('+'));
    }
    if (myDonationTerms.length > 0) {
      sheet.getRange('B12').setFormula('=' + myDonationTerms.join('+'));
    }
    if (wifeDonationTerms.length > 0) {
      sheet.getRange('B13').setFormula('=' + wifeDonationTerms.join('+'));
    }
    return;
  }
  
  // NEW METHOD: Use category summary rows
  // B5: Sum of all summary rows' Column C (My total for this category)
  var myTotalTerms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    myTotalTerms.push('C' + summaryRows[i]);
  }
  if (myTotalTerms.length > 0) {
    sheet.getRange('B5').setFormula('=' + myTotalTerms.join('+'));
  }
  
  // B6: Sum of all summary rows' Column E (Wife's total for this category)
  var wifeTotalTerms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    wifeTotalTerms.push('E' + summaryRows[i]);
  }
  if (wifeTotalTerms.length > 0) {
    sheet.getRange('B6').setFormula('=' + wifeTotalTerms.join('+'));
  }
  
  // B12: Sum of all summary rows' Column G (My donations for this category)
  var myDonationTerms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    myDonationTerms.push('G' + summaryRows[i]);
  }
  if (myDonationTerms.length > 0) {
    sheet.getRange('B12').setFormula('=' + myDonationTerms.join('+'));
  }
  
  // B13: Sum of all summary rows' Column I (Wife's donations for this category)
  var wifeDonationTerms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    wifeDonationTerms.push('I' + summaryRows[i]);
  }
  if (wifeDonationTerms.length > 0) {
    sheet.getRange('B13').setFormula('=' + wifeDonationTerms.join('+'));
  }
}

/**
 * Helper function to convert column number to letter
 */
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

