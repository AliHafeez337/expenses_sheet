/**
 * Diagnostic Script for Category Formulas
 * 
 * This script checks formulas in a specific category to identify any errors.
 * Run this function from the script editor or add it to a custom menu.
 */
function diagnoseCategoryFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Ask which category to diagnose
  var categoryResponse = ui.prompt(
    'Diagnose Category Formulas',
    'Enter the category name to check:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (categoryResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('Operation cancelled.');
    return;
  }
  
  var categoryName = categoryResponse.getResponseText().trim();
  if (categoryName === '') {
    ui.alert('Category name cannot be empty!');
    return;
  }
  
  // Find the category in the sheet
  var lastRow = sheet.getLastRow();
  var categoryHeaderRow = -1;
  var categoryTotalRow = -1;
  
  for (var row = lastRow; row >= 27; row--) {
    var cellValue = sheet.getRange(row, 1).getValue();
    var note = sheet.getRange(row, 1).getNote();
    
    // Check if this is the category total row
    if (cellValue === categoryName + ' TOTAL' && note === '[CategoryTotal]') {
      categoryTotalRow = row;
    }
    
    // Check if this is the category header
    if (cellValue === categoryName && !note) {
      categoryHeaderRow = row;
      break;
    }
  }
  
  if (categoryHeaderRow === -1 || categoryTotalRow === -1) {
    ui.alert('Error', 'Category "' + categoryName + '" not found! Please check the spelling.', ui.ButtonSet.OK);
    return;
  }
  
  // Collect all issues
  var issues = [];
  var subcategoryCount = 0;
  
  // Find all subcategories in this category
  var subcategoryRows = [];
  for (var row = categoryHeaderRow + 1; row < categoryTotalRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[Totals]') {
      subcategoryRows.push({
        totalsRow: row,
        meRow: row + 1,
        wifeRow: row + 2,
        commentRow: row + 3,
        name: sheet.getRange(row, 1).getValue()
      });
      subcategoryCount++;
    }
  }
  
  if (subcategoryRows.length === 0) {
    ui.alert('No subcategories found in category "' + categoryName + '"');
    return;
  }
  
  // Check each subcategory's formulas
  for (var i = 0; i < subcategoryRows.length; i++) {
    var subcat = subcategoryRows[i];
    var subcatName = subcat.name || ('Subcategory #' + (i + 1));
    
    // Check [Totals] row formulas
    checkTotalsRowFormulas(subcat.totalsRow, subcat.meRow, subcat.wifeRow, sheet, issues, subcatName);
    
    // Check [Me] row formulas
    checkMeRowFormulas(subcat.meRow, sheet, issues, subcatName);
    
    // Check [Wife] row formulas
    checkWifeRowFormulas(subcat.wifeRow, sheet, issues, subcatName);
  }
  
  // Check category total row formulas
  checkCategoryTotalFormulas(categoryTotalRow, subcategoryRows, sheet, issues, categoryName);
  
  // Check category summary row formulas (if it exists)
  var categorySummaryRow = categoryTotalRow + 1;
  var summaryNote = sheet.getRange(categorySummaryRow, 1).getNote();
  if (summaryNote === '[CategorySummary]') {
    checkCategorySummaryRowFormulas(categorySummaryRow, subcategoryRows, sheet, issues, categoryName);
  }
  
  // Report results
  reportResults(issues, categoryName, subcategoryCount, ui);
}

/**
 * Diagnostic Script for Global Formulas (Control Panel & Grand Total)
 * 
 * This script checks formulas that depend on ALL categories:
 * - Control Panel summaries (B5, B6, B12, B13)
 * - Grand Total Row (Row 26)
 */
function diagnoseGlobalFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Collect all issues
  var issues = [];
  
  // Check control panel formulas
  checkControlPanelFormulas(sheet, issues);
  
  // Check grand total formulas
  checkGrandTotalFormulas(sheet, issues);
  
  // Report results
  if (issues.length === 0) {
    ui.alert('Diagnosis Complete', 
      '✓ All global formulas are correct!\n\n' +
      'Checked:\n' +
      '• Control Panel summaries (B5, B6, B12, B13)\n' +
      '• Grand Total Row (Row 26)',
      ui.ButtonSet.OK);
    return;
  }
  
  // Build detailed report
  var report = 'GLOBAL FORMULAS DIAGNOSTIC REPORT\n';
  report += '====================================\n\n';
  report += 'Issues found: ' + issues.length + '\n\n';
  report += 'DETAILED ISSUES:\n';
  report += '================\n\n';
  
  for (var i = 0; i < issues.length; i++) {
    var issue = issues[i];
    report += (i + 1) + '. ' + issue.type + ': ' + issue.location + '\n';
    report += '   Row: ' + issue.row + ', Column: ' + getColumnLetter(issue.col) + issue.col + '\n';
    report += '   Expected: ' + issue.expected + '\n';
    report += '   Actual: ' + issue.actual + '\n\n';
  }
  
  // Show in alert (truncated if too long)
  if (report.length > 1000) {
    report = report.substring(0, 1000) + '\n\n... (showing first 1000 chars, see logs for full report)';
  }
  
  Logger.log('=== GLOBAL FORMULAS DIAGNOSTIC REPORT ===\n' + report);
  
  var response = ui.alert('Issues Found', 
    'Found ' + issues.length + ' formula error(s) in global formulas.\n\n' +
    'Check the execution log (View > Logs) for detailed information.\n\n' +
    'Would you like to fix these errors automatically?',
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    fixGlobalFormulas(issues, ui);
  }
}

/**
 * Check formulas in a [Totals] row
 */
function checkTotalsRowFormulas(totalsRow, meRow, wifeRow, sheet, issues, subcatName) {
  // Check monthly total (Column B)
  var expectedMonthlyFormula = buildMonthlyTotalFormula(totalsRow);
  var actualFormula = sheet.getRange(totalsRow, 2).getFormula();
  if (!formulasAreEquivalent(actualFormula, expectedMonthlyFormula)) {
    issues.push({
      type: 'ERROR',
      location: subcatName + ' [Totals] row, Column B (Monthly Total)',
      expected: expectedMonthlyFormula,
      actual: actualFormula || '(empty)',
      row: totalsRow,
      col: 2
    });
  }
  
  // Check each day's formulas
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total (baseCol) = Personal + Family + Donation (all from same row)
    var expectedDayTotal = '=' + getColumnLetter(baseCol + 1) + totalsRow + 
                          '+' + getColumnLetter(baseCol + 2) + totalsRow + 
                          '+' + getColumnLetter(baseCol + 3) + totalsRow;
    var actualDayTotal = sheet.getRange(totalsRow, baseCol).getFormula();
    if (!formulasAreEquivalent(actualDayTotal, expectedDayTotal)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Totals] row, Day ' + day + ' Total',
        expected: expectedDayTotal,
        actual: actualDayTotal || '(empty)',
        row: totalsRow,
        col: baseCol
      });
    }
    
    // Personal Total (baseCol + 1) = Me Personal + Wife Personal
    var expectedPersonal = '=' + getColumnLetter(baseCol + 1) + meRow + 
                          '+' + getColumnLetter(baseCol + 1) + wifeRow;
    var actualPersonal = sheet.getRange(totalsRow, baseCol + 1).getFormula();
    if (!formulasAreEquivalent(actualPersonal, expectedPersonal)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Totals] row, Day ' + day + ' Personal',
        expected: expectedPersonal,
        actual: actualPersonal || '(empty)',
        row: totalsRow,
        col: baseCol + 1
      });
    }
    
    // Family Total (baseCol + 2) = Me Family + Wife Family
    var expectedFamily = '=' + getColumnLetter(baseCol + 2) + meRow + 
                        '+' + getColumnLetter(baseCol + 2) + wifeRow;
    var actualFamily = sheet.getRange(totalsRow, baseCol + 2).getFormula();
    if (!formulasAreEquivalent(actualFamily, expectedFamily)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Totals] row, Day ' + day + ' Family',
        expected: expectedFamily,
        actual: actualFamily || '(empty)',
        row: totalsRow,
        col: baseCol + 2
      });
    }
    
    // Donation Total (baseCol + 3) = Me Donation + Wife Donation
    var expectedDonation = '=' + getColumnLetter(baseCol + 3) + meRow + 
                          '+' + getColumnLetter(baseCol + 3) + wifeRow;
    var actualDonation = sheet.getRange(totalsRow, baseCol + 3).getFormula();
    if (!formulasAreEquivalent(actualDonation, expectedDonation)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Totals] row, Day ' + day + ' Donation',
        expected: expectedDonation,
        actual: actualDonation || '(empty)',
        row: totalsRow,
        col: baseCol + 3
      });
    }
  }
}

/**
 * Check formulas in a [Me] row
 */
function checkMeRowFormulas(meRow, sheet, issues, subcatName) {
  // Check monthly total (Column B)
  var expectedMonthlyFormula = buildMonthlyTotalFormula(meRow);
  var actualFormula = sheet.getRange(meRow, 2).getFormula();
  if (!formulasAreEquivalent(actualFormula, expectedMonthlyFormula)) {
    issues.push({
      type: 'ERROR',
      location: subcatName + ' [Me] row, Column B (Monthly Total)',
      expected: expectedMonthlyFormula,
      actual: actualFormula || '(empty)',
      row: meRow,
      col: 2
    });
  }
  
  // Check each day's total formula
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total = Personal + Family + Donation (all from same row)
    var expectedDayTotal = '=' + getColumnLetter(baseCol + 1) + meRow + 
                          '+' + getColumnLetter(baseCol + 2) + meRow + 
                          '+' + getColumnLetter(baseCol + 3) + meRow;
    var actualDayTotal = sheet.getRange(meRow, baseCol).getFormula();
    if (!formulasAreEquivalent(actualDayTotal, expectedDayTotal)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Me] row, Day ' + day + ' Total',
        expected: expectedDayTotal,
        actual: actualDayTotal || '(empty)',
        row: meRow,
        col: baseCol
      });
    }
  }
}

/**
 * Check formulas in a [Wife] row
 */
function checkWifeRowFormulas(wifeRow, sheet, issues, subcatName) {
  // Check monthly total (Column B)
  var expectedMonthlyFormula = buildMonthlyTotalFormula(wifeRow);
  var actualFormula = sheet.getRange(wifeRow, 2).getFormula();
  if (!formulasAreEquivalent(actualFormula, expectedMonthlyFormula)) {
    issues.push({
      type: 'ERROR',
      location: subcatName + ' [Wife] row, Column B (Monthly Total)',
      expected: expectedMonthlyFormula,
      actual: actualFormula || '(empty)',
      row: wifeRow,
      col: 2
    });
  }
  
  // Check each day's total formula
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total = Personal + Family + Donation (all from same row)
    var expectedDayTotal = '=' + getColumnLetter(baseCol + 1) + wifeRow + 
                          '+' + getColumnLetter(baseCol + 2) + wifeRow + 
                          '+' + getColumnLetter(baseCol + 3) + wifeRow;
    var actualDayTotal = sheet.getRange(wifeRow, baseCol).getFormula();
    if (!formulasAreEquivalent(actualDayTotal, expectedDayTotal)) {
      issues.push({
        type: 'ERROR',
        location: subcatName + ' [Wife] row, Day ' + day + ' Total',
        expected: expectedDayTotal,
        actual: actualDayTotal || '(empty)',
        row: wifeRow,
        col: baseCol
      });
    }
  }
}

/**
 * Check formulas in category total row
 */
function checkCategoryTotalFormulas(categoryTotalRow, subcategoryRows, sheet, issues, categoryName) {
  // Build expected monthly formula
  var expectedMonthlyTerms = [];
  for (var i = 0; i < subcategoryRows.length; i++) {
    expectedMonthlyTerms.push('B' + subcategoryRows[i].totalsRow);
  }
  var expectedMonthlyFormula = '=' + expectedMonthlyTerms.join('+');
  var actualMonthlyFormula = sheet.getRange(categoryTotalRow, 2).getFormula();
  
  if (!formulasAreEquivalent(actualMonthlyFormula, expectedMonthlyFormula)) {
    issues.push({
      type: 'ERROR',
      location: categoryName + ' TOTAL row, Column B (Monthly Total)',
      expected: expectedMonthlyFormula,
      actual: actualMonthlyFormula || '(empty)',
      row: categoryTotalRow,
      col: 2
    });
  }
  
  // Check each day's formulas
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total
    var expectedDayTotalTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      expectedDayTotalTerms.push(getColumnLetter(baseCol) + subcategoryRows[i].totalsRow);
    }
    var expectedDayTotal = '=' + expectedDayTotalTerms.join('+');
    var actualDayTotal = sheet.getRange(categoryTotalRow, baseCol).getFormula();
    if (!formulasAreEquivalent(actualDayTotal, expectedDayTotal)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' TOTAL row, Day ' + day + ' Total',
        expected: expectedDayTotal,
        actual: actualDayTotal || '(empty)',
        row: categoryTotalRow,
        col: baseCol
      });
    }
    
    // Personal Total
    var expectedPersonalTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      expectedPersonalTerms.push(getColumnLetter(baseCol + 1) + subcategoryRows[i].totalsRow);
    }
    var expectedPersonal = '=' + expectedPersonalTerms.join('+');
    var actualPersonal = sheet.getRange(categoryTotalRow, baseCol + 1).getFormula();
    if (!formulasAreEquivalent(actualPersonal, expectedPersonal)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' TOTAL row, Day ' + day + ' Personal',
        expected: expectedPersonal,
        actual: actualPersonal || '(empty)',
        row: categoryTotalRow,
        col: baseCol + 1
      });
    }
    
    // Family Total
    var expectedFamilyTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      expectedFamilyTerms.push(getColumnLetter(baseCol + 2) + subcategoryRows[i].totalsRow);
    }
    var expectedFamily = '=' + expectedFamilyTerms.join('+');
    var actualFamily = sheet.getRange(categoryTotalRow, baseCol + 2).getFormula();
    if (!formulasAreEquivalent(actualFamily, expectedFamily)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' TOTAL row, Day ' + day + ' Family',
        expected: expectedFamily,
        actual: actualFamily || '(empty)',
        row: categoryTotalRow,
        col: baseCol + 2
      });
    }
    
    // Donation Total
    var expectedDonationTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      expectedDonationTerms.push(getColumnLetter(baseCol + 3) + subcategoryRows[i].totalsRow);
    }
    var expectedDonation = '=' + expectedDonationTerms.join('+');
    var actualDonation = sheet.getRange(categoryTotalRow, baseCol + 3).getFormula();
    if (!formulasAreEquivalent(actualDonation, expectedDonation)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' TOTAL row, Day ' + day + ' Donation',
        expected: expectedDonation,
        actual: actualDonation || '(empty)',
        row: categoryTotalRow,
        col: baseCol + 3
      });
    }
  }
}

/**
 * Check formulas in category summary row
 */
function checkCategorySummaryRowFormulas(summaryRow, subcategoryRows, sheet, issues, categoryName) {
  // Find all [Me] and [Wife] rows in this category
  var meRows = [];
  var wifeRows = [];
  
  for (var i = 0; i < subcategoryRows.length; i++) {
    meRows.push(subcategoryRows[i].meRow);
    wifeRows.push(subcategoryRows[i].wifeRow);
  }
  
  // Check Column C: My total for this category
  if (meRows.length > 0) {
    var expectedMeTotalTerms = [];
    for (var i = 0; i < meRows.length; i++) {
      expectedMeTotalTerms.push('B' + meRows[i]);
    }
    var expectedMeTotal = '=' + expectedMeTotalTerms.join('+');
    var actualMeTotal = sheet.getRange(summaryRow, 3).getFormula();
    if (!formulasAreEquivalent(actualMeTotal, expectedMeTotal)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' Summary Row - My Total (C)',
        expected: expectedMeTotal,
        actual: actualMeTotal || '(empty)',
        row: summaryRow,
        col: 3
      });
    }
  }
  
  // Check Column E: Wife's total for this category
  if (wifeRows.length > 0) {
    var expectedWifeTotalTerms = [];
    for (var i = 0; i < wifeRows.length; i++) {
      expectedWifeTotalTerms.push('B' + wifeRows[i]);
    }
    var expectedWifeTotal = '=' + expectedWifeTotalTerms.join('+');
    var actualWifeTotal = sheet.getRange(summaryRow, 5).getFormula();
    if (!formulasAreEquivalent(actualWifeTotal, expectedWifeTotal)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' Summary Row - Wife\'s Total (E)',
        expected: expectedWifeTotal,
        actual: actualWifeTotal || '(empty)',
        row: summaryRow,
        col: 5
      });
    }
  }
  
  // Check Column G: My donations for this category
  if (meRows.length > 0) {
    var expectedMyDonationTerms = [];
    for (var i = 0; i < meRows.length; i++) {
      var meRow = meRows[i];
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        expectedMyDonationTerms.push(getColumnLetter(baseCol + 3) + meRow);
      }
    }
    var expectedMyDonation = '=' + expectedMyDonationTerms.join('+');
    var actualMyDonation = sheet.getRange(summaryRow, 7).getFormula();
    if (!formulasAreEquivalent(actualMyDonation, expectedMyDonation)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' Summary Row - My Donations (G)',
        expected: expectedMyDonation,
        actual: actualMyDonation || '(empty)',
        row: summaryRow,
        col: 7
      });
    }
  }
  
  // Check Column I: Wife's donations for this category
  if (wifeRows.length > 0) {
    var expectedWifeDonationTerms = [];
    for (var i = 0; i < wifeRows.length; i++) {
      var wifeRow = wifeRows[i];
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        expectedWifeDonationTerms.push(getColumnLetter(baseCol + 3) + wifeRow);
      }
    }
    var expectedWifeDonation = '=' + expectedWifeDonationTerms.join('+');
    var actualWifeDonation = sheet.getRange(summaryRow, 9).getFormula();
    if (!formulasAreEquivalent(actualWifeDonation, expectedWifeDonation)) {
      issues.push({
        type: 'ERROR',
        location: categoryName + ' Summary Row - Wife\'s Donations (H)',
        expected: expectedWifeDonation,
        actual: actualWifeDonation || '(empty)',
        row: summaryRow,
        col: 8
      });
    }
  }
}

/**
 * Build expected monthly total formula (sum of all day totals)
 */
function buildMonthlyTotalFormula(row) {
  var terms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    terms.push(getColumnLetter(baseCol) + row);
  }
  return '=' + terms.join('+');
}

/**
 * Check control panel summary formulas (B5, B6, B12, B13)
 * NOW USES CATEGORY SUMMARY ROWS instead of individual [Me]/[Wife] rows
 */
function checkControlPanelFormulas(sheet, issues) {
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
    // Fallback: Check using old method (individual [Me]/[Wife] rows)
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
    
    // Check B5: My Monthly Total
    if (meRows.length > 0) {
      var expectedB5 = '=' + meRows.join('+');
      var actualB5 = sheet.getRange('B5').getFormula();
      if (!formulasAreEquivalent(actualB5, expectedB5)) {
        issues.push({
          type: 'ERROR',
          location: 'Control Panel - My Monthly Total (B5)',
          expected: expectedB5,
          actual: actualB5 || '(empty)',
          row: 5,
          col: 2
        });
      }
    }
    
    // Check B6: Wife's Monthly Total
    if (wifeRows.length > 0) {
      var expectedB6 = '=' + wifeRows.join('+');
      var actualB6 = sheet.getRange('B6').getFormula();
      if (!formulasAreEquivalent(actualB6, expectedB6)) {
        issues.push({
          type: 'ERROR',
          location: 'Control Panel - Wife\'s Monthly Total (B6)',
          expected: expectedB6,
          actual: actualB6 || '(empty)',
          row: 6,
          col: 2
        });
      }
    }
    
    // Check B12: My Total Donation
    if (myDonationTerms.length > 0) {
      var expectedB12 = '=' + myDonationTerms.join('+');
      var actualB12 = sheet.getRange('B12').getFormula();
      if (!formulasAreEquivalent(actualB12, expectedB12)) {
        issues.push({
          type: 'ERROR',
          location: 'Control Panel - My Total Donation (B12)',
          expected: expectedB12,
          actual: actualB12 || '(empty)',
          row: 12,
          col: 2
        });
      }
    }
    
    // Check B13: Wife's Total Donation
    if (wifeDonationTerms.length > 0) {
      var expectedB13 = '=' + wifeDonationTerms.join('+');
      var actualB13 = sheet.getRange('B13').getFormula();
      if (!formulasAreEquivalent(actualB13, expectedB13)) {
        issues.push({
          type: 'ERROR',
          location: 'Control Panel - Wife\'s Total Donation (B13)',
          expected: expectedB13,
          actual: actualB13 || '(empty)',
          row: 13,
          col: 2
        });
      }
    }
    return;
  }
  
  // NEW METHOD: Use category summary rows
  // Check B5: My Monthly Total (sum of all summary rows' Column C)
  var expectedB5Terms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    expectedB5Terms.push('C' + summaryRows[i]);
  }
  var expectedB5 = '=' + expectedB5Terms.join('+');
  var actualB5 = sheet.getRange('B5').getFormula();
  if (!formulasAreEquivalent(actualB5, expectedB5)) {
    issues.push({
      type: 'ERROR',
      location: 'Control Panel - My Monthly Total (B5)',
      expected: expectedB5,
      actual: actualB5 || '(empty)',
      row: 5,
      col: 2
    });
  }
  
  // Check B6: Wife's Monthly Total (sum of all summary rows' Column E)
  var expectedB6Terms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    expectedB6Terms.push('E' + summaryRows[i]);
  }
  var expectedB6 = '=' + expectedB6Terms.join('+');
  var actualB6 = sheet.getRange('B6').getFormula();
  if (!formulasAreEquivalent(actualB6, expectedB6)) {
    issues.push({
      type: 'ERROR',
      location: 'Control Panel - Wife\'s Monthly Total (B6)',
      expected: expectedB6,
      actual: actualB6 || '(empty)',
      row: 6,
      col: 2
    });
  }
  
  // Check B12: My Total Donation (sum of all summary rows' Column G)
  var expectedB12Terms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    expectedB12Terms.push('G' + summaryRows[i]);
  }
  var expectedB12 = '=' + expectedB12Terms.join('+');
  var actualB12 = sheet.getRange('B12').getFormula();
  if (!formulasAreEquivalent(actualB12, expectedB12)) {
    issues.push({
      type: 'ERROR',
      location: 'Control Panel - My Total Donation (B12)',
      expected: expectedB12,
      actual: actualB12 || '(empty)',
      row: 12,
      col: 2
    });
  }
  
  // Check B13: Wife's Total Donation (sum of all summary rows' Column I)
  var expectedB13Terms = [];
  for (var i = 0; i < summaryRows.length; i++) {
    expectedB13Terms.push('I' + summaryRows[i]);
  }
  var expectedB13 = '=' + expectedB13Terms.join('+');
  var actualB13 = sheet.getRange('B13').getFormula();
  if (!formulasAreEquivalent(actualB13, expectedB13)) {
    issues.push({
      type: 'ERROR',
      location: 'Control Panel - Wife\'s Total Donation (B13)',
      expected: expectedB13,
      actual: actualB13 || '(empty)',
      row: 13,
      col: 2
    });
  }
}

/**
 * Check grand total row formulas (row 26)
 */
function checkGrandTotalFormulas(sheet, issues) {
  var lastRow = sheet.getLastRow();
  
  // Find all category total rows
  var categoryTotalRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    return; // No categories found
  }
  
  // Check B26: Monthly Grand Total
  var expectedMonthlyTerms = [];
  for (var i = 0; i < categoryTotalRows.length; i++) {
    expectedMonthlyTerms.push('B' + categoryTotalRows[i]);
  }
  var expectedB26 = '=' + expectedMonthlyTerms.join('+');
  var actualB26 = sheet.getRange('B26').getFormula();
  
  if (!formulasAreEquivalent(actualB26, expectedB26)) {
    issues.push({
      type: 'ERROR',
      location: 'Grand Total Row - Monthly Total (B26)',
      expected: expectedB26,
      actual: actualB26 || '(empty)',
      row: 26,
      col: 2
    });
  }
  
  // Check each day's grand total formulas
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1;
    
    // Day Total
    var expectedDayTotalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      expectedDayTotalTerms.push(getColumnLetter(baseCol) + categoryTotalRows[i]);
    }
    var expectedDayTotal = '=' + expectedDayTotalTerms.join('+');
    var actualDayTotal = sheet.getRange(26, baseCol).getFormula();
    if (!formulasAreEquivalent(actualDayTotal, expectedDayTotal)) {
      issues.push({
        type: 'ERROR',
        location: 'Grand Total Row - Day ' + day + ' Total',
        expected: expectedDayTotal,
        actual: actualDayTotal || '(empty)',
        row: 26,
        col: baseCol
      });
    }
    
    // Personal Total
    var expectedPersonalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      expectedPersonalTerms.push(getColumnLetter(baseCol + 1) + categoryTotalRows[i]);
    }
    var expectedPersonal = '=' + expectedPersonalTerms.join('+');
    var actualPersonal = sheet.getRange(26, baseCol + 1).getFormula();
    if (!formulasAreEquivalent(actualPersonal, expectedPersonal)) {
      issues.push({
        type: 'ERROR',
        location: 'Grand Total Row - Day ' + day + ' Personal',
        expected: expectedPersonal,
        actual: actualPersonal || '(empty)',
        row: 26,
        col: baseCol + 1
      });
    }
    
    // Family Total
    var expectedFamilyTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      expectedFamilyTerms.push(getColumnLetter(baseCol + 2) + categoryTotalRows[i]);
    }
    var expectedFamily = '=' + expectedFamilyTerms.join('+');
    var actualFamily = sheet.getRange(26, baseCol + 2).getFormula();
    if (!formulasAreEquivalent(actualFamily, expectedFamily)) {
      issues.push({
        type: 'ERROR',
        location: 'Grand Total Row - Day ' + day + ' Family',
        expected: expectedFamily,
        actual: actualFamily || '(empty)',
        row: 26,
        col: baseCol + 2
      });
    }
    
    // Donation Total
    var expectedDonationTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      expectedDonationTerms.push(getColumnLetter(baseCol + 3) + categoryTotalRows[i]);
    }
    var expectedDonation = '=' + expectedDonationTerms.join('+');
    var actualDonation = sheet.getRange(26, baseCol + 3).getFormula();
    if (!formulasAreEquivalent(actualDonation, expectedDonation)) {
      issues.push({
        type: 'ERROR',
        location: 'Grand Total Row - Day ' + day + ' Donation',
        expected: expectedDonation,
        actual: actualDonation || '(empty)',
        row: 26,
        col: baseCol + 3
      });
    }
  }
}

/**
 * Report diagnostic results
 */
function reportResults(issues, categoryName, subcategoryCount, ui) {
  if (issues.length === 0) {
    ui.alert('Diagnosis Complete', 
      'Category "' + categoryName + '" has ' + subcategoryCount + ' subcategories.\n\n' +
      '✓ All formulas are correct! No issues found.',
      ui.ButtonSet.OK);
    return;
  }
  
  // Build detailed report
  var report = 'Category: ' + categoryName + '\n';
  report += 'Subcategories checked: ' + subcategoryCount + '\n';
  report += 'Issues found: ' + issues.length + '\n\n';
  report += 'DETAILED ISSUES:\n';
  report += '================\n\n';
  
  for (var i = 0; i < issues.length; i++) {
    var issue = issues[i];
    report += (i + 1) + '. ' + issue.type + ': ' + issue.location + '\n';
    report += '   Row: ' + issue.row + ', Column: ' + getColumnLetter(issue.col) + issue.col + '\n';
    report += '   Expected: ' + issue.expected + '\n';
    report += '   Actual: ' + issue.actual + '\n\n';
  }
  
  // Show in alert (truncated if too long)
  if (report.length > 1000) {
    report = report.substring(0, 1000) + '\n\n... (showing first 1000 chars, see logs for full report)';
  }
  
  Logger.log('=== DIAGNOSTIC REPORT ===\n' + report);
  
  var response = ui.alert('Issues Found', 
    'Found ' + issues.length + ' formula error(s) in category "' + categoryName + '".\n\n' +
    'Check the execution log (View > Logs) for detailed information.\n\n' +
    'Would you like to fix these errors automatically?',
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    fixCategoryFormulas(issues, categoryName, ui);
  }
}

/**
 * Fix all identified global formula issues
 */
function fixGlobalFormulas(issues, ui) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var fixedCount = 0;
  var errorCount = 0;
  
  try {
    for (var i = 0; i < issues.length; i++) {
      var issue = issues[i];
      try {
        sheet.getRange(issue.row, issue.col).setFormula(issue.expected);
        fixedCount++;
        
        // Add small delay every 50 fixes to avoid rate limiting
        if (fixedCount % 50 === 0) {
          SpreadsheetApp.flush();
          Utilities.sleep(100);
        }
      } catch (error) {
        errorCount++;
        Logger.log('Error fixing ' + issue.location + ': ' + error.toString());
      }
    }
    
    SpreadsheetApp.flush();
    
    if (errorCount === 0) {
      ui.alert('Repair Complete', 
        'Successfully fixed ' + fixedCount + ' global formula error(s).',
        ui.ButtonSet.OK);
    } else {
      ui.alert('Repair Partially Complete', 
        'Fixed ' + fixedCount + ' formula error(s).\n' +
        errorCount + ' error(s) could not be fixed automatically.\n\n' +
        'Check the execution log for details.',
        ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Repair Error', 
      'An error occurred while fixing formulas: ' + error.toString() + '\n\n' +
      'Fixed ' + fixedCount + ' formulas before the error.',
      ui.ButtonSet.OK);
  }
}

/**
 * Fix all identified formula issues in a category
 */
function fixCategoryFormulas(issues, categoryName, ui) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var fixedCount = 0;
  var errorCount = 0;
  
  try {
    for (var i = 0; i < issues.length; i++) {
      var issue = issues[i];
      try {
        sheet.getRange(issue.row, issue.col).setFormula(issue.expected);
        fixedCount++;
        
        // Add small delay every 50 fixes to avoid rate limiting
        if (fixedCount % 50 === 0) {
          SpreadsheetApp.flush();
          Utilities.sleep(100);
        }
      } catch (error) {
        errorCount++;
        Logger.log('Error fixing ' + issue.location + ': ' + error.toString());
      }
    }
    
    SpreadsheetApp.flush();
    
    if (errorCount === 0) {
      ui.alert('Repair Complete', 
        'Successfully fixed ' + fixedCount + ' formula error(s) in category "' + categoryName + '".',
        ui.ButtonSet.OK);
    } else {
      ui.alert('Repair Partially Complete', 
        'Fixed ' + fixedCount + ' formula error(s).\n' +
        errorCount + ' error(s) could not be fixed automatically.\n\n' +
        'Check the execution log for details.',
        ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Repair Error', 
      'An error occurred while fixing formulas: ' + error.toString() + '\n\n' +
      'Fixed ' + fixedCount + ' formulas before the error.',
      ui.ButtonSet.OK);
  }
}

/**
 * Standalone function to fix a specific category (can be called directly)
 * This function re-applies all formulas in a category to fix any issues
 */
function fixCategoryFormulasByName() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var categoryResponse = ui.prompt(
    'Fix Category Formulas',
    'Enter the category name to fix:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (categoryResponse.getSelectedButton() != ui.Button.OK) {
    return;
  }
  
  var categoryName = categoryResponse.getResponseText().trim();
  
  // Find the category
  var lastRow = sheet.getLastRow();
  var categoryHeaderRow = -1;
  var categoryTotalRow = -1;
  
  for (var row = lastRow; row >= 27; row--) {
    var cellValue = sheet.getRange(row, 1).getValue();
    var note = sheet.getRange(row, 1).getNote();
    
    if (cellValue === categoryName + ' TOTAL' && note === '[CategoryTotal]') {
      categoryTotalRow = row;
    }
    
    if (cellValue === categoryName && !note) {
      categoryHeaderRow = row;
      break;
    }
  }
  
  if (categoryHeaderRow === -1 || categoryTotalRow === -1) {
    ui.alert('Error', 'Category "' + categoryName + '" not found!', ui.ButtonSet.OK);
    return;
  }
  
  // Find all subcategories
  var subcategoryRows = [];
  for (var row = categoryHeaderRow + 1; row < categoryTotalRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[Totals]') {
      subcategoryRows.push({
        totalsRow: row,
        meRow: row + 1,
        wifeRow: row + 2
      });
    }
  }
  
  try {
    var fixedCount = 0;
    
    // Fix subcategory formulas
    for (var i = 0; i < subcategoryRows.length; i++) {
      var subcat = subcategoryRows[i];
      
      // Fix [Totals] row
      fixTotalsRowFormulas(subcat.totalsRow, subcat.meRow, subcat.wifeRow, sheet);
      fixedCount++;
      
      // Fix [Me] row
      fixMeRowFormulas(subcat.meRow, sheet);
      fixedCount++;
      
      // Fix [Wife] row
      fixWifeRowFormulas(subcat.wifeRow, sheet);
      fixedCount++;
      
      // Add delay every 10 subcategories
      if ((i + 1) % 10 === 0) {
        SpreadsheetApp.flush();
        Utilities.sleep(100);
      }
    }
    
    // Fix category total formulas
    fixCategoryTotalRowFormulas(categoryTotalRow, subcategoryRows, sheet);
    fixedCount++;
    
    // Fix category summary row formulas (if it exists)
    var categorySummaryRow = categoryTotalRow + 1;
    var summaryNote = sheet.getRange(categorySummaryRow, 1).getNote();
    if (summaryNote === '[CategorySummary]') {
      fixCategorySummaryRowFormulas(categorySummaryRow, subcategoryRows, sheet);
      fixedCount++;
    }
    
    SpreadsheetApp.flush();
    
    ui.alert('Success', 
      'All formulas in category "' + categoryName + '" have been re-applied successfully!\n\n' +
      'Fixed ' + fixedCount + ' formula sets.',
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 
      'Error fixing formulas: ' + error.toString(),
      ui.ButtonSet.OK);
  }
}

/**
 * Fix formulas in a [Totals] row
 */
function fixTotalsRowFormulas(totalsRow, meRow, wifeRow, sheet) {
  // Monthly Total (Column B)
  var monthlyFormula = buildMonthlyTotalFormula(totalsRow);
  sheet.getRange(totalsRow, 2).setFormula(monthlyFormula);
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total
    sheet.getRange(totalsRow, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + totalsRow + 
      '+' + getColumnLetter(baseCol + 2) + totalsRow + 
      '+' + getColumnLetter(baseCol + 3) + totalsRow
    );
    
    // Personal Total
    sheet.getRange(totalsRow, baseCol + 1).setFormula(
      '=' + getColumnLetter(baseCol + 1) + meRow + 
      '+' + getColumnLetter(baseCol + 1) + wifeRow
    );
    
    // Family Total
    sheet.getRange(totalsRow, baseCol + 2).setFormula(
      '=' + getColumnLetter(baseCol + 2) + meRow + 
      '+' + getColumnLetter(baseCol + 2) + wifeRow
    );
    
    // Donation Total
    sheet.getRange(totalsRow, baseCol + 3).setFormula(
      '=' + getColumnLetter(baseCol + 3) + meRow + 
      '+' + getColumnLetter(baseCol + 3) + wifeRow
    );
  }
}

/**
 * Fix formulas in a [Me] row
 */
function fixMeRowFormulas(meRow, sheet) {
  // Monthly Total
  var monthlyFormula = buildMonthlyTotalFormula(meRow);
  sheet.getRange(meRow, 2).setFormula(monthlyFormula);
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    sheet.getRange(meRow, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + meRow + 
      '+' + getColumnLetter(baseCol + 2) + meRow + 
      '+' + getColumnLetter(baseCol + 3) + meRow
    );
  }
}

/**
 * Fix formulas in a [Wife] row
 */
function fixWifeRowFormulas(wifeRow, sheet) {
  // Monthly Total
  var monthlyFormula = buildMonthlyTotalFormula(wifeRow);
  sheet.getRange(wifeRow, 2).setFormula(monthlyFormula);
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    sheet.getRange(wifeRow, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + wifeRow + 
      '+' + getColumnLetter(baseCol + 2) + wifeRow + 
      '+' + getColumnLetter(baseCol + 3) + wifeRow
    );
  }
}

/**
 * Fix formulas in category summary row
 */
function fixCategorySummaryRowFormulas(summaryRow, subcategoryRows, sheet) {
  // Find all [Me] and [Wife] rows in this category
  var meRows = [];
  var wifeRows = [];
  
  for (var i = 0; i < subcategoryRows.length; i++) {
    meRows.push(subcategoryRows[i].meRow);
    wifeRows.push(subcategoryRows[i].wifeRow);
  }
  
  // Column C: Sum of all [Me] rows' Column B
  if (meRows.length > 0) {
    var meTotalTerms = [];
    for (var i = 0; i < meRows.length; i++) {
      meTotalTerms.push('B' + meRows[i]);
    }
    sheet.getRange(summaryRow, 3).setFormula('=' + meTotalTerms.join('+'));
  }
  
  // Column E: Sum of all [Wife] rows' Column B
  if (wifeRows.length > 0) {
    var wifeTotalTerms = [];
    for (var i = 0; i < wifeRows.length; i++) {
      wifeTotalTerms.push('B' + wifeRows[i]);
    }
    sheet.getRange(summaryRow, 5).setFormula('=' + wifeTotalTerms.join('+'));
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
      sheet.getRange(summaryRow, 7).setFormula('=' + myDonationTerms.join('+'));
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
      sheet.getRange(summaryRow, 9).setFormula('=' + wifeDonationTerms.join('+'));
    }
  }
}

/**
 * Fix formulas in category total row
 */
function fixCategoryTotalRowFormulas(categoryTotalRow, subcategoryRows, sheet) {
  // Monthly Total (Column B)
  var monthlyTerms = [];
  for (var i = 0; i < subcategoryRows.length; i++) {
    monthlyTerms.push('B' + subcategoryRows[i].totalsRow);
  }
  sheet.getRange(categoryTotalRow, 2).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total
    var dayTotalTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + subcategoryRows[i].totalsRow);
    }
    sheet.getRange(categoryTotalRow, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    // Personal Total
    var personalTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + subcategoryRows[i].totalsRow);
    }
    sheet.getRange(categoryTotalRow, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    // Family Total
    var familyTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + subcategoryRows[i].totalsRow);
    }
    sheet.getRange(categoryTotalRow, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    // Donation Total
    var donationTerms = [];
    for (var i = 0; i < subcategoryRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + subcategoryRows[i].totalsRow);
    }
    sheet.getRange(categoryTotalRow, baseCol + 3).setFormula('=' + donationTerms.join('+'));
  }
}

/**
 * Normalize a formula by extracting and sorting cell references
 * This makes formulas with same cells in different order equivalent
 * Example: "=B378+B382+B386" and "=B386+B378+B382" both become "=B378+B382+B386"
 * Also handles chunked formulas like "=(B378+B382)+(B386+B390)"
 */
function normalizeFormula(formula) {
  if (!formula || formula === '' || !formula.startsWith('=')) {
    return formula || '';
  }
  
  // Remove the leading '='
  var formulaBody = formula.substring(1);
  
  // Remove all parentheses and whitespace to extract just the cell references
  // This handles both simple formulas and chunked formulas
  formulaBody = formulaBody.replace(/[()]/g, '');
  formulaBody = formulaBody.replace(/\s+/g, '');
  
  // Extract cell references using regex (matches patterns like B378, AA123, etc.)
  // Pattern: one or more letters followed by one or more digits
  var cellRefPattern = /([A-Z]+)(\d+)/gi;
  var matches = [];
  var match;
  
  // Reset regex lastIndex to ensure we get all matches
  cellRefPattern.lastIndex = 0;
  
  while ((match = cellRefPattern.exec(formulaBody)) !== null) {
    matches.push(match[0].toUpperCase()); // Full match like "B378", convert to uppercase
  }
  
  if (matches.length === 0) {
    // No cell references found, return as-is (might be a constant or function)
    return formula;
  }
  
  // Remove duplicates (in case formula has same cell twice)
  var uniqueMatches = [];
  for (var i = 0; i < matches.length; i++) {
    if (uniqueMatches.indexOf(matches[i]) === -1) {
      uniqueMatches.push(matches[i]);
    }
  }
  
  // Sort cell references alphabetically/numerically
  // First sort by column (letters), then by row (numbers)
  uniqueMatches.sort(function(a, b) {
    // Extract column and row from each reference
    var aMatch = a.match(/([A-Z]+)(\d+)/i);
    var bMatch = b.match(/([A-Z]+)(\d+)/i);
    
    if (!aMatch || !bMatch) return 0;
    
    var aCol = aMatch[1];
    var bCol = bMatch[1];
    var aRow = parseInt(aMatch[2]);
    var bRow = parseInt(bMatch[2]);
    
    // Compare columns first (lexicographically)
    if (aCol < bCol) return -1;
    if (aCol > bCol) return 1;
    
    // If columns are same, compare rows
    return aRow - bRow;
  });
  
  // Rebuild formula with sorted references
  return '=' + uniqueMatches.join('+');
}

/**
 * Compare two formulas for equivalence (ignoring order of terms)
 */
function formulasAreEquivalent(formula1, formula2) {
  if (!formula1 && !formula2) return true;
  if (!formula1 || !formula2) return false;
  
  var normalized1 = normalizeFormula(formula1);
  var normalized2 = normalizeFormula(formula2);
  
  return normalized1 === normalized2;
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

