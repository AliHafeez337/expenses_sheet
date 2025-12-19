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
  
  // Report results
  reportResults(issues, categoryName, subcategoryCount, ui);
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

