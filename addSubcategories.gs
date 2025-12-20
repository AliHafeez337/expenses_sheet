function addSubcategoriesToExisting() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Step 1: Ask for the existing Category Name
  var categoryResponse = ui.prompt(
    'Add Subcategories to Existing Category',
    'In which category to add?',
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
  
  // Step 2: Find the category in the sheet (search from bottom to top for speed)
  var lastRow = sheet.getLastRow();
  var categoryHeaderRow = -1;
  var categoryTotalRow = -1;
  
  for (var row = lastRow; row >= 27; row--) {
    var cellValue = sheet.getRange(row, 1).getValue();
    var note = sheet.getRange(row, 1).getNote();
    
    // Check if this is the category total row (matches name + " TOTAL")
    if (cellValue === categoryName + ' TOTAL' && note === '[CategoryTotal]') {
      categoryTotalRow = row;
    }
    
    // Check if this is the category header (matches name, no note)
    if (cellValue === categoryName && !note) {
      categoryHeaderRow = row;
      break; // Found both, we can stop
    }
  }
  
  if (categoryHeaderRow === -1 || categoryTotalRow === -1) {
    ui.alert('Error', 'Category "' + categoryName + '" not found! Please check the spelling and try again.', ui.ButtonSet.OK);
    return;
  }
  
  // Step 3: Ask for new Subcategories (comma-separated)
  var subcatResponse = ui.prompt(
    'Add Subcategories',
    'Enter the comma-separated names of subcategories:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (subcatResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('Operation cancelled.');
    return;
  }
  
  var subcategoriesText = subcatResponse.getResponseText().trim();
  if (subcategoriesText === '') {
    ui.alert('At least one subcategory is required!');
    return;
  }
  
  // Parse subcategories
  var subcategories = subcategoriesText.split(',').map(function(s) { 
    return s.trim(); 
  }).filter(function(s) { 
    return s !== ''; 
  });
  
  if (subcategories.length === 0) {
    ui.alert('No valid subcategories provided!');
    return;
  }
  
  // Step 4: Calculate how many rows we need to insert (4 rows per subcategory)
  var totalRowsNeeded = subcategories.length * 4;
  
  // Step 5: Insert rows BEFORE the category total row
  var insertPosition = categoryTotalRow - 1;
  sheet.insertRowsAfter(insertPosition, totalRowsNeeded);
  
  // The category total row has now moved down
  categoryTotalRow = categoryTotalRow + totalRowsNeeded;
  
  var currentRow = insertPosition + 1;
  
  // Step 6: Add all subcategories
  for (var i = 0; i < subcategories.length; i++) {
    var subcat = subcategories[i];
    
    // Totals row
    sheet.getRange(currentRow, 1).setValue(subcat);
    sheet.getRange(currentRow, 1).setNote('[Totals]');
    currentRow++;
    
    // Me row
    sheet.getRange(currentRow, 1).setNote('[Me]');
    currentRow++;
    
    // Wife row
    sheet.getRange(currentRow, 1).setNote('[Wife]');
    currentRow++;
    
    // Comment row
    sheet.getRange(currentRow, 1).setNote('[Comment]');
    currentRow++;
  }
  
  // Step 7: Update the category total row note (in case it got lost)
  sheet.getRange(categoryTotalRow, 1).setValue(categoryName + ' TOTAL');
  sheet.getRange(categoryTotalRow, 1).setNote('[CategoryTotal]');
  
  // Step 7b: Find and update category summary row (should be right after category total)
  var categorySummaryRow = categoryTotalRow + 1;
  var summaryNote = sheet.getRange(categorySummaryRow, 1).getNote();
  if (summaryNote !== '[CategorySummary]') {
    // Summary row doesn't exist, create it
    sheet.insertRowsAfter(categoryTotalRow, 1);
    categorySummaryRow = categoryTotalRow + 1;
    sheet.getRange(categorySummaryRow, 1).setNote('[CategorySummary]');
    sheet.getRange(categorySummaryRow, 2).setValue('My total for this category');
    sheet.getRange(categorySummaryRow, 4).setValue('Wife\'s total for this category');
    sheet.getRange(categorySummaryRow, 6).setValue('My donations for this category');
    sheet.getRange(categorySummaryRow, 8).setValue('Wife\'s donations for this category');
    // Format summary row
    sheet.getRange(categorySummaryRow, 1, 1, 9)
      .setBackground('#e8f4f8')
      .setFontWeight('normal')
      .setFontSize(9);
  }
  
  // Step 8: Apply formatting to the new subcategory rows
  var maxCols = sheet.getMaxColumns();
  var firstNewRow = insertPosition + 1;
  var lastNewRow = categoryTotalRow - 1;
  var subcatRowCount = lastNewRow - firstNewRow + 1;
  
  if (subcatRowCount > 0) {
    var subcatRange = sheet.getRange(firstNewRow, 1, subcatRowCount, maxCols);
    subcatRange.setBackground(null);
    subcatRange.setFontWeight('normal');
  }
  
  // Reapply category total row formatting
  var totalRange = sheet.getRange(categoryTotalRow, 1, 1, maxCols);
  totalRange.setBackground('#b8b8b8');
  totalRange.setFontWeight('bold');
  totalRange.setFontSize(10);
  
  // Step 9: Apply formulas to new rows
  for (var row = firstNewRow; row <= lastNewRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    
    if (note === '[Totals]') {
      applyTotalsRowFormulas(row, sheet);
    } else if (note === '[Me]') {
      applyMeRowFormulas(row, sheet);
    } else if (note === '[Wife]') {
      applyWifeRowFormulas(row, sheet);
    }
  }
  
  // Step 10: Re-apply category total formulas (to include new subcategories)
  applyCategoryTotalsRowFormulas(categoryTotalRow, sheet);
  
  // Step 10b: Re-apply category summary row formulas (to include new subcategories)
  applyCategorySummaryRowFormulas(categorySummaryRow, sheet);
  
  // Step 11: Apply cell coloring and text validation for comment cells
  applyInputCellColors(firstNewRow, lastNewRow, sheet);
  
  // Step 12: Apply number formatting
  applyNumberFormatting(firstNewRow, categorySummaryRow, sheet);
  
  // Step 13: Update control panel summaries
  updateControlPanelSummaries();
  
  // Step 14: Update grand total formulas (with error handling)
  var grandTotalError = false;
  try {
    applyGrandTotalFormulas();
  } catch (error) {
    grandTotalError = true;
    Logger.log('Warning: Grand total formulas update failed: ' + error.toString());
    // Continue anyway - subcategories were added successfully
  }
  
  SpreadsheetApp.flush();
  
  if (grandTotalError) {
    ui.alert('Partial Success', 
      subcategories.length + ' subcategories have been added to "' + categoryName + '" successfully!\n\n' +
      'However, the grand total formulas (row 26) could not be updated automatically.\n' +
      'You may need to manually update them or run the script again.', 
      ui.ButtonSet.OK);
  } else {
    ui.alert('Success!', subcategories.length + ' subcategories have been added to "' + categoryName + '" successfully!', ui.ButtonSet.OK);
  }
}

/**
 * Apply formulas to category summary row
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
      sheet.getRange(row, 8).setFormula('=' + wifeDonationTerms.join('+'));
    }
  }
}

// Helper function to convert column number to letter
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function applyInputCellColors(startRow, endRow, sheet) {
  var dayColors = [
    ['#E6E6FF', '#D6D6FF', '#C6C6FF'],
    ['#E6F3FF', '#D6E3FF', '#C6D3FF'],
    ['#E6FFE6', '#D6FFD6', '#C6FFC6'],
    ['#FFF0E6', '#FFE0D6', '#FFD0C6'],
    ['#F0E6FF', '#E0D6FF', '#D0C6FF']
  ];
  
  for (var row = startRow; row <= endRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    
    if (note === '[Me]' || note === '[Wife]') {
      for (var day = 1; day <= 31; day++) {
        var colorSet = dayColors[(day - 1) % dayColors.length];
        var baseCol = 3 + (day - 1) * 4;
        
        sheet.getRange(row, baseCol + 1).setBackground(colorSet[0]);
        sheet.getRange(row, baseCol + 2).setBackground(colorSet[1]);
        sheet.getRange(row, baseCol + 3).setBackground(colorSet[2]);
      }
    }
    
    // Apply text formatting to comment rows
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        
        sheet.getRange(row, baseCol + 1).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 2).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 3).setBackground('#FFFEF0');
        
        sheet.getRange(row, baseCol + 1).setNumberFormat('@');
        sheet.getRange(row, baseCol + 2).setNumberFormat('@');
        sheet.getRange(row, baseCol + 3).setNumberFormat('@');
        
        // Add data validation to force text input (triggers text keyboard on mobile)
        var textValidation = SpreadsheetApp.newDataValidation()
          .requireFormulaSatisfied('=TRUE')
          .setAllowInvalid(true)
          .setHelpText('Enter comment text')
          .build();
        
        sheet.getRange(row, baseCol + 1).setDataValidation(textValidation);
        sheet.getRange(row, baseCol + 2).setDataValidation(textValidation);
        sheet.getRange(row, baseCol + 3).setDataValidation(textValidation);
      }
    }
  }
}

function applyNumberFormatting(startRow, endRow, sheet) {
  sheet.getRange(startRow, 2, endRow - startRow + 1, 1).setNumberFormat('"PKR "#,##0.00');
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    sheet.getRange(startRow, baseCol, endRow - startRow + 1, 4).setNumberFormat('"PKR "#,##0.00');
  }
  
  // RE-APPLY text formatting to comment rows (after currency formatting above)
  for (var row = startRow; row <= endRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        sheet.getRange(row, baseCol + 1).setNumberFormat('@');
        sheet.getRange(row, baseCol + 2).setNumberFormat('@');
        sheet.getRange(row, baseCol + 3).setNumberFormat('@');
      }
    }
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  var totalsRows = [];
  
  for (var r = row - 1; r >= 27; r--) {
    var categoryCell = sheet.getRange(r, 1).getValue();
    var note = sheet.getRange(r, 1).getNote();
    // Stop if we hit a category header (has value but no note)
    if (categoryCell !== '' && !note) {
      break;
    }
    
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  if (totalsRows.length === 0) {
    return;
  }
  
  // Monthly Total (Column B) = Sum of all subcategory monthly totals
  var monthlyTerms = [];
  for (var i = 0; i < totalsRows.length; i++) {
    monthlyTerms.push('B' + totalsRows[i]);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    var dayTotalTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    var personalTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    var familyTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    var donationTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 3).setFormula('=' + donationTerms.join('+'));
  }
}

function applyTotalsRowFormulas(row, sheet) {
  var meRow = row + 1;
  var wifeRow = row + 2;
  
  // Monthly Total (Column B) = Sum of all day totals for this subcategory
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day (31 days)
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Day Total (Col 0) = Personal + Family + Donation
    sheet.getRange(row, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + row + 
      '+' + getColumnLetter(baseCol + 2) + row + 
      '+' + getColumnLetter(baseCol + 3) + row
    );
    
    // Personal Total (Col 1) = Me Personal + Wife Personal
    sheet.getRange(row, baseCol + 1).setFormula(
      '=' + getColumnLetter(baseCol + 1) + meRow + 
      '+' + getColumnLetter(baseCol + 1) + wifeRow
    );
    
    // Family Total (Col 2) = Me Family + Wife Family
    sheet.getRange(row, baseCol + 2).setFormula(
      '=' + getColumnLetter(baseCol + 2) + meRow + 
      '+' + getColumnLetter(baseCol + 2) + wifeRow
    );
    
    // Donation Total (Col 3) = Me Donation + Wife Donation
    sheet.getRange(row, baseCol + 3).setFormula(
      '=' + getColumnLetter(baseCol + 3) + meRow + 
      '+' + getColumnLetter(baseCol + 3) + wifeRow
    );
  }
}

function applyMeRowFormulas(row, sheet) {
  // Monthly Total (Column B) = Sum of all my day totals
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // My Day Total = My Personal + My Family + My Donation
    sheet.getRange(row, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + row + 
      '+' + getColumnLetter(baseCol + 2) + row + 
      '+' + getColumnLetter(baseCol + 3) + row
    );
  }
}

function applyWifeRowFormulas(row, sheet) {
  // Monthly Total (Column B) = Sum of all wife's day totals
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4;
    
    // Wife's Day Total = Wife's Personal + Wife's Family + Wife's Donation
    sheet.getRange(row, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + row + 
      '+' + getColumnLetter(baseCol + 2) + row + 
      '+' + getColumnLetter(baseCol + 3) + row
    );
  }
}

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

function applyGrandTotalFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  var categoryTotalRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    return;
  }
  
  // Monthly Grand Total (Column B) = Sum of all category totals
  var monthlyTerms = [];
  for (var i = 0; i < categoryTotalRows.length; i++) {
    monthlyTerms.push('B' + categoryTotalRows[i]);
  }
  sheet.getRange('B26').setFormula('=' + monthlyTerms.join('+'));
  
  // For each day, sum all category totals
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1;
    
    // Day Total
    var dayTotalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    // Personal Total
    var personalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    // Family Total
    var familyTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    // Donation Total
    var donationTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 3).setFormula('=' + donationTerms.join('+'));
  }
}

