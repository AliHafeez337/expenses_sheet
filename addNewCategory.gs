function addNewCategory() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Step 1: Ask for Category Name
  var categoryResponse = ui.prompt(
    'Add New Category',
    'Enter the category name:',
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
  
  // Step 2: Ask for Subcategories (comma-separated)
  var subcatResponse = ui.prompt(
    'Add Subcategories',
    'Enter subcategories separated by commas (e.g., Item1, Item2, Item3):',
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
  
  // Step 3: Find the last row with data
  var lastRow = sheet.getLastRow();
  
  // Calculate how many rows we need to insert
  var totalRowsNeeded = 1 + (subcategories.length * 4) + 1;
  
  // Step 4: Insert all rows at once
  sheet.insertRowsAfter(lastRow, totalRowsNeeded);
  
  var currentRow = lastRow + 1;
  
  // Step 5: Create Category Header Row
  var categoryHeaderRow = currentRow;
  sheet.getRange(categoryHeaderRow, 1).setValue(categoryName);
  currentRow++;
  
  // Step 6: Add all subcategories - NOW IN COLUMN A
  for (var i = 0; i < subcategories.length; i++) {
    var subcat = subcategories[i];
    
    sheet.getRange(currentRow, 1).setValue(subcat); // CHANGED: column 2 to 1
    sheet.getRange(currentRow, 1).setNote('[Totals]'); // CHANGED: column 2 to 1
    currentRow++;
    
    sheet.getRange(currentRow, 1).setNote('[Me]'); // CHANGED: column 2 to 1
    currentRow++;
    
    sheet.getRange(currentRow, 1).setNote('[Wife]'); // CHANGED: column 2 to 1
    currentRow++;
    
    sheet.getRange(currentRow, 1).setNote('[Comment]'); // CHANGED: column 2 to 1
    currentRow++;
  }
  
  // Step 7: Add Category Total Row
  var categoryTotalRow = currentRow;
  sheet.getRange(categoryTotalRow, 1).setValue(categoryName + ' TOTAL'); // CHANGED: column 2 to 1
  sheet.getRange(categoryTotalRow, 1).setNote('[CategoryTotal]'); // CHANGED: column 2 to 1
  
  // Step 8-9: Apply formatting
  var maxCols = sheet.getMaxColumns();
  var headerRange = sheet.getRange(categoryHeaderRow, 1, 1, maxCols);
  headerRange.setBackground('#d9d9d9');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  sheet.getRange(categoryHeaderRow, 2, 1, maxCols - 1).clearContent(); // CHANGED: column 3 to 2
  
  var totalRange = sheet.getRange(categoryTotalRow, 1, 1, maxCols);
  totalRange.setBackground('#b8b8b8');
  totalRange.setFontWeight('bold');
  totalRange.setFontSize(10);
  
  var subcatRowCount = categoryTotalRow - categoryHeaderRow - 1;
  if (subcatRowCount > 0) {
    var subcatRange = sheet.getRange(categoryHeaderRow + 1, 1, subcatRowCount, maxCols);
    subcatRange.setBackground(null);
    subcatRange.setFontWeight('normal');
  }
  
  // Step 10-11: Apply formulas
  var firstSubcatRow = categoryHeaderRow + 1;
  var lastSubcatRow = categoryTotalRow - 1;
  
  for (var row = firstSubcatRow; row <= lastSubcatRow; row++) {
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    
    if (note === '[Totals]') {
      applyTotalsRowFormulas(row, sheet);
    } else if (note === '[Me]') {
      applyMeRowFormulas(row, sheet);
    } else if (note === '[Wife]') {
      applyWifeRowFormulas(row, sheet);
    }
  }
  
  applyCategoryTotalsRowFormulas(categoryTotalRow, sheet);
  
  // Step 12: Apply cell coloring and text validation for comment cells
  applyInputCellColors(firstSubcatRow, lastSubcatRow, sheet);
  
  // Step 13: Apply number formatting
  applyNumberFormatting(firstSubcatRow, categoryTotalRow, sheet);
  
  // Step 14: Update control panel summaries
  updateControlPanelSummaries();
  
  // Step 15: Update grand total formulas
  applyGrandTotalFormulas();
  
  SpreadsheetApp.flush();
  
  ui.alert('Success!', 'Category "' + categoryName + '" with ' + subcategories.length + ' subcategories has been added successfully!', ui.ButtonSet.OK);
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
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    
    if (note === '[Me]' || note === '[Wife]') {
      for (var day = 1; day <= 31; day++) {
        var colorSet = dayColors[(day - 1) % dayColors.length];
        var baseCol = 3 + (day - 1) * 4; // SHIFTED: Day 1 starts at column 3 (C) - was 4 (D)
        
        sheet.getRange(row, baseCol + 1).setBackground(colorSet[0]);
        sheet.getRange(row, baseCol + 2).setBackground(colorSet[1]);
        sheet.getRange(row, baseCol + 3).setBackground(colorSet[2]);
      }
    }
    
    // Apply text formatting to comment rows
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
        
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
  sheet.getRange(startRow, 2, endRow - startRow + 1, 1).setNumberFormat('"PKR "#,##0.00'); // CHANGED: column 3 to 2
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    sheet.getRange(startRow, baseCol, endRow - startRow + 1, 4).setNumberFormat('"PKR "#,##0.00');
  }
  
  // RE-APPLY text formatting to comment rows (after currency formatting above)
  // This ensures comment cells show alphabetic keyboard on mobile devices
  for (var row = startRow; row <= endRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        // Force text format - this triggers alphabetic keyboard on mobile
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
    var note = sheet.getRange(r, 1).getNote(); // CHANGED: column 2 to 1
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
    monthlyTerms.push('B' + totalsRows[i]); // CHANGED: C to B
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    
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
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  // For each day (31 days)
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: Day 1 at column C (3) - was D (4)
    
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
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    
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
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    
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
  
  var meRows = [];
  var wifeRows = [];
  var myDonationTerms = [];
  var wifeDonationTerms = [];
  
  for (var row = 28; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    
    if (note === '[Me]') {
      meRows.push('B' + row); // CHANGED: C to B
      
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
        myDonationTerms.push(getColumnLetter(baseCol + 3) + row);
      }
    } else if (note === '[Wife]') {
      wifeRows.push('B' + row); // CHANGED: C to B
      
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
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
}

function applyGrandTotalFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  var categoryTotalRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    return;
  }
  
  var monthlyTerms = [];
  for (var i = 0; i < categoryTotalRows.length; i++) {
    monthlyTerms.push('B' + categoryTotalRows[i]); // CHANGED: C to B
  }
  sheet.getRange('B26').setFormula('=' + monthlyTerms.join('+')); // CHANGED: C26 to B26
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    
    var dayTotalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    var personalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    var familyTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    var donationTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 3).setFormula('=' + donationTerms.join('+'));
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