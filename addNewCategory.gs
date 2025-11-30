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
  // 1 category header + (4 rows per subcategory) + 1 category total
  var totalRowsNeeded = 1 + (subcategories.length * 4) + 1;
  
  // Step 4: Insert all rows at once
  sheet.insertRowsAfter(lastRow, totalRowsNeeded);
  
  var currentRow = lastRow + 1;
  
  // Step 5: Create Category Header Row (FIRST)
  var categoryHeaderRow = currentRow;
  sheet.getRange(categoryHeaderRow, 1).setValue(categoryName);
  currentRow++;
  
  // Step 6: Add all subcategories
  for (var i = 0; i < subcategories.length; i++) {
    var subcat = subcategories[i];
    
    // Row 1: [Totals] - Subcategory name
    sheet.getRange(currentRow, 2).setValue(subcat);
    sheet.getRange(currentRow, 2).setNote('[Totals]');
    currentRow++;
    
    // Row 2: [Me]
    sheet.getRange(currentRow, 2).setNote('[Me]');
    currentRow++;
    
    // Row 3: [Wife]
    sheet.getRange(currentRow, 2).setNote('[Wife]');
    currentRow++;
    
    // Row 4: [Comment]
    sheet.getRange(currentRow, 2).setNote('[Comment]');
    currentRow++;
  }
  
  // Step 7: Add Category Total Row (LAST)
  var categoryTotalRow = currentRow;
  sheet.getRange(categoryTotalRow, 2).setValue(categoryName + ' TOTAL');
  sheet.getRange(categoryTotalRow, 2).setNote('[CategoryTotal]');
  
  // Step 8: Apply formatting to category header row ONLY
  var maxCols = sheet.getMaxColumns();
  var headerRange = sheet.getRange(categoryHeaderRow, 1, 1, maxCols);
  headerRange.setBackground('#d9d9d9');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  
  // Clear columns C onwards for category header
  sheet.getRange(categoryHeaderRow, 3, 1, maxCols - 2).clearContent();
  
  // Step 9: Apply formatting to category total row ONLY
  var totalRange = sheet.getRange(categoryTotalRow, 1, 1, maxCols);
  totalRange.setBackground('#b8b8b8');
  totalRange.setFontWeight('bold');
  totalRange.setFontSize(10);
  
  // Step 9.5: Clear background AND remove bold from all subcategory rows
  var subcatRowCount = categoryTotalRow - categoryHeaderRow - 1;
  if (subcatRowCount > 0) {
    var subcatRange = sheet.getRange(categoryHeaderRow + 1, 1, subcatRowCount, maxCols);
    subcatRange.setBackground(null); // Clear any background
    subcatRange.setFontWeight('normal'); // Remove bold
  }
  
  // Step 10: Apply formulas to all subcategory rows
  var firstSubcatRow = categoryHeaderRow + 1;
  var lastSubcatRow = categoryTotalRow - 1;
  
  for (var row = firstSubcatRow; row <= lastSubcatRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Totals]') {
      applyTotalsRowFormulas(row, sheet);
    } else if (note === '[Me]') {
      applyMeRowFormulas(row, sheet);
    } else if (note === '[Wife]') {
      applyWifeRowFormulas(row, sheet);
    }
  }
  
  // Step 11: Apply formulas to category total row
  applyCategoryTotalsRowFormulas(categoryTotalRow, sheet);
  
  // Step 12: Apply cell coloring ONLY to [Me] and [Wife] input cells
  applyInputCellColors(firstSubcatRow, lastSubcatRow, sheet);
  
  // Step 13: Apply number formatting
  applyNumberFormatting(firstSubcatRow, categoryTotalRow, sheet);
  
  // Step 14: Update control panel summaries
  updateControlPanelSummaries();
  
  SpreadsheetApp.flush();
  
  ui.alert('Success!', 'Category "' + categoryName + '" with ' + subcategories.length + ' subcategories has been added successfully!', ui.ButtonSet.OK);
}

function applyInputCellColors(startRow, endRow, sheet) {
  // Color scheme for days
  var dayColors = [
    ['#E6E6FF', '#D6D6FF', '#C6C6FF'], // Purple
    ['#E6F3FF', '#D6E3FF', '#C6D3FF'], // Blue  
    ['#E6FFE6', '#D6FFD6', '#C6FFC6'], // Green
    ['#FFF0E6', '#FFE0D6', '#FFD0C6'], // Orange
    ['#F0E6FF', '#E0D6FF', '#D0C6FF']  // Light purple
  ];
  
  // Only apply colors to [Me] and [Wife] rows
  for (var row = startRow; row <= endRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    // ONLY color [Me] and [Wife] rows
    if (note === '[Me]' || note === '[Wife]') {
      for (var day = 1; day <= 31; day++) {
        var colorSet = dayColors[(day - 1) % dayColors.length];
        var baseCol = 4 + (day - 1) * 4; // Day 1 starts at column 4 (D)
        
        // Color Personal, Family, Donation input cells
        sheet.getRange(row, baseCol + 1).setBackground(colorSet[0]);   // Personal
        sheet.getRange(row, baseCol + 2).setBackground(colorSet[1]);   // Family
        sheet.getRange(row, baseCol + 3).setBackground(colorSet[2]);   // Donation
      }
    }
  }
}

function applyNumberFormatting(startRow, endRow, sheet) {
  // Format Column C (Monthly Total) as currency
  sheet.getRange(startRow, 3, endRow - startRow + 1, 1).setNumberFormat('"PKR "#,##0.00');
  
  // Format all day columns (31 days × 4 columns each)
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    sheet.getRange(startRow, baseCol, endRow - startRow + 1, 4).setNumberFormat('"PKR "#,##0.00');
  }
}

function applyTotalsRowFormulas(row, sheet) {
  var meRow = row + 1;
  var wifeRow = row + 2;
  
  // Monthly Total (Column C) = Sum of all day totals for this subcategory
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day (31 days)
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4; // Day 1 at column D (4)
    
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
  // Monthly Total (Column C) = Sum of all my day totals
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    
    // My Day Total = My Personal + My Family + My Donation
    sheet.getRange(row, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + row + 
      '+' + getColumnLetter(baseCol + 2) + row + 
      '+' + getColumnLetter(baseCol + 3) + row
    );
  }
}

function applyWifeRowFormulas(row, sheet) {
  // Monthly Total (Column C) = Sum of all wife's day totals
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    
    // Wife's Day Total = Wife's Personal + Wife's Family + Wife's Donation
    sheet.getRange(row, baseCol).setFormula(
      '=' + getColumnLetter(baseCol + 1) + row + 
      '+' + getColumnLetter(baseCol + 2) + row + 
      '+' + getColumnLetter(baseCol + 3) + row
    );
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  // Find all [Totals] rows above this row until we hit a category header
  var totalsRows = [];
  
  for (var r = row - 1; r >= 26; r--) {
    var categoryCell = sheet.getRange(r, 1).getValue();
    // Stop if we hit a category header (has value in column A)
    if (categoryCell !== '') {
      break;
    }
    
    var note = sheet.getRange(r, 2).getNote();
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  if (totalsRows.length === 0) {
    return; // No subcategories found
  }
  
  // Monthly Total (Column C) = Sum of all subcategory monthly totals
  var monthlyTerms = [];
  for (var i = 0; i < totalsRows.length; i++) {
    monthlyTerms.push('C' + totalsRows[i]);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    
    // Day Total
    var dayTotalTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    // Personal Total
    var personalTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    // Family Total
    var familyTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    // Donation Total
    var donationTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + totalsRows[i]);
    }
    sheet.getRange(row, baseCol + 3).setFormula('=' + donationTerms.join('+'));
  }
}

function updateControlPanelSummaries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Find all [Me] and [Wife] rows
  var meRows = [];
  var wifeRows = [];
  var myDonationTerms = [];
  var wifeDonationTerms = [];
  
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]') {
      meRows.push('C' + row);
      
      // Add all donation cells for this row
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
        myDonationTerms.push(getColumnLetter(baseCol + 3) + row);
      }
    } else if (note === '[Wife]') {
      wifeRows.push('C' + row);
      
      // Add all donation cells for this row
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
        wifeDonationTerms.push(getColumnLetter(baseCol + 3) + row);
      }
    }
  }
  
  // Update MY MONTHLY TOTAL (B5)
  if (meRows.length > 0) {
    sheet.getRange('B5').setFormula('=' + meRows.join('+'));
  }
  
  // Update WIFE'S MONTHLY TOTAL (B6)
  if (wifeRows.length > 0) {
    sheet.getRange('B6').setFormula('=' + wifeRows.join('+'));
  }
  
  // Update MY TOTAL DONATION (B12)
  if (myDonationTerms.length > 0) {
    sheet.getRange('B12').setFormula('=' + myDonationTerms.join('+'));
  }
  
  // Update WIFE'S TOTAL DONATION (B13)
  if (wifeDonationTerms.length > 0) {
    sheet.getRange('B13').setFormula('=' + wifeDonationTerms.join('+'));
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