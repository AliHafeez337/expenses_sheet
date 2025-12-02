// Updated scripts with row 27 fixes

// ============================================
// addNewCategory.gs - UPDATED
// ============================================

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
  
  // Step 6: Add all subcategories
  for (var i = 0; i < subcategories.length; i++) {
    var subcat = subcategories[i];
    
    sheet.getRange(currentRow, 2).setValue(subcat);
    sheet.getRange(currentRow, 2).setNote('[Totals]');
    currentRow++;
    
    sheet.getRange(currentRow, 2).setNote('[Me]');
    currentRow++;
    
    sheet.getRange(currentRow, 2).setNote('[Wife]');
    currentRow++;
    
    sheet.getRange(currentRow, 2).setNote('[Comment]');
    currentRow++;
  }
  
  // Step 7: Add Category Total Row
  var categoryTotalRow = currentRow;
  sheet.getRange(categoryTotalRow, 2).setValue(categoryName + ' TOTAL');
  sheet.getRange(categoryTotalRow, 2).setNote('[CategoryTotal]');
  
  // Step 8-9: Apply formatting
  var maxCols = sheet.getMaxColumns();
  var headerRange = sheet.getRange(categoryHeaderRow, 1, 1, maxCols);
  headerRange.setBackground('#d9d9d9');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  sheet.getRange(categoryHeaderRow, 3, 1, maxCols - 2).clearContent();
  
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
    var note = sheet.getRange(row, 2).getNote();
    
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
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]' || note === '[Wife]') {
      for (var day = 1; day <= 31; day++) {
        var colorSet = dayColors[(day - 1) % dayColors.length];
        var baseCol = 4 + (day - 1) * 4;
        
        sheet.getRange(row, baseCol + 1).setBackground(colorSet[0]);
        sheet.getRange(row, baseCol + 2).setBackground(colorSet[1]);
        sheet.getRange(row, baseCol + 3).setBackground(colorSet[2]);
      }
    }
    
    // Apply text formatting to comment rows
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
        
        sheet.getRange(row, baseCol + 1).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 2).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 3).setBackground('#FFFEF0');
        
        sheet.getRange(row, baseCol + 1).setNumberFormat('@STRING@');
        sheet.getRange(row, baseCol + 2).setNumberFormat('@STRING@');
        sheet.getRange(row, baseCol + 3).setNumberFormat('@STRING@');
      }
    }
  }
}

function applyNumberFormatting(startRow, endRow, sheet) {
  sheet.getRange(startRow, 3, endRow - startRow + 1, 1).setNumberFormat('"PKR "#,##0.00');
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    sheet.getRange(startRow, baseCol, endRow - startRow + 1, 4).setNumberFormat('"PKR "#,##0.00');
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  var totalsRows = [];
  
  // FIXED: Changed from 26 to 27
  for (var r = row - 1; r >= 27; r--) {
    var categoryCell = sheet.getRange(r, 1).getValue();
    if (categoryCell !== '') {
      break;
    }
    
    var note = sheet.getRange(r, 2).getNote();
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  if (totalsRows.length === 0) {
    return;
  }
  
  var monthlyTerms = [];
  for (var i = 0; i < totalsRows.length; i++) {
    monthlyTerms.push('C' + totalsRows[i]);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    
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

function updateControlPanelSummaries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  var meRows = [];
  var wifeRows = [];
  var myDonationTerms = [];
  var wifeDonationTerms = [];
  
  // FIXED: Changed from 27 to 28
  for (var row = 28; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]') {
      meRows.push('C' + row);
      
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
        myDonationTerms.push(getColumnLetter(baseCol + 3) + row);
      }
    } else if (note === '[Wife]') {
      wifeRows.push('C' + row);
      
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
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

// ============================================
// setupMonthlyTrigger.gs - UPDATED
// ============================================

function setupMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createMonthlySheetAuto') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('createMonthlySheetAuto')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
  
  SpreadsheetApp.getUi().alert('Monthly trigger set up successfully!\n\nA new sheet will be created automatically on the 1st of each month at 6:00 AM.');
}

function createMonthlySheetAuto() {
  try {
    createMonthlySheet();
  } catch (e) {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'Error creating monthly expense sheet',
      body: 'There was an error creating the monthly expense sheet:\n\n' + e.toString()
    });
  }
}

function createMonthlySheetManual() {
  createMonthlySheet();
  SpreadsheetApp.getUi().alert('Monthly sheet created successfully!');
}

function createMonthlySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('Template');
  
  if (!templateSheet) {
    throw new Error('Template sheet not found! Please create a sheet named "Template".');
  }
  
  var now = new Date();
  var monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy');
  
  if (ss.getSheetByName(monthName)) {
    Logger.log('Sheet "' + monthName + '" already exists. Skipping creation.');
    return;
  }
  
  var newSheet = templateSheet.copyTo(ss);
  newSheet.setName(monthName);
  ss.moveActiveSheet(2);
  
  carryForwardShortfalls(newSheet, ss);
  clearInputCells(newSheet);
  
  newSheet.getRange('A1').setValue('Expenses - ' + monthName);
  
  Logger.log('Successfully created sheet: ' + monthName);
}

function carryForwardShortfalls(newSheet, ss) {
  var previousSheet = getPreviousMonthSheet(ss);
  
  if (!previousSheet) {
    newSheet.getRange('B21').setValue(0);
    newSheet.getRange('B22').setValue(0);
    return;
  }
  
  var myRemainingNeed = previousSheet.getRange('B18').getValue() || 0;
  var wifeRemainingNeed = previousSheet.getRange('B19').getValue() || 0;
  
  newSheet.getRange('B21').setValue(Math.max(0, myRemainingNeed));
  newSheet.getRange('B22').setValue(Math.max(0, wifeRemainingNeed));
}

function getPreviousMonthSheet(ss) {
  var sheets = ss.getSheets();
  var monthSheets = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName !== 'Template' && isMonthSheet(sheetName)) {
      monthSheets.push({
        name: sheetName,
        sheet: sheets[i],
        date: parseMonthSheet(sheetName)
      });
    }
  }
  
  if (monthSheets.length === 0) {
    return null;
  }
  
  monthSheets.sort(function(a, b) {
    return b.date.getTime() - a.date.getTime();
  });
  
  return monthSheets[0].sheet;
}

function isMonthSheet(sheetName) {
  var monthPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/;
  return monthPattern.test(sheetName);
}

function parseMonthSheet(sheetName) {
  var parts = sheetName.split(' ');
  var month = parts[0];
  var year = parseInt(parts[1]);
  
  var months = {
    'January': 0, 'February': 1, 'March': 2, 'April': 3,
    'May': 4, 'June': 5, 'July': 6, 'August': 7,
    'September': 8, 'October': 9, 'November': 10, 'December': 11
  };
  
  return new Date(year, months[month], 1);
}

function clearInputCells(sheet) {
  var lastRow = sheet.getLastRow();
  
  // FIXED: Changed from 27 to 28
  for (var row = 28; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]' || note === '[Wife]' || note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4 + 1;
        
        sheet.getRange(row, baseCol + 1).clearContent();
        sheet.getRange(row, baseCol + 2).clearContent();
        sheet.getRange(row, baseCol + 3).clearContent();
      }
    }
  }
  
  sheet.getRange('B2').setValue(0);
  sheet.getRange('B3').setValue(0);
}

function checkTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createMonthlySheetAuto') {
      found = true;
      var triggerSource = triggers[i].getTriggerSource();
      var eventType = triggers[i].getEventType();
      
      SpreadsheetApp.getUi().alert(
        'Monthly trigger is ACTIVE\n\n' +
        'Function: createMonthlySheetAuto\n' +
        'Runs on: 1st of each month at 6:00 AM\n' +
        'Trigger Source: ' + triggerSource + '\n' +
        'Event Type: ' + eventType
      );
      break;
    }
  }
  
  if (!found) {
    SpreadsheetApp.getUi().alert(
      'No monthly trigger found!\n\n' +
      'Run setupMonthlyTrigger() to create one.'
    );
  }
}

function deleteMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createMonthlySheetAuto') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted = true;
    }
  }
  
  if (deleted) {
    SpreadsheetApp.getUi().alert('Monthly trigger deleted successfully!');
  } else {
    SpreadsheetApp.getUi().alert('No monthly trigger found to delete.');
  }
}

// New function to add control panel formulas after all data is set
function updateControlPanelSummaries() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Build column lists for summation
  var dayColumns = {
    personal: [],
    family: [],
    donation: []
  };
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    dayColumns.personal.push(getColumnLetter(baseCol + 1));
    dayColumns.family.push(getColumnLetter(baseCol + 2));
    dayColumns.donation.push(getColumnLetter(baseCol + 3));
  }
  
  var meRows = [];
  var wifeRows = [];
  
  for (var row = 28; row <= lastRow; row++) { // CHANGED FROM 27 TO 28
    var note = sheet.getRange(row, 2).getNote();
    if (note === '[Me]') {
      meRows.push('C' + row);
    } else if (note === '[Wife]') {
      wifeRows.push('C' + row);
    }
  }
  
  if (meRows.length > 0) {
    sheet.getRange('B5').setFormula('=' + meRows.join('+'));
  }
  
  if (wifeRows.length > 0) {
    sheet.getRange('B6').setFormula('=' + wifeRows.join('+'));
  }
  
  // MY TOTAL DONATION (B12) - sum donation column for [Me] rows
  var myDonationTerms = [];
  var wifeDonationTerms = [];
  
  for (var row = 28; row <= lastRow; row++) { // CHANGED FROM 27 TO 28
    var note = sheet.getRange(row, 2).getNote();
    if (note === '[Me]') {
      dayColumns.donation.forEach(function(col) {
        myDonationTerms.push(col + row);
      });
    } else if (note === '[Wife]') {
      dayColumns.donation.forEach(function(col) {
        wifeDonationTerms.push(col + row);
      });
    }
  }
  
  if (myDonationTerms.length > 0) {
    sheet.getRange('B12').setFormula('=' + myDonationTerms.join('+'));
  }
  
  if (wifeDonationTerms.length > 0) {
    sheet.getRange('B13').setFormula('=' + wifeDonationTerms.join('+'));
  }
}

// NEW FUNCTION: Apply grand total formulas for row 26
function applyGrandTotalFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Find all category total rows
  var categoryTotalRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    return; // No category totals found
  }
  
  // Monthly Grand Total (Column C) = Sum of all category totals
  var monthlyTerms = [];
  for (var i = 0; i < categoryTotalRows.length; i++) {
    monthlyTerms.push('C' + categoryTotalRows[i]);
  }
  sheet.getRange('C26').setFormula('=' + monthlyTerms.join('+'));
  
  // For each day, sum all category totals
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Day 1 starts at column D (4)
    
    // Day Total = Sum of all category day totals
    var dayTotalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      dayTotalTerms.push(getColumnLetter(baseCol) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol).setFormula('=' + dayTotalTerms.join('+'));
    
    // Personal Total = Sum of all category personal totals
    var personalTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      personalTerms.push(getColumnLetter(baseCol + 1) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 1).setFormula('=' + personalTerms.join('+'));
    
    // Family Total = Sum of all category family totals
    var familyTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      familyTerms.push(getColumnLetter(baseCol + 2) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 2).setFormula('=' + familyTerms.join('+'));
    
    // Donation Total = Sum of all category donation totals
    var donationTerms = [];
    for (var i = 0; i < categoryTotalRows.length; i++) {
      donationTerms.push(getColumnLetter(baseCol + 3) + categoryTotalRows[i]);
    }
    sheet.getRange(26, baseCol + 3).setFormula('=' + donationTerms.join('+'));
  }
}

function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var headers = ['Category', 'Subcategory', 'Monthly Total'];
  
  // Add headers for 31 days - 4 columns per day
  for (var day = 1; day <= 31; day++) {
    headers.push('Day ' + day + ' Total');
    headers.push('Day ' + day + ' Personal');
    headers.push('Day ' + day + ' Family');
    headers.push('Day ' + day + ' Donation');
  }
  
  // Set headers in row 1
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Category
  sheet.setColumnWidth(2, 150); // Subcategory
  sheet.setColumnWidth(3, 120); // Monthly Total
  
  // Set width for all day columns
  for (var col = 4; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 85);
  }
}

function setupControlPanel() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Income Section (Rows 2-4)
  sheet.getRange('A2').setValue('MY INCOME:');
  sheet.getRange('B2').setValue(0);
  sheet.getRange('A3').setValue('WIFE\'S INCOME:');
  sheet.getRange('B3').setValue(0);
  
  // Monthly Totals (Rows 5-7)
  sheet.getRange('A5').setValue('MY MONTHLY TOTAL:');
  sheet.getRange('A6').setValue('WIFE\'S MONTHLY TOTAL:');
  sheet.getRange('A7').setValue('COMBINED MONTHLY TOTAL:');
  
  // Donation Targets (Rows 9-10)
  sheet.getRange('A9').setValue('MY TARGET %:');
  sheet.getRange('B9').setValue(10);
  sheet.getRange('A10').setValue('WIFE\'S TARGET %:');
  sheet.getRange('B10').setValue(10);
  
  // Donation Totals (Rows 12-13)
  sheet.getRange('A12').setValue('MY TOTAL DONATION:');
  sheet.getRange('A13').setValue('WIFE\'S TOTAL DONATION:');
  
  // Donation Percentages (Rows 15-16)
  sheet.getRange('A15').setValue('MY DONATION %:');
  sheet.getRange('A16').setValue('WIFE\'S DONATION %:');
  
  // Remaining Need (Rows 18-19)
  sheet.getRange('A18').setValue('MY REMAINING NEED:');
  sheet.getRange('A19').setValue('WIFE\'S REMAINING NEED:');
  
  // Previous Shortfall (Rows 21-22)
  sheet.getRange('A21').setValue('MY PREVIOUS SHORTFALL:');
  sheet.getRange('B21').setValue(0);
  sheet.getRange('A22').setValue('WIFE\'S PREVIOUS SHORTFALL:');
  sheet.getRange('B22').setValue(0);
  
  // Adjusted Target (Rows 24-25)
  sheet.getRange('A24').setValue('MY ADJUSTED TARGET:');
  sheet.getRange('A25').setValue('WIFE\'S ADJUSTED TARGET:');
  
  // NEW: Grand Total Per Day (Row 26)
  sheet.getRange('A26').setValue('GRAND TOTAL PER DAY:');
  sheet.getRange('A26').setNote('[GrandTotal]');
  
  // Format control panel (now including row 26)
  sheet.getRange('A2:B26')
    .setBackground('#f3f3f3')
    .setFontWeight('bold');
  
  // Highlight grand total row differently
  sheet.getRange('A26:B26')
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setFontSize(11);
  
  // Mark input cells
  var inputCells = ['B2', 'B3', 'B9', 'B10', 'B21', 'B22'];
  for (var i = 0; i < inputCells.length; i++) {
    sheet.getRange(inputCells[i])
      .setBackground('#ffffff')
      .setBorder(true, true, true, true, false, false, '#0066cc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function setupCategories() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 27; // CHANGED FROM 26 TO 27
  var currentRow = startRow;
  
  var categories = [
    {name: 'Personal Expenses', subcats: ['Groceries', 'Restaurants', 'Bring food item', 'Clothes', 'Fruits', 'Other Food', 'Personal care', 'AK personal food', 'AG food', 'Picnic', 'Other']},
    {name: 'Children', subcats: ['Clothing', 'Skin Care', 'Diaper', 'Other']},
    {name: 'Gifts', subcats: ['Gifts', 'Donations', 'Other']},
    {name: 'Health/medical', subcats: ['Doctors/dental/vision', 'Test', 'Pharmacy', 'Emergency', 'Other']},
    {name: 'Home', subcats: ['Wife', 'Iron helper', 'Other']},
    {name: 'Transportation', subcats: ['Fuel', 'Car maintenance', 'Toll tax', 'Public transport', 'Other']},
    {name: 'Utilities', subcats: ['Mobile Packages', 'Other']}
  ];
  
  for (var i = 0; i < categories.length; i++) {
    var cat = categories[i];
    
    // Category header row - highlight entire row and CLEAR columns C, D, E onward
    sheet.getRange(currentRow, 1).setValue(cat.name);
    sheet.getRange(currentRow, 1, 1, sheet.getMaxColumns())
      .setBackground('#d9d9d9')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Clear columns C, D, E and all day columns for category header
    sheet.getRange(currentRow, 3, 1, sheet.getMaxColumns() - 2).clearContent();
    
    currentRow++;
    
    // Each subcategory takes 4 rows
    for (var j = 0; j < cat.subcats.length; j++) {
      // Row 1: Subcategory name with [Totals]
      sheet.getRange(currentRow, 2).setValue(cat.subcats[j]);
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
    
    // Add Category Totals Row
    sheet.getRange(currentRow, 2).setValue(cat.name + ' TOTAL');
    sheet.getRange(currentRow, 2).setNote('[CategoryTotal]');
    sheet.getRange(currentRow, 1, 1, sheet.getMaxColumns())
      .setBackground('#b8b8b8')
      .setFontWeight('bold')
      .setFontSize(10);
    currentRow++;
  }
}

function applyFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Apply control panel formulas
  applyControlPanelFormulas();
  
  // Apply data section formulas (starting row 27) - CHANGED FROM 26 TO 27
  applyDataFormulas(27, lastRow);
}

function applyControlPanelFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // COMBINED MONTHLY TOTAL (B7)
  sheet.getRange('B7').setFormula('=B5+B6');
  
  // MY DONATION % (B15)
  sheet.getRange('B15').setFormula('=IF(B2>0,(B12/B2)*100,0)');
  
  // WIFE'S DONATION % (B16)
  sheet.getRange('B16').setFormula('=IF(B3>0,(B13/B3)*100,0)');
  
  // MY REMAINING NEED (B18)
  sheet.getRange('B18').setFormula('=(B9*B2/100)-B12');
  
  // WIFE'S REMAINING NEED (B19)
  sheet.getRange('B19').setFormula('=(B10*B3/100)-B13');
  
  // MY ADJUSTED TARGET (B24)
  sheet.getRange('B24').setFormula('=MAX(0,B9+(B21/B2)*100)');
  
  // WIFE'S ADJUSTED TARGET (B25)
  sheet.getRange('B25').setFormula('=MAX(0,B10+(B22/B3)*100)');
  
  // Set placeholder values for B5, B6, B12, B13
  sheet.getRange('B5').setValue(0);
  sheet.getRange('B6').setValue(0);
  sheet.getRange('B12').setValue(0);
  sheet.getRange('B13').setValue(0);
}

function applyDataFormulas(startRow, lastRow) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  SpreadsheetApp.flush();
  
  var batchSize = 20;
  for (var batchStart = startRow; batchStart <= lastRow; batchStart += batchSize) {
    var batchEnd = Math.min(batchStart + batchSize - 1, lastRow);
    
    for (var row = batchStart; row <= batchEnd; row++) {
      var subcatVal = sheet.getRange(row, 2).getValue();
      var note = sheet.getRange(row, 2).getNote();
      
      if (note === '[Totals]' && subcatVal !== '') {
        applyTotalsRowFormulas(row, sheet);
      } else if (note === '[Me]') {
        applyMeRowFormulas(row, sheet);
      } else if (note === '[Wife]') {
        applyWifeRowFormulas(row, sheet);
      } else if (note === '[CategoryTotal]') {
        applyCategoryTotalsRowFormulas(row, sheet);
      }
    }
    
    SpreadsheetApp.flush();
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  var totalsRows = [];
  
  for (var r = row - 1; r >= 27; r--) { // CHANGED FROM 26 TO 27
    var cellVal = sheet.getRange(r, 1).getValue();
    if (cellVal !== '') {
      break;
    }
    
    var note = sheet.getRange(r, 2).getNote();
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  if (totalsRows.length > 0) {
    var monthlyFormula = '=';
    var monthlyTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      monthlyTerms.push('C' + totalsRows[i]);
    }
    sheet.getRange(row, 3).setFormula(monthlyFormula + monthlyTerms.join('+'));
  }
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    
    if (totalsRows.length > 0) {
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
}

function applyTotalsRowFormulas(row, sheet) {
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    var meRow = row + 1;
    var wifeRow = row + 2;
    
    var dayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(dayTotalFormula);
    
    var personalFormula = '=' + 
      getColumnLetter(baseCol + 1) + meRow + '+' + 
      getColumnLetter(baseCol + 1) + wifeRow;
    sheet.getRange(row, baseCol + 1).setFormula(personalFormula);
    
    var familyFormula = '=' + 
      getColumnLetter(baseCol + 2) + meRow + '+' + 
      getColumnLetter(baseCol + 2) + wifeRow;
    sheet.getRange(row, baseCol + 2).setFormula(familyFormula);
    
    var donationFormula = '=' + 
      getColumnLetter(baseCol + 3) + meRow + '+' + 
      getColumnLetter(baseCol + 3) + wifeRow;
    sheet.getRange(row, baseCol + 3).setFormula(donationFormula);
  }
}

function applyMeRowFormulas(row, sheet) {
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    
    var myDayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(myDayTotalFormula);
  }
}

function applyWifeRowFormulas(row, sheet) {
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    
    var wifeDayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(wifeDayTotalFormula);
  }
}

function applyFormatting() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  var dayColors = [
    ['#E6E6FF', '#D6D6FF', '#C6C6FF'],
    ['#E6F3FF', '#D6E3FF', '#C6D3FF'],
    ['#E6FFE6', '#D6FFD6', '#C6FFC6'],
    ['#FFF0E6', '#FFE0D6', '#FFD0C6'],
    ['#F0E6FF', '#E0D6FF', '#D0C6FF']
  ];
  
  // Apply colors to input cells ([Me] and [Wife] rows) - CHANGED FROM 26 TO 27
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]' || note === '[Wife]') {
      for (var day = 1; day <= 31; day++) {
        var colorSet = dayColors[(day - 1) % dayColors.length];
        var baseCol = 4 + (day - 1) * 4;
        
        sheet.getRange(row, baseCol + 1).setBackground(colorSet[0]);
        sheet.getRange(row, baseCol + 2).setBackground(colorSet[1]);
        sheet.getRange(row, baseCol + 3).setBackground(colorSet[2]);
      }
    }
    
    // Apply text formatting to comment rows
    if (note === '[Comment]') {
      for (var day = 1; day <= 31; day++) {
        var baseCol = 4 + (day - 1) * 4;
        
        // Set background color for comment cells (light yellow)
        sheet.getRange(row, baseCol + 1).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 2).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 3).setBackground('#FFFEF0');
        
        // Set number format to plain text - this forces text input
        sheet.getRange(row, baseCol + 1).setNumberFormat('@STRING@');
        sheet.getRange(row, baseCol + 2).setNumberFormat('@STRING@');
        sheet.getRange(row, baseCol + 3).setNumberFormat('@STRING@');
      }
    }
  }
  
  // Protect category header rows - CHANGED FROM 26 TO 27
  for (var row = 27; row <= lastRow; row++) {
    var categoryVal = sheet.getRange(row, 1).getValue();
    if (categoryVal !== '') {
      sheet.getRange(row, 3, 1, lastCol - 2).setBackground('#d9d9d9');
    }
  }
  
  // Format grand total row (row 26) with light green and make it stand out
  sheet.getRange(26, 3, 1, lastCol - 2)
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  // Format totals as currency - CHANGED FROM 27 TO 28
  sheet.getRange(28, 3, lastRow - 27, 1).setNumberFormat('"PKR "#,##0.00');
  
  // Number formatting for grand total row
  sheet.getRange(26, 3, 1, lastCol - 2).setNumberFormat('"PKR "#,##0.00');
  
  // Number formatting for all day columns
  for (var day = 1; day <= 31; day++) {
    var baseCol = 4 + (day - 1) * 4;
    sheet.getRange(27, baseCol, lastRow - 26, 4).setNumberFormat('"PKR "#,##0.00');
  }
  
  // Number formatting for control panel
  sheet.getRange('B2:B26').setNumberFormat('"PKR "#,##0.00');
  sheet.getRange('B15:B16').setNumberFormat('0.00"%"');
}

function setFreezePanes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Freeze only row 1 (header)
  sheet.setFrozenRows(1);
  
  // Freeze only columns A-B (Category, Subcategory)
  sheet.setFrozenColumns(2);
}

// Helper function to get column letter from number
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}