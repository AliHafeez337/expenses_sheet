function completeSetup() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear entire sheet first
  sheet.clear();
  
  // Step 1: Setup Headers at Row 1
  setupHeaders();
  SpreadsheetApp.flush();
  
  // Step 2: Setup Control Panel (Rows 2-25)
  setupControlPanel();
  SpreadsheetApp.flush();
  
  // Step 3: Setup Categories starting at Row 26
  setupCategories();
  SpreadsheetApp.flush();
  
  // Step 4: Apply all formulas
  applyFormulas();
  SpreadsheetApp.flush();
  
  // Step 5: Apply formatting and colors
  applyFormatting();
  SpreadsheetApp.flush();
  
  // Step 6: Set freeze panes
  setFreezePanes();
  
  // Step 7: Update control panel summary formulas
  updateControlPanelSummaries();
  
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
  
  // MY MONTHLY TOTAL (B5) - sum all [Me] rows
  var myFormula = '=';
  var myTerms = [];
  dayColumns.personal.forEach(function(col) { myTerms.push('SUM(' + col + '27:' + col + lastRow + ')'); });
  dayColumns.family.forEach(function(col) { myTerms.push('SUM(' + col + '27:' + col + lastRow + ')'); });
  dayColumns.donation.forEach(function(col) { myTerms.push('SUM(' + col + '27:' + col + lastRow + ')'); });
  
  // Filter only [Me] rows - we'll sum every 4th row starting from row 27
  // Actually, simpler: sum all C column values where row has [Me] note
  // Even simpler: sum the monthly totals (column C) for [Me] rows
  var meRows = [];
  var wifeRows = [];
  
  for (var row = 27; row <= lastRow; row++) {
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
  
  // MY TOTAL DONATION (B11) - sum donation column for [Me] rows
  var myDonationTerms = [];
  var wifeDonationTerms = [];
  
  for (var row = 27; row <= lastRow; row++) {
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
    sheet.getRange('B11').setFormula('=' + myDonationTerms.join('+'));
  }
  
  if (wifeDonationTerms.length > 0) {
    sheet.getRange('B12').setFormula('=' + wifeDonationTerms.join('+'));
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
  
  // Monthly Totals (Rows 5-6)
  sheet.getRange('A5').setValue('MY MONTHLY TOTAL:');
  sheet.getRange('A6').setValue('WIFE\'S MONTHLY TOTAL:');
  
  // Donation Targets (Rows 8-9)
  sheet.getRange('A8').setValue('MY TARGET %:');
  sheet.getRange('B8').setValue(10);
  sheet.getRange('A9').setValue('WIFE\'S TARGET %:');
  sheet.getRange('B9').setValue(10);
  
  // Donation Totals (Rows 11-12)
  sheet.getRange('A11').setValue('MY TOTAL DONATION:');
  sheet.getRange('A12').setValue('WIFE\'S TOTAL DONATION:');
  
  // Donation Percentages (Rows 14-15)
  sheet.getRange('A14').setValue('MY DONATION %:');
  sheet.getRange('A15').setValue('WIFE\'S DONATION %:');
  
  // Remaining Need (Rows 17-18)
  sheet.getRange('A17').setValue('MY REMAINING NEED:');
  sheet.getRange('A18').setValue('WIFE\'S REMAINING NEED:');
  
  // Previous Shortfall (Rows 20-21)
  sheet.getRange('A20').setValue('MY PREVIOUS SHORTFALL:');
  sheet.getRange('B20').setValue(0);
  sheet.getRange('A21').setValue('WIFE\'S PREVIOUS SHORTFALL:');
  sheet.getRange('B21').setValue(0);
  
  // Adjusted Target (Rows 23-24)
  sheet.getRange('A23').setValue('MY ADJUSTED TARGET:');
  sheet.getRange('A24').setValue('WIFE\'S ADJUSTED TARGET:');
  
  // Format control panel
  sheet.getRange('A2:B24')
    .setBackground('#f3f3f3')
    .setFontWeight('bold');
  
  // Mark input cells
  var inputCells = ['B2', 'B3', 'B8', 'B9', 'B20', 'B21'];
  for (var i = 0; i < inputCells.length; i++) {
    sheet.getRange(inputCells[i])
      .setBackground('#ffffff')
      .setBorder(true, true, true, true, false, false, '#0066cc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function setupCategories() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 26;
  var currentRow = startRow;
  
  var categories = [
    {name: 'Personal Expenses', subcats: ['Groceries', 'Restaurants', 'Bring food item', 'Clothes', 'Fruits', 'Other Food', 'Personal care', 'AK personal food', 'AG food', 'Picnic', 'Other']},
    {name: 'Children', subcats: ['Clothing', 'Skin Care', 'Diaper', 'Other']},
    {name: 'Gifts', subcats: ['Gifts', 'Donations', 'Other']},
    {name: 'Health/medical', subcats: ['Doctors/dental/vision', 'Test', 'Pharmacy', 'Emergency', 'Other']},
    {name: 'Home', subcats: ['Wife', 'Iron helper', 'Other']},
    {name: 'Transportation', subcats: ['Fuel', 'Car maintenance', 'Toll tax', 'Public transport', 'Other']},
    {name: 'Utilities', subcats: ['Mobile Packages', 'Other']},
    {name: 'Family Expenses', subcats: ['Groceries', 'Restaurants', 'Bring food item', 'Clothes', 'AG food', 'Picnic', 'Fruits', 'Other']}
  ];
  
  for (var i = 0; i < categories.length; i++) {
    var cat = categories[i];
    
    // Category header row - highlight entire row
    sheet.getRange(currentRow, 1).setValue(cat.name);
    sheet.getRange(currentRow, 1, 1, sheet.getMaxColumns())
      .setBackground('#d9d9d9')
      .setFontWeight('bold')
      .setFontSize(11);
    var categoryHeaderRow = currentRow;
    currentRow++;
    
    // Store the first subcategory row for totals calculation
    var categoryStartRow = currentRow;
    
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
  
  // Apply data section formulas (starting row 26)
  applyDataFormulas(26, lastRow);
}

function applyControlPanelFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Simple approach: We'll update these formulas after all data is set up
  // For now, just set basic formulas that won't cause timeout
  
  // MY DONATION % (B14)
  sheet.getRange('B14').setFormula('=IF(B2>0,(B11/B2)*100,0)');
  
  // WIFE'S DONATION % (B15)
  sheet.getRange('B15').setFormula('=IF(B3>0,(B12/B3)*100,0)');
  
  // MY REMAINING NEED (B17)
  sheet.getRange('B17').setFormula('=(B8*B2/100)-B11');
  
  // WIFE'S REMAINING NEED (B18)
  sheet.getRange('B18').setFormula('=(B9*B3/100)-B12');
  
  // MY ADJUSTED TARGET (B23)
  sheet.getRange('B23').setFormula('=MAX(0,B8+(B20/B2)*100)');
  
  // WIFE'S ADJUSTED TARGET (B24)
  sheet.getRange('B24').setFormula('=MAX(0,B9+(B21/B3)*100)');
  
  // Set placeholder values for B5, B6, B11, B12
  // These will be updated by the helper function after data setup
  sheet.getRange('B5').setValue(0);
  sheet.getRange('B6').setValue(0);
  sheet.getRange('B11').setValue(0);
  sheet.getRange('B12').setValue(0);
}

function applyDataFormulas(startRow, lastRow) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Process in batches to avoid timeout
  SpreadsheetApp.flush(); // Commit any pending changes
  
  var batchSize = 20;
  for (var batchStart = startRow; batchStart <= lastRow; batchStart += batchSize) {
    var batchEnd = Math.min(batchStart + batchSize - 1, lastRow);
    
    // Process each row in this batch
    for (var row = batchStart; row <= batchEnd; row++) {
      var subcatVal = sheet.getRange(row, 2).getValue();
      var note = sheet.getRange(row, 2).getNote();
      
      // Only process subcategory rows (those with notes)
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
    
    // Flush after each batch
    SpreadsheetApp.flush();
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  // Find all [Totals] rows above this row until we hit a category header
  var totalsRows = [];
  
  for (var r = row - 1; r >= 26; r--) {
    var cellVal = sheet.getRange(r, 1).getValue();
    // If we hit a category header row, stop
    if (cellVal !== '') {
      break;
    }
    
    var note = sheet.getRange(r, 2).getNote();
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  // Monthly Total (Column C) = Sum of all subcategory totals
  if (totalsRows.length > 0) {
    var monthlyFormula = '=';
    var monthlyTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      monthlyTerms.push('C' + totalsRows[i]);
    }
    sheet.getRange(row, 3).setFormula(monthlyFormula + monthlyTerms.join('+'));
  }
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Column D for Day 1
    
    if (totalsRows.length > 0) {
      // Day Total = Sum of all subcategory day totals
      var dayTotalTerms = [];
      for (var i = 0; i < totalsRows.length; i++) {
        dayTotalTerms.push(getColumnLetter(baseCol) + totalsRows[i]);
      }
      sheet.getRange(row, baseCol).setFormula('=' + dayTotalTerms.join('+'));
      
      // Personal Total = Sum of all subcategory personal totals
      var personalTerms = [];
      for (var i = 0; i < totalsRows.length; i++) {
        personalTerms.push(getColumnLetter(baseCol + 1) + totalsRows[i]);
      }
      sheet.getRange(row, baseCol + 1).setFormula('=' + personalTerms.join('+'));
      
      // Family Total = Sum of all subcategory family totals
      var familyTerms = [];
      for (var i = 0; i < totalsRows.length; i++) {
        familyTerms.push(getColumnLetter(baseCol + 2) + totalsRows[i]);
      }
      sheet.getRange(row, baseCol + 2).setFormula('=' + familyTerms.join('+'));
      
      // Donation Total = Sum of all subcategory donation totals
      var donationTerms = [];
      for (var i = 0; i < totalsRows.length; i++) {
        donationTerms.push(getColumnLetter(baseCol + 3) + totalsRows[i]);
      }
      sheet.getRange(row, baseCol + 3).setFormula('=' + donationTerms.join('+'));
    }
  }
}

function applyTotalsRowFormulas(row, sheet) {
  // Monthly Total (Column C) = Sum of Me + Wife for all days
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Day 1 starts at D (column 4)
    monthlyTerms.push(getColumnLetter(baseCol) + row); // Day Total
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Column D for Day 1
    var meRow = row + 1;
    var wifeRow = row + 2;
    
    // Day Total (baseCol) = Personal + Family + Donation
    var dayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(dayTotalFormula);
    
    // Personal Total = Me Personal + Wife Personal
    var personalFormula = '=' + 
      getColumnLetter(baseCol + 1) + meRow + '+' + 
      getColumnLetter(baseCol + 1) + wifeRow;
    sheet.getRange(row, baseCol + 1).setFormula(personalFormula);
    
    // Family Total = Me Family + Wife Family
    var familyFormula = '=' + 
      getColumnLetter(baseCol + 2) + meRow + '+' + 
      getColumnLetter(baseCol + 2) + wifeRow;
    sheet.getRange(row, baseCol + 2).setFormula(familyFormula);
    
    // Donation Total = Me Donation + Wife Donation
    var donationFormula = '=' + 
      getColumnLetter(baseCol + 3) + meRow + '+' + 
      getColumnLetter(baseCol + 3) + wifeRow;
    sheet.getRange(row, baseCol + 3).setFormula(donationFormula);
  }
}

function applyMeRowFormulas(row, sheet) {
  // Monthly Total (Column C) = Sum of all my spending for all days
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    monthlyTerms.push(getColumnLetter(baseCol) + row); // My Day Total
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Column D for Day 1
    
    // My Day Total = My Personal + My Family + My Donation
    var myDayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(myDayTotalFormula);
  }
}

function applyWifeRowFormulas(row, sheet) {
  // Monthly Total (Column C) = Sum of all wife's spending for all days
  var monthlyTerms = [];
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    monthlyTerms.push(getColumnLetter(baseCol) + row); // Wife's Day Total
  }
  sheet.getRange(row, 3).setFormula('=' + monthlyTerms.join('+'));
  
  // For each day
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1; // Column D for Day 1
    
    // Wife's Day Total = Wife's Personal + Wife's Family + Wife's Donation
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
  
  // Simple color scheme from the old script - much faster
  var dayColors = [
    ['#E6E6FF', '#D6D6FF', '#C6C6FF'], // Day 1: Purple theme
    ['#E6F3FF', '#D6E3FF', '#C6D3FF'], // Day 2: Blue theme  
    ['#E6FFE6', '#D6FFD6', '#C6FFC6'], // Day 3: Green theme
    ['#FFF0E6', '#FFE0D6', '#FFD0C6'], // Day 4: Orange theme
    ['#F0E6FF', '#E0D6FF', '#D0C6FF']  // Day 5: Light purple
  ];
  
  // Apply day colors to input columns only
  for (var day = 1; day <= 31; day++) {
    var colorSet = dayColors[(day - 1) % dayColors.length];
    var baseCol = 3 + (day - 1) * 4 + 1;
    
    // Personal, Family, Donation columns (input columns only)
    sheet.getRange(27, baseCol + 1, 100, 1).setBackground(colorSet[0]);   // Personal
    sheet.getRange(27, baseCol + 2, 100, 1).setBackground(colorSet[1]);   // Family
    sheet.getRange(27, baseCol + 3, 100, 1).setBackground(colorSet[2]);   // Donation
  }
  
  // Format totals as currency
  sheet.getRange(27, 3, 100, 1).setNumberFormat('"PKR "#,##0.00');
  
  // Number formatting for all day columns
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4 + 1;
    sheet.getRange(26, baseCol, lastRow - 25, 4).setNumberFormat('"PKR "#,##0.00');
  }
  
  // Number formatting for control panel
  sheet.getRange('B2:B24').setNumberFormat('"PKR "#,##0.00');
  sheet.getRange('B14:B15').setNumberFormat('0.00"%"');
}

function setFreezePanes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Freeze rows 1-2 (header + control panel start)
  sheet.setFrozenRows(2);
  
  // Freeze columns A-C (Category, Subcategory, Monthly Total)
  sheet.setFrozenColumns(3);
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