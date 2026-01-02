function completeSetup() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear entire sheet first
  sheet.clear();
  
  // Step 1: Setup Headers at Row 1
  setupHeaders();
  SpreadsheetApp.flush();
  
  // Step 2: Setup Control Panel (Rows 2-26) - NOW INCLUDES GRAND TOTAL ROW
  setupControlPanel();
  SpreadsheetApp.flush();
  
  // Step 3: Setup Categories starting at Row 27 (CHANGED FROM 26)
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
  
  // Step 8: Apply grand total formulas
  applyGrandTotalFormulas();
  
  SpreadsheetApp.getUi().alert('Complete setup finished successfully!');
}

// New function to add control panel formulas after all data is set
// NOW USES CATEGORY SUMMARY ROWS instead of individual [Me]/[Wife] rows
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
    
    for (var day = 1; day <= 31; day++) {
      var baseCol = 3 + (day - 1) * 4;
      // Build donation column references
    }
    
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

// NEW FUNCTION: Apply grand total formulas for row 26
function applyGrandTotalFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Find all category total rows
  var categoryTotalRows = [];
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    if (note === '[CategoryTotal]') {
      categoryTotalRows.push(row);
    }
  }
  
  if (categoryTotalRows.length === 0) {
    return; // No category totals found
  }
  
  // Monthly Grand Total (Column B) = Sum of all category totals - CHANGED: was C, now B
  var monthlyTerms = [];
  for (var i = 0; i < categoryTotalRows.length; i++) {
    monthlyTerms.push('B' + categoryTotalRows[i]); // CHANGED: C to B
  }
  sheet.getRange('B26').setFormula('=' + monthlyTerms.join('+')); // CHANGED: C26 to B26
  
  // For each day, sum all category totals
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: Day 1 starts at column C (3) - was D (4)
    
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
  
  var headers = ['Category/Subcategory', 'Monthly Total'];
  
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
  sheet.setColumnWidth(1, 180); // Category/Subcategory (wider now)
  sheet.setColumnWidth(2, 120); // Monthly Total
  
  // Set width for all day columns
  for (var col = 3; col <= headers.length; col++) {
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
  
  // var categories = [
  //   {name: 'Personal Expenses', subcats: ['Groceries', 'Restaurants', 'Bring food item', 'Clothes', 'Fruits', 'Other Food', 'Personal care', 'AK personal food', 'AG food', 'Picnic', 'Other']},
  //   {name: 'Children', subcats: ['Clothing', 'Skin Care', 'Diaper', 'Other']},
  //   {name: 'Gifts', subcats: ['Gifts', 'Donations', 'Other']},
  //   {name: 'Health/medical', subcats: ['Doctors/dental/vision', 'Test', 'Pharmacy', 'Emergency', 'Other']},
  //   {name: 'Home', subcats: ['Wife', 'Iron helper', 'Other']},
  //   {name: 'Transportation', subcats: ['Fuel', 'Car maintenance', 'Toll tax', 'Public transport', 'Other']},
  //   {name: 'Utilities', subcats: ['Mobile Packages', 'Other']}
  // ];

  var categories = [
    // Most Frequent Daily Expenses
    {name: 'Food & Groceries', subcats: ['Groceries', 'Fruits', 'Vegetables', 'Meat/Chicken', 'Dairy', 'Bakery', 'Dry Fruits', 'Snacks', 'Other']},
    {name: 'Dining & Orders', subcats: ['Restaurants', 'Fast Food', 'Online Orders', 'Bring Home', 'Cafe/Tea', 'Picnic Food', 'Other']},
    
  //   // Personal & Routine
  //   {name: 'Personal Care', subcats: ['Clothes', 'Shoes', 'Barber/Salon', 'Cosmetics', 'Accessories', 'Laundry', 'Other']},
  //   {name: 'Fragrances', subcats: ['Perfumes', 'Attars', 'Body Sprays', 'Room Sprays', 'Air Fresheners', 'Incense', 'Other']},
    
  //   // Household
  //   {name: 'Household Essentials', subcats: ['Drinking Water', 'Batteries', 'Cleaning Supplies', 'Toiletries', 'Kitchen Items', 'Detergents', 'Tissues/Paper', 'Other']},
  //   {name: 'Home Maintenance', subcats: ['Repairs', 'Plumbing', 'Electrical', 'Painting', 'Furniture', 'Appliances', 'Decorations', 'Other']},
    
  //   // Family & Pets
  //   {name: 'Children', subcats: ['Clothing', 'Diapers', 'Baby Food', 'Toys', 'Skin Care', 'School Supplies', 'Activities', 'Other']},
  //   {name: 'Cat Care', subcats: ['Cat Food', 'Litter', 'Vet Visits', 'Grooming', 'Toys', 'Medications', 'Accessories', 'Other']},
  //   {name: 'Pocket Money', subcats: ['Mother', 'Wife', 'Children', 'Maid', 'Helper', 'Other']},
    
  //   // Health
  //   {name: 'Health & Medical', subcats: ['Doctor Visits', 'Dental/Vision', 'Lab Tests', 'Pharmacy', 'Emergency', 'Vitamins/Supplements', 'Medical Equipment', 'Other']},
    
  //   // Transportation & Utilities
  //   {name: 'Transportation', subcats: ['Fuel/Petrol', 'Car Maintenance', 'Car Wash', 'Toll Tax', 'Parking', 'Public Transport', 'Ride Sharing', 'Other']},
  //   {name: 'Utilities & Bills', subcats: ['Electricity', 'Gas', 'Water Bill', 'Internet', 'Mobile Packages', 'TV/Cable', 'Landline', 'Other']},
    
  //   // Gifts & Charity
  //   {name: 'Gifts & Charity', subcats: ['Birthday Gifts', 'Wedding Gifts', 'Festival Gifts', 'Charity/Donations', 'Zakat', 'Sadaqah', 'Other']},
    
  //   // Less Frequent
  //   {name: 'Books & Learning', subcats: ['Books', 'Magazines', 'Courses', 'Online Learning', 'Stationery', 'Other']},
  //   {name: 'Education', subcats: ['School Fees', 'Tuition', 'Uniforms', 'Transport', 'School Supplies', 'Other']},
  //   {name: 'Technology', subcats: ['Electronics', 'Gadgets', 'Accessories', 'Repairs', 'Software', 'Apps', 'Cloud Storage', 'Other']},
  //   {name: 'Subscriptions', subcats: ['Streaming Services', 'Online Services', 'Insurance Premium', 'Memberships', 'Other']},
    
  //   // Rare/Occasional
  //   {name: 'Miscellaneous', subcats: ['Emergency Expenses', 'Unexpected', 'Lost/Damaged Items', 'Fines/Penalties', 'Miscellaneous', 'Other']}
  ];
  
  for (var i = 0; i < categories.length; i++) {
    var cat = categories[i];
    
    // Category header row - highlight entire row and CLEAR columns B, C onward
    sheet.getRange(currentRow, 1).setValue(cat.name);
    sheet.getRange(currentRow, 1, 1, sheet.getMaxColumns())
      .setBackground('#d9d9d9')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // Clear columns B, C and all day columns for category header
    sheet.getRange(currentRow, 2, 1, sheet.getMaxColumns() - 1).clearContent();
    
    currentRow++;
    
    // Each subcategory takes 4 rows
    for (var j = 0; j < cat.subcats.length; j++) {
      // Row 1: Subcategory name with [Totals] - NOW IN COLUMN A
      sheet.getRange(currentRow, 1).setValue(cat.subcats[j]);
      sheet.getRange(currentRow, 1).setNote('[Totals]');
      currentRow++;
      
      // Row 2: [Me] - NOW IN COLUMN A
      sheet.getRange(currentRow, 1).setNote('[Me]');
      currentRow++;
      
      // Row 3: [Wife] - NOW IN COLUMN A
      sheet.getRange(currentRow, 1).setNote('[Wife]');
      currentRow++;
      
      // Row 4: [Comment] - NOW IN COLUMN A
      sheet.getRange(currentRow, 1).setNote('[Comment]');
      currentRow++;
    }
    
    // Add Category Totals Row
    sheet.getRange(currentRow, 1).setValue(cat.name + ' TOTAL');
    sheet.getRange(currentRow, 1).setNote('[CategoryTotal]');
    sheet.getRange(currentRow, 1, 1, sheet.getMaxColumns())
      .setBackground('#b8b8b8')
      .setFontWeight('bold')
      .setFontSize(10);
    currentRow++;
    
    // Add Category Summary Row (for control panel formulas)
    sheet.getRange(currentRow, 1).setNote('[CategorySummary]');
    sheet.getRange(currentRow, 2).setValue('My total for this category');
    sheet.getRange(currentRow, 4).setValue('Wife\'s total for this category');
    sheet.getRange(currentRow, 6).setValue('My donations for this category');
    sheet.getRange(currentRow, 8).setValue('Wife\'s donations for this category');
    // Format summary row
    sheet.getRange(currentRow, 1, 1, 9)
      .setBackground('#e8f4f8')
      .setFontWeight('normal')
      .setFontSize(9);
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
  // Formula: This month's remaining need - Previous month's shortfall
  // B18 = This month's remaining need = (B9*B2/100) - B12
  // B21 = Previous shortfall:
  //   - Negative (e.g., -5) = under-donated by $5 (need to add it, so subtracting negative adds)
  //   - Positive (e.g., +5) = over-donated by $5 (need to subtract it)
  // Result: How much more you need to donate (can be negative if you've over-donated)
  sheet.getRange('B24').setFormula('=B18-B21');
  
  // WIFE'S ADJUSTED TARGET (B25)
  // Formula: This month's remaining need - Previous month's shortfall
  // B19 = This month's remaining need = (B10*B3/100) - B13
  // B22 = Previous shortfall:
  //   - Negative (e.g., -5) = under-donated by $5 (need to add it, so subtracting negative adds)
  //   - Positive (e.g., +5) = over-donated by $5 (need to subtract it)
  // Result: How much more you need to donate (can be negative if you've over-donated)
  sheet.getRange('B25').setFormula('=B19-B22');
  
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
      var subcatVal = sheet.getRange(row, 1).getValue(); // CHANGED: column 2 to 1
      var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
      
      if (note === '[Totals]' && subcatVal !== '') {
        applyTotalsRowFormulas(row, sheet);
      } else if (note === '[Me]') {
        applyMeRowFormulas(row, sheet);
      } else if (note === '[Wife]') {
        applyWifeRowFormulas(row, sheet);
      } else if (note === '[CategoryTotal]') {
        applyCategoryTotalsRowFormulas(row, sheet);
      } else if (note === '[CategorySummary]') {
        applyCategorySummaryRowFormulas(row, sheet);
      }
    }
    
    SpreadsheetApp.flush();
  }
}

function applyCategoryTotalsRowFormulas(row, sheet) {
  var totalsRows = [];
  
  for (var r = row - 1; r >= 27; r--) {
    var cellVal = sheet.getRange(r, 1).getValue(); // Check column A
    if (cellVal !== '') {
      // Check if this is a category header (not a subcategory)
      var note = sheet.getRange(r, 1).getNote();
      if (!note || note === '') {
        // This is a category header with no note, stop here
        break;
      }
    }
    
    var note = sheet.getRange(r, 1).getNote(); // CHANGED: column 2 to 1
    if (note === '[Totals]') {
      totalsRows.push(r);
    }
  }
  
  if (totalsRows.length > 0) {
    var monthlyFormula = '=';
    var monthlyTerms = [];
    for (var i = 0; i < totalsRows.length; i++) {
      monthlyTerms.push('B' + totalsRows[i]); // CHANGED: C to B
    }
    sheet.getRange(row, 2).setFormula(monthlyFormula + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  }
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    
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
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: Column C for Day 1
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
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    
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
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    monthlyTerms.push(getColumnLetter(baseCol) + row);
  }
  sheet.getRange(row, 2).setFormula('=' + monthlyTerms.join('+')); // CHANGED: column 3 to 2
  
  for (var day = 1; day <= 31; day++) {
    var baseCol = 2 + (day - 1) * 4 + 1; // SHIFTED: was 3, now 2
    
    var wifeDayTotalFormula = '=' + 
      getColumnLetter(baseCol + 1) + row + '+' + 
      getColumnLetter(baseCol + 2) + row + '+' + 
      getColumnLetter(baseCol + 3) + row;
    sheet.getRange(row, baseCol).setFormula(wifeDayTotalFormula);
  }
}

/**
 * Apply formulas to category summary row
 * This row contains summaries for control panel formulas
 * Column A: "My total for this category" (label)
 * Column B: Sum of all [Me] rows' Column B in this category
 * Column C: "Wife's total for this category" (label)
 * Column D: Sum of all [Wife] rows' Column B in this category
 * Column E: "My donations for this category" (label)
 * Column F: Sum of all [Me] rows' donation columns in this category
 * Column G: "Wife's donations for this category" (label)
 * Column H: Sum of all [Wife] rows' donation columns in this category
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
  
  // Apply colors to input cells ([Me] and [Wife] rows)
  for (var row = 27; row <= lastRow; row++) {
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
        
        // Set background color for comment cells (light yellow)
        sheet.getRange(row, baseCol + 1).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 2).setBackground('#FFFEF0');
        sheet.getRange(row, baseCol + 3).setBackground('#FFFEF0');
        
        // Set number format to plain text
        sheet.getRange(row, baseCol + 1).setNumberFormat('@');
        sheet.getRange(row, baseCol + 2).setNumberFormat('@');
        sheet.getRange(row, baseCol + 3).setNumberFormat('@');
        
        // Add data validation to force text input (this triggers text keyboard on mobile)
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
  
  // Protect category header rows
  for (var row = 27; row <= lastRow; row++) {
    var categoryVal = sheet.getRange(row, 1).getValue();
    var note = sheet.getRange(row, 1).getNote();
    // Category headers have values but no notes
    if (categoryVal !== '' && !note) {
      sheet.getRange(row, 2, 1, lastCol - 1).setBackground('#d9d9d9'); // CHANGED: column 3 to 2
    }
  }
  
  // Format grand total row (row 26) with light green and make it stand out
  sheet.getRange(26, 2, 1, lastCol - 1) // CHANGED: column 3 to 2
    .setBackground('#d9ead3')
    .setFontWeight('bold');
  
  // Format totals as currency
  sheet.getRange(28, 2, lastRow - 27, 1).setNumberFormat('"PKR "#,##0.00'); // CHANGED: column 3 to 2
  
  // Number formatting for grand total row
  sheet.getRange(26, 2, 1, lastCol - 1).setNumberFormat('"PKR "#,##0.00'); // CHANGED: column 3 to 2
  
  // Number formatting for all day columns
  for (var day = 1; day <= 31; day++) {
    var baseCol = 3 + (day - 1) * 4; // SHIFTED: was 4, now 3
    sheet.getRange(27, baseCol, lastRow - 26, 4).setNumberFormat('"PKR "#,##0.00');
  }
  
  // RE-APPLY text formatting to comment rows (after currency formatting above)
  // This ensures comment cells show alphabetic keyboard on mobile devices
  for (var row = 27; row <= lastRow; row++) {
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
  
  // Number formatting for control panel
  sheet.getRange('B2:B26').setNumberFormat('"PKR "#,##0.00');
  sheet.getRange('B15:B16').setNumberFormat('0.00"%"');
}

function setFreezePanes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Freeze only row 1 (header)
  sheet.setFrozenRows(1);
  
  // Freeze only column A (Category/Subcategory) - CHANGED: was 2, now 1
  sheet.setFrozenColumns(1);
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