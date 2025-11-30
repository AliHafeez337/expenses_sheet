# Expenses<year> Google Sheet

## Overview
This Google Sheet is designed for comprehensive expense tracking with advanced categorization and monthly reporting capabilities. The sheet automatically creates new monthly sheets from a template and provides detailed insights into personal, family, and donation expenses.

## UPDATED LAYOUT STRUCTURE

### New Row Organization:
- **Row 1**: Days Header (frozen)
- **Rows 2-25**: Control Panel with income, totals, and graphs
- **Rows 26+**: Data Input Section with categories and daily expenses

## DUMMY TABLE VISUALIZATION (For Your Reference Only)
| Category     | Subcategory | Subcat Total   | Day 1 total    | Day 1 - Personal | Day 1 - Family | Day 1 - Donation | Day 2 total    | Day 2 - Personal | Day 2 - Family   | Day 2 - Donation |
|--------------|-------------|----------------|----------------|------------------|----------------|------------------|----------------|------------------|------------------|------------------|
|--------------|-------------|----------------|----------------|------------------|----------------|------------------|----------------|------------------|------------------|------------------|
| Children     |             |$197+$120+...=..|$95+$111+..=$304| $35+$21+...=$88  | $35+$31+...=$12| $25+$41+...=$62  | Day 1 total    | $30+$51+...=$15  | $34+$61+...=$98  | $38+$71+...=$83  |
|--------------| Diapers     | 110 + 87 = $197| $45 + $50 = $95| $15 + $20 = $35  | $25 + $10 = $35| $10 + $15 = $25  | $42+$60 = $102 | $18 + $12 = $30  | $20 + $14 = $34  | $22 + $16 = $38  |
|              |             | $87(all days)  | $20+$10+$15=$45| [Me: $20]        | [Me: $10]      | [Me: $15]        | $12+$14+$16=$42| [Me: $12]        | [Me: $14]        | [Me: $16]        |
|              |             | $110(all days) | $15+$25+$10=$50| [Wife: $15]      | [Wife: $25]    | [Wife: $10]      | $18+$20+$22=$60| [Wife: $18]      | [Wife: $20]      | [Wife: $22]      |
|              |             |                |                | "Size 3"         | "Bulk"         | "Shelter"        | Day 1 total    | ""               | ""               | "comment"        |
|              | Toys        | $120           | $111           | $21              | $31            | $41              | Day 1 total    | $51              | $61              | $71              |

## CALCULATION EXPLANATION (From Your Example):
1) [Me: $20] / [Wife: $15] : Individual entries per subcategory per day per tier
2) $15 + $20 = $35 : total per subcategory per day per tier (total of me and my wife spent for ourself on day1 on diapers)
3) $25 + $10 = $35 : total of me and my wife spent on family on day1 on diapers
4) $18 + $12 = $30 : total of me and my wife donated diapers on day1
5) $20+$10+$15=$45 : total of I spent on diapers on day 1
6) $15+$25+$10=$50 : total of my wife spent on diapers on day 1
7) $45 + $50 = $95 : total of me and my wife spend on diapers on day 1 (on ourselfs, on family and on donation)
8) $21 : the total of me and my wife spent on ourselfs on day1 on toys (same as point 2 but for toys subcategory)
9) $31 : total of me and my wife spent on family toys on day1
10) $41 : total of me and my wife donated toys on day 1
11) $35+$21+...=$88 : total of me and my wife spent on all categories for ourselfs on day 1 (adding point 2 and point 8 and so on... for all subcategories)
12) $35+$31+...=$12 : total of me and my wife spent on all categories for family on day 1 (adding 3 and 9 and so on)
13) $25+$41+...=$62 : total of me and my wife donated all categories on day 1 (adding 4 and 10 and so on)
14) $111 : total of me and my spent on toys on day 1 (just like point 7 but for toys)
15) $95+$111+..=$304 : total of me and my wife spending on all subcategories (of Children category) on day 1
16) $88 + $12 + $62 = $162 : the total me and my wife spent on day 1 on all the subcategories of Children category
17) $87(all days) : total of me spending on diapers in a whole month
18) $110(all days) : total of my wife spending on diapers in a whole month
19) 110 + 87 = $197 : total of me and my wife both spending on diapers in a whole month
20) $120 : total of me and my wife spending on toys in a whole month
21) $197+$120+...=.. : total of me and my wife spending on all subcategories of Children in a whole month

## IMPORTANT NOTE
Cell references (A1, B5, etc.) are conceptual only. Actual positions may shift due to graphs, new rows, or layout changes. All formulas use relative referencing and named ranges.

## FILE STRUCTURE

### NEW ROW STRUCTURE:
- **Row 1**: Days Header (frozen)
- **Rows 2-25**: CONTROL PANEL SECTION (income, totals, graphs)
- **Rows 26+**: DATA INPUT SECTION (categories & subcategories)

### Column Structure (31 Days):
A: Category
B: Subcategory  
C: Subcategory Monthly Total (Me + Wife)
F: Day 1 Total (subcategory)
G: Day 1 - Personal Total
H: Day 1 - Family Total
I: Day 1 - Donation Total
J: Day 1 - Personal [Me]
K: Day 1 - Personal [Wife]
L: Day 1 - Personal [Comment]
M: Day 1 - Family [Me]
N: Day 1 - Family [Wife]
O: Day 1 - Family [Comment]
P: Day 1 - Donation [Me]
Q: Day 1 - Donation [Wife]
R: Day 1 - Donation [Comment]
S: Day 2 Total (subcategory)
... (repeats pattern G-R for each subsequent day, up to Day 31)

## UPDATED CONTROL PANEL SECTION (Rows 2-25)

### Income & Individual Totals:
[My Income]: [Input cell]
[Wife's Income]: [Input cell]
[My Monthly Total]: =SUM(J26+J33+J40+..., M26+M33+M40+..., P26+P33+P40+...) [sums all my spending across entire sheet]
[Wife's Monthly Total]: =SUM(K26+K33+K40+..., N26+N33+N40+..., Q26+Q33+Q40+...) [sums all wife's spending across entire sheet]

### Donation Targets & Progress:
[My Target %]: [Input cell - default 10%]
[Wife's Target %]: [Input cell - default 10%]
[My Total Donation]: =SUM(P26+P33+P40+...) [sums all my donation cells across all categories]
[Wife's Total Donation]: =SUM(Q26+Q33+Q40+...) [sums all wife's donation cells across all categories]
[My Donation %]: =([My Total Donation] / [My Income]) × 100
[Wife's Donation %]: =([Wife's Total Donation] / [Wife's Income]) × 100
[My Remaining Need]: =([My Target %] × [My Income]) - [My Total Donation]
[Wife's Remaining Need]: =([Wife's Target %] × [Wife's Income]) - [Wife's Total Donation]

### Donation Carry-Over:
[My Previous Month Shortfall]: [Auto-filled from previous month]
[Wife's Previous Month Shortfall]: [Auto-filled from previous month]
[My Adjusted Target]: =MAX(0, [Current Target] + [Previous Shortfall])
[Wife's Adjusted Target]: =MAX(0, [Current Target] + [Previous Shortfall])

### Progress Visualization Area
- Donation Progress Bars (My & Wife separately)
- Monthly Spending Pie Chart by Category
- Donation Target Achievement Trend Line
- Category-wise Breakdown Chart
- Monthly Comparison Graphs

## DATA SECTION CALCULATIONS (Starting Row 26)

### Individual Monthly Totals (in CONTROL PANEL):
[My Monthly Total]: =SUM(all my spending cells across all categories and all days)
[Wife's Monthly Total]: =SUM(all wife's spending cells across all categories and all days)

### Subcategory Row Formulas:
[Subcat Monthly Total]: = [My Monthly Total] + [Wife's Monthly Total]
[My Monthly Total]: =SUM(J26+J33+J40+...) [sums all 'Me' cells across all days of month]
[Wife's Monthly Total]: =SUM(K26+K33+K40+...) [sums all 'Wife' cells across all days of month]
[Day X Total]: = [Personal Total] + [Family Total] + [Donation Total]
[Personal Total]: = [Me Personal] + [Wife Personal]
[Family Total]: = [Me Family] + [Wife Family]
[Donation Total]: = [Me Donation] + [Wife Donation]

### Individual Daily Calculations (per subcategory):
[My Day Total]: = [Me Personal] + [Me Family] + [Me Donation]
[Wife's Day Total]: = [Wife Personal] + [Wife Family] + [Wife Donation]

### Category Total Row:
[Category Total]: =SUM(all subcategory C column totals under this category)
[Category Day X Total]: =SUM(all subcategory Day X totals under this category)
[Category Personal Total]: =SUM(all subcategory Personal totals under this category)
[Category Family Total]: =SUM(all subcategory Family totals under this category)
[Category Donation Total]: =SUM(all subcategory Donation totals under this category)

## PRE-DEFINED CATEGORY STRUCTURE

### Personal Expenses:
- Groceries
- Restaurants
- Bring food item
- Clothes
- Fruits
- Other Food
- Personal care (barber etc)
- AK personal food
- AG food
- Picnic
- Other

### Children:
- Clothing
- Skin Care
- Diaper
- Other

### Gifts:
- Gifts
- Donations (charity)
- Other

### Health/medical:
- Doctors/dental/vision
- Test
- Pharmacy
- Emergency
- Other

### Home:
- Wife
- Iron helper
- Other

### Transportation:
- Fuel
- Car maintenance
- Toll tax
- Public transport
- Other

### Utilities:
- Mobile Packages
- Other

### Family Expenses:
- Groceries
- Restaurants
- Bring food item
- Clothes
- AG food
- Picnic
- Fruits
- Other

## TECHNICAL FEATURES

## UPDATED FREEZE CONFIGURATION:
- Freeze Rows: 1-2 (Days header and control panel start)
- Freeze Columns: A-B (Category, Subcategory)

### Input/Output Cell Management:
- Input Cells: White background with colored borders (editable)
- Output Cells: Light gray background, protected from editing
- Daily Columns: Alternating light colors for visual separation
- Category Rows: Different background colors for visual separation

### Data Validation:
- All amount cells: Decimal numbers only (up to 2 decimal places)
- PKR currency format - pure numerical values with "PKR" display
- Comment fields: Text input only (max 100 characters)
- Percentage fields: 0-100% range validation

### Protection Settings:
- All formula cells protected from accidental editing
- Input cells clearly marked and color-coded
- Template structure locked
- Only data entry cells are editable

### Color Coding Scheme:
- Personal Expenses: Light Blue (#E6F3FF)
- Family Expenses: Light Green (#E6FFE6)
- Children: Light Pink (#FFE6E6)
- Health/Medical: Light Orange (#FFE6CC)
- Gifts: Light Purple (#F0E6FF)
- Home: Light Yellow (#FFFFE6)
- Transportation: Light Cyan (#E6FFFF)
- Utilities: Light Gray (#F0F0F0)
- Input Cell Borders: Blue for Personal, Green for Family, Yellow for Donation
- Output Cells: Light Gray (#F8F8F8)

## MOBILE-FRIENDLY FEATURES:
✅ Days header always visible (frozen row 1)
✅ Category names always visible (frozen columns A-B)  
✅ Quick summary always visible (frozen rows 1-2)
✅ Easy scrolling to graphs
✅ Mobile-optimized column widths

## AUTOMATION FEATURES

### Monthly Sheet Creation Script:
```javascript
function createMonthlySheet() {
  // Triggers: 1st day of each month at 6:00 AM
  // 1. Get current month name
  // 2. Duplicate template sheet
  // 3. Rename to current month
  // 4. Carry forward donation shortfalls from previous month
  // 5. Reset daily columns but preserve monthly formulas
  // 6. Update all date references
  // 7. Apply color formatting
  // 8. Set protection on formula cells
}

// Donation Carry-Over Logic:
function calculateDonationCarryOver() {
  // At month-end: Calculate [Remaining Need] for both persons
  const myRemainingNeed = (myTargetPercent * myIncome) - myTotalDonation;
  const wifeRemainingNeed = (wifeTargetPercent * wifeIncome) - wifeTotalDonation;
  
  // If [Remaining Need] > 0, carry to next month as [Previous Shortfall]
  const myPreviousShortfall = myRemainingNeed > 0 ? myRemainingNeed : 0;
  const wifePreviousShortfall = wifeRemainingNeed > 0 ? wifeRemainingNeed : 0;
  
  // If [Remaining Need] <= 0 (target met or exceeded), set [Previous Shortfall] to 0
  // This is handled in the ternary above
  
  // [Adjusted Target] = [Current Target %] + ([Previous Shortfall] / [Income]) × 100
  const myAdjustedTarget = myTargetPercent + (myPreviousShortfall / myIncome) * 100;
  const wifeAdjustedTarget = wifeTargetPercent + (wifePreviousShortfall / wifeIncome) * 100;
  
  return {
    myPreviousShortfall,
    wifePreviousShortfall,
    myAdjustedTarget,
    wifeAdjustedTarget
  };
}

// Progress Tracking Automation:
function updateProgressTracking() {
  // Real-time donation percentage calculation
  const myDonationPercent = (myTotalDonation / myIncome) * 100;
  const wifeDonationPercent = (wifeTotalDonation / wifeIncome) * 100;
  
  // Automatic color coding based on progress status
  function getProgressStatus(donationPercent, monthElapsedPercent) {
    if (donationPercent >= 100) {
      return 'Completed'; // Blue
    } else if (monthElapsedPercent > 90 && donationPercent < 50) {
      return 'Critical'; // Red
    } else if (monthElapsedPercent > 75 && donationPercent < 75) {
      return 'Behind Target'; // Yellow
    } else {
      return 'On Track'; // Green
    }
  }
  
  const myStatus = getProgressStatus(myDonationPercent, monthElapsedPercent);
  const wifeStatus = getProgressStatus(wifeDonationPercent, monthElapsedPercent);
  
  // Monthly summary generation
  generateMonthlySummary();
  
  // Shortfall accumulation tracking
  trackShortfallAccumulation();
}

// Monthly Sheet Creation Script:
function createMonthlySheet() {
  // Triggers: 1st day of each month at 6:00 AM
  const triggerTime = new Date();
  triggerTime.setDate(1);
  triggerTime.setHours(6, 0, 0, 0);
  
  // 1. Get current month name
  const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy');
  
  // 2. Duplicate template sheet
  const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Template');
  const newSheet = templateSheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  
  // 3. Rename to current month
  newSheet.setName(currentMonth);
  
  // 4. Carry forward donation shortfalls from previous month
  carryForwardShortfalls(newSheet);
  
  // 5. Reset daily columns but preserve monthly formulas
  resetDailyColumns(newSheet);
  
  // 6. Update all date references
  updateDateReferences(newSheet, currentMonth);
  
  // 7. Apply color formatting
  applyColorFormatting(newSheet);
  
  // 8. Set protection on formula cells
  setFormulaProtection(newSheet);
}

// Helper Functions:
function carryForwardShortfalls(sheet) {
  const lastMonthData = getLastMonthData();
  sheet.getRange('B20').setValue(lastMonthData.myPreviousShortfall); // My Previous Month Shortfall
  sheet.getRange('B21').setValue(lastMonthData.wifePreviousShortfall); // Wife's Previous Month Shortfall
}

function resetDailyColumns(sheet) {
  const lastRow = sheet.getLastRow();
  // Clear input cells for all days (columns J-R, and repeating pattern)
  for (let day = 1; day <= 31; day++) {
    const dayStartCol = 6 + (day - 1) * 13; // F column for Day 1, then +13 for each subsequent day
    const inputRanges = [
      {col: dayStartCol + 4, label: 'P[Me]'},      // J column for Day 1
      {col: dayStartCol + 5, label: 'P[Wife]'},    // K column for Day 1
      {col: dayStartCol + 7, label: 'F[Me]'},      // M column for Day 1
      {col: dayStartCol + 8, label: 'F[Wife]'},    // N column for Day 1
      {col: dayStartCol + 10, label: 'D[Me]'},     // P column for Day 1
      {col: dayStartCol + 11, label: 'D[Wife]'}    // Q column for Day 1
    ];
    
    inputRanges.forEach(range => {
      sheet.getRange(26, range.col, lastRow - 25, 1).clearContent();
    });
    
    // Clear comment columns
    const commentRanges = [
      {col: dayStartCol + 6, label: 'P[Comment]'},  // L column for Day 1
      {col: dayStartCol + 9, label: 'F[Comment]'},  // O column for Day 1
      {col: dayStartCol + 12, label: 'D[Comment]'}  // R column for Day 1
    ];
    
    commentRanges.forEach(range => {
      sheet.getRange(26, range.col, lastRow - 25, 1).clearContent();
    });
  }
}

function updateDateReferences(sheet, currentMonth) {
  // Update any date references in the control panel
  sheet.getRange('A1').setValue(`Expenses - ${currentMonth}`);
}

function applyColorFormatting(sheet) {
  // Reapply color formatting to ensure consistency
  const categoryColors = {
    'Personal Expenses': '#E6F3FF',
    'Family Expenses': '#E6FFE6',
    'Children': '#FFE6E6',
    'Health/medical': '#FFE6CC',
    'Gifts': '#F0E6FF',
    'Home': '#FFFFE6',
    'Transportation': '#E6FFFF',
    'Utilities': '#F0F0F0'
  };
  
  const lastRow = sheet.getLastRow();
  for (let row = 26; row <= lastRow; row++) {
    const category = sheet.getRange(row, 1).getValue();
    if (categoryColors[category]) {
      sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(categoryColors[category]);
    }
  }
}

function setFormulaProtection(sheet) {
  const protection = sheet.protect();
  protection.setWarningOnly(true);
  protection.setDescription('Protected cells contain formulas - only input cells are editable');
}

function generateMonthlySummary() {
  // Generate monthly summary report
  const summary = {
    totalSpending: myMonthlyTotal + wifeMonthlyTotal,
    totalDonation: myTotalDonation + wifeTotalDonation,
    donationRate: ((myTotalDonation + wifeTotalDonation) / (myIncome + wifeIncome)) * 100,
    categoryBreakdown: getCategoryBreakdown(),
    monthlyTrend: getMonthlyTrend()
  };
  
  return summary;
}

function trackShortfallAccumulation() {
  // Track shortfall accumulation over multiple months
  const shortfallHistory = JSON.parse(PropertiesService.getDocumentProperties().getProperty('shortfallHistory') || '[]');
  
  shortfallHistory.push({
    timestamp: new Date(),
    myShortfall: myPreviousShortfall,
    wifeShortfall: wifePreviousShortfall,
    myAdjustedTarget: myAdjustedTarget,
    wifeAdjustedTarget: wifeAdjustedTarget
  });
  
  PropertiesService.getDocumentProperties().setProperty('shortfallHistory', JSON.stringify(shortfallHistory));
}# expenses_sheet
