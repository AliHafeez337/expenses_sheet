/**
 * Migration Script: Fix Adjusted Target Formulas
 * 
 * This script fixes the Adjusted Target formulas in existing sheets.
 * 
 * The old formula was incorrect:
 * OLD: =MAX(0,B9+(B21/B2)*100)  (calculated percentage)
 * 
 * The new formula uses this month's remaining need - last month's shortfall:
 * NEW: =B18-B21
 * 
 * Shortfall can be positive or negative:
 * - Negative (e.g., -5) = under-donated by $5 (needed $5 more)
 *   Formula: B18 - (-5) = B18 + 5 (adds to remaining need)
 * - Positive (e.g., +5) = over-donated by $5 (paid $5 more than needed)
 *   Formula: B18 - 5 (subtracts from remaining need)
 * 
 * Example 1 (under-donated):
 * - This month's remaining (B18): $4
 * - Previous shortfall: -$5 (needed $5 more last month)
 * - Adjusted target: 4 - (-5) = $9 (need to donate $9 more)
 * 
 * Example 2 (over-donated):
 * - This month's remaining (B18): $4
 * - Previous shortfall: +$5 (over-donated by $5 last month)
 * - Adjusted target: 4 - 5 = -$1 (already over-donated, no need to donate more)
 */
function fixAdjustedTargetFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    'Fix Adjusted Target Formulas',
    'This will update the Adjusted Target formulas (B24 and B25) to correctly handle negative shortfalls.\n\n' +
    'Old formula: =MAX(0,B9+(B21/B2)*100)\n' +
    'New formula: =MAX(0,B9+(ABS(B21)/B2)*100)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Fix My Adjusted Target (B24)
    var currentB24 = sheet.getRange('B24').getFormula();
    // Check if formula needs fixing - new formula should be B18-B21
    if (currentB24 && currentB24.indexOf('B18-B21') === -1 && currentB24.indexOf('B21') >= 0) {
      sheet.getRange('B24').setFormula('=B18-B21');
      Logger.log('Fixed B24: My Adjusted Target');
    } else {
      Logger.log('B24 already has correct formula or is empty');
    }
    
    // Fix Wife's Adjusted Target (B25)
    var currentB25 = sheet.getRange('B25').getFormula();
    // Check if formula needs fixing - new formula should be B19-B22
    if (currentB25 && currentB25.indexOf('B19-B22') === -1 && currentB25.indexOf('B22') >= 0) {
      sheet.getRange('B25').setFormula('=B19-B22');
      Logger.log('Fixed B25: Wife\'s Adjusted Target');
    } else {
      Logger.log('B25 already has correct formula or is empty');
    }
    
    SpreadsheetApp.flush();
    
    ui.alert('Fix Complete!', 
      'Adjusted Target formulas have been updated.\n\n' +
      'B24 (My Adjusted Target): Now uses B18 - B21\n' +
      'B25 (Wife\'s Adjusted Target): Now uses B19 - B22\n\n' +
      'Formula: This month\'s remaining need - Previous month\'s shortfall\n\n' +
      'Example 1 (under-donated):\n' +
      'This month\'s remaining = $4, shortfall = -$5\n' +
      'Adjusted Target = 4 - (-5) = $9 (need to donate more)\n\n' +
      'Example 2 (over-donated):\n' +
      'This month\'s remaining = $4, shortfall = +$5\n' +
      'Adjusted Target = 4 - 5 = -$1 (already over-donated)',
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('Error', 
      'An error occurred while fixing formulas:\n\n' + error.toString(),
      ui.ButtonSet.OK);
    Logger.log('Error fixing Adjusted Target formulas: ' + error.toString());
  }
}

/**
 * Check if Adjusted Target formulas need fixing
 */
function checkAdjustedTargetFormulas() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var b24Formula = sheet.getRange('B24').getFormula();
  var b25Formula = sheet.getRange('B25').getFormula();
  
  // Check if formulas need fixing - new formula should be B18-B21
  var b24NeedsFix = b24Formula && b24Formula.indexOf('B18-B21') === -1 && b24Formula.indexOf('B21') >= 0;
  var b25NeedsFix = b25Formula && b25Formula.indexOf('B19-B22') === -1 && b25Formula.indexOf('B22') >= 0;
  
  var status = 'ADJUSTED TARGET FORMULA STATUS\n';
  status += '================================\n\n';
  
  if (b24NeedsFix) {
    status += '❌ B24 (My Adjusted Target): Needs fix\n';
    status += '   Current: ' + b24Formula + '\n';
    status += '   Should be: =B18-B21\n\n';
  } else {
    status += '✅ B24 (My Adjusted Target): OK\n';
    if (b24Formula) {
      status += '   Formula: ' + b24Formula + '\n\n';
    } else {
      status += '   (Empty)\n\n';
    }
  }
  
  if (b25NeedsFix) {
    status += '❌ B25 (Wife\'s Adjusted Target): Needs fix\n';
    status += '   Current: ' + b25Formula + '\n';
    status += '   Should be: =B19-B22\n\n';
  } else {
    status += '✅ B25 (Wife\'s Adjusted Target): OK\n';
    if (b25Formula) {
      status += '   Formula: ' + b25Formula + '\n\n';
    } else {
      status += '   (Empty)\n\n';
    }
  }
  
  if (b24NeedsFix || b25NeedsFix) {
    status += '\n💡 Run "Fix Adjusted Target Formulas" from the menu to fix.';
  } else {
    status += '\n✅ All formulas are correct!';
  }
  
  ui.alert('Adjusted Target Status', status, ui.ButtonSet.OK);
}

