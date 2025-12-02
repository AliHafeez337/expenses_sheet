// Script to automatically create monthly sheets from template

// SETUP INSTRUCTIONS:
// 1. Make sure your template sheet is named "Template"
// 2. Run setupMonthlyTrigger() once to set up automatic monthly execution
// 3. The script will run automatically on the 1st of each month at 6:00 AM

function setupMonthlyTrigger() {
  // Delete existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createMonthlySheetAuto') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger for 1st of every month at 6:00 AM
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
    // Send email notification if something goes wrong
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: 'Error creating monthly expense sheet',
      body: 'There was an error creating the monthly expense sheet:\n\n' + e.toString()
    });
  }
}

function createMonthlySheetManual() {
  // For manual testing - just run this function
  createMonthlySheet();
  SpreadsheetApp.getUi().alert('Monthly sheet created successfully!');
}

function createMonthlySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('Template');
  
  if (!templateSheet) {
    throw new Error('Template sheet not found! Please create a sheet named "Template".');
  }
  
  // Get current month and year
  var now = new Date();
  var monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
  
  // Check if sheet already exists
  if (ss.getSheetByName(monthName)) {
    Logger.log('Sheet "' + monthName + '" already exists. Skipping creation.');
    return;
  }
  
  // Duplicate the template
  var newSheet = templateSheet.copyTo(ss);
  newSheet.setName(monthName);
  
  // Move new sheet to the beginning (after template)
  ss.moveActiveSheet(2);
  
  // Carry forward donation shortfalls from previous month
  carryForwardShortfalls(newSheet, ss);
  
  // Clear all input cells (Me, Wife, Comment rows)
  clearInputCells(newSheet);
  
  // Update month name in sheet
  newSheet.getRange('A1').setValue('Expenses - ' + monthName);
  
  Logger.log('Successfully created sheet: ' + monthName);
}

function carryForwardShortfalls(newSheet, ss) {
  // Get previous month's sheet
  var previousSheet = getPreviousMonthSheet(ss);
  
  if (!previousSheet) {
    // No previous month, set shortfalls to 0
    newSheet.getRange('B21').setValue(0);
    newSheet.getRange('B22').setValue(0);
    return;
  }
  
  // Get remaining needs from previous month (B18 and B19)
  var myRemainingNeed = previousSheet.getRange('B18').getValue() || 0;
  var wifeRemainingNeed = previousSheet.getRange('B19').getValue() || 0;
  
  // Set as shortfalls in new sheet (only if positive)
  newSheet.getRange('B21').setValue(Math.max(0, myRemainingNeed)); // My Previous Shortfall
  newSheet.getRange('B22').setValue(Math.max(0, wifeRemainingNeed)); // Wife's Previous Shortfall
}

function getPreviousMonthSheet(ss) {
  var sheets = ss.getSheets();
  var monthSheets = [];
  
  // Find all sheets with month names (exclude Template)
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
  
  // Sort by date descending
  monthSheets.sort(function(a, b) {
    return b.date.getTime() - a.date.getTime();
  });
  
  // Return the most recent month
  return monthSheets[0].sheet;
}

function isMonthSheet(sheetName) {
  // Check if sheet name matches "Month Year" pattern
  var monthPattern = /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/;
  return monthPattern.test(sheetName);
}

function parseMonthSheet(sheetName) {
  // Parse "Month Year" format into Date object
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
  
  // FIXED: Changed from row 27 to row 28 (first data row after grand total)
  for (var row = 28; row <= lastRow; row++) {
    var note = sheet.getRange(row, 2).getNote();
    
    if (note === '[Me]' || note === '[Wife]' || note === '[Comment]') {
      // Clear all day columns for this row
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4 + 1;
        
        // Clear Personal, Family, Donation columns
        sheet.getRange(row, baseCol + 1).clearContent();
        sheet.getRange(row, baseCol + 2).clearContent();
        sheet.getRange(row, baseCol + 3).clearContent();
      }
    }
  }
  
  // Reset income values to 0 (user can update these)
  sheet.getRange('B2').setValue(0); // My Income
  sheet.getRange('B3').setValue(0); // Wife's Income
}

// Utility function to check trigger status
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

// Function to delete the trigger if needed
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