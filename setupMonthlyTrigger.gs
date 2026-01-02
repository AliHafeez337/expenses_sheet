// ADD THIS CONSTANT AT THE TOP OF YOUR SCRIPT
var SPREADSHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';

// Script to automatically create monthly sheets from template

// SETUP INSTRUCTIONS:
// 1. Make sure your template sheet is named "TEMPLATE"
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
  
  Logger.log('Monthly trigger set up successfully!');
  
  // Only show UI alert if running manually (not from trigger)
  try {
    SpreadsheetApp.getUi().alert('Monthly trigger set up successfully!\n\nA new sheet will be created automatically on the 1st of each month at 6:00 AM.');
  } catch (e) {
    // If no UI context (running from trigger), just log it
    Logger.log('Trigger setup completed from automated context');
  }
}

function createMonthlySheetAuto() {
  try {
    createMonthlySheet();
    Logger.log('Monthly sheet created successfully via trigger');
  } catch (e) {
    // Log the error
    Logger.log('Error creating monthly sheet: ' + e.toString());
    
    // Try to send email notification if something goes wrong
    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: 'Error creating monthly expense sheet',
        body: 'There was an error creating the monthly expense sheet:\n\n' + e.toString()
      });
    } catch (mailError) {
      Logger.log('Could not send error email: ' + mailError.toString());
    }
  }
}

function createMonthlySheetManual() {
  try {
    createMonthlySheet();
    SpreadsheetApp.getUi().alert('Monthly sheet created successfully!');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error creating sheet: ' + e.toString());
  }
}

function createMonthlySheet() {
  var ss = getSpreadsheet();
  var templateSheet = ss.getSheetByName('TEMPLATE');
  
  if (!templateSheet) {
    throw new Error('TEMPLATE sheet not found! Please create a sheet named "TEMPLATE".');
  }
  
  // Get current month name only (e.g., "January")
  var now = new Date();
  var monthName = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
  
  // Check if sheet already exists - JUST MONTH NAME
  if (ss.getSheetByName(monthName)) {
    Logger.log('Sheet "' + monthName + '" already exists. Skipping creation.');
    SpreadsheetApp.getUi().alert('Sheet "' + monthName + '" already exists!');
    return;
  }
  
  // Duplicate the template
  var newSheet = templateSheet.copyTo(ss);
  newSheet.setName(monthName); // Just month name, no year
  
  // Move new sheet to the beginning (after TEMPLATE)
  var templateIndex = templateSheet.getIndex();
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(templateIndex + 1);
  
  // Carry forward donation shortfalls from previous month
  carryForwardShortfalls(newSheet, ss);
  
  // Update month name in sheet header
  newSheet.getRange('A1').setValue('Expenses - ' + monthName);
  
  Logger.log('Successfully created sheet: ' + monthName);
  return monthName;
}

// NEW FUNCTION: Get spreadsheet by ID or active
function getSpreadsheet() {
  try {
    // First try to get active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss;
  } catch (e) {
    // If no active spreadsheet, use the ID
    if (SPREADSHEET_ID && SPREADSHEET_ID !== 'PASTE_YOUR_SHEET_ID_HERE') {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      throw new Error('Please set your SPREADSHEET_ID in the script or run from an active spreadsheet.');
    }
  }
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
  
  // Set as shortfalls in new sheet
  // Positive remaining need = under-donated (store as negative, e.g., -5 means $5 shortfall)
  // Negative remaining need = over-donated (store as positive, e.g., +5 means $5 over-donation)
  newSheet.getRange('B21').setValue(-myRemainingNeed); // My Previous Shortfall
  newSheet.getRange('B22').setValue(-wifeRemainingNeed); // Wife's Previous Shortfall
}

function getPreviousMonthSheet(ss) {
  var sheets = ss.getSheets();
  var monthSheets = [];
  
  // Find all sheets with month names (exclude TEMPLATE)
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName !== 'TEMPLATE' && isMonthName(sheetName)) {
      monthSheets.push({
        name: sheetName,
        sheet: sheets[i],
        monthIndex: getMonthIndex(sheetName)
      });
    }
  }
  
  if (monthSheets.length === 0) {
    return null;
  }
  
  // Sort by month index (January=0, February=1, etc.)
  monthSheets.sort(function(a, b) {
    return b.monthIndex - a.monthIndex;
  });
  
  // Return the most recent month
  return monthSheets[0].sheet;
}

function isMonthName(sheetName) {
  // Check if sheet name is just a month name
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months.includes(sheetName);
}

function getMonthIndex(monthName) {
  var months = {
    'January': 0, 'February': 1, 'March': 2, 'April': 3,
    'May': 4, 'June': 5, 'July': 6, 'August': 7,
    'September': 8, 'October': 9, 'November': 10, 'December': 11
  };
  return months[monthName];
}

function clearInputCells(sheet) {
  var lastRow = sheet.getLastRow();
  
  // Clear input cells starting from row 27 (first data row)
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote(); // CHANGED: column 2 to 1
    
    if (note === '[Me]' || note === '[Wife]' || note === '[Comment]') {
      // Clear all day columns for this row
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4; // Day columns start at column 3
        
        if (note === '[Me]' || note === '[Wife]') {
          // Clear Personal, Family, Donation input columns
          sheet.getRange(row, baseCol + 1).clearContent(); // Personal
          sheet.getRange(row, baseCol + 2).clearContent(); // Family
          sheet.getRange(row, baseCol + 3).clearContent(); // Donation
        } else if (note === '[Comment]') {
          // Clear comment columns
          sheet.getRange(row, baseCol + 1).clearContent(); // Personal comment
          sheet.getRange(row, baseCol + 2).clearContent(); // Family comment
          sheet.getRange(row, baseCol + 3).clearContent(); // Donation comment
        }
      }
    }
  }
  
  // Reset income values to 0 (user can update these)
  sheet.getRange('B2').setValue(0); // My Income
  sheet.getRange('B3').setValue(0); // Wife's Income
}

// NEW: Custom menu for easy access
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Monthly Sheets')
    .addItem('📅 Create This Month\'s Sheet', 'createMonthlySheetManual')
    .addItem('⏰ Set Up Auto-Creation', 'setupMonthlyTrigger')
    .addItem('🔍 Check Trigger Status', 'checkTriggerStatus')
    .addItem('🗑️ Delete Trigger', 'deleteMonthlyTrigger')
    .addItem('ℹ️ Show Sheet ID', 'showSheetId')
    .addToUi();
}

// NEW: Show your Sheet ID
function showSheetId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = ss.getId();
  
  SpreadsheetApp.getUi().alert(
    'Your Sheet ID:\n\n' + 
    sheetId + '\n\n' +
    'Copy this and paste it in the SPREADSHEET_ID variable at the top of the script.'
  );
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
        '✅ Monthly trigger is ACTIVE\n\n' +
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
      '❌ No monthly trigger found!\n\n' +
      'Run "Set Up Auto-Creation" from the menu to create one.'
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
    SpreadsheetApp.getUi().alert('✅ Monthly trigger deleted successfully!');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ No monthly trigger found to delete.');
  }
}