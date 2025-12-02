function setFreezePanes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Freeze only row 1 (header)
  sheet.setFrozenRows(1);
  
  // Freeze only column A (Category/Subcategory) - CHANGED: was 2, now 1
  sheet.setFrozenColumns(1);
}

// ============================================
// PROTECTION FUNCTIONS
// ============================================

/**
 * Protects all formula cells from accidental editing
 * MODE: 'warning' = Shows warning but allows override
 *       'strict' = Completely blocks editing (even for owner)
 */
function protectFormulaCells(mode) {
  mode = mode || 'warning'; // Default to warning mode
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // Step 1: Protect the entire sheet first
  var protection = sheet.protect().setDescription('Formula cells are protected');
  
  // Step 2: Build list of unprotected ranges (input cells only)
  var unprotectedRanges = [];
  
  // Control Panel Input Cells (rows 2-25)
  unprotectedRanges.push(sheet.getRange('B2')); // My Income
  unprotectedRanges.push(sheet.getRange('B3')); // Wife's Income
  unprotectedRanges.push(sheet.getRange('B9')); // My Target %
  unprotectedRanges.push(sheet.getRange('B10')); // Wife's Target %
  unprotectedRanges.push(sheet.getRange('B21')); // My Previous Shortfall
  unprotectedRanges.push(sheet.getRange('B22')); // Wife's Previous Shortfall
  
  // Data Section Input Cells (rows 27+)
  for (var row = 27; row <= lastRow; row++) {
    var note = sheet.getRange(row, 1).getNote();
    
    if (note === '[Me]' || note === '[Wife]') {
      // Unprotect all day input cells (Personal, Family, Donation)
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        unprotectedRanges.push(sheet.getRange(row, baseCol + 1)); // Personal
        unprotectedRanges.push(sheet.getRange(row, baseCol + 2)); // Family
        unprotectedRanges.push(sheet.getRange(row, baseCol + 3)); // Donation
      }
    } else if (note === '[Comment]') {
      // Unprotect all comment cells
      for (var day = 1; day <= 31; day++) {
        var baseCol = 3 + (day - 1) * 4;
        unprotectedRanges.push(sheet.getRange(row, baseCol + 1));
        unprotectedRanges.push(sheet.getRange(row, baseCol + 2));
        unprotectedRanges.push(sheet.getRange(row, baseCol + 3));
      }
    }
  }
  
  // Step 3: Set unprotected ranges
  protection.setUnprotectedRanges(unprotectedRanges);
  
  // Step 4: Apply protection mode
  if (mode === 'warning') {
    protection.setWarningOnly(true);
    SpreadsheetApp.getUi().alert(
      'Protection Applied (Warning Mode)',
      'Formula cells are now protected with warnings.\n\n' +
      'You will see a warning if you try to edit them, but can override it.\n\n' +
      'To change to strict mode, run: protectFormulaCells(\'strict\')',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else if (mode === 'strict') {
    protection.setWarningOnly(false);
    // Remove all editors except the current user (makes it strict even for owner)
    var me = Session.getEffectiveUser();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(me);
    
    SpreadsheetApp.getUi().alert(
      'Protection Applied (Strict Mode)',
      'Formula cells are now FULLY PROTECTED.\n\n' +
      'You cannot edit them even if you try.\n\n' +
      'To change to warning mode, run: protectFormulaCells(\'warning\')',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  SpreadsheetApp.flush();
}

/**
 * Removes all protection from the sheet
 */
function removeAllProtection() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }
  
  SpreadsheetApp.getUi().alert(
    'Protection Removed',
    'All cell protection has been removed from this sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  SpreadsheetApp.flush();
}