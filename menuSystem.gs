/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Expense Tracker')
    .addItem('🔧 Complete Setup', 'completeSetup')
    .addItem('➕ Add New Category', 'addNewCategory')
    .addItem('📝 Add Subcategories to Existing', 'addSubcategoriesToExisting')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 Diagnostics & Repair')
      .addItem('🔎 Diagnose Category Formulas', 'diagnoseCategoryFormulas')
      .addItem('🔧 Fix Category Formulas', 'fixCategoryFormulasByName')
      .addSeparator()
      .addItem('📊 Diagnose Global Formulas', 'diagnoseGlobalFormulas')
      .addSeparator()
      .addItem('📋 Check Migration Status', 'checkMigrationStatus')
      .addItem('🔄 Step 1: Add Summary Rows', 'migrateStep1_AddSummaryRows')
      .addItem('🔄 Step 2: Apply Formulas', 'migrateStep2_ApplyFormulas')
      .addItem('🔄 Step 3: Update Control Panel', 'migrateStep3_UpdateControlPanel'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔒 Cell Protection')
      .addItem('⚠️ Warning Mode (Recommended)', 'applyWarningProtection')
      .addItem('🔐 Strict Mode (Full Lock)', 'applyStrictProtection')
      .addItem('🔓 Remove Protection', 'removeAllProtection'))
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

/**
 * Apply warning-only protection (shows warning but allows override)
 */
function applyWarningProtection() {
  protectFormulaCells('warning');
}

/**
 * Apply strict protection (completely blocks editing)
 */
function applyStrictProtection() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Strict Protection Mode',
    'This will FULLY LOCK all formula cells.\n\n' +
    'You will NOT be able to edit them even if you try.\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    protectFormulaCells('strict');
  }
}

/**
 * Show help dialog
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  
  var helpText = 
    '📊 EXPENSE TRACKER HELP\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '🔒 CELL PROTECTION MODES:\n\n' +
    '⚠️ WARNING MODE (Recommended):\n' +
    '   • Shows a warning when you try to edit formula cells\n' +
    '   • You can still override and edit if needed\n' +
    '   • Works on both mobile and desktop\n' +
    '   • Good for preventing accidental changes\n\n' +
    '🔐 STRICT MODE:\n' +
    '   • FULLY BLOCKS editing of formula cells\n' +
    '   • Cannot be overridden (even by owner)\n' +
    '   • Maximum protection against accidental changes\n' +
    '   • Use this if you frequently make mistakes\n\n' +
    '🔓 REMOVE PROTECTION:\n' +
    '   • Removes all protection\n' +
    '   • Use this if you need to manually edit formulas\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '✅ EDITABLE CELLS:\n' +
    '   • Control Panel: Income, Target %, Previous Shortfall\n' +
    '   • Data Rows: [Me] and [Wife] rows (white cells)\n' +
    '   • Comment Rows: [Comment] rows (yellow cells)\n\n' +
    '🚫 PROTECTED CELLS:\n' +
    '   • All cells with formulas (gray background)\n' +
    '   • Category headers\n' +
    '   • Totals rows\n' +
    '   • Monthly total columns\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '➕ ADDING CATEGORIES:\n' +
    '   • Add New Category: Creates a complete new category\n' +
    '   • Add Subcategories: Adds more items to existing category\n' +
    '   • Tip: Use subcategories feature to avoid timeouts\n\n' +
    '🔍 DIAGNOSTICS & REPAIR:\n' +
    '   • Diagnose Category Formulas: Check formulas in a specific category\n' +
    '   • Fix Category Formulas: Automatically repair category formula errors\n' +
    '   • Diagnose Global Formulas: Check control panel & grand total formulas\n' +
    '   • Use these if formulas seem incorrect after adding categories\n\n' +
    '💡 TIP: Protection is automatically applied when you:\n' +
    '   • Run Complete Setup\n' +
    '   • Add a new category\n' +
    '   • Add subcategories\n\n' +
    'You can change protection mode anytime from the menu.';
  
  ui.alert('Help', helpText, ui.ButtonSet.OK);
}