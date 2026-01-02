# Monthly Trigger Script: setupMonthlyTrigger.gs

## Overview
This script sets up automatic monthly sheet creation. It will automatically create a new expense sheet on the 1st of each month.

## Purpose
Automate the creation of new monthly expense sheets so you don't have to manually create them each month.

## Prerequisites
1. **Sheet named "TEMPLATE"** (case-sensitive) in your spreadsheet
2. **TEMPLATE sheet must be clean** (input cells empty, formulas intact, income cells at 0)
3. **Script permissions** granted to create sheets and send emails

## How to Set Up

### Initial Setup
1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Copy and paste the full script (including your Sheet ID)
4. Save the script
5. Run `setupMonthlyTrigger()` function once
6. Authorize the script when prompted

**Note:** The trigger is set to run on the 1st of each month at 6:00 AM.

### Custom Menu (Recommended)
The script includes an `onOpen()` function that creates a menu. After saving the script:
1. Refresh your Google Sheet
2. You'll see a new menu: **💰 Monthly Sheets**
3. Use this menu for all operations

## What Happens Automatically
On the 1st of each month at 6:00 AM:
1. Script creates a new sheet by duplicating the TEMPLATE
2. Names it with the current month only (e.g., "January", "February")
3. Moves it right after the TEMPLATE sheet
4. Carries forward donation shortfalls from the previous month
5. Updates the header to show current month
6. Leaves all input cells empty (ready for new month's data)

## Manual Sheet Creation

Two ways to create a monthly sheet manually:

### Method 1: Using the Menu
1. Click **💰 Monthly Sheets → 📅 Create This Month's Sheet**
2. The sheet will be created immediately

### Method 2: Using Script Editor
1. Go to **Extensions → Apps Script**
2. Select `createMonthlySheetManual()` from the dropdown
3. Click **Run**

## Template Requirements

Your TEMPLATE sheet should:
- Be named exactly **"TEMPLATE"** (uppercase)
- Have all formulas and calculations set up correctly
- Have **empty input cells** (no data from previous months)
- Have income cells (B2, B3) set to 0 or empty
- Include all categories and subcategories
- Be properly formatted

**Important:** The TEMPLATE is copied exactly - ensure it's clean before using!

## What Gets Carried Forward

From the previous month's sheet:
- **My Previous Shortfall** (from previous month's B18 to new sheet's B21)
- **Wife's Previous Shortfall** (from previous month's B19 to new sheet's B22)

## What the Script Does NOT Do
- Does NOT clear input cells (TEMPLATE should already be clean)
- Does NOT reset formulas or formatting
- Does NOT add year to sheet names (only month names like "January")

## Check Trigger Status

To verify the trigger is set up:
1. Use menu: **💰 Monthly Sheets → 🔍 Check Trigger Status**
2. Or run `checkTriggerStatus()` in script editor
3. It will show if the trigger is active and when it runs

## Delete Trigger

To stop automatic sheet creation:
1. Use menu: **💰 Monthly Sheets → 🗑️ Delete Trigger**
2. Or run `deleteMonthlyTrigger()` in script editor
3. The trigger will be removed (no more automatic sheets)

## Show Sheet ID

To find your Sheet ID for the script:
1. Use menu: **💰 Monthly Sheets → ℹ️ Show Sheet ID**
2. Copy the ID shown
3. Paste it in the `SPREADSHEET_ID` variable at the top of the script

## Error Handling

### If Sheet Creation Fails:
- You'll receive an email notification at your Google account email
- Check your email for error details
- The script logs all errors to **View → Logs** in script editor

### Common Issues & Solutions:

#### 1. "TEMPLATE sheet not found!"
- Make sure you have a sheet named exactly **"TEMPLATE"**
- Check case sensitivity

#### 2. "Sheet already exists!"
- If a sheet for the current month already exists, the script skips creation
- This prevents duplicate sheets

#### 3. "Cannot call SpreadsheetApp.getUi()"
- This happens when running from time-based triggers (normal)
- UI functions only work when manually triggered

#### 4. Trigger doesn't run
- Check execution logs: **View → Logs** in script editor
- Verify trigger is set: **💰 Monthly Sheets → 🔍 Check Trigger Status**
- Ensure script has proper permissions

## Year Transition Consideration

**Important:** Since sheets are named only by month (no year), you cannot have both "January 2025" and "January 2026" in the same spreadsheet.

**Options:**
1. **Archive old year**: At year-end, move previous year's sheets to a different spreadsheet
2. **Delete old months**: Remove previous year's sheets manually
3. **Contact developer**: To modify script to handle year transitions

## Related Scripts
- `completeSetup.gs` - Create your initial TEMPLATE sheet structure
- `addNewCategory.gs` - Add new categories to TEMPLATE
- `addSubcategories.gs` - Add new subcategories to TEMPLATE

## Notes
- The trigger runs automatically - no manual intervention needed
- New sheets are exact copies of TEMPLATE - keep TEMPLATE updated
- If you modify your expense tracking system, update TEMPLATE first
- All email notifications go to your Google account email
- Script logs can be viewed in **View → Logs** in the script editor

## Support
If you encounter issues:
1. Check the logs first
2. Verify TEMPLATE sheet exists and is clean
3. Ensure script has proper permissions
4. Check trigger status using the menu