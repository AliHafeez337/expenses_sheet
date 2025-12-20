# Monthly Trigger Script: setupMonthlyTrigger.gs

## Overview
This script sets up automatic monthly sheet creation. It will automatically create a new expense sheet on the 1st of each month.

## Purpose
Automate the creation of new monthly expense sheets so you don't have to manually create them each month.

## How to Set Up

### Initial Setup
1. Make sure you have a sheet named "Template" in your spreadsheet
2. The Template sheet should have the complete expense tracking structure
3. Go to the script editor and run `setupMonthlyTrigger()` function
4. Or create a custom menu item to run it (if available)

**Note:** The trigger is set to run on the 1st of each month at 6:00 AM.

### What Happens Automatically
On the 1st of each month at 6:00 AM:
1. Script creates a new sheet from the Template
2. Names it with the current month (e.g., "January", "February")
3. Carries forward donation shortfalls from the previous month
4. Clears all input cells (ready for new month's data)
5. Resets income values to 0

## Manual Sheet Creation

If you want to create a monthly sheet manually:
1. Go to the script editor
2. Run `createMonthlySheetManual()` function
3. A new sheet will be created immediately

## Template Requirements

Your Template sheet should:
- Have the complete expense tracking structure
- Include all categories and subcategories
- Have all formulas set up correctly
- Be properly formatted

**Tip:** Use `completeSetup.gs` to create your Template sheet, then save it as "Template".

## What Gets Carried Forward

From the previous month's sheet:
- **My Previous Shortfall** (B21) - Carried to new sheet's B21
- **Wife's Previous Shortfall** (B22) - Carried to new sheet's B22

Only positive values are carried forward (negative values become 0).

## What Gets Cleared

In the new monthly sheet:
- All [Me] row input cells (Personal, Family, Donation columns)
- All [Wife] row input cells (Personal, Family, Donation columns)
- All [Comment] row cells
- Income values (B2, B3) - Reset to 0

**Note:** Formulas and structure remain intact.

## Check Trigger Status

To verify the trigger is set up:
1. Go to the script editor
2. Run `checkTriggerStatus()` function
3. It will show if the trigger is active and when it runs

## Delete Trigger

To stop automatic sheet creation:
1. Go to the script editor
2. Run `deleteMonthlyTrigger()` function
3. The trigger will be removed

## Troubleshooting

### Sheet Not Created Automatically
- Check if the trigger is set up: Run `checkTriggerStatus()`
- Verify you have a "Template" sheet
- Check execution logs for errors
- Make sure the script has permission to create sheets

### Email Notifications
- If sheet creation fails, you'll receive an email notification
- Check your email for error details

### Template Not Found
- Make sure you have a sheet named exactly "Template"
- The name is case-sensitive

### Sheet Already Exists
- If a sheet for the current month already exists, the script skips creation
- This prevents duplicate sheets

## Related Scripts
- `completeSetup.gs` - Create your Template sheet
- `addNewCategory.gs` - Update Template with new categories
- `addSubcategories.gs` - Update Template with new subcategories

## Notes
- The trigger runs automatically - you don't need to do anything
- New sheets are created from the Template, so keep Template updated
- If you add new categories, update the Template sheet first

