# Migration Script: migrateToCategorySummaryRows.gs

## Overview
This script migrates existing expense sheets to use the new Category Summary Row structure. The migration is broken into 3 separate steps to avoid script timeout errors.

## Purpose
Adds category summary rows to existing sheets that were created before this feature was implemented. These summary rows help prevent "Service Spreadsheets failed" errors by keeping control panel formulas short.

## How to Run

### Step 1: Add Summary Rows (Structure Only)
1. Open your Google Sheet
2. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔄 Step 1: Add Summary Rows**
3. The script will show how many categories it found
4. Enter how many categories to process (e.g., "5" or "10") or type "all" for all categories
5. Click OK
6. The script will add summary row structure (labels and formatting) after each category total
7. **Note**: Formulas will be added in Step 2

**Tips:**
- Start with a small batch (5-10 categories) to test
- If you have many categories, run this step multiple times
- The script skips categories that already have summary rows

### Step 2: Apply Formulas
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔄 Step 2: Apply Formulas**
2. The script will show how many summary rows need formulas
3. Enter how many to process (e.g., "5" or "10") or type "all" for all
4. Click OK
5. The script will add formulas to the summary rows created in Step 1

**Tips:**
- Run this step after Step 1 is complete
- You can run it multiple times if needed
- The script skips rows that already have formulas

### Step 3: Update Control Panel
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔄 Step 3: Update Control Panel**
2. Click Yes to confirm
3. The script will update B5, B6, B12, B13 to use the new summary rows
4. This is a quick operation (takes seconds)

**Note**: This step should only be run once, after Steps 1 and 2 are complete.

## Check Migration Status
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **📋 Check Migration Status**
2. This shows:
   - Total categories found
   - How many summary rows have been added
   - How many have formulas
   - Which steps are complete

## What Gets Created

Each category summary row contains:
- **Column A**: Empty (contains note `[CategorySummary]`)
- **Column B**: "My total for this category" (label)
- **Column C**: Formula summing all [Me] rows' Column B in this category
- **Column D**: "Wife's total for this category" (label)
- **Column E**: Formula summing all [Wife] rows' Column B in this category
- **Column F**: "My donations for this category" (label)
- **Column G**: Formula summing all [Me] rows' donation columns (all 31 days)
- **Column H**: "Wife's donations for this category" (label)
- **Column I**: Formula summing all [Wife] rows' donation columns (all 31 days)

## Troubleshooting

### Script Times Out
- Use smaller batch sizes (3-5 categories at a time)
- Run each step multiple times until all categories are processed

### Summary Rows Already Exist
- The script will skip categories that already have summary rows
- If you want to recreate them, you'll need to manually delete the existing summary rows first

### Formulas Not Applied
- Make sure Step 1 completed successfully
- Run Step 2 again - it will only process rows that need formulas

### Control Panel Still Shows Old Formulas
- Make sure Steps 1 and 2 are complete
- Run Step 3 to update the control panel formulas

## Related Scripts
- `completeSetup.gs` - Creates new sheets with summary rows automatically
- `addNewCategory.gs` - Adds new categories with summary rows automatically
- `diagnoseCategoryFormulas.gs` - Can check and fix summary row formulas

