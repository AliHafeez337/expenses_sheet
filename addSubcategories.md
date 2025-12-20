# Add Subcategories Script: addSubcategories.gs

## Overview
This script adds new subcategories to an existing category in your expense sheet.

## Purpose
Use this when you want to add more expense items to a category that already exists, without recreating the entire category.

## How to Run
1. Open your expense tracking Google Sheet
2. Go to **📊 Expense Tracker** menu → **📝 Add Subcategories to Existing**
3. Enter the category name when prompted
4. Enter new subcategories (comma-separated) when prompted
5. The script will:
   - Find the category
   - Add new subcategory rows before the category total
   - Update category total formulas
   - Update category summary row formulas
   - Update control panel formulas
   - Update grand total formulas

## Input Format

**Category Name:**
```
Groceries
```

**New Subcategories (comma-separated):**
```
Snacks, Beverages, Frozen Foods
```

## What Gets Created

For each new subcategory, the script creates 4 rows:

1. **[Totals] Row**
   - Subcategory name in Column A
   - Contains formulas for:
     - Monthly total (Column B)
     - Day totals (all 31 days)
     - Personal, Family, Donation totals per day

2. **[Me] Row**
   - Note `[Me]` in Column A
   - Input cells for your expenses (white background)
   - Column B: Monthly total formula
   - Day columns: Personal, Family, Donation inputs

3. **[Wife] Row**
   - Note `[Wife]` in Column A
   - Input cells for wife's expenses (white background)
   - Column B: Monthly total formula
   - Day columns: Personal, Family, Donation inputs

4. **[Comment] Row**
   - Note `[Comment]` in Column A
   - Text input cells (yellow background)
   - For notes about expenses

## Features
- Automatically finds the correct category
- Inserts subcategories in the right location
- Updates all related formulas automatically
- Maintains proper formatting
- Handles category summary row updates

## Notes
- Subcategories are added before the category total row
- All formulas are automatically updated
- The script updates the category summary row to include new subcategories
- Control panel formulas are updated to reflect new data

## Troubleshooting

### Category Not Found
- Make sure you enter the exact category name as it appears in the sheet
- Category names are case-sensitive

### Script Times Out
- Try adding fewer subcategories at once
- You can run the script multiple times to add more

### Formulas Not Updated
- Run the diagnostic script to check for errors
- Use "Fix Category Formulas" from the menu

## Related Scripts
- `addNewCategory.gs` - Add a completely new category
- `diagnoseCategoryFormulas.gs` - Check and fix formula errors
- `completeSetup.gs` - Create a new sheet from scratch

