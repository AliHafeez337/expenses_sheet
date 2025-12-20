# Diagnostic Scripts: diagnoseCategoryFormulas.gs

## Overview
This script provides diagnostic and repair tools for checking and fixing formula errors in your expense sheet.

## Available Functions

### 1. Diagnose Category Formulas
Checks all formulas within a specific category for errors.

**How to Run:**
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔎 Diagnose Category Formulas**
2. Enter the category name when prompted
3. The script will check:
   - [Totals] row formulas for each subcategory
   - [Me] row formulas
   - [Wife] row formulas
   - Category total row formulas
   - Category summary row formulas
4. Review the results
5. Choose to fix errors automatically if any are found

**What It Checks:**
- Monthly total formulas (Column B)
- Day total formulas (all 31 days)
- Personal, Family, Donation formulas per day
- Category-level aggregations
- Summary row formulas

### 2. Fix Category Formulas
Automatically fixes all formula errors in a specific category.

**How to Run:**
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔧 Fix Category Formulas**
2. Enter the category name when prompted
3. The script will:
   - Diagnose the category
   - Fix all identified errors
   - Report how many formulas were fixed

**Note:** This function re-applies all formulas in the category, so it can fix issues even if diagnostics didn't catch them.

### 3. Diagnose Global Formulas
Checks control panel and grand total formulas that use all categories.

**How to Run:**
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **📊 Diagnose Global Formulas**
2. The script will check:
   - B5: My Monthly Total
   - B6: Wife's Monthly Total
   - B12: My Total Donations
   - B13: Wife's Total Donations
   - Row 26: Grand Total formulas (all 31 days)
3. Review the results
4. Choose to fix errors automatically if any are found

**What It Checks:**
- Control panel formulas use category summary rows correctly
- Grand total row includes all category totals
- Formulas are not missing any categories

## Formula Comparison
The diagnostic script uses **normalized formula comparison**, which means:
- Formulas with terms in different orders are considered equivalent
- Example: `=B10+B20+B30` is the same as `=B30+B10+B20`
- This prevents false positives when formulas are logically correct but ordered differently

## When to Use

### Use Category Diagnostics When:
- You've manually edited formulas and want to verify they're correct
- After adding subcategories, to ensure formulas were updated
- When totals seem incorrect for a specific category
- After migrating to summary rows

### Use Global Diagnostics When:
- Control panel totals seem incorrect
- Grand total row seems wrong
- After adding new categories
- After migrating to summary rows

## Troubleshooting

### No Issues Found
- If diagnostics show no errors but totals still seem wrong, check:
  - Are input cells filled correctly?
  - Are there any manual overrides in formula cells?
  - Run "Fix Category Formulas" to re-apply all formulas

### Many Errors Found
- Don't panic! The fix function can repair all errors automatically
- Click "Yes" when prompted to fix
- The script will update all formulas correctly

### Script Times Out
- For very large sheets, diagnostics might timeout
- Try running diagnostics on individual categories instead
- Use "Fix Category Formulas" which is more efficient

## Related Scripts
- `migrateToCategorySummaryRows.gs` - Migrate to use summary rows
- `addNewCategory.gs` - Add new categories
- `addSubcategories.gs` - Add subcategories

