# Add New Category Script: addNewCategory.gs

## Overview
This script adds a complete new category with subcategories to an existing expense sheet.

## Purpose
Use this when you want to add a new expense category to your sheet without recreating everything.

## How to Run
1. Open your expense tracking Google Sheet
2. Go to **📊 Expense Tracker** menu → **➕ Add New Category**
3. Enter the category name when prompted
4. Enter subcategories (comma-separated) when prompted
5. The script will:
   - Add the category header
   - Create all subcategory rows ([Totals], [Me], [Wife], [Comment])
   - Add category total row
   - Add category summary row
   - Update control panel formulas
   - Update grand total formulas

## Input Format

**Category Name:**
```
Utilities & Bills
```

**Subcategories (comma-separated):**
```
Water, Mobile Packages, TV/Cable, Electricity
```

## What Gets Created

For each new category, the script creates:

1. **Category Header Row**
   - Category name in Column A
   - Formatted with category header style

2. **Subcategory Structure** (for each subcategory):
   - **[Totals] Row**: Contains all formulas for calculations
   - **[Me] Row**: Input cells for your expenses (white background)
   - **[Wife] Row**: Input cells for wife's expenses (white background)
   - **[Comment] Row**: Text input cells (yellow background)

3. **Category Total Row**
   - Sums all subcategory totals
   - Contains formulas for monthly total and all 31 days

4. **Category Summary Row**
   - Aggregates [Me] and [Wife] totals for the category
   - Used by control panel formulas (B5, B6, B12, B13)

## Features
- Automatically inserts in the correct location
- Creates all necessary formulas
- Updates global formulas (control panel and grand total)
- Applies proper formatting
- Handles cell protection

## Notes
- The category is added at the end of existing categories
- All formulas are automatically updated
- The script handles formula length issues by using category summary rows
- Cell protection is automatically applied

## Troubleshooting

### Script Times Out
- Try adding fewer subcategories at once
- You can add more subcategories later using `addSubcategories.gs`

### Formulas Not Working
- Run the diagnostic script to check for errors
- Use "Fix Category Formulas" from the menu

## Related Scripts
- `addSubcategories.gs` - Add more subcategories to an existing category
- `completeSetup.gs` - Create a new sheet from scratch
- `diagnoseCategoryFormulas.gs` - Check and fix formula errors

