# Complete Setup Script: completeSetup.gs

## Overview
This script sets up a brand new expense tracking sheet from scratch. It creates the entire structure including the control panel, all categories, subcategories, and formulas.

## Purpose
Use this script when creating a new monthly expense sheet or when setting up the sheet structure for the first time.

## How to Run
1. Open a new or empty Google Sheet
2. Go to **📊 Expense Tracker** menu → **🔧 Complete Setup**
3. The script will prompt you to enter categories and subcategories
4. Follow the prompts to add each category and its subcategories
5. The script will create the entire sheet structure automatically

## What It Creates

### Control Panel (Rows 2-26)
- Income inputs (My Income, Wife's Income)
- Monthly totals (My Monthly Total, Wife's Monthly Total)
- Target percentages
- Remaining needs calculations
- Previous shortfalls
- Donation totals
- Grand total row (Row 26)

### Data Section (Row 27+)
For each category:
1. **Category Header Row** - Category name
2. **Subcategory Rows** (for each subcategory):
   - [Totals] row - Contains formulas
   - [Me] row - Input cells for your expenses
   - [Wife] row - Input cells for wife's expenses
   - [Comment] row - Text input for comments
3. **Category Total Row** - Sums all subcategories in the category
4. **Category Summary Row** - Aggregates [Me] and [Wife] totals for control panel

## Input Format

When prompted, enter categories and subcategories in this format:
```
Category Name
Subcategory1, Subcategory2, Subcategory3
```

Example:
```
Groceries
Vegetables, Fruits, Dairy, Meat
```

## Features
- Automatically creates all formulas
- Sets up proper formatting and colors
- Applies cell protection (warning mode)
- Creates category summary rows for efficient control panel formulas
- Handles 31 days of expense tracking

## Notes
- This script will overwrite existing data in the sheet
- Use this only on new/empty sheets or when you want to start fresh
- The script creates a complete, ready-to-use expense tracking sheet

## Related Scripts
- `addNewCategory.gs` - Add a new category to an existing sheet
- `addSubcategories.gs` - Add subcategories to an existing category
- `migrateToCategorySummaryRows.gs` - Migrate existing sheets to use summary rows

