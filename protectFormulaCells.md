# Cell Protection Script: protectFormulaCells.gs

## Overview
This script protects formula cells from accidental editing while allowing you to edit input cells.

## Purpose
Prevents accidental changes to formulas that could break your expense tracking calculations.

## How to Run

### Warning Mode (Recommended)
1. Go to **📊 Expense Tracker** menu → **🔒 Cell Protection** → **⚠️ Warning Mode (Recommended)**
2. The script will protect all formula cells
3. When you try to edit a protected cell, you'll see a warning
4. You can still override and edit if needed

**Best for:** Most users - provides protection with flexibility

### Strict Mode (Full Lock)
1. Go to **📊 Expense Tracker** menu → **🔒 Cell Protection** → **🔐 Strict Mode (Full Lock)**
2. Confirm you want strict protection
3. The script will fully lock all formula cells
4. You will NOT be able to edit protected cells (even as owner)

**Best for:** Users who frequently make mistakes and need maximum protection

### Remove Protection
1. Go to **📊 Expense Tracker** menu → **🔒 Cell Protection** → **🔓 Remove Protection**
2. All protection will be removed
3. You can now edit any cell

**Use when:** You need to manually edit formulas or make structural changes

## What Gets Protected

### Protected Cells (Formula Cells):
- All cells with formulas
- Category headers
- [Totals] rows
- Category total rows
- Category summary rows
- Control panel formula cells (B5, B6, B12, B13, etc.)
- Grand total row (Row 26)
- Monthly total columns (Column B)

### Editable Cells (Input Cells):
- Control Panel: Income inputs (B2, B3)
- Control Panel: Target percentages (B7, B8)
- Control Panel: Previous shortfall inputs (B21, B22)
- Data Rows: [Me] rows (white cells)
- Data Rows: [Wife] rows (white cells)
- Data Rows: [Comment] rows (yellow cells)

## Protection Modes Explained

### Warning Mode
- Shows a warning dialog when you try to edit
- You can click "Edit anyway" to proceed
- Works on both desktop and mobile
- Good balance of protection and flexibility

### Strict Mode
- Completely blocks editing of protected cells
- No override option available
- Maximum protection against accidental changes
- Use only if you really need strict protection

## When Protection is Applied

Protection is automatically applied when you:
- Run **Complete Setup**
- Add a new category
- Add subcategories to existing category

You can change the protection mode anytime from the menu.

## Notes
- Protection only affects formula cells
- Input cells remain editable in all modes
- You can switch between modes or remove protection anytime
- Protection helps prevent accidental formula corruption

## Troubleshooting

### Can't Edit a Cell I Should Be Able To
- Make sure you're editing an input cell ([Me], [Wife], or [Comment] row)
- If it's a formula cell, remove protection first if you need to edit it

### Want to Edit Formulas
- Remove protection first
- Make your changes
- Re-apply protection if desired

### Protection Not Working
- Make sure you ran the protection script
- Check that the cell actually has a formula (formula cells are protected)
- Try running the protection script again

## Related Scripts
- `completeSetup.gs` - Automatically applies protection
- `addNewCategory.gs` - Automatically applies protection
- `addSubcategories.gs` - Automatically applies protection

