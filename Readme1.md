# Expenses<year> Google Sheet - UPDATED

## Overview
This Google Sheet is designed for comprehensive expense tracking with advanced categorization and monthly reporting capabilities. The sheet automatically creates new monthly sheets from a template and provides detailed insights into personal, family, and donation expenses.

## CRITICAL STRUCTURE CLARIFICATION

### Each Subcategory Takes 4 ROWS (not columns):
For every subcategory (e.g., "Diapers"), there are **4 consecutive rows**:
1. **[Totals] Row**: Contains subcategory name + all calculated totals
2. **[Me] Row**: Input cells for my spending
3. **[Wife] Row**: Input cells for wife's spending
4. **[Comment] Row**: Text input for comments

### Each Day Has 4 COLUMNS:
- Column 1: **Day Total** (formula)
- Column 2: **Personal** (formula for totals row, input for Me/Wife rows)
- Column 3: **Family** (formula for totals row, input for Me/Wife rows)
- Column 4: **Donation** (formula for totals row, input for Me/Wife rows)

## VISUAL EXAMPLE: Diapers Subcategory for Day 1

```
Row Type    | Category | Subcategory | Monthly | Day1 Total | Day1 Personal | Day1 Family | Day1 Donation |
------------|----------|-------------|---------|------------|---------------|-------------|---------------|
[Totals]    |          | Diapers     | $197    | $95        | $35           | $35         | $25           |
[Me]        |          |             | $87     | $45        | $15           | $25         | $10           |
[Wife]      |          |             | $110    | $50        | $20           | $10         | $15           |
[Comment]   |          |             |         |            | "Size 3"      | "Bulk"      | "Shelter"     |
```

### Calculation Flow (Using Above Example):

**For [Me] Row:**
- Input: $15 (Personal), $25 (Family), $10 (Donation)
- Formula: Day1 Total = $15 + $25 + $10 = $45
- Formula: Monthly Total = Sum of all 31 days = $87

**For [Wife] Row:**
- Input: $20 (Personal), $10 (Family), $15 (Donation)
- Formula: Day1 Total = $20 + $10 + $15 = $50
- Formula: Monthly Total = Sum of all 31 days = $110

**For [Totals] Row:**
- Formula: Day1 Personal = Me Personal + Wife Personal = $15 + $20 = $35
- Formula: Day1 Family = Me Family + Wife Family = $25 + $10 = $35
- Formula: Day1 Donation = Me Donation + Wife Donation = $10 + $15 = $25
- Formula: Day1 Total = $35 + $35 + $25 = $95
- Formula: Monthly Total = Me Monthly + Wife Monthly = $87 + $110 = $197

**For [Comment] Row:**
- Input: Text only (e.g., "Size 3", "Bulk", "Shelter")

## UPDATED LAYOUT STRUCTURE

### Row Organization:
- **Row 1**: Column Headers (frozen)
- **Rows 2-25**: Control Panel with income, totals, and graphs
- **Rows 26+**: Data Input Section with categories and subcategories

### Column Structure:
- **Column A**: Category name
- **Column B**: Subcategory name (only in [Totals] row)
- **Column C**: Monthly Total (formulas)
- **Columns D-G**: Day 1 (Total, Personal, Family, Donation)
- **Columns H-K**: Day 2 (Total, Personal, Family, Donation)
- ... continues for all 31 days ...
- **Columns 125-128**: Day 31 (Total, Personal, Family, Donation)

Total columns: 3 + (31 days × 4 columns) = **127 columns**

## DETAILED CALCULATION EXPLANATION

Using the Diapers example from above:

1. **Point 7**: $95 = Total spent on Diapers on Day 1 (by both me and wife)
2. **Point 2**: $35 = Total we spent on ourselves (Personal) on Day 1 for Diapers
3. **Point 3**: $35 = Total we spent on family on Day 1 for Diapers
4. **Point 4**: $25 = Total we donated on Day 1 for Diapers
5. **Point 5**: $45 = My total spending on Diapers on Day 1
6. **Point 6**: $50 = Wife's total spending on Diapers on Day 1
7. **Point 17**: $87 = My total spending on Diapers for entire month
8. **Point 18**: $110 = Wife's total spending on Diapers for entire month
9. **Point 19**: $197 = Combined total spending on Diapers for entire month

### 4x4 Grid Per Subcategory Per Day:

```
        | Day Total (Col1) | Personal (Col2) | Family (Col3) | Donation (Col4) |
--------|------------------|-----------------|---------------|-----------------|
Totals  | Point 7          | Point 2         | Point 3       | Point 4         |
[Me]    | Point 5          | [INPUT]         | [INPUT]       | [INPUT]         |
[Wife]  | Point 6          | [INPUT]         | [INPUT]       | [INPUT]         |
[Comment]| null            | [TEXT]          | [TEXT]        | [TEXT]          |
```

## CONTROL PANEL SECTION (Rows 2-25)

### Income & Individual Totals:
- **My Income** (B2): Input cell
- **Wife's Income** (B3): Input cell
- **My Monthly Total** (B5): =SUM(all my Personal, Family, Donation inputs across all days)
- **Wife's Monthly Total** (B6): =SUM(all wife's Personal, Family, Donation inputs across all days)

### Donation Targets & Progress:
- **My Target %** (B8): Input cell (default 10%)
- **Wife's Target %** (B9): Input cell (default 10%)
- **My Total Donation** (B11): =SUM(all my Donation inputs across all subcategories and days)
- **Wife's Total Donation** (B12): =SUM(all wife's Donation inputs across all subcategories and days)
- **My Donation %** (B14): =(B11 / B2) × 100
- **Wife's Donation %** (B15): =(B12 / B3) × 100
- **My Remaining Need** (B17): =(B8 × B2 / 100) - B11
- **Wife's Remaining Need** (B18): =(B9 × B3 / 100) - B12

### Donation Carry-Over:
- **My Previous Month Shortfall** (B20): Input/Auto-filled from previous month
- **Wife's Previous Month Shortfall** (B21): Input/Auto-filled from previous month
- **My Adjusted Target** (B23): =MAX(0, B8 + (B20/B2) × 100)
- **Wife's Adjusted Target** (B24): =MAX(0, B9 + (B21/B3) × 100)

### Progress Visualization Area (Rows 25+):
- Donation Progress Bars (My & Wife separately)
- Monthly Spending Pie Chart by Category
- Donation Target Achievement Trend Line
- Category-wise Breakdown Chart
- Monthly Comparison Graphs

## DATA SECTION FORMULA PATTERNS (Row 26+)

### [Totals] Row Formulas:
```javascript
Monthly Total (C) = Sum of all Day Totals for this subcategory
Day X Total (Col1) = Day X Personal + Day X Family + Day X Donation
Day X Personal (Col2) = [Me] Personal + [Wife] Personal
Day X Family (Col3) = [Me] Family + [Wife] Family
Day X Donation (Col4) = [Me] Donation + [Wife] Donation
```

### [Me] Row Formulas:
```javascript
Monthly Total (C) = Sum of all My Day Totals (31 days)
Day X Total (Col1) = My Personal + My Family + My Donation
// Personal, Family, Donation (Col2-4) = INPUT CELLS
```

### [Wife] Row Formulas:
```javascript
Monthly Total (C) = Sum of all Wife's Day Totals (31 days)
Day X Total (Col1) = Wife's Personal + Wife's Family + Wife's Donation
// Personal, Family, Donation (Col2-4) = INPUT CELLS
```

### [Comment] Row:
```javascript
// All cells except Category/Subcategory = TEXT INPUT CELLS (no formulas)
```

## PRE-DEFINED CATEGORY STRUCTURE

### Personal Expenses:
- Groceries, Restaurants, Bring food item, Clothes, Fruits, Other Food
- Personal care (barber etc), AK personal food, AG food, Picnic, Other

### Children:
- Clothing, Skin Care, Diaper, Other

### Gifts:
- Gifts, Donations (charity), Other

### Health/medical:
- Doctors/dental/vision, Test, Pharmacy, Emergency, Other

### Home:
- Wife, Iron helper, Other

### Transportation:
- Fuel, Car maintenance, Toll tax, Public transport, Other

### Utilities:
- Mobile Packages, Other

### Family Expenses:
- Groceries, Restaurants, Bring food item, Clothes, AG food, Picnic, Fruits, Other

## TECHNICAL FEATURES

### Freeze Configuration:
- **Freeze Rows**: 1-2 (Header and control panel start)
- **Freeze Columns**: A-C (Category, Subcategory, Monthly Total)

### Cell Types:
- **Input Cells**: White background (#ffffff) - editable by user
- **Formula Cells**: Light gray background (#f8f8f8) - protected
- **Comment Cells**: Light yellow background (#fffef0) - text input only
- **Category Headers**: Dark gray background (#d9d9d9) - entire row

### Data Validation:
- Amount cells: Decimal numbers (up to 2 decimal places)
- Currency format: "PKR #,##0.00"
- Comment fields: Text input only (max 100 characters)
- Percentage fields: 0-100% range

### Color Coding by Category:
- Personal Expenses: Light Blue (#E6F3FF)
- Family Expenses: Light Green (#E6FFE6)
- Children: Light Pink (#FFE6E6)
- Health/Medical: Light Orange (#FFE6CC)
- Gifts: Light Purple (#F0E6FF)
- Home: Light Yellow (#FFFFE6)
- Transportation: Light Cyan (#E6FFFF)
- Utilities: Light Gray (#F0F0F0)

## MOBILE-FRIENDLY FEATURES:
✅ Column headers always visible (frozen row 1)
✅ Category/Subcategory names always visible (frozen columns A-C)
✅ Quick summary always visible (frozen rows 1-2)
✅ Easy horizontal scrolling through days
✅ Optimized column widths (85px for day columns)

## AUTOMATION FEATURES

### Monthly Sheet Creation:
- Triggers: 1st day of each month at 6:00 AM
- Duplicates template sheet
- Carries forward donation shortfalls
- Resets all input cells (Me, Wife, Comment rows)
- Preserves all formulas (Totals rows)
- Updates month name and references
- Applies all formatting and protection

### Real-time Progress Tracking:
- Automatic donation percentage calculation
- Color-coded progress status (Completed/On Track/Behind/Critical)
- Monthly summary generation
- Shortfall accumulation tracking across months

## IMPORTANT NOTES

1. **Each subcategory = 4 rows**: Don't insert rows within a subcategory group
2. **4 columns per day**: Total, Personal, Family, Donation
3. **Input cells are ONLY in [Me], [Wife], and [Comment] rows**
4. **All other cells contain formulas** - do not edit manually
5. **Category rows span entire width** for easy visual separation
6. Cell references in formulas use relative positioning
7. Named ranges may be used for complex formulas