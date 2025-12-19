# Expenses<year> Google Sheet - Complete Documentation

## Overview
This Google Sheet is designed for comprehensive expense tracking with advanced categorization and monthly reporting capabilities. The sheet automatically creates new monthly sheets from a template and provides detailed insights into personal, family, and donation expenses.

## CRITICAL STRUCTURE CLARIFICATION

### Each Subcategory Takes 4 ROWS:
For every subcategory (e.g., "Diapers"), there are **4 consecutive rows**:
1. **[Totals] Row**: Contains subcategory name in Column A + all calculated totals
2. **[Me] Row**: Input cells for my spending (note `[Me]` in Column A)
3. **[Wife] Row**: Input cells for wife's spending (note `[Wife]` in Column A)
4. **[Comment] Row**: Text input for comments (note `[Comment]` in Column A)

### Each Day Has 4 COLUMNS:
Starting from Column C (Day 1), each day has 4 columns:
- **Column 1**: Day Total (formula)
- **Column 2**: Personal (formula for totals row, input for Me/Wife rows)
- **Column 3**: Family (formula for totals row, input for Me/Wife rows)
- **Column 4**: Donation (formula for totals row, input for Me/Wife rows)

### Column Structure:
- **Column A**: Category/Subcategory name (and notes: `[Totals]`, `[Me]`, `[Wife]`, `[Comment]`, `[CategoryTotal]`)
- **Column B**: Monthly Total (formulas)
- **Columns C-DV**: 31 days × 4 columns = 124 columns
  - Day 1: Columns C, D, E, F (Total, Personal, Family, Donation)
  - Day 2: Columns G, H, I, J (Total, Personal, Family, Donation)
  - ... continues for all 31 days ...
  - Day 31: Columns 123-126 (Total, Personal, Family, Donation)
- **Total columns**: 2 + 124 = **126 columns**

### Row Structure:
- **Row 1**: Column Headers (frozen)
- **Rows 2-26**: Control Panel with income, totals, and graphs
  - Row 26: Grand Total Per Day (sums all category totals)
- **Row 27+**: Data Input Section with categories and subcategories

## VISUAL EXAMPLE: Diapers Subcategory for Day 1

```
Row Type    | Column A | Column B | Column C | Column D | Column E | Column F |
------------|----------|----------|----------|----------|----------|----------|
            | Category | Monthly  | Day1 Tot | Day1 Pers| Day1 Fam | Day1 Don |
[Totals]    | Diapers  | $197     | $95      | $35      | $35      | $25      |
[Me]        | [Me]     | $87      | $45      | $15      | $25      | $10      |
[Wife]      | [Wife]   | $110     | $50      | $20      | $10      | $15      |
[Comment]   | [Comment]|          |          | "Size 3" | "Bulk"   | "Shelter"|
```

### Calculation Flow (Using Above Example):

**For [Me] Row:**
- Input: $15 (Personal), $25 (Family), $10 (Donation)
- Formula: Day1 Total (C) = $15 + $25 + $10 = $45
- Formula: Monthly Total (B) = Sum of all 31 day totals = $87

**For [Wife] Row:**
- Input: $20 (Personal), $10 (Family), $15 (Donation)
- Formula: Day1 Total (C) = $20 + $10 + $15 = $50
- Formula: Monthly Total (B) = Sum of all 31 day totals = $110

**For [Totals] Row:**
- Formula: Day1 Personal (D) = Me Personal + Wife Personal = $15 + $20 = $35
- Formula: Day1 Family (E) = Me Family + Wife Family = $25 + $10 = $35
- Formula: Day1 Donation (F) = Me Donation + Wife Donation = $10 + $15 = $25
- Formula: Day1 Total (C) = $35 + $35 + $25 = $95
- Formula: Monthly Total (B) = Me Monthly + Wife Monthly = $87 + $110 = $197

**For [Comment] Row:**
- Input: Text only (e.g., "Size 3", "Bulk", "Shelter")

## CONTROL PANEL SECTION (Rows 2-26)

### Income & Individual Totals:
- **My Income** (B2): Input cell
- **Wife's Income** (B3): Input cell
- **My Monthly Total** (B5): =SUM(all [Me] rows' Column B across entire sheet)
- **Wife's Monthly Total** (B6): =SUM(all [Wife] rows' Column B across entire sheet)
- **Combined Monthly Total** (B7): =B5+B6

### Donation Targets & Progress:
- **My Target %** (B9): Input cell (default 10%)
- **Wife's Target %** (B10): Input cell (default 10%)
- **My Total Donation** (B12): =SUM(all [Me] rows' donation columns across all days)
- **Wife's Total Donation** (B13): =SUM(all [Wife] rows' donation columns across all days)
- **My Donation %** (B15): =IF(B2>0,(B12/B2)*100,0)
- **Wife's Donation %** (B16): =IF(B3>0,(B13/B3)*100,0)
- **My Remaining Need** (B18): =(B9*B2/100)-B12
- **Wife's Remaining Need** (B19): =(B10*B3/100)-B13

### Donation Carry-Over:
- **My Previous Month Shortfall** (B21): Input/Auto-filled from previous month
- **Wife's Previous Month Shortfall** (B22): Input/Auto-filled from previous month
- **My Adjusted Target** (B24): =MAX(0,B9+(B21/B2)*100)
- **Wife's Adjusted Target** (B25): =MAX(0,B10+(B22/B3)*100)

### Grand Total Per Day (Row 26):
- **Monthly Grand Total** (B26): =SUM(all category totals' Column B)
- **Day X Total**: =SUM(all category totals' Day X Total column)
- **Day X Personal**: =SUM(all category totals' Day X Personal column)
- **Day X Family**: =SUM(all category totals' Day X Family column)
- **Day X Donation**: =SUM(all category totals' Day X Donation column)

## DATA SECTION FORMULA PATTERNS (Row 27+)

### Category Structure:
Each category consists of:
1. **Category Header Row**: Category name in Column A (no note)
2. **Subcategories**: 4 rows each ([Totals], [Me], [Wife], [Comment])
3. **Category Total Row**: "CategoryName TOTAL" in Column A with note `[CategoryTotal]`

### [Totals] Row Formulas:
```javascript
Monthly Total (B) = Sum of all Day Totals for this subcategory (31 days)
Day X Total (baseCol) = Day X Personal + Day X Family + Day X Donation
Day X Personal (baseCol+1) = [Me] Personal + [Wife] Personal
Day X Family (baseCol+2) = [Me] Family + [Wife] Family
Day X Donation (baseCol+3) = [Me] Donation + [Wife] Donation

Where baseCol = 3 + (day - 1) * 4
```

### [Me] Row Formulas:
```javascript
Monthly Total (B) = Sum of all My Day Totals (31 days)
Day X Total (baseCol) = My Personal + My Family + My Donation
// Personal, Family, Donation (baseCol+1, baseCol+2, baseCol+3) = INPUT CELLS
```

### [Wife] Row Formulas:
```javascript
Monthly Total (B) = Sum of all Wife's Day Totals (31 days)
Day X Total (baseCol) = Wife's Personal + Wife's Family + Wife's Donation
// Personal, Family, Donation (baseCol+1, baseCol+2, baseCol+3) = INPUT CELLS
```

### [Comment] Row:
```javascript
// All cells except Category/Subcategory = TEXT INPUT CELLS (no formulas)
// Format: Text only, triggers text keyboard on mobile devices
```

### Category Total Row Formulas:
```javascript
Monthly Total (B) = Sum of all [Totals] rows' Column B in this category
Day X Total (baseCol) = Sum of all [Totals] rows' Day X Total in this category
Day X Personal (baseCol+1) = Sum of all [Totals] rows' Day X Personal in this category
Day X Family (baseCol+2) = Sum of all [Totals] rows' Day X Family in this category
Day X Donation (baseCol+3) = Sum of all [Totals] rows' Day X Donation in this category
```

## PRE-DEFINED CATEGORY STRUCTURE

Based on `categoryNames.txt`:

### Most Frequent Daily Expenses:
- **Food & Groceries**: Groceries, Fruits, Vegetables, Meat/Chicken, Dairy, Bakery, Dry Fruits, Snacks, Other
- **Dining & Orders**: Restaurants, Fast Food, Online Orders, Bring Home, Cafe/Tea, Picnic Food, Other

### Personal & Routine:
- **Personal Care**: Clothes, Shoes, Barber/Salon, Cosmetics, Accessories, Laundry, Other
- **Fragrances**: Perfumes, Attars, Body Sprays, Room Sprays, Air Fresheners, Incense, Other

### Household:
- **Household Essentials**: Drinking Water, Batteries, Cleaning Supplies, Toiletries, Kitchen Items, Detergents, Tissues/Paper, Other
- **Home Maintenance**: Repairs, Plumbing, Electrical, Painting, Furniture, Appliances, Decorations, Other

### Family & Pets:
- **Children**: Clothing, Diapers, Baby Food, Toys, Skin Care, School Supplies, Activities, Other
- **Cat Care**: Cat Food, Litter, Vet Visits, Grooming, Toys, Medications, Accessories, Other
- **Pocket Money**: Mother, Wife, Children, Maid, Helper, Other

### Health:
- **Health & Medical**: Doctor Visits, Dental/Vision, Lab Tests, Pharmacy, Emergency, Vitamins/Supplements, Medical Equipment, Other

### Transportation & Utilities:
- **Transportation**: Fuel/Petrol, Car Maintenance, Car Wash, Toll Tax, Parking, Public Transport, Ride Sharing, Other
- **Utilities & Bills**: Electricity Father, Electricity Mother, Gas, Water, Mobile Packages, TV/Cable, Landline/Internet, Other

### Gifts & Charity:
- **Gifts & Charity**: Birthday Gifts, Wedding Gifts, Festival Gifts, Charity/Donations, Zakat, Sadaqah, Other

### Less Frequent:
- **Books & Learning**: Books, Magazines, Courses, Online Learning, Stationery, Other
- **Education**: School Fees, Tuition, Uniforms, Transport, School Supplies, Other
- **Technology**: Electronics, Gadgets, Accessories, Repairs, Software, Apps, Cloud Storage, Other
- **Subscriptions**: Streaming Services, Online Services, Insurance Premium, Memberships, Other

### Rare/Occasional:
- **Miscellaneous**: Emergency Expenses, Unexpected, Lost/Damaged Items, Fines/Penalties, Miscellaneous, Other

## TECHNICAL FEATURES

### Freeze Configuration:
- **Freeze Rows**: 1-2 (Header and control panel start)
- **Freeze Columns**: A-B (Category/Subcategory, Monthly Total)

### Cell Types:
- **Input Cells**: White background (#ffffff) - editable by user
  - Income cells (B2, B3)
  - Target % cells (B9, B10)
  - Previous Shortfall cells (B21, B22)
  - [Me] and [Wife] rows: Personal, Family, Donation columns
- **Formula Cells**: Light gray background (#f8f8f8) - protected
  - All totals and calculated values
  - [Totals] rows
  - Category total rows
  - Control panel formulas
- **Comment Cells**: Light yellow background (#fffef0) - text input only
  - [Comment] rows: Personal, Family, Donation columns
- **Category Headers**: Dark gray background (#d9d9d9) - entire row
- **Category Totals**: Medium gray background (#b8b8b8) - entire row
- **Grand Total Row**: Light green background (#d9ead3) - row 26

### Data Validation:
- Amount cells: Decimal numbers (up to 2 decimal places)
- Currency format: "PKR #,##0.00"
- Comment fields: Text input only (max 100 characters)
- Percentage fields: 0-100% range
- Comment cells: Data validation to force text input (triggers text keyboard on mobile)

### Color Coding by Day:
Input cells for [Me] and [Wife] rows use alternating color sets:
- Day 1, 6, 11, 16, 21, 26, 31: #E6E6FF, #D6D6FF, #C6C6FF
- Day 2, 7, 12, 17, 22, 27: #E6F3FF, #D6E3FF, #C6D3FF
- Day 3, 8, 13, 18, 23, 28: #E6FFE6, #D6FFD6, #C6FFC6
- Day 4, 9, 14, 19, 24, 29: #FFF0E6, #FFE0D6, #FFD0C6
- Day 5, 10, 15, 20, 25, 30: #F0E6FF, #E0D6FF, #D0C6FF

### Column Widths:
- Column A (Category/Subcategory): 180px
- Column B (Monthly Total): 120px
- Day columns (C-DV): 85px each

## MOBILE-FRIENDLY FEATURES:
✅ Column headers always visible (frozen row 1)
✅ Category/Subcategory names always visible (frozen columns A-B)
✅ Quick summary always visible (frozen rows 1-2)
✅ Easy horizontal scrolling through days
✅ Optimized column widths (85px for day columns)
✅ Text validation on comment cells triggers text keyboard on mobile

## AUTOMATION FEATURES

### Available Scripts:
1. **completeSetup**: Initial setup of the entire sheet
2. **addNewCategory**: Add a new category with subcategories
3. **addSubcategoriesToExisting**: Add subcategories to an existing category
4. **diagnoseCategoryFormulas**: Check formulas in a specific category
5. **diagnoseGlobalFormulas**: Check control panel and grand total formulas
6. **fixCategoryFormulasByName**: Fix formulas in a specific category
7. **protectFormulaCells**: Apply protection to formula cells (warning or strict mode)
8. **setupMonthlyTrigger**: Set up automatic monthly sheet creation

### Monthly Sheet Creation:
- Triggers: 1st day of each month at 6:00 AM
- Duplicates template sheet
- Carries forward donation shortfalls
- Resets all input cells ([Me], [Wife], [Comment] rows)
- Preserves all formulas ([Totals] rows, category totals)
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
6. **Notes in Column A**: Used to identify row types (`[Totals]`, `[Me]`, `[Wife]`, `[Comment]`, `[CategoryTotal]`)
7. **Column B**: Monthly totals for all rows
8. **Day columns start at Column C**: baseCol = 3 + (day - 1) * 4
9. **Data starts at Row 27**: Control panel is rows 2-26
10. **Grand Total Row**: Row 26 sums all category totals

## KNOWN LIMITATIONS & FUTURE IMPROVEMENTS

### Current Issue:
- Control panel formulas (B5, B6, B12, B13) and grand total formulas (Row 26) can become very long with many categories
- This can cause "Service Spreadsheets failed" errors when formulas exceed ~50,000 characters

### Planned Solution:
- Add a **Category Summary Row** after each category total row
- This row will contain 8 cells:
  - Column A: "My total for this category" (label)
  - Column B: Sum of all [Me] rows' Column B in this category
  - Column C: "Wife's total for this category" (label)
  - Column D: Sum of all [Wife] rows' Column B in this category
  - Column E: "My donations for this category" (label)
  - Column F: Sum of all [Me] rows' donation columns in this category
  - Column G: "Wife's donations for this category" (label)
  - Column H: Sum of all [Wife] rows' donation columns in this category
- Then B5, B6, B12, B13, and Row 26 will sum these summary rows instead of all individual rows
- This will keep formulas short and manageable

## DIAGNOSTIC TOOLS

### Category Formula Diagnostics:
- Checks all formulas within a specific category
- Verifies [Totals], [Me], [Wife] row formulas
- Verifies category total row formulas
- Offers automatic repair

### Global Formula Diagnostics:
- Checks control panel summaries (B5, B6, B12, B13)
- Checks grand total row (Row 26)
- Verifies formulas include all categories
- Offers automatic repair

Both diagnostic tools use normalized formula comparison (ignores term order) to avoid false positives.
