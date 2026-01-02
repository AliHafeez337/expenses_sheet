# Fix Adjusted Target Formulas: fixAdjustedTargetFormulas.gs

## Overview
This script fixes the Adjusted Target formulas in existing sheets to correctly handle negative shortfall values.

## The Problem
The original formula was incorrect:
```
OLD: =MAX(0,B9+(B21/B2)*100)
```

**Issue:** When shortfall (B21) is negative (e.g., -10), dividing by income and multiplying by 100 gives a negative percentage, which REDUCES the target instead of increasing it.

## The Solution
The corrected formula uses `ABS()` to convert negative shortfalls to positive:
```
NEW: =MAX(0,B9+(ABS(B21)/B2)*100)
```

**How it works:**
- Shortfalls are stored as negative numbers (e.g., -10 means $10 shortfall)
- `ABS()` converts -10 to 10
- The shortfall amount is then converted to a percentage and ADDED to the base target

## Example

**Scenario:**
- Base Target: 10%
- Income: $100
- Previous Shortfall: -$10 (needed $10 more last month)

**Calculation:**
- Shortfall as percentage: (ABS(-10) / 100) × 100 = 10%
- Adjusted Target: 10% + 10% = **20%**

This means you need to donate $20 this month:
- $10 for this month's 10% target
- $10 to catch up on last month's shortfall

## How to Run

### Check Status First
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔍 Check Adjusted Target**
2. This will show:
   - Current formulas in B24 and B25
   - Whether they need fixing
   - What the correct formulas should be

### Fix the Formulas
1. Go to **📊 Expense Tracker** menu → **🔍 Diagnostics & Repair** → **🔧 Fix Adjusted Target**
2. Click Yes to confirm
3. The script will:
   - Update B24 (My Adjusted Target) if needed
   - Update B25 (Wife's Adjusted Target) if needed
   - Show a confirmation message

## What Gets Fixed

- **B24 (My Adjusted Target)**: Updated to use `ABS(B21)` instead of `B21`
- **B25 (Wife's Adjusted Target)**: Updated to use `ABS(B22)` instead of `B22`

## When to Use

- **After migrating from old sheets**: If you have sheets created before this fix
- **If Adjusted Targets seem wrong**: If your adjusted targets are lower than your base targets when you have shortfalls
- **As a preventive measure**: Run the check to verify your formulas are correct

## Notes

- The script only fixes formulas that don't already have `ABS()` in them
- If formulas are already correct, the script will report that no changes are needed
- This fix is automatically included in new sheets created with `completeSetup.gs`

## Related

- `completeSetup.gs` - New sheets automatically use the correct formula
- `Readme.md` - See "Donation Carry-Over" section for formula explanation

