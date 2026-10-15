# Georgia sales tax workbook

The monthly Georgia sales tax returns, off-road and on-road, in one
workbook. You paste in the month's two reports. Every county's rate comes
from the state's own import template, so no rate is typed anywhere. Two
buttons fill the state's import file for each return.

## Each month

1. Open GA Sales Tax Workbook.xlsm. Save a copy for the month.
2. Paste the month's AR sales tax report into the AR SALES tab at A1,
   headers and all. Paste the TERSE into the TERSE tab the same way.
3. On RECAP, type the month as its first day, like 9/1/2026, and the
   collected general ledger amount.
4. Press Get rates and pick the state's ST-3 Excel import template for
   that month's quarter. You only do this once a quarter. The template
   comes from dor.georgia.gov/sales-use-tax-import, and a new one comes
   out each quarter. The rate table hidden inside it goes onto the RATES
   tab.
5. Look at the light at the top of RECAP. If it's red, open REVIEW. Every
   line marked LOOK says what to do.
6. Press Make the state's files. It saves two filled copies of the
   template next to the workbook, like GA ST-3 off-road 2026-09.xls and
   GA ST-3 on-road 2026-09.xls. A box shows each copy's Main tab lines
   next to RECAP's. Every line should match. Then import each copy on its
   return in GTC.

You can also key RECAP's lines 1 to 13 into the state site by hand, the
way you always have.

## Once

On RECAP, type both account numbers, off-road and on-road, in the yellow
cells next to "Account number". Save. Make the state's files puts them on
each file's Main tab.

## What it checks

REVIEW flags these. If an AR line was charged a rate that isn't the
state's, it shows the state's rate, and the return uses the state's. It
also flags a county with no state code, a code with no rate for the
month, rates loaded for the wrong quarter, stale orders, untaxed fuel
with no exemption, and gallons or sales that didn't land on a return.

## Tested

August 2026 comes out to the penny of both returns as filed, lines 1 to
13, on RECAP and on the template's own Main tab after Make the state's
files fills it. That was tested in real Excel against the 3Q26 template.

## Files

GA Sales Tax Workbook.xlsm is the workbook, with the two buttons and the
3Q26 rates already in it. GASalesTax.bas is the same macro as text. If
the .xlsm won't download, open GA Sales Tax Workbook - BLANK.xlsx and
follow the steps at the top of the .bas to paste the macro in. Then Alt
F8 runs GetRates and MakeStateFiles. The state import files folder holds
the 3Q26 template. recalc.ps1 recalculates a workbook in real Excel and
strips the author details. static_lists.json is the state's code for each
county and city area, from the template's motor fuel county list.

## Open

No upload has gone into GTC yet. September 2026's return is the first
one: check GTC's return against RECAP before you submit. The 4Q26
template isn't posted yet, and October's return needs it. One question
for DOR (1-877-423-6711) still stands. The template's off-road motor fuel
tab claims a flat 3% vendor's compensation, and the state takes the
tiered rate. The workbook matches what the state does and leaves that
tab empty.
