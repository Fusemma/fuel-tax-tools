# Georgia sales tax workbook

The monthly Georgia sales tax returns, off-road and on-road, in one
workbook. You paste in the month's two reports. Every county's rate comes
from the state's own import template, so no rate is typed anywhere. Two
buttons fill the state's import file for each return.

## Once

1. Download GA Sales Tax Workbook.xlsm and keep it in a folder on the PC.
   If Excel blocks its macros, close it, right-click the file,
   Properties, tick Unblock, OK, and open it again.
2. On RECAP, type the off-road account number in C15 and the on-road one
   in D15. Save. Make the state's files puts them on each file's Main
   tab.

## Each month

1. Open your copy of the workbook. Save As with the month in the name,
   like GA Sales Tax 2026-09.xlsm. The rates and account numbers come
   along.
2. Paste the month's AR sales tax report into the AR SALES tab at A1,
   headers and all, over last month's. Paste the TERSE into the TERSE tab
   the same way.
3. On RECAP, type the month as its first day, like 9/1/2026, and the
   collected general ledger amount.
4. First month of a quarter only: download the state's ST-3 Excel import
   template for the quarter from dor.georgia.gov/sales-use-tax-import.
   Press Get rates and pick it. The rate table hidden inside it goes onto
   the RATES tab and stays there for the rest of the quarter. If you skip
   this, REVIEW says the rates don't cover the month.
5. Look at the light at the top of RECAP. If it's red, open REVIEW. Every
   line marked LOOK says what to do.
6. Press Make the state's files. It saves two filled copies of the
   template, like GA ST-3 off-road 2026-09.xls and GA ST-3 on-road
   2026-09.xls. The box says which folder they're in and shows each
   copy's Main tab lines next to RECAP's. Every line should match. Import
   the two files the box names, each on its own return in GTC.

You can also key RECAP's lines 1 to 13 into the state site by hand, the
way you always have.

If a box says Nothing made or Nothing saved, it says why, and where it
stopped. Nothing half-made is left behind. The status bar at the bottom of
Excel shows each step while it works.

## What it checks

REVIEW flags these. If an AR line was charged a rate that isn't the
state's, it shows the state's rate, and the return uses the state's. It
also flags a county with no state code, a code with no rate for the
month, rates loaded for the wrong quarter, an AR line whose tax name says
neither off road nor on road, stale orders, untaxed fuel with no
exemption, and gallons or sales that didn't land on a return.

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
the 3Q26 template.

## College Park, East Point and Hapeville

On the on-road return these three cities go on line 4 as sales tax. That's
how College Park has always been filed. They share one 1% city tax, so
they go together, and College Park's Clayton side goes with them. The
state's bulletin on that tax, Policy Bulletin SUT-2022-01 (August 30,
2022), calls it prepaid local tax, which would be line 6. The dollars due
are the same either way. Ask DOR. If they say line 6, the build moves all
of them in one line.

## Open

No upload has gone into GTC yet. September 2026's return is the first
one: check GTC's return against RECAP before you submit. The 4Q26
template isn't posted yet, and October's return needs it. One question
for DOR (1-877-423-6711) still stands. The template's off-road motor fuel
tab claims a flat 3% vendor's compensation, and the state takes the
tiered rate. The workbook matches what the state does and leaves that
tab empty.
