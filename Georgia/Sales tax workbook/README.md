# Georgia sales tax workbook

The monthly Georgia sales tax returns, off-road and on-road, in one
workbook. You paste in the month's two reports. Every county's rate comes
from the state's own import template, so no rate is typed anywhere. Two
buttons fill the state's import file for each return.

## Once

1. Download GA Sales Tax Workbook.xlsm and keep it in a folder on the PC.
   If Excel blocks its macros, close it, right-click the file,
   Properties, check Unblock, OK, and open it again.
2. On RECAP, type the off-road account number in C15 and the on-road one
   in D15. Save. Make the state's files puts them on each file's Main
   tab.

## Each month

The steps are on the RECAP tab, under the two buttons. The workbook's
READ ME tab has the rest: what each tab is for, the rates and other
numbers, and notes.

The state's ST-3 Excel import template comes from
dor.georgia.gov/sales-use-tax-import, once a quarter. Make the state's
files saves two filled copies of it, like GA ST-3 off-road 2026-09.xls,
and a box shows each copy's lines next to RECAP's. If a box says Nothing
made or Nothing saved, it says why. Nothing half-made is left behind.

## What it checks

REVIEW flags these, and each line says how it decides. If an AR line was
charged a rate that isn't the state's, it shows the state's rate, and the
return uses the state's. It also flags a county with no state code, a
code with no rate for the month, rates loaded for the wrong quarter, rows
left from another month under the paste, an AR line whose tax name says
neither off road nor on road, stale orders, untaxed fuel with no
exemption, and gallons or sales that didn't land on a return.

## Files

GA Sales Tax Workbook.xlsm is the workbook, with the two buttons and the
3Q26 rates already in it. GASalesTax.bas is the same macro as text. If
the .xlsm won't download, open GA Sales Tax Workbook - BLANK.xlsx and
paste all of GASalesTax.bas into a new module (Alt F11, Insert, Module,
Ctrl A, Ctrl V), as the steps at the top of the .bas say. Then Alt F8
runs GetRates and MakeStateFiles.
The state import files folder holds the 3Q26 template.

## College Park, East Point and Hapeville

On the on-road return these three cities go on line 4 as sales tax. That's
how College Park has always been filed. They share one 1% city tax, so
they go together, and College Park's Clayton side goes with them. The
state's bulletin on that tax, Policy Bulletin SUT-2022-01 (August 30,
2022), calls it prepaid local tax, which would be line 6. The dollars due
are the same either way. Ask DOR. If they say line 6, all of them move there
in one change to the workbook.

## Open

The 4Q26 template isn't posted yet, and October's return needs it. One
question for DOR (1-877-423-6711) still stands. The template's off-road motor fuel
tab claims a flat 3% vendor's compensation, and the state takes the
tiered rate. The workbook matches what the state does and leaves that
tab empty.
