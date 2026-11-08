# Ohio PAT recap

The quarterly Ohio PAT (Petroleum Activity Tax) recap in one workbook. A
tab for each month holds the lines, QTR RECAP adds the three months up
down to GL RESIDUAL and works out what you're paying on the quarter
return, and RATES holds each quarter's rates. A status light at the top of
QTR RECAP says whether every check is OK.

Two buttons come with the macro: Add row puts one more empty row at the
bottom of the section you're in, and Tidy up takes out the empty rows.
The workbook works the same without the macro. Every total is a formula;
the buttons only save time.

## Putting the macro in once

1. Open OH PAT Recap - BLANK.xlsx.
2. Alt F11, then Insert, Module.
3. Open PATRecap.bas as text, press Ctrl A and Ctrl C, click in the new
   module, then Ctrl A and Ctrl V to paste all of it in.
4. Close that window. Save the workbook as a macro-enabled workbook
   (.xlsm).
5. Alt F8, pick TidyUp, Run. The first run puts the buttons on: Add row
   and Tidy up on each month tab, Tidy up on QTR RECAP. After that, press
   the buttons.

## Each quarter

1. Open the workbook. On QTR RECAP, type the first day of the quarter in
   A1, like 1/1/2027. Every tab follows it, and A2 shows the quarter.
2. On RATES, type the quarter's average wholesale diesel and gas prices
   from the Ohio Tax Commissioner's notice. The PAT rate and the tax per
   gallon work themselves out. Add the notice's source line under the note.
3. On QTR RECAP, type the account number in D2. Each month tab shows it.
4. On each month tab, type Collected, Paid through AP and AR in price in
   the shaded cells at the top, and AP in price, per AP in Q4.
5. Paste or type each line's gallons in G and product in H. Pick the
   product from the dropdown: NO2ULS, NO2ULR, KERULS or GASOHOL. Gallons
   are always positive. The section decides the sign. A note about a line
   goes in D, beside it.
6. If a section runs out of rows, click a line in it and press Add row.
7. Press Tidy up to take out the empty rows. The box says how many came
   out on each tab.
8. Read the light on QTR RECAP. Green, All checks OK, means every check
   passes. Red says how many checks to look at; the CHECKS block on
   QTR RECAP and the checks on each month tab show which.
9. EFT AMOUNT and Paying on Quarter return are on QTR RECAP, from K4 down.

## Updating the macro

Open the workbook, Alt F11, open its module under Modules, press Ctrl A,
and paste the new copy of PATRecap.bas over the old one.

## What to know

- Add row works from any line in a section, or from its running total.
  Click somewhere else and the box says to click a line in the section
  first. Nothing changes.
- Add row and Tidy up change only the line rows. Every running total,
  section total and check counts the rows that are there, so no total moves
  when Tidy up runs.
- Tidy up works on all three month tabs. It takes out a row when nothing
  is typed in D, G or H, and it keeps one empty row in a section with
  nothing typed.
- Each category rounds to whole dollars before the EFT AMOUNT adds them.
- The checks: AP in price against the In price lines, the rounding
  difference on the return (within $5), the gallons in the three buckets
  against the gallons flagged, every line inside its section, and that the
  quarter's rates are filled in. A product that isn't one of the four shows
  up as CHECK PRODUCT CODES.
- Gridlines are on, the way they come.
- If Excel blocks the macros, close the workbook, right-click it,
  Properties, tick Unblock, OK, and open it again.

## Open

- RATES has one row for each quarter and no year. The rates check can't
  tell last year's prices from this year's, so type the new prices each
  quarter.
