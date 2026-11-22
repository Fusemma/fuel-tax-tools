# Ohio PAT recap

## What it does

The quarterly Ohio PAT (Petroleum Activity Tax) recap in one workbook. A
tab for each month holds the lines, QTR RECAP adds the three months up
down to GL RESIDUAL and works out what you're paying on the quarter
return, RATES holds each quarter's rates and the product codes, and
CHECKS lists every check, one row each.

- The light at the top of QTR RECAP says All clear in green, or LOOK in
  red with how many checks to look at and what to do first. Nothing pops
  up. CHECKS shows the rest.
- What you type is yellow. Everything else is a formula.
- Two buttons come on the workbook: Add row puts one more empty row at the
  bottom of the section you're in, and Tidy up takes out the empty rows.
  The workbook works the same without the macro. Every total is a formula;
  the buttons only save time.
- A line's product comes from a list on RATES, each code with its return
  line: CLR DIESEL and DYED DIESEL (Highway Diesel and NonHighway Diesel)
  and GASOHOL (Highway Gas). The return line decides which bucket a code's
  DIV OWE and Additive owed gallons go in and its tax a gallon (gas for
  Highway Gas, diesel for the other two). A new code goes in the first
  empty yellow row, with its return line. The codes are text, so one like
  87E10 shows as typed. A product that isn't on the list turns the light
  red and says which row.
- AP in price, per AP only turns red when it is $1.00 or more off the In
  price lines, either way.
- Each category rounds to whole dollars before the EFT AMOUNT adds them.
- Take empty rows out with Tidy up, not by deleting whole rows.

## Once

Download OH PAT Recap.xlsm and keep it in a folder. If Excel blocks the
macros, close it, right-click the file, Properties, tick Unblock, OK, and
open it again. A new version comes as a new OH PAT Recap.xlsm; start the
next quarter in it.

## Each quarter

Open the workbook and follow the numbered steps on QTR RECAP.

## Open

- RATES has one row for each quarter and no year. The rates check can't
  tell last year's prices from this year's, so type the new prices each
  quarter.
- Gridlines are on, the way they come.
