# New York sales tax workbook

What it does: the New York sales and use tax return for part-quarterly
(monthly) filers, NY Sales Tax Workbook - BLANK.xlsx, one workbook a
quarter. Each month takes five reports, pasted whole at A1 on that
month's tabs: AR Sales Taxes By County, the Terse Sales Report, AR
Excise Taxes run twice (the PPD one and the OTHER one), and the
Purchase Report For PPD Sales. The sheet sorts every purchase row into
its group by order number and fills in its % column. It fills each
month's ST-809 with Schedule FR, Local Sales By Jurisdiction below it,
the month's ST-810 with its Calculate step below it, and the month's
RECAP. At quarter end the QTR tabs add the three months into the
ST-810. Gray columns beside every report say which return line each
row feeds. One light on MAIN reads All clear, or names the first
problem and what to do. It turns red, instead of leaving lines out
without a word, when a report is on the wrong tab, is missing, runs
past the gray columns or has dates outside the month; when the gray
columns are gone after a whole-sheet paste; when a cell shows an error
like #N/A; when a product type, tax name, exemption, county or rate
isn't on LISTS; when an order with a minus line doesn't cancel out;
when the TERSE and AR tax disagree; and when the recap is $1.00 or
more off. Plain formulas, no macro, any Excel from 2010 on.

Setup: download NY Sales Tax Workbook - BLANK.xlsx and save a copy for
the quarter. Type the quarter's first day in C8 on MAIN, like 9/1/2026
for September to November. The how-to is on MAIN beside the light.

Each month: run the five reports for New York for the month and paste
each on its tab, the way MAIN says. On the month's RECAP, type
Collected A/R GL in E6, and if an order crosses months, its order
number in H18. Type the other yellow lines when they come up. When the
light is green, file from the month's two return tabs: ST-809, FR and
Local (the FR starts at row 25, the Local at row 67), and ST-810 and
Calculate (the Calculate starts at row 13).

Open:
- Rates are from Publication 718-F, effective September 1, 2026. When
  a rate changes, change it on LISTS.
- Jurisdiction codes not on LISTS yet: Chemung, Dutchess, Orange,
  Oswego, Putnam and Rockland on FR steps 4 and 4A, and Ogdensburg
  (city) on step 5. Type each in its Code cell on LISTS when you have
  it.
- When the light names something new, like a product type, tax name,
  exemption, or county spelling, add it on LISTS in the first empty
  row of that table.
- An order with both gas and diesel on one AR line can't be split
  between step 5's gas and diesel columns. The light names it.
- RECAP B3 holds a made-up account line. Type yours over it.
