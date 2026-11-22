# Virginia ST-1 worksheet

The monthly Virginia ST-1 sales and use tax return from two pasted
reports. Paste the TERSE report and the AR sales tax by county report,
once each, as downloaded. The ST-1 tab fills in the return, ST-1R&B fills
in the ST-1R and ST-1B, and ENTRY lists what to type on Virginia Tax's
eForms site. The VA sales recap has its own tab, RECAP.
A status light at the top of ST-1 says whether every check is OK.

No macro. Every number is a formula.

## Each month

1. Open VA ST-1 Worksheet - BLANK.xlsx and save a copy for the month.
2. Download the Terse Sales Report: change Display As to download,
   Ship-to state VA.
3. Download A/R Sales Tax by County: change Display As to download,
   Ship-to state VA, Tax Class Sales Tax.
4. In each downloaded report, click the box above row 1 and left of
   column A, and press Ctrl C. In the worksheet, click the same box on the
   TERSE tab (the AR report: the AR-SALES-TAX-BY-CNTY tab) and press
   Ctrl V. Don't filter or delete anything, and don't delete rows or
   columns on TERSE or AR. Pasting over last month is fine.
5. Read the light at the top of ST-1. All checks OK means every check
   passes. LOOK says how many checks to look at: open CHECKS, and each
   LOOK line says what to do.
6. Type the recap's yellow cells on the RECAP tab. Collected is the AR
   GL amount.
7. On the eForms site, open ENTRY. For each line, pick the locality in
   Select Locality and type its Gross general sales and Exempt sales and
   deductions, the green cells. Only localities with sales show, in the
   site's order and under the site's names.

## Cell colors

One legend line under the light on ST-1.

- Light green: type it on the site.
- Light yellow: type it here.

## The tabs

- **INSTRUCTIONS:** these steps, and where each rate comes from, each
  source a link.
- **ST-1:** the return, and the rates box to its right. Every rate the
  formulas use is in the rates box; change one there and every formula
  follows.
- **RECAP:** the VA sales recap.
- **ENTRY:** the ST-1A lines to type on the eForms site.
- **ST-1R&B:** one row for each locality on the site's Select Locality
  list, 133 rows, under the site's names. Each has its region marks (a
  yellow X under NVR, HRR, CVR, HTR or LOCAL), gross, exempt and taxable
  sales, the state's rate from its January 2026 rate file, and the rate
  the marks add up to. The regions table at the top right gives ST-1
  lines 10A to 10D and 12A.
- **CHECKS:** 15 checks, the product type table (each type's Fuel and
  whether it counts in sales), and the list of orders check 12 finds off
  (row 131, folded away).
- **TERSE** and **AR-SALES-TAX-BY-CNTY:** the paste tabs.

## The recap

The month sits top left, then the title with the account number under
it, then the lines, with a subtotal after Paid Thru AP and after Paid on
DSL Return.

| Line | Where it comes from |
|---|---|
| Account number | type it, under the title |
| Collected | type it: the AR GL amount |
| Paid Thru AP | type it |
| (subtotal) | Collected plus Paid Thru AP |
| Paid on DSL Return | ST-1 line 18, or the site's total typed to its right when the two differ |
| (subtotal) | the subtotal above plus Paid on DSL Return |
| Discount | ST-1 line 8 |
| Crossing Months | the crossing months' sales typed beside it, x 0.007 |
| AR In Price | type it |
| Rounding | the second subtotal plus Discount, Crossing Months and AR In Price |
| Order number Crossing months | type it, to the right of the label |

The account number cell, the site's total cell and the crossing months
sales cell each carry a note saying what goes there.

## What it checks

On the CHECKS tab. A difference of 0.01 or less is OK unless the line
says otherwise.

1. The columns the formulas read are where they were.
2. Every invoice date is in the month.
3. TERSE is all VA and AR is all Sales Tax.
4. Every TERSE product type is on the product type table, once, with its
   Fuel and Yes or No under Counts in sales. The On LOOK words name the
   first type that isn't.
5. Every sale landed on one of the site's localities.
6. Every INDEPENDENT CITIES line found its city.
7. Every taxable sale landed on one of the site's localities.
8. Dyed diesel on TERSE matches AR's 5.3% line, within 1.00.
9. The regions from the marks match what AR charged, one row for each
   ST-1 line.
10. AR charged no tax at a rate or name the sheet doesn't know.
11. Every locality's marks add up to the state's rate for it.
12. Every dyed order is on both reports and taxed right for where it
    shipped. On LOOK, open the order list at row 131 (click the + at the
    left). Each line gives the order number and what is off on it.
13. Lines below the helper rows: the helper rows for check 12 stop at
    TERSE row 3001 and AR row 1501 and hold 600 orders. LOOK says which
    tab ran past, and its row.
14. No cell in either pasted report shows an error, like #N/A.
15. No TERSE line is on the tab twice. TERSE gives each line its own line
    ID (column AR, cnt Order Item ID), so a repeat means old lines from an
    earlier paste are still under the report. Paste the TERSE again by the
    box above row 1, which clears them. It can't catch a re-paste that
    lost only its last lines: the old lines left under the report are
    then exactly the lines it dropped, so no line repeats.

## Check 12, order by order

TERSE has one line per product on an order. AR has one line per tax on
an order, so an order in a region is on AR twice: its 5.3% line and its
0.7% line. Check 12 counts each order once, from its 5.3% line, and
checks every other AR line against where the order shipped.

- Orders match on Order Number, on both reports.
- An order's locality is AR's County, found on ST-1R&B. Its marks there
  say which lines it needs: NVR or HRR, the NVTC/PRTC & HRPDC 0.7% line;
  CVR, Central Virginia 0.7%; HTR, Historic Triangle 1%; LOCAL, a county
  1% line.
- Exempt orders need the same lines, marked exempt.
- What it can say about an order: Not on AR; Not on TERSE; Amounts
  differ (over 1.00, the same as check 8); No 5.3% line; Two 5.3% lines;
  Locality not on ST-1R&B; Missing or Extra (Northern Virginia or Hampton
  Roads 0.7%, Central Virginia 0.7%, Historic Triangle 1%, local 1%); A
  0.7% or 1% line's amount differs from the 5.3% line; Paid and exempt
  mixed; Tax doesn't match the rate; TERSE and AR tax differ. An order
  with more than one problem lists them all.
- It holds 3,000 TERSE lines, 1,500 AR lines and 600 orders. Past that,
  check 12 says LOOK: "More lines than the order check holds. Check
  orders by hand this month." Check 13 names the tab and row.

## What to know

- The product type table on CHECKS lists each type with its Fuel (Clear
  diesel, Dyed diesel or Gas) and Yes or No under Counts in sales. Only a
  type on the table with Yes and a Fuel counts in gross sales; a type that
  is not on the table stays out of the numbers, and check 4 names it.
  Dyed diesel in check 8 and check 12 means the types marked Dyed diesel
  there.
- A TERSE line with county INDEPENDENT CITIES counts under its Ship To
  City.
- If the state changes a locality's rate, change its marks on ST-1R&B and
  type the new rate in its State's rate cell.
- If the site's Select Locality list changes, ST-1R&B needs the same
  change.
- It uses INDEX and MATCH, no FILTER or SORT, so it works in any Excel.
- Gridlines are on, the way they come.

## Open

- A new product type: add it there in the first empty yellow row, then
  tell whoever keeps this workbook, so the next copy comes with it.
- A re-paste that lost only its last lines isn't caught (check 15).
