# Virginia ST-1 worksheet

The monthly Virginia ST-1 sales and use tax return from two pasted
reports. Paste the TERSE report and the AR sales tax by county report,
once each, as downloaded. The ST-1 tab fills in the return, ST-1R&B fills
in the ST-1R and ST-1B, and ENTRY lists what to type on Virginia Tax's
eForms site. The VA sales recap sits on the ST-1 tab, right of the form.
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
6. Type the recap's yellow cells on ST-1, columns H to J.
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
- **ST-1:** the return, the VA sales recap to its right, and the rates
  box under the recap. Every rate the formulas use is in the rates box;
  change one there and every formula follows.
- **ENTRY:** the ST-1A lines to type on the eForms site.
- **ST-1R&B:** one row for each locality on the site's Select Locality
  list, 133 rows, under the site's names. Each has its region marks (a
  yellow X under NVR, HRR, CVR, HTR or LOCAL), gross, exempt and taxable
  sales, the state's rate from its January 2026 rate file, and the rate
  the marks add up to. The regions table at the top right gives ST-1
  lines 10A to 10D and 12A.
- **CHECKS:** 11 checks, and the product types that don't count in gross
  sales.
- **TERSE** and **AR-SALES-TAX-BY-CNTY:** the paste tabs.

## The recap

| Line | Where it comes from |
|---|---|
| Account | type it |
| Collected | AR's Tax Paid |
| Paid Thru AP | type it |
| Paid on DSL Return | ST-1 line 18, or the site's total typed beside it when the two differ |
| Discount | ST-1 line 8 |
| Crossing Months | the crossing months' sales typed beside it, x 0.007 |
| AR In Price | type it |
| Rounding | the sum of the lines above |
| Order number Crossing months | type it |

The two cells beside Paid on DSL Return and Crossing Months carry a note
saying what goes there.

## What it checks

On the CHECKS tab. A difference of 0.01 or less is OK unless the line
says otherwise.

1. The columns the formulas read are where they were.
2. Every invoice date is in the month.
3. TERSE is all VA and AR is all Sales Tax.
4. Every TERSE product type is on the product type table.
5. Every sale landed on one of the site's localities.
6. Every INDEPENDENT CITIES line found its city.
7. Every taxable sale landed on one of the site's localities.
8. Dyed diesel on TERSE matches AR's 5.3% line, within 1.00.
9. The regions from the marks match what AR charged, one row for each
   ST-1 line.
10. AR charged no tax at a rate or name the sheet doesn't know.
11. Every locality's marks add up to the state's rate for it.

## What to know

- The product types that don't count in gross sales (F, S, O and Z) are
  marked No on CHECKS. Change Yes or No there. When check 4 asks, add the
  new type in the next empty row.
- A TERSE line with county INDEPENDENT CITIES counts under its Ship To
  City.
- If the state changes a locality's rate, change its marks on ST-1R&B and
  type the new rate in its State's rate cell.
- If the site's Select Locality list changes, ST-1R&B needs the same
  change.
- It uses INDEX and MATCH, no FILTER or SORT, so it works in any Excel.
- Gridlines are on, the way they come.

## Open

- The dealer discount uses 1.6, 1.2 and 0.8 percent of the first 3
  percent of the state tax, from Ruling of the Tax Commissioner 13-103
  (2013). Va. Code § 58.1-622 reads 4, 3 and 2 percent. Whether 1.6, 1.2
  and 0.8 still hold in 2026 isn't confirmed.
- Crossing Months is the sales x 0.007. Which 0.7% that is (Northern
  Virginia, Hampton Roads or Central Virginia) isn't written down yet, so
  it doesn't read the rates box.
