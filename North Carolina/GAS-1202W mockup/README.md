# North Carolina GAS-1202W mockup

North Carolina's suspension worksheets for the GAS-1202, October and
November 2026, filled in with made-up gallons. It shows where each
number comes from and how the worksheet's last line gets to GAS-1202
Line 19, on paper and in the EDI file.

No macro. Every number is a formula. Type over the yellow cells and the
rest fills itself, so you can also run your own month through it.

## Use it

1. Open NC GAS-1202W mockup.xlsx. Start on Your numbers.
2. The yellow cells are each schedule's gallons, split by Date Shipped:
   October 1 to 9, then October 10 to 31.
3. GAS-1202 Oct fills Lines 7 to 20 the normal way, full rate for the
   whole month. Its check at the bottom should say OK.
4. GAS-1202W OCT is the worksheet. Line 28, in green, goes on GAS-1202
   Line 19.
5. EDI Line 19 splits Line 28 by product, for the S03A lines in the EDI
   file.
6. GAS-1202W NOV is November's worksheet, with its own yellow cells.

Keep the worksheet for your records; NC says don't file it with the
return.

## The tabs

- Your numbers: the made-up gallons, the rates, and the color legend.
- GAS-1202 Oct: page 1, Lines 7 to 20, and a check that Line 20 equals
  road tax on October 1 to 9, plus inspection tax for the month, minus
  the corrected discount.
- GAS-1202W OCT: Parts 1, 2 and 3, each line with where its number
  comes from.
- EDI Line 19: Line 28 by product, and each product's TFS and TIA lines
  written out.
- GAS-1202W NOV: November's GAS-1202 lines and the worksheet's Lines 1
  to 4.
- Sources: NC's own words for each rate and rule, with the form and
  link.

## Open

- EDI Line 19 (S03A): NC's EDI guide doesn't say whether the amount goes
  in with a minus sign, or whether NC wants it split by product. Ask the
  EDI Coordinator, EDICoordinator@ncdor.gov, (919) 707-7526.
