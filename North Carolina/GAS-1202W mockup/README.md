# North Carolina GAS-1202W mockup

North Carolina's suspension worksheets for the GAS-1202, October and
November 2026, filled in with made-up gallons. It shows where each
number comes from and how each worksheet's last line gets to GAS-1202
Line 19, on paper and in the EDI file.

No macro. Every number is a formula. Type over the yellow cells and the
rest fills itself, so you can also run your own month through it. Gray
is a box NC shades on its form. Each tab with math has a Check row at
the top that should say OK.

Left out on purpose: the Jet Fuel and Aviation Gasoline columns (we
don't sell them), and made-up numbers in Schedule 5F, dyed sold for
taxable purposes (we don't sell dyed for road use). Schedule 5F still
flows if you type there.

## Use it

1. Open NC GAS-1202W mockup.xlsx. Start on Your numbers.
2. The yellow cells are each schedule's gallons, split by Date Shipped:
   October 1 to 9, then October 10 to 31. The whole-month block under
   them is GAS-1202 page 2, and it fills itself.
3. GAS-1202 Oct fills GAS-1202 Lines 7 to 20 the normal way, full rate
   for the whole month.
4. GAS-1202W OCT is the worksheet. GAS-1202W OCT Line 28, in green,
   goes on GAS-1202 Line 19.
5. EDI (GAS-1202 Line 19) writes GAS-1202 Line 19 out by fuel, as the
   S03A lines for the EDI file.
6. GAS-1202W NOV is November's worksheet, with its own yellow cells.

Keep both worksheets for your records; NC says don't file them with the
return.

## The tabs

- Your numbers: the made-up gallons by schedule and date, GAS-1202
  page 2 for the whole month, the rates, and the color legend.
- GAS-1202 Oct: GAS-1202 Lines 7 to 20, with a box for each fuel on
  GAS-1202 Line 19. Its check: GAS-1202 Line 20 equals road tax on
  October 1 to 9, plus inspection tax for the month, minus the
  corrected discount.
- GAS-1202W OCT: Parts 1, 2 and 3. GAS-1202W OCT Lines 2 and 14 say
  which tax-paid loads go where by Date Shipped. Its check flags the
  last two open questions below.
- EDI (GAS-1202 Line 19): GAS-1202 Line 19 by fuel, and each fuel's TFS
  and TIA lines written out. Its check: the fuels add up to GAS-1202W
  OCT Line 28.
- GAS-1202W NOV: November's GAS-1202 lines and GAS-1202W NOV Lines 1
  to 4.
- Sources: NC's own words for each rate and rule, with the form, date
  and link; where NC's pages disagree; and the mockup's own choices.

## Open

- EDI (S03A): whether GAS-1202 Line 19 goes in with a minus sign. Ask
  the EDI Coordinator, EDICoordinator@ncdor.gov, (919) 707-7526.
- EDI (S03A): NC's layout is by fuel. What NC hasn't said is how the one
  GAS-1202W OCT Line 28 total splits across the fuels. The split in the
  mockup is its own math. Ask the EDI Coordinator.
- Whether a below-zero GAS-1202W OCT Part 1 stands: tax-paid receipts
  shipped Oct 1 to 9 that top Oct 1 to 9 sales. Ask the Excise Tax
  Division, 1-877-308-9092.
- Whether GAS-1202W OCT Line 27 may top $8,000. GAS-1202W OCT Lines 12
  and 24 each say "may not exceed $8,000", and GAS-1202W OCT Line 27
  adds them. Ask the Excise Tax Division.
