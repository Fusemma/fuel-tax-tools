# North Carolina simple recap

One North Carolina month, October 2026, laid out plain. The recap's
STATE, INSPECTION and Recap Summary numbers come first, cleaned up.
Then GAS-1202 and GAS-1202W OCT are filled out. Every number has a note
beside it that says what it is and where to look: a cell on the recap,
one filter on one report, NC's rate, or math on numbers on the same tab.

There are two files, laid out the same:

- NC Simple Recap.xlsx has made-up numbers. They still tie, and each
  note says where a real month's number comes from.
- NC Simple Recap - October 2026.xlsx has the month's own numbers. Its
  government gallons are an estimate, since the October government
  sales report wasn't pulled, and every line they reach says ESTIMATE.

No macro, and nothing calculates: every number is typed. It shows how a
suspension month ties out, and why the recap and the return won't match.

## Use it

1. Open either file. Start on Recap.
2. Recap sections 1 to 3 are the recap's STATE and INSPECTION tabs,
   down to the EFT total.
3. Recap section 4 walks the recap's EFT total to GAS-1202 Line 20. It
   adds the tare, adds the October 10 to 31 road tax, takes it back off
   with GAS-1202 Line 19, and adds the inspection tax difference. The
   table under it shows the gallons that won't match because of the
   suspension.
4. Recap section 5 is the Recap Summary tab's GL accounts.
5. GAS-1202 and GAS-1202W OCT are NC's lines by fuel, gray where NC
   shades. GAS-1202W OCT Line 28 goes on GAS-1202 Line 19.
6. To check a report number:
   1. Open the report.
   2. Click its SUPPLIER NAME header cell and press Ctrl+Shift+End.
   3. Click Data > Filter, and set the filters the note names.
   4. Select the gallons column and read Sum at the bottom of Excel.

Keep the GAS-1202W OCT worksheet with the month's records. NC says
not to file it with the return.

## The tabs

- Recap: the recap's numbers in five sections, each with the cell it
  comes from.
- GAS-1202: page 1, Lines 7 to 23, and page 2, Lines 24 to 37. At the
  bottom is the box "Where the gallons come from": each fuel's gallons
  by B/L date, one report filter per number.
- GAS-1202W OCT: Parts 1, 2 and 3, with the same box at the bottom.

## Open

- GAS-1202 Line 19 in the EDI file: whether it goes in with a minus
  sign, and how one GAS-1202W OCT Line 28 total splits across the
  fuels. Ask the EDI Coordinator, EDICoordinator@ncdor.gov,
  (919) 707-7526.
- Whether a below-zero GAS-1202W OCT Part 1 stands. Ask the Excise Tax
  Division, 1-877-308-9092.
- Whether GAS-1202W OCT Line 27 may top $8,000. Ask the Excise Tax
  Division.
- Whether pipeline batches go on GAS-1202 Line 28 (Schedule 6X). Ask
  the Excise Tax Division.
- NC doesn't say how tare rounds. Both files round it to whole gallons.
- Both files split the government gallons between gas and undyed diesel
  by what each sold October 1 to 9.
