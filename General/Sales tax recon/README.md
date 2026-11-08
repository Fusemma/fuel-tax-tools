# Sales tax recon

What it does: an Excel macro (SalesTaxRecon.bas) that reconciles the
TERSE report against the AR sales report. Paste the two reports onto
their own tabs, run CompareReports, and it shows only what does not agree.
It finds columns by their header text and totals the lines per order
number before comparing.

Setup: in the workbook that holds the two reports, paste all of
SalesTaxRecon.bas into a new module (Alt F11, Insert, Module, Ctrl A,
Ctrl V) and run CompareReports.
