# Sales tax recon

What it does: an Excel macro (SalesTaxRecon.bas) that reconciles the
TERSE report against the AR sales report. Paste the two reports onto
their own tabs, run CompareReports, and it shows only what does not agree.
It finds columns by their header text and totals the lines per order
number before comparing.

Setup: bring SalesTaxRecon.bas into the workbook that holds the two
reports (Alt F11, File, Import File) and run CompareReports.
