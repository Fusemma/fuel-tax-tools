# Georgia sales tax workbook

What it does: the Georgia sales tax return workbook, GA Sales Tax
Workbook - BLANK.xlsx. Its OFF ROAD UPLOAD and ON ROAD UPLOAD tabs fill
the state's Excel import template, so a return can go in as a file
instead of being keyed; state import files/ holds the 3Q26 template.
recalc.ps1 recalculates a workbook in real Excel and strips author
details. test_upload.py checks each new quarter's template against a
filled workbook.

At work: download GA Sales Tax Workbook - BLANK.xlsx and paste the month's
AR sales, TERSE and recap reports onto their tabs. Shipped and in use.
To file by upload, download the quarter's template from
dor.georgia.gov/sales-use-tax-import, one copy per return, and follow the
upload tab: type the few Main cells it lists, paste the shaded rows as
values, check the Main totals it shows, then import.

Open: no upload has gone into GTC yet. The first real one, September
2026's return, is the test; compare GTC's return with RECAP before
submitting. Someday: the state takes tiered vendor's compensation, not
the 3% flat its template gives off-road motor fuel. A call to DOR
(1-877-423-6711) settles it; the workbook matches what the state does.
