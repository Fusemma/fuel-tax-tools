# California rates check

What it does: an Excel macro (CARates.bas) that reads the CDTFA's
quarterly rates spreadsheet and shows what changed in the city and county
sales and use tax rates. It draws its own How to use tab. Tested on the
state's real files, 228 checks.

At work: paste CARates.bas into a new workbook, save as CA Rates Check
.xlsm, run CheckCaliforniaRates once; it makes the sheets and the button
and asks for the rates file. Load October's list first as the baseline.

Open: setup at work. The January 1, 2027 list should post December 1,
2026. Watch whether the work PC opens an old .xls from the macro; the
.xlsx is the one to use anyway.
