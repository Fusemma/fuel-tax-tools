# Alabama rates check

What it does: an Excel macro (ALRates.bas) that reads the state's monthly
taxrates.csv, the big file with the whole history, and shows which of the
local rates you watch changed going into this month and into next month.
Tested on the state's real files, 188 checks.

At work: paste ALRates.bas into a new workbook (Alt F11, Insert, Module),
save as .xlsm, and follow the steps at the top of the macro. Each month,
download taxrates.csv from the DOR's local rates page and click Check new
rates file.

Open: setup at work, then load October's file once the state posts it.
Watch whether opening the workbook redraws the months by itself behind the
Enable Content bar.
