# Illinois rates check

What it does: an Excel macro (ILRates.bas) that reads the two machine
readable rate files the Department of Revenue posts (ordmache-current.txt
and ordmacha-current.txt) and shows which rates you watch changed since
the last files loaded: place, county, old rate, new rate, start date.
Tested on the state's real files, 302 checks.

At work: paste ILRates.bas into a new workbook, save as .xlsm, follow the
steps at the top of the macro; il_rates_notes.txt is the short card. Each
time the state posts new files, download the top two links and click Check
new rates file, picking both.

Open: setup at work. Ctrl click for two files in the open box is untried
there; one file at a time works the same.
