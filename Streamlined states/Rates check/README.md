# Streamlined rates check

What it does: one Excel macro (SSTRates.bas) for every Streamlined state;
you keep one workbook per state (Ohio, Washington, Nebraska, Oklahoma),
and each shows what changed in that state's sales tax rates since its last
file. The place names inside the macro come from the Census and state
lists.

Setup: in a new workbook, paste all of SSTRates.bas into a new module
(Alt F11, Insert, Module, Ctrl A, Ctrl V), save it four times as SST
Rates Check OH, WA, NE and OK, and in each run CheckNewRatesFile. The rate files come from
streamlinedsalestax.org, Rates and Boundary Files.

Open: Washington's file lists each city change twice, so What changed
shows it twice.
