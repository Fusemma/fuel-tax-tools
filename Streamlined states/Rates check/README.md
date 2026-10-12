# Streamlined rates check

What it does: one Excel macro (SSTRates.bas) for every Streamlined state;
you keep one workbook per state (Ohio, Washington, Nebraska, Oklahoma),
and each shows what changed in that state's sales tax rates since its last
file. build_names.py made the place-name table inside the macro from the
Census and state lists; the inputs are under state files/.

At work: paste SSTRates.bas into a new workbook, save it four times as
SST Rates Check OH, WA, NE and OK, and in each run CheckNewRatesFile. The
rate files come from streamlinedsalestax.org, Rates and Boundary Files.

Open: setup at work. Whether the Streamlined file list loads there (the
update page's link pointed at a bare address that timed out Oct 4; the
Rates page itself works). Washington lists each city change twice; that
is the file.
