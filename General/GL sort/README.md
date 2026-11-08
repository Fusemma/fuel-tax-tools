# GL sort

What it does: two Excel macros in GLSort.bas for the month-end general
ledger. SortGL pulls the AP rows whose document number says ORD, or whose
description says BANK or MISC, below a gap, one group per keyword, and
sums the block on top. ARGLTotals sums the AR GL. Safe to run twice.

Setup: bring GLSort.bas into the workbook (Alt F11, File, Import
File; remove the old copy first) and run SortGL on the AP sheet,
ARGLTotals on the AR sheet.
