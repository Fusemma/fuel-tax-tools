# GL sort

What it does: two Excel macros in GLSort.bas for the month-end general
ledger. SortGL pulls the AP rows whose document number says ORD, or whose
description says BANK or MISC, below a gap, one group per keyword, and
sums the block on top. ARGLTotals sums the AR GL. Safe to run twice.

Setup: in the workbook, paste all of GLSort.bas into a new module
(Alt F11, Insert, Module, Ctrl A, Ctrl V; to update, paste the new copy
over the old one the same way) and run SortGL on the AP sheet,
ARGLTotals on the AR sheet.
