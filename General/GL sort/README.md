# GL sort

What it does: two Excel macros in GLSort.bas for the month-end general
ledger. SortGL pulls the AP rows whose reference says ORD, or whose AP
reference says BANK or MISC, below a gap in keyword groups, and sums the
block on top. ARGLTotals sums the AR GL. Safe to run twice.

At work: paste GLSort.bas into the workbook's module and run SortGL on
the AP sheet, ARGLTotals on the AR sheet. Live at work since September.

Open: a real-data run of the newest version. By design, the AP sheet's
totals only sum the top block.
