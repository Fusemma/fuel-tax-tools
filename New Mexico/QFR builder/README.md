# New Mexico QFR builder

What it does: an Excel macro (NMQFR.bas) that turns the NON-EXCISE
purchase report for New Mexico into rows for NMED's quarterly fuel report
upload template (CTFP-ARCS, Liquids tab): an Upload sheet in the
template's column order and a CHECK sheet showing every report row and
what happened to it.

Setup: in a new workbook, paste all of NMQFR.bas into a new module
(Alt F11, Insert, Module, Ctrl A, Ctrl V), save as NM QFR Builder.xlsm,
and run BuildQFRRows once; the steps are at the top of the macro. Each quarter, paste the report onto
Report and run it, then paste Upload as values into the template
downloaded from CTFP-ARCS.

Defaults, each one line to change: NO2ULCNRT counted as clear diesel,
dyed loads left out, the business partner as the terminal name with
FEIN blank.
