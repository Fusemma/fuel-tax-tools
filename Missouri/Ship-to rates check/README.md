# Missouri ship-to rates check

What it does: an Access database (MORatesCheck.bas) that loads the DOR's
quarterly boundary and total-rate files, keeps the last two quarters,
re-matches every ship-to against the boundary file, and shows where the
use tax rate changed, which ship-tos need fixing (one scrolling list; a
click on a ship-to's 100879/4 copies it), and a lookup by code. Each
quarter's state data goes in its own file on the PC's own drive (the
DataFolder setting, %LOCALAPPDATA%\MO Rates Check unless changed), so the
database stays a few MB and never compacts. mo_changes.py, mo_match.py
and mo_shiptos.py are the same logic in Python, for testing.

At work: in a new blank Access database, Alt F11, Insert, Module, paste
the whole of MORatesCheck.bas, save it as MORatesCode, run SetUp once
(mo_rates_notes.txt has the clicks; it is only the setup). Then
everything runs from the MO Rates Check form, and its How to use it
panel at the right is the whole guide, so nothing else comes along to
work: where to get the DOR's two files each quarter (DOR Boundary
Database and DOR Jurisdiction and Total Rate Database, under the quarter
in the Remote Seller and Marketplace Facilitator box on dor.mo.gov >
Taxation > Business > Sales/Use Tax, skipping DOR Rate Database), that
one pick of the .csv loads the .xlsx beside it, and every step after.
The panel's text is HowToText in the .bas; change it there.

How the state data sits: one file per quarter in the data folder, named
by its date (Quarter 2026-10-01.accdb), the two newest kept and an older
one deleted. A load builds Loading.accdb first, so a bad file changes
nothing. The Use tax changes answer lives in the newer quarter's file;
the database holds only the ship-tos, settings, forms and the code. The
quarter files are opened directly, not linked, because the matching
seeks on their indexes.

Tests: python test_match.py 5000, then python test_access.py (about ten
minutes; it loads test_hooks.bas, the test-only module, beside the
code). They start their own hidden Access, so a test copy left open on
the screen is fine. python make_test_copy.py builds the
desktop test copy and its data folder beside it.

Open: how the ship-to list comes. The okayed-pending defaults in
mo_shiptos.py's docstring.
