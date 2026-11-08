# Missouri ship-to rates check

What it does: an Access database (MORatesCheck.bas) that loads the DOR's
quarterly boundary and total-rate files, keeps the last two quarters,
re-matches every ship-to against the boundary file, and shows where the
use tax rate changed, which ship-tos need fixing (one scrolling list; a
click on a ship-to's 100879/4 copies it), and the ones to look up on the
state's map. Each quarter's state data goes in its own file on the PC's
own drive (the DataFolder setting, %LOCALAPPDATA%\MO Rates Check unless
changed), so the database stays a few MB and never needs compacting.

Setup: in a new blank Access database, paste all of MORatesCheck.bas
into a new module (Alt F11, Insert, Module, Ctrl A, Ctrl V), save it as
MORatesCode, run SetUp once
(mo_rates_notes.txt has the clicks). Then everything runs from the MO
Rates Check form, and its How to use it panel at the right is the whole
guide: where to get the DOR's two files each quarter (DOR Boundary
Database and DOR Jurisdiction and Total Rate Database, under the quarter
in the Remote Seller and Marketplace Facilitator box on dor.mo.gov >
Taxation > Business > Sales/Use Tax, skipping DOR Rate Database), that
picking the .csv loads the .xlsx beside it, and every step after.
The panel's text is HowToText in the .bas; change it there.

Where the state data goes: one file per quarter in the data folder,
named by its date (Quarter 2026-10-01.accdb), the two newest kept and an
older one deleted. A load builds Loading.accdb first, so a bad file
changes nothing. The Use tax changes results are in the newer quarter's
file; the database holds only the ship-tos, settings, forms and the code.

Open: how the ship-to list comes isn't settled. Three rules aren't
confirmed yet: an active, non-exempt ship-to with Tax empty is flagged
as no rate set up; Tax that matches the sales rate is flagged as the
sales rate; the list's county is a cross-check, and a different county
goes to the map.
