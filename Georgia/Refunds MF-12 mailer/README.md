# Georgia refunds, the MF-12 mailer

What it does: an Excel macro (GARefunds.bas) that lives in the GA
tracking workbook and works Refunds Due as the running list of every
refund still out, whatever its period. Build packets makes a folder for
each period, refund type and supplier, files the pulled invoices into
them from INBOX, and fills each supplier's MF-12 on the state's own form
(mf12_template.pdf). Make drafts makes one Outlook draft per supplier per
period; it never sends. Every click stamps MF-12 Sent and MF-12 Signed on
the rows. ENTRY is the typing sheet for the portal, one group per period
and refund type, and Mark filed stamps Filed Date once it's keyed. Rows
with a Refund Receive Date move to Refunds Received.

At work: paste GARefunds.bas into the tracking workbook, save it as
.xlsm, and save mf12_template.pdf in the same folder. Run
Refunds1_BuildPackets once (Alt F8). It asks for the company's name and
address, the folder the packets go in, and the oldest period to work,
once each, then puts the HOW TO tab with the three buttons at the front.
You don't make any folders.

Before that first click, any row on Refunds Due already filed by hand
needs its Filed Date, unless it's from before the oldest period you pick.
Otherwise it gets a packet again.

The headers it needs, any order, any capitals, the name alone in each
cell: on Refunds Due, Refund Type (or Schedule), Period, Corp, License
Number, BOL, INV, Inv Date, Gallons, Refund, Allowance, TTL Refund, Fuel,
plus MF-12 Sent, MF-12 Signed, Filed Date and Refund Receive Date, which
it adds if they're missing. On Corp License Number, Corp, License Number,
Contact name, Email address, Group Email Address, and Additional Email if
it's there (it goes on CC), plus Short Name and NET, which the first
click adds. A blank Short Name gets the Corp name's first word. A paid
row moves under Refunds Received's own headers; a column that tab doesn't
have, like FUEL, stays behind.

mf12_template.pdf is the state's 2017 MF-12 with room set aside in each
box for the macro to write in. If Georgia revises the form, the template
has to be rebuilt from the new blank before the macro can fill it.

Open: first run at work, drafts only.
