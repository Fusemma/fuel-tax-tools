# Georgia refunds, the MF-12 mailer

What it does: an Excel macro in two pieces that lives in the GA
tracking workbook and works Refunds Due as the running list of every
refund still out, whatever its period. GARefundsReader.bas reads the
tabs; GARefundsWorker.bas does the folders, forms and drafts and has the
three buttons. Build packets makes a folder for each period, refund type
and supplier, files the pulled invoices into them from INBOX, and fills
each supplier's MF-12 on the state's own form (mf12_template.pdf). Make
drafts makes one Outlook draft per supplier per period; it never sends.
The letter's words are on SETUP: it says Hi and the contact's first
name (Hello when there's none), then gives each refund type's MF-12s and
invoices, one line each, or in the one sentence when there's one type.
Your Outlook signature stays under it. A SETUP still holding the last
version's letter gets the new one; a letter you changed is kept. Every
click stamps MF-12 Sent and MF-12 Signed on the rows. CHECK opens with
one bar, red with a count when something needs you and green when
nothing does, then each fix under Needs you, a line per email (where it
stands) and a line per packet (loads, MF-12s, invoices in); it adds up
the tax on each MF-12 and lists one that doesn't match its rows. ENTRY
is the typing sheet for the portal, one filing per period and refund
type, in the portal's columns with a Keyed column and a running total,
and Mark filed stamps Filed Date once it's keyed. Rows with a Refund Receive
Date move to Refunds Received; a row already there isn't copied twice.
A locked file, a file it can't read or a name too long for its box holds
only that packet or email, with a line on CHECK saying how to fix it;
the rest of the click goes on.

Names: a folder for each period like 2026-08 August, a folder for each
packet like Acme Licensed or Acme Government (the name comes from the
refund type list on SETUP), MF-12s like Acme Licensed MF-12 August 2026
(1 of 2).pdf, and filed invoices like Acme BOL 123456 Invoice 7890
VI.pdf. The Short Name is used as typed; the MF-12 prints the Corp name as
typed. Each packet folder holds a small packet.txt saying whose packet it
is (Short Name, license number, refund type, period, and the subject its
email went under). The macro finds the folder by it, so changing a Short
Name renames that supplier's open folders and files on the next click,
and an email sent under the old name still counts as sent. Filed packets
keep their names. The first click after an update renames the folders
and files the last version made (LIC-ACME and the like), and CHECK
lists every rename.

At work, the first time and each time the macro is updated: open the
tracking workbook and press Alt F11. First take out the old piece(s):
under Modules, right-click each one that holds this macro (GARefunds,
GARefundsReader, GARefundsWorker, or a Module1 it was pasted into),
click Remove, and No when it asks to export it. Import File never
replaces a piece with the same name; it adds a second copy, and two
copies stop the macro working. Then File, Import File, once for
GARefundsReader.bas and once for GARefundsWorker.bas. Save the workbook
as .xlsm, and save mf12_template.pdf in the same folder. Run
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
click adds. A blank Short Name gets the Corp name's first word, as typed. A paid
row moves under Refunds Received's own headers; a column that tab doesn't
have, like FUEL, stays behind.

mf12_template.pdf is the state's 2017 MF-12 with room set aside in each
box for the macro to write in. If Georgia revises the form, the template
has to be rebuilt from the new blank before the macro can fill it.

Open: first run at work, drafts only.
