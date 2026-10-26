# Georgia refunds, the MF-12 mailer

What it does: an Excel macro (GARefunds.bas) that lives in the GA
tracking workbook. Each cycle it makes the month's folders, renames and
files the pulled invoices into packets, fills each supplier's MF-12 on the
state's own form (mf12_template.pdf), lays it out on CHECK, draws the
ENTRY sheet for the portal, and makes one Outlook draft per supplier. It
never sends.

At work: paste GARefunds.bas into the tracking workbook, save as .xlsm,
run Refunds1_BuildPackets once (it adds the HOW TO tab with the buttons
and the SETUP sheet), fill the yellow cells, and save mf12_template.pdf
where the workbook can reach it. The steps are at the top of the macro.

The headers it needs, any order, any capitals, the name alone in each
cell: Refunds Due, on row 4, Refund Type (or Schedule), Period, Corp,
License Number, BOL, INV, Inv Date, Gallons, Refund, Allowance, TTL
Refund, Fuel; Corp License Number, on row 1, Corp, License Number,
Contact name, Email address, Group Email Address, plus Short Name and
NET, which the first run adds for you to fill.

mf12_template.pdf is the state's 2017 MF-12 with room set aside in each
box for the macro to write in. If Georgia revises the form, the template
has to be rebuilt from the new blank before the macro can fill it.

Open: first run at work, drafts only. Part 2, a running-list version,
is not built yet.
