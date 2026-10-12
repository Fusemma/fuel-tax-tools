# Power Automate cards

What it does: the cards for the flows built by hand at work, one card
per flow. notes.txt is card 1, mail filing: a state or city tax mail
arrives, its attachments land in a dated OneDrive folder, the mail is
flagged, and one row goes into a log table. pa-license-list.txt is card
4: every morning Pennsylvania's distributor list lands in OneDrive as a
dated file, with a mail if it didn't. Nothing here needs a paid plan.

At work: open the card, fill in what it asks for, and follow its steps.
Both cards are cloud flows in the web designer (make.powerautomate.com).
Changes come from the first try.

What the web version can do: cloud flows move mail, files and rows
between Microsoft's own apps (Outlook, OneDrive, SharePoint, Teams,
Excel on the web, Forms, To Do, Planner) and run on a schedule, at no
cost past the work license. An outside connector (HTTP, SQL, Dataverse,
custom) makes the whole flow premium, a paid plan per user. Microsoft's
Power Automate licensing FAQ is the binding word; checked Oct 9, 2026.
A flow is bad at logic: past a few branches the work belongs in VBA.

Limits on the work PC: a cloud flow only sees files on OneDrive or
SharePoint, never a mapped drive. The Excel connector needs a real table
and locks the file while it writes. IT can switch a connector off; a
greyed-out one in the designer is that. IT blocks the Power Automate
desktop app, so nothing that clicks through a program or a portal is
possible.

Open: cards 1 and 4 are written and not yet built at work.
