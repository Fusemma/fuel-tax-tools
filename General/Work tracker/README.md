# Work tracker

What it does: an Excel macro (Tracker.bas) that keeps everything you are
tracking at work in one list, with classic Outlook doing the popups. You add
a thing with one line of plain words, like "call Huntsville back 11/17 9am"
or "GA refund mailer monthly on the 5th", and it lands on the right day. A
flagged email comes in by itself, and Pull from Outlook brings in your
calendar's repeats. A thing with a set time pops up then; the rest come in
one list at 8:30 AM on workdays. How to use it is on the Today sheet, beside
the buttons. This copy replaces v2.

At work, once: new blank workbook, paste Tracker.bas into a new module, run
TrackerSetup (Alt F8), and save it as a macro workbook (.xlsm). Days off
comes filled in with the usual 8 for this year and next; change it to match
what the company gives. Then put it in Startup so it opens when you sign in:
in File Explorer, right-click the saved file and pick Copy, press the Windows
key and R, type shell:startup and press Enter, then right-click in that
folder and pick Paste shortcut (on Windows 11 it's under Show more options).
To update, paste a newer copy over the old one and run TrackerSetup again;
every row stays. If popups from the Tracker folder never come up, run
TrackerUseMainTasks (Alt F8) and they come from the main Tasks folder
instead.

Open, for the first day at work:
- Macros run.
- Add works (if not, IT has VBScript off).
- Outlook stays classic.
- Popups come up on top.
- Popups from the Tracker folder fire.
- Right-click shows Mark Complete.
- No security question.
