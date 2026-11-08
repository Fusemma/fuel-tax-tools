# Georgia refunds, the MF-12 mailer

An Excel macro for the GA tracking workbook that builds each supplier's
MF-12 packet from Refunds Due and makes an Outlook draft to send it.

## Putting the macros in

The first time:

1. Open the GA tracking workbook and press Alt F11.
2. Paste all of GARefundsReader.bas into a new module (Insert, Module,
   Ctrl A, Ctrl V).
3. Paste all of GARefundsWorker.bas into its own new module the same way.
4. Save the workbook as .xlsm. Save mf12_template.pdf in the same folder.

To update, paste the new copy over the old one the same way: open each
module under Modules, press Ctrl A, and paste the new GARefundsReader.bas
over the reader and the new GARefundsWorker.bas over the worker. The top
lines of each say which one it is.

## The first run

Press Alt F8 and run Refunds1_BuildPackets. It asks for the company's
name and address and the packets folder, once each, then which month to
work. Then it puts the HOW TO tab at the front, with the three buttons
and the rest of the steps.

It works one month at a time: every button asks which month, and only
that month's rows are touched. A row with MF-12 Sent filled in and no
packet from this macro counts as sent by hand and is left alone.
