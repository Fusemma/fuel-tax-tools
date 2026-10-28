# Georgia refunds, the MF-12 mailer

An Excel macro for the GA tracking workbook that builds each supplier's
MF-12 packet from Refunds Due and makes an Outlook draft to send it.

## Putting the macros in

Do this the first time and each time the macro is updated.

1. Open the GA tracking workbook and press Alt F11.
2. Take out the old copy. Under Modules, right-click each one that holds
   this macro (GARefunds, GARefundsReader, GARefundsWorker, or a Module1
   it was pasted into), click Remove, then No. Two copies stop it working.
3. File, Import File, GARefundsReader.bas.
4. File, Import File, GARefundsWorker.bas.
5. Save the workbook as .xlsm. Save mf12_template.pdf in the same folder.

## The first one to run

Press Alt F8 and run Refunds1_BuildPackets. It asks for the company's
name and address and the packets folder, once each, then which month to
work. Then it puts the HOW TO tab at the front, with the three buttons
and the rest of the steps.

It works one month at a time: every button asks which month, and only
that month's rows are touched. A row with MF-12 Sent filled in and no
packet from this macro counts as sent by hand and is left alone.
