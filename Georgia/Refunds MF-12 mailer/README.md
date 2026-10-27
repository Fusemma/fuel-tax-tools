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
name and address, the packets folder and the oldest period, once each.
Then it puts the HOW TO tab at the front, with the three buttons and
the rest of the steps.
