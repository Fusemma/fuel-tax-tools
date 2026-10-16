# EFT recap fill

Fills the four things that change on an EFT form each month: the tax
period, the amount, the draft date and the due date. You type them once in
Excel and press Fill. Fill makes a new copy of the EFT form with today's
date in its name and those four filled in. Your master EFT form never
changes, and the rest of it, the parts that never change, stay as you
typed them.

It doesn't matter what the state's recap looks like. You read the numbers
off the recap and type them into Excel. You never touch the macro.

## Once per state: set it up

Do this one time for each state. After that, skip to Each month.

1. Make a folder for the state, like Georgia EFT. Put a copy of
   Recap Fill.xlsm in it.
2. Put the state's EFT form, the Word file, in that same folder. It has to
   be a .docx or .doc. Its name needs EFT in it, like Georgia EFT.docx, and
   must not end in a date, since Fill skips those as its own old copies.
   Rename it if it needs it. The sample is called EFT recap.docx and
   stands in for your EFT form.
3. Open the Word file. Find the amount. Delete the whole amount, dollar
   sign and all, and type this in its place:

       [[amount]]

   That's two square brackets, the word, two square brackets. No spaces.
4. Do the same for the other three:

       [[period]]   where the tax period goes, like September 2026
       [[draft]]    where the draft date goes
       [[due]]      where the due date goes

   A tag works anywhere in the form: the text, a table, a header, a
   footer, a text box. The same tag can go in more than one spot and
   each one gets filled.

5. Save the Word file and close it.

That Word file is now the master. Don't type numbers into it again. If
something on it that never changes needs fixing, like a bank line, fix it
in the master and save.

## Each month

1. Open Recap Fill.xlsm in the state's folder.
2. On the Fields sheet, column B, type over the four values. Type them,
   don't paste. If you do paste from the recap, right-click and pick
   Paste Values, so the cell keeps its look:
   - period: the month the return is for. Type 9/1/2026 and it shows
     September 2026.
   - amount: the payment, off the state's recap. Type 66601.02 and it
     shows $66,601.02.
   - draft: the date the payment goes through. Type 10/17/2026.
   - due: the due date. Type 10/20/2026.
3. Press Fill.
4. The box lists what it filled. Check the four against the recap.
5. A new Word file is next to the master, with today's date in its name,
   like Georgia EFT 2026-10-14.docx. That's this month's EFT. Use that one.

If the box warns about anything, the copy it saved may still have a
[[tag]] in it. Delete that copy, fix the problem, and press Fill again.
The new copy ends in (2). Use the newest one.

Each value goes into Word the way its cell on Fields shows it. To make one
look different, like 10/17/2026 instead of October 17, 2026, or no dollar
sign, right-click the cell on Fields, Format Cells, and pick the look.

## Adding a fifth blank later

Say you want the invoice number filled too. In the master Word file, type
[[invoice]] where it goes and save. On Fields, in the first empty row
under the others, type invoice in column A and the value in column B.
That's it. Any word works as long as Word and Fields spell it the same.
Capitals don't matter.

A new row has no look of its own. If the new blank is a date or dollars,
click the cell above that already looks right (draft for a date, amount
for dollars), click Format Painter on the Home tab, then click your new
cell.

## If the box says something went wrong

**In the Word file but not on Fields.** Word has a tag that Fields
doesn't. Usually a spelling difference, like [[ammount]]. Make them match.

**Not in the Word file.** Fields has a row that Word doesn't. Same fix, or
delete the row on Fields if you don't need it.

**No Word file has EFT in its name.** The Word file isn't in the same
folder as the workbook, or EFT isn't in its name. Move it or rename it.
EFT right after a letter doesn't count, so Left recap and Theft log are
skipped.

**More than one Word file has EFT in its name.** Fill lists them. Next to
Word file at the top of Fields, type the exact name of the one to fill.

**Fill skipped every Word file, each one ends in a date.** Only the dated
copies are there, not the master. Put the master back in the folder.

**On Fields twice, the first one was used.** Two rows on Fields have the
same tag. Delete the one you don't want.

**Fill stopped.** Something outside Fill went wrong, usually Word. The
rest of the box says what. Close Word if it's stuck and press Fill again.

**Fix these on Fields first.** A value on Fields shows an error, like
#DIV/0!. Fix that cell and press Fill again.

**Excel blocks the macros.** Close the workbook, right-click it,
Properties, tick Unblock, OK, and open it again.

**Can't find this workbook's folder on the PC.** Usually OneDrive. Move
the workbook and the Word file to a folder on the PC itself.

## Try it first

Put Recap Fill.xlsm and the sample EFT recap.docx in one folder, open the
workbook and press Fill. A copy with today's date shows up beside the
sample, all four filled.

## What else to know

Fill only reads and writes in the workbook's own folder. It opens the
Word file in a hidden Word of its own, so whatever you have open in Word
stays as it is. Save the master before you press Fill.

If the .xlsm won't download, RecapFill.bas is the same macro as text. The
steps at its top make the same workbook from a blank one.

Tested in real Excel and Word, 132 checks.
