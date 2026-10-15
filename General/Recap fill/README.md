# EFT recap fill

What it does: fills the EFT recap in Word from Excel, so the copy you
hand over to be keyed in already has the values in it. In the Word file,
a tag like [[amount]] marks each spot a value goes, in the text, a
table, a header or a footer. Recap Fill.xlsm lists each tag on its
Fields sheet with its value: a number, text, or a formula pointing at
another sheet. The Fill button opens the one Word file in the
workbook's own folder with EFT in its name, swaps every tag for its
value, and saves a filled copy beside it with the date, like
EFT recap 2026-10-14.docx. The Word file itself never changes. Numbers
and dates go in the way their cell shows them, $1,234.56 or October 20,
2026. Then a box says what it filled and names any tag it couldn't
match. Tested in real Excel and Word, 128 checks.

Fill never reads or writes outside the workbook's own folder, and it
only works in a folder that holds the workbook itself. The box next to
Word file on Fields says EFT, so Fill looks there for a Word file with
EFT in its name, capitals or not, and no letter right before it. So
EFT recap and IN EFT recap count, but Left recap and Theft log don't.
It skips the dated copies it saved before and Word's ~$ temp files. If
no Word file there has EFT in its name, or more than one does, Fill
stops and lists them. If the only ones are dated copies, it says so and
names them. Type the exact name of the one you want next to Word file.
Fill turns away a name with a folder in it.

Try it first: put Recap Fill.xlsm and the sample EFT recap.docx in one
folder, open the workbook and press Fill. A copy with today's date shows
up beside the sample, three values filled.

At work:

1. Put Recap Fill.xlsm and the recap's Word file in one folder. Give the
   Word file a name with EFT in it, or type its exact name next to Word
   file on Fields.
2. In the Word file, type a tag wherever a value goes, like [[amount]]
   or [[period]], and save it.
3. On Fields, under Tag, list each tag without the brackets and its
   value beside it, formatted the way the recap should show it. Replace
   the sample's three rows.
4. Press Fill.

If Excel blocks the macros, close the workbook, right-click it,
Properties, tick Unblock, OK, and open it again. If the .xlsm won't
download at all, RecapFill.bas is the same code as text, and the steps
at its top make the same workbook from a blank one.

A second fill the same day saves EFT recap 2026-10-14 (2).docx and
keeps the first. Fill reads the saved Word file in a Word of its own
you never see, so anything open in Word stays as it is; save your
tagging before you press Fill.

Open: the real recap's tags and values, which go on Fields at work.
Whether one filled copy per payment does the job, or you want one
document for all of them. Whether macros run at work. A workbook in a
OneDrive folder is handled, but only tested on made-up folders; if Fill
says it can't find its folder, move both files to a folder on the PC.
