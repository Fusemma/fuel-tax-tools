# NY upload builder

What it does: an Excel macro (NYUpload.bas) that turns New York PBT
schedule exports from Crystal into one Web File upload file. Pick one
export, or the A and B exports of a schedule together. It does the three
Notepad steps: each comma in the data becomes a space, 999999999 becomes
nine spaces when it fills a whole field, and each * becomes a comma.
Line breaks are cleaned up, the A file goes first, and nothing comes
after the last line. It saves "<name> upload.txt" next to the first file
you picked. The files you pick are never changed.

Nothing pops up. The CHECK sheet opens with one line: green when it's
good to go, red when something needs a look. Red lists only the lines to
look at, like a line with an extra field, or 999999999 inside a longer
number, which it leaves alone. Or it says why nothing was saved, like a
file that's open in another program; fix that and run it again. Every
line, and what changed in it, is folded away under the + on the left.

Setup: copy NYUpload.bas and follow "Setting it up, once" at the top:
in a new workbook, paste all of NYUpload.bas into a new module (Alt F11,
Insert, Module, Ctrl A, Ctrl V), then save it as NY Upload Builder.xlsm. After that, each time: Alt F8, BuildNYUpload, Run,
pick the files.

Defaults, each one line to change at the top of the macro: the dummy
FEINs it turns to spaces; what a comma becomes; the upload file is a
.txt, which Web File takes the same as a .csv and which opens in Notepad
instead of Excel; a file whose first line has no * is taken as already
clean, so its commas stay.

Also in this folder: New York upload rules.txt (what Web File takes,
with the state's own words) and Crystal Reports.txt (the versions and
the export box).
