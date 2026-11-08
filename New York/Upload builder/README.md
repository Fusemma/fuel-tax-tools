# NY upload builder

What it does: an Excel macro (NYUpload.bas) that turns New York PBT
schedule exports from Crystal into one Web File upload file. Pick one
export, or the A and B exports of a schedule together. It does the three
Notepad steps: each comma in the data becomes a space, 999999999 becomes
nine spaces when it fills a whole field, and each * becomes a comma. If
999999999 is still in a line inside a longer value, the macro leaves it
alone, says so, and CHECK marks the line. Line breaks are cleaned up,
the A file goes first, and nothing comes after the last line. It saves "<name>
upload.txt" next to the first file you picked and lists every line, and
what changed in it, on a CHECK sheet. The files you pick are never
changed.

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
