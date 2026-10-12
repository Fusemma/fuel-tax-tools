"""Fills the state's ST-3 Excel import template from the DRAFT's upload tabs, the way it gets
done at work, and checks the template's Main tab against the numbers the upload tabs expect.

    python test_upload.py ["<template .xls>"]

With no argument it uses the newest template in "state import files". Needs Excel. Run it on
each new quarter's template: a moved column, a changed county code or a changed rate shows up
as a FAIL. The filled copies are saved to the temp folder for a look; nothing here is changed.
"""
import glob, os, shutil, sys, tempfile
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
DRAFT = os.path.join(HERE, "GA Sales Tax Workbook - DRAFT.xlsx")
VALUES = -4163  # xlPasteValues

if len(sys.argv) > 1:
    TEMPLATE = os.path.abspath(sys.argv[1])
else:
    TEMPLATE = max(glob.glob(os.path.join(HERE, "state import files", "ST-3 import template *.xls")), key=os.path.getmtime)

# Main tab cells: what gets typed, and what the upload tabs say it should show
TYPED = {"D10": "D10", "D11": "D11", "D16": "D12"}     # template cell: upload tab cell
CHECKS = {
    "OFF ROAD UPLOAD": [("Line 3", "D12", "D22"), ("Line 4", "D13", "D23"), ("Line 8", "D17", "D24"),
                        ("Line 9", "D18", "D25"), ("Line 10", "D19", "D26"), ("Total Tax Due", "D20", "D27")],
    "ON ROAD UPLOAD": [("Line 4", "D13", "D25"), ("Line 6", "D15", "D26"), ("Line 8", "D17", "D27"),
                       ("Line 9", "D18", "D28"), ("Line 10", "D19", "D29"), ("Total Tax Due", "D20", "D30")],
}
# (count cell, first row, last column, template tab) for each paste block
BLOCKS = {
    "OFF ROAD UPLOAD": [("D18", 30, "D", "Sales Tax Distribution")],
    "ON ROAD UPLOAD": [("D15", 18, "D", "Sales Tax Distribution"), ("D21", 34, "J", "MFD On-Road Worksheet")],
}


def column(ws, col, first, last):
    return [ws.Range(f"{col}{r}").Value for r in range(first, last + 1)]


def main():
    print("template:", TEMPLATE)
    out_dir = tempfile.mkdtemp(prefix="ga upload test ")
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3  # the template's macros stay off, as they would be at work
    fails = 0
    try:
        draft = xl.Workbooks.Open(DRAFT, 0, True)
        xl.CalculateFull()
        for tab in ("OFF ROAD UPLOAD", "ON ROAD UPLOAD"):
            up = draft.Worksheets(tab)
            path = os.path.join(out_dir, f"{tab.title()} test.xls")
            shutil.copy(TEMPLATE, path)
            tpl = xl.Workbooks.Open(path)
            main_ws = tpl.Worksheets("Main")
            period_end = main_ws.Range("O6").Value          # the template's own quarter end
            main_ws.Range("D5").Value = period_end
            main_ws.Range("D6").Value = "Monthly"
            main_ws.Range("K6").Value = "Yes"
            for cell, src in TYPED.items():
                main_ws.Range(cell).Value = up.Range(src).Value
            if tab == "OFF ROAD UPLOAD":
                tpl.Worksheets("PART B").Range("D6").Value = up.Range("D15").Value

            # every code pasted has to be one the template's own dropdown offers
            lists = tpl.Worksheets("Lists2")
            allowed = {"Sales Tax Distribution": set(column(lists, "F", 2, 810)) - {None},
                       "MFD On-Road Worksheet": set(column(lists, "S", 2, 300)) - {None}}
            for count_cell, first, last_col, target in BLOCKS[tab]:
                n = int(up.Range(count_cell).Value or 0)
                if not n:
                    continue
                bad = [c for c in column(up, "A", first, first + n - 1) if c not in allowed[target]]
                if bad:
                    fails += 1
                    print(f"FAIL {tab}: codes the template doesn't offer on {target}: {bad}")
                up.Range(f"A{first}:{last_col}{first + n - 1}").Copy()
                tpl.Worksheets(target).Range("A3").PasteSpecial(VALUES)
                xl.CutCopyMode = False
            xl.CalculateFull()

            print(f"\n{tab}, period end {period_end:%m/%d/%Y}")
            for label, tcell, ucell in CHECKS[tab]:
                got = round(main_ws.Range(tcell).Value or 0, 2)
                want = round(up.Range(ucell).Value or 0, 2)
                ok = abs(got - want) < 0.005
                fails += not ok
                print(f"  {'ok  ' if ok else 'FAIL'} {label:<14} template {got:>14,.2f}   upload tab {want:>14,.2f}")
            tpl.Save()
            tpl.Close(False)
        draft.Close(False)
    finally:
        xl.Quit()
    print(f"\nfilled copies: {out_dir}")
    print("ok" if not fails else f"{fails} FAIL")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
