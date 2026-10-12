"""Tests State Info.xlsx in real Excel, on a throwaway copy.

Checks what the State page shows for Georgia, for an empty state, with
more rows than the page has room for, and that nothing shows an error.
Prints PASS or FAIL for each check and a count at the end.
"""
import os
import shutil

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "State Info.xlsx")
COPY = os.path.join(HERE, "test copy.xlsx")
results = []


def check(name, ok, got=None):
    results.append(ok)
    print(("PASS " if ok else "FAIL ") + name + ("" if ok or got is None else f"  (got {got!r})"))


def find_row(ws, starts):
    for r in range(1, 120):
        if str(ws.Range(f"B{r}").Text).startswith(starts):
            return r
    return None


def no_errors(ws, label):
    bad = [c.Address for c in ws.UsedRange if str(c.Text).startswith("#")]
    check(f"no error values on the State page ({label})", not bad, bad[:5])


shutil.copyfile(SRC, COPY)
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(COPY)
    page = wb.Worksheets("State")
    returns = wb.Worksheets("Returns")

    # Georgia, as the file opens
    check("opens on GA", page.Range("C3").Text == "GA", page.Range("C3").Text)
    check("shows the name Georgia", page.Range("D3").Text == "Georgia", page.Range("D3").Text)
    basics = {str(page.Range(f"B{r}").Text): page.Range(f"C{r}") for r in range(6, 15)}
    check("tax agency", basics["Tax agency"].Text == "Georgia Department of Revenue",
          basics["Tax agency"].Text)
    check("website shows as a link", basics["Website"].Text == "https://dor.georgia.gov"
          and "HYPERLINK" in basics["Website"].Formula, basics["Website"].Formula)
    check("main phone", basics["Main phone"].Text == "(877) 423-6711", basics["Main phone"].Text)
    check("bond required left blank", basics["Bond required"].Text == "", basics["Bond required"].Text)
    check("last checked is a date", basics["Last checked"].Text == "10/1/2026",
          basics["Last checked"].Text)

    r = find_row(page, "Returns (")
    check("returns header counts 4", page.Range(f"B{r}").Text == "Returns (4)", page.Range(f"B{r}").Text)
    check("first return", page.Range(f"B{r + 2}").Text == "Sales and use tax, on-road",
          page.Range(f"B{r + 2}").Text)
    check("fourth return", page.Range(f"B{r + 5}").Text == "Refund claim", page.Range(f"B{r + 5}").Text)
    check("fifth slot empty", page.Range(f"B{r + 6}").Text == "", page.Range(f"B{r + 6}").Text)
    check("blank form shows blank, not 0", page.Range(f"C{r + 4}").Text == "",
          page.Range(f"C{r + 4}").Text)
    check("no overflow line at 4", page.Range(f"B{r + 10}").Text == "", page.Range(f"B{r + 10}").Text)
    check("slot rows are two lines tall", page.Range(f"B{r + 2}").RowHeight == 27,
          page.Range(f"B{r + 2}").RowHeight)

    r = find_row(page, "Current rates (")
    check("current rates header counts 4", page.Range(f"B{r}").Text == "Current rates (4)",
          page.Range(f"B{r}").Text)
    check("gasoline excise rate", page.Range(f"D{r + 2}").Value == 0.333, page.Range(f"D{r + 2}").Value)
    check("suspended rate shows 0, not blank", page.Range(f"D{r + 4}").Text == "0",
          page.Range(f"D{r + 4}").Text)
    check("suspension start date", page.Range(f"F{r + 4}").Text == "9/29/2026",
          page.Range(f"F{r + 4}").Text)
    check("suspension end date", page.Range(f"G{r + 4}").Text == "10/29/2026",
          page.Range(f"G{r + 4}").Text)

    for name, n in [("Licenses", 1), ("Contacts", 1), ("Tax setups", 0)]:
        r = find_row(page, name + " (")
        check(f"{name.lower()} header counts {n}", page.Range(f"B{r}").Text == f"{name} ({n})",
              page.Range(f"B{r}").Text)
    r = find_row(page, "Contacts (")
    check("contact phone", page.Range(f"C{r + 2}").Text == "(877) 423-6711", page.Range(f"C{r + 2}").Text)
    no_errors(page, "GA")

    # An empty state
    page.Unprotect()
    page.Range("C3").Value = "AL"
    xl.Calculate()
    check("shows the name Alabama", page.Range("D3").Text == "Alabama", page.Range("D3").Text)
    check("basics blank for an empty state", all(page.Range(f"C{r}").Text == "" for r in range(6, 15)))
    for name in ["Returns", "Licenses", "Current rates", "Contacts", "Tax setups"]:
        r = find_row(page, name + " (")
        check(f"{name.lower()} counts 0 for AL", page.Range(f"B{r}").Text == f"{name} (0)",
              page.Range(f"B{r}").Text)
    no_errors(page, "AL")

    # More Georgia returns than the page has room for (8)
    first_empty = 6
    for i in range(12):
        returns.Range(f"A{first_empty + i}").Value = "GA"
        returns.Range(f"B{first_empty + i}").Value = f"Test return {i + 1}"
    page.Range("C3").Value = "GA"
    xl.Calculate()
    r = find_row(page, "Returns (")
    check("returns header counts 16", page.Range(f"B{r}").Text == "Returns (16)", page.Range(f"B{r}").Text)
    check("eighth slot filled", page.Range(f"B{r + 9}").Text == "Test return 4", page.Range(f"B{r + 9}").Text)
    check("overflow line says 8 more", page.Range(f"B{r + 10}").Text == "Plus 8 more on the Returns tab",
          page.Range(f"B{r + 10}").Text)
    no_errors(page, "GA with 16 returns")

    # An old rate (To date in the past) stays on the Rates tab, off the page
    rates = wb.Worksheets("Rates")
    rates.Range("A6").Value = "GA"
    rates.Range("B6").Value = "Old test rate"
    rates.Range("G6").Value = "1/1/2020"
    xl.Calculate()
    rr = find_row(page, "Current rates (")
    check("an ended rate drops off the page", page.Range(f"B{rr}").Text == "Current rates (4)",
          page.Range(f"B{rr}").Text)
    check("and isn't listed", all(page.Range(f"B{rr + 2 + k}").Text != "Old test rate" for k in range(8)))

    # A row for another state in the middle of the list stays out of Georgia's
    returns.Range("A4").Value = "IN"
    xl.Calculate()
    check("a row changed to IN drops out of GA", page.Range(f"B{r}").Text == "Returns (15)",
          page.Range(f"B{r}").Text)
    page.Range("C3").Value = "IN"
    xl.Calculate()
    check("and shows up under IN", page.Range(f"B{r + 2}").Text == "Motor fuel excise, distributor return",
          page.Range(f"B{r + 2}").Text)

    # The page itself
    page.Protect()
    check("the State page is locked", page.ProtectContents)
    check("the yellow box can still be typed in", not page.Range("C3").Locked)
    check("the yellow box is a state dropdown", page.Range("C3").Validation.Formula1 == "=StateList",
          page.Range("C3").Validation.Formula1)
    check("prints one page wide", page.PageSetup.FitToPagesWide == 1, page.PageSetup.FitToPagesWide)
    check("returns State column is a dropdown",
          returns.Range("A500").Validation.Formula1 == "=StateList")
    check("States lists all 50 plus DC",
          xl.WorksheetFunction.CountA(wb.Worksheets("States").Range("A2:A100")) == 51)
    wb.Close(False)
finally:
    xl.Quit()
    os.remove(COPY)

print(f"{sum(results)} of {len(results)} checks passed")
