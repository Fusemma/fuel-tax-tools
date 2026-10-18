# Supplier license check

What it does: an Excel macro (SupplierLicenseCheck.bas) for the Texas
files where every seller not licensed as a supplier has to come out. It
reads the state's list of licensed fuels taxpayers and moves each line
whose seller holds no supplier license for that fuel, diesel or
gasoline, below the lines to keep, under a banner, for you to check and
delete. Red: licensed, but not as a supplier of that fuel. Orange: the
number isn't on the state's list. Grey: no number, or the supplier
license is suspended. Each line it moves gets the licenses the seller
does hold, in the column just past the data. A permissive supplier
counts as a supplier. Tested in Excel on the state's list of Oct 18,
2026 (6,531 licenses).

The state's list: on comptroller.texas.gov go to Taxes, Motor Fuels
IFTA and Interstate Truckers, then Diesel Fuel or Gasoline. Under
Additional Resources click Licensed Fuels Taxpayer Search, then the
Download button. It saves permitted-fuels-taxpayers.csv. The state
updates it nightly.

At work: paste SupplierLicenseCheck.bas into a module in the workbook
that holds the Texas file, click on that sheet, and run
CheckSupplierLicenses. It asks you to click a cell in the column with
the seller's Texas taxpayer number, then for the first row of data, D
or G for the fuel, and the list you downloaded. It works on a copy of
the sheet named with SUPPLIERS on the end, so the original is not
touched. On a file that also goes through the credit rebill matcher,
run this one first and delete what it moves down.

To look up a list by hand instead, no macro: in Excel, Data, From Web.
For the link, right-click the Download button on the search page and
pick Copy link. Load it, then name the table Fuels (Table Design, Table
Name box). Put your taxpayer numbers in column A of another sheet, and
in B2, filled down:

    =TEXTJOIN(", ",TRUE,FILTER(Fuels[Permit Type],Fuels[Taxpayer ID]&""=A2&"","NOT ON LIST"))

Each row shows every Texas license that number holds. To refresh by
itself: Data, Queries & Connections, right-click the query, Properties,
tick Refresh data when opening the file.

Open: not yet run on a real file. Which column on each of the two files
holds the seller's taxpayer number, and that it is the 11-digit Texas
number and not a federal ID (if not one number matches, the macro says
so).
