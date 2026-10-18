# Credit rebill matcher

What it does: an Excel macro (CreditRebillMatcher.bas) for the Texas
file. Texas rejects negative lines, so it pulls every credit and rebill
that net to zero on the same BOL out of the order list and leaves the
lines to upload at the top. A credit with no matching rebill turns red;
that one may mean amending a prior return. A line with no BOL or no
gallons turns grey, for you to check by hand.

On each BOL it takes, in order: the whole BOL if it nets to zero; one
credit against one rebill of the same size; one credit against several
rebills, or one rebill against several credits; then whatever is left
if it nets to zero. Each matched group gets its own colour and moves
below the upload lines, under a banner, for you to check and delete.

At work: paste CreditRebillMatcher.bas into a module in the workbook
that holds the Texas file, click on that sheet, and run
MatchCreditsAndRebills. It asks you to click a cell in the BOL column,
then one in the gallons column, then for the first row of data. It
works on a copy of the sheet named with MATCHED on the end, so the
original is not touched.

Open: a BOL with more than 20 lines to try is left for you to check by
hand, and the message at the end can count it more than once.
