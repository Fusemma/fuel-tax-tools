# Fuel tax tools

Excel macros, workbooks and Word templates for state fuel and sales tax
work. One folder per state. A tool for several states sits under
Streamlined states, and one for no state under General. Each tool's
folder has a README: what it does, how to set it up, what is open.

To get a file: click its name below, then the download button at the
top right of the file. A .bas file is a macro; open it, select all,
copy, and paste it into a module (Alt F11, Insert, Module).

| State | Tool | What it does | Download |
|---|---|---|---|
| Alabama | [Rates check](Alabama/Rates%20check/) | Shows which local rates changed going into this month and next, from the state's taxrates.csv. | [ALRates.bas](Alabama/Rates%20check/ALRates.bas) |
| California | [Rates check](California/Rates%20check/) | Shows what changed in city and county rates, from the CDTFA's quarterly rates file. | [CARates.bas](California/Rates%20check/CARates.bas) |
| Federal | [Forms templates](Federal/Forms%20templates/) | Word templates for the 637 notification certificate and the excise exemption certificate. | [Notification Certificate - TEMPLATE.docx](Federal/Forms%20templates/Notification%20Certificate%20-%20TEMPLATE.docx), [Exemption Certificate - TEMPLATE.docx](Federal/Forms%20templates/Exemption%20Certificate%20-%20TEMPLATE.docx) |
| Georgia | [Refunds MF-12 mailer](Georgia/Refunds%20MF-12%20mailer/) | Files the pulled invoices into packets, fills each supplier's MF-12 and makes one Outlook draft each. | [GARefunds.bas](Georgia/Refunds%20MF-12%20mailer/GARefunds.bas), [mf12_template.pdf](Georgia/Refunds%20MF-12%20mailer/mf12_template.pdf) |
| Georgia | [Sales tax workbook](Georgia/Sales%20tax%20workbook/) | The monthly sales tax return workbook; paste the reports in, key or upload the return. | [GA Sales Tax Workbook - BLANK.xlsx](Georgia/Sales%20tax%20workbook/GA%20Sales%20Tax%20Workbook%20-%20BLANK.xlsx) |
| Illinois | [Rates check](Illinois/Rates%20check/) | Shows which rates changed since the last machine-readable files you loaded. | [ILRates.bas](Illinois/Rates%20check/ILRates.bas) |
| Missouri | [Sales tax workbook](Missouri/Sales%20tax%20workbook/) | The quarterly use tax worksheet; paste the month's TERSE and AR reports in, and one light says whether the month checks out. | [MO Use Tax Workbook - BLANK.xlsx](Missouri/Sales%20tax%20workbook/MO%20Use%20Tax%20Workbook%20-%20BLANK.xlsx) |
| Missouri | [Ship-to rates check](Missouri/Ship-to%20rates%20check/) | An Access database that re-matches every ship-to against the DOR's boundary file each quarter. | [MORatesCheck.bas](Missouri/Ship-to%20rates%20check/MORatesCheck.bas) |
| Nebraska | [Streamlined rates check](Streamlined%20states/Rates%20check/) | Shows what changed in the state's rates since its last Streamlined file. | [SSTRates.bas](Streamlined%20states/Rates%20check/SSTRates.bas) |
| New Mexico | [QFR builder](New%20Mexico/QFR%20builder/) | Turns the non-excise purchase report into rows for NMED's quarterly fuel report upload. | [NMQFR.bas](New%20Mexico/QFR%20builder/NMQFR.bas) |
| New York | [Upload builder](New%20York/Upload%20builder/) | Turns PBT schedule exports into one Web File upload file. | [NYUpload.bas](New%20York/Upload%20builder/NYUpload.bas) |
| Ohio | [Streamlined rates check](Streamlined%20states/Rates%20check/) | Shows what changed in the state's rates since its last Streamlined file. | [SSTRates.bas](Streamlined%20states/Rates%20check/SSTRates.bas) |
| Oklahoma | [Streamlined rates check](Streamlined%20states/Rates%20check/) | Shows what changed in the state's rates since its last Streamlined file. | [SSTRates.bas](Streamlined%20states/Rates%20check/SSTRates.bas) |
| Washington | [Streamlined rates check](Streamlined%20states/Rates%20check/) | Shows what changed in the state's rates since its last Streamlined file. | [SSTRates.bas](Streamlined%20states/Rates%20check/SSTRates.bas) |
| General | [File reorg](General/File%20reorg/) | Writes a shared drive's folder tree to a text file, names only. | [ListShareFolders.bas](General/File%20reorg/ListShareFolders.bas) |
| General | [GL sort](General/GL%20sort/) | Sorts and totals the month-end AP and AR general ledger. | [GLSort.bas](General/GL%20sort/GLSort.bas) |
| General | [Power Automate cards](General/Power%20Automate%20cards/) | Step-by-step cards for two cloud flows: tax mail filing and Pennsylvania's distributor list. | [notes.txt](General/Power%20Automate%20cards/notes.txt), [pa-license-list.txt](General/Power%20Automate%20cards/pa-license-list.txt) |
| General | [Recap fill](General/Recap%20fill/) | Fills the EFT recap in Word from Excel and saves a dated copy beside it. | [Recap Fill.xlsm](General/Recap%20fill/Recap%20Fill.xlsm), [EFT recap.docx](General/Recap%20fill/EFT%20recap.docx), [RecapFill.bas](General/Recap%20fill/RecapFill.bas) |
| General | [Sales tax recon](General/Sales%20tax%20recon/) | Reconciles the TERSE report against the AR sales report and shows only what disagrees. | [SalesTaxRecon.bas](General/Sales%20tax%20recon/SalesTaxRecon.bas) |
| General | [State info](General/State%20info/) | One workbook for every state's returns, licenses, rates and contacts, one state per page. | [State Info.xlsx](General/State%20info/State%20Info.xlsx) |
| General | [Work tracker](General/Work%20tracker/) | One list of everything you track, with a dated log and pop-up reminders. | [Tracker.bas](General/Work%20tracker/Tracker.bas) |
