# Opens a workbook in Excel, recalculates, strips author details, saves.
# Run after build_workbook.py so the saved file carries values.
# Usage: powershell -NoProfile -File recalc.ps1 "<full path to .xlsx>"
param([Parameter(Mandatory=$true)][string]$Path)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $xl.CalculateFull()
    $wb.RemovePersonalInformation = $true
    $wb.Save()
    $wb.Close($false)
    Write-Output "recalculated $Path"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
