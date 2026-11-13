# Opens the built workbook in desktop Excel, recalculates every formula and saves,
# so the cached values can be checked against Python (tests/test_ipo_case.py).
# Requires Excel for Windows.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$path = Join-Path $root "output\Dhoot_IPO_Case.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($path)
    $excel.CalculateFull()
    $wb.Worksheets.Item("Valuation").Activate()
    $wb.Save()
    $wb.Close($false)
    Write-Output "recalculated $path"
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
