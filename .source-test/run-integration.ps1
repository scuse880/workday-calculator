$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$bookPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'actual-copy.xlsm')).Path
$sourceFolder = (Resolve-Path -LiteralPath (Join-Path $root 'testsource')).Path + [IO.Path]::DirectorySeparatorChar
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($bookPath, 0, $false)
    $macro = "'" + $book.Name + "'!TempSmoke.TestRealSourceLoaders"
    $excel.Run($macro, $sourceFolder)
    $evidence = $book.Worksheets.Item('TestEvidence')
    foreach ($row in 2..10) {
        $label = [string]$evidence.Cells.Item($row, 1).Value2
        $value = [string]$evidence.Cells.Item($row, 2).Value2
        Write-Output ($label + '=' + $value)
    }
    $book.Save()
}
finally {
    if ($null -ne $book) { $book.Close($false) | Out-Null }
    if ($null -ne $excel) { $excel.Quit() | Out-Null }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
