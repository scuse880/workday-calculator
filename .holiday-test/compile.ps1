$ErrorActionPreference = 'Stop'
$p = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '근무일수계산-test.xlsm')).Path
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($p,0,$true)
    $vbe = $book.VBProject.VBE
    $project = $book.VBProject
    $project.VBComponents.Item('main_공휴일').Activate()
    Write-Output ("VBE object=" + [string]($null -ne $vbe) + " CommandBars=" + [string]($null -ne $vbe.CommandBars))
    $compile = $vbe.CommandBars.FindControl(1,578)
    Write-Output ('Project=' + $project.Name + ' CompileEnabledBefore=' + $compile.Enabled)
    $compile.Execute()
    Write-Output ('CompileEnabledAfter=' + $compile.Enabled)
    if ($compile.Enabled) { throw 'Compile command remains enabled; inspect VBE for a compile error.' }
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
