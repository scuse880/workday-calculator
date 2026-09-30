param([string]$WorkbookPath = (Join-Path $PSScriptRoot '근무일수계산-test.xlsm'))
$ErrorActionPreference = 'Stop'
$log = Join-Path $PSScriptRoot 'ui-test.log'
Remove-Item -LiteralPath $log -ErrorAction SilentlyContinue
function Log([string]$text) { Add-Content -LiteralPath $log -Value ((Get-Date -Format 'HH:mm:ss.fff') + ' ' + $text) -Encoding UTF8 }
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $excel.WindowState = -4137
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $false)
    $ws = $book.Worksheets.Item('공휴일')
    $ws.Range('A2:AF2').ClearContents()
    Log ('Excel HWND=' + $excel.Hwnd)
    Log 'FIRST_MODAL_STARTED'
    $excel.Run('main_공휴일.공휴일설정')
    Log ('FIRST_MODAL_RETURNED A2=' + [string]$ws.Range('A2').Text + ' B2=' + [string]$ws.Range('B2').Text + ' C2=' + [string]$ws.Range('C2').Text)
    $book.Save()
    Log 'SECOND_MODAL_STARTED'
    $excel.Run('main_공휴일.공휴일설정')
    Log ('SECOND_MODAL_RETURNED A2=' + [string]$ws.Range('A2').Text + ' B2=' + [string]$ws.Range('B2').Text + ' C2=' + [string]$ws.Range('C2').Text)
    $book.Save()
    Log 'COMPLETE'
} catch {
    Log ('ERROR ' + $_.Exception.ToString())
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    if ($null -ne $book) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) } catch {} }
    if ($null -ne $excel) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
