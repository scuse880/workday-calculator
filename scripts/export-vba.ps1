$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$xlsm = Join-Path $root "workbook\근무일수계산.xlsm"
$outDir = Join-Path $root "vba"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($xlsm)
    $project = $workbook.VBProject

    foreach ($component in $project.VBComponents) {

        # 1 = Standard Module (.bas)
        if ($component.Type -eq 1) {

            $path = Join-Path $outDir ($component.Name + ".bas")

            if (Test-Path $path) {
                Remove-Item $path
            }

            $component.Export($path)

            Write-Host "Exported: $($component.Name)"
        }
    }

    $workbook.Close($false)
}
finally {
    if ($workbook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}