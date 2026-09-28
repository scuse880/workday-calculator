$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$xlsm = Join-Path $root "workbook\근무일수계산.xlsm"
$sourceDir = Join-Path $root "vba"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($xlsm)
    $project = $workbook.VBProject

    foreach ($file in Get-ChildItem $sourceDir -Filter "*.bas") {

        $moduleName = $file.BaseName

        # 기존 모듈이 있으면 제거
        $existing = $null

        foreach ($component in $project.VBComponents) {
            if ($component.Name -eq $moduleName -and $component.Type -eq 1) {
                $existing = $component
                break
            }
        }

        if ($existing) {
            $project.VBComponents.Remove($existing)
            Write-Host "Removed: $moduleName"
        }

        # 새 파일 import
        $project.VBComponents.Import($file.FullName)

        Write-Host "Imported: $moduleName"
    }

    $workbook.Save()
    $workbook.Close($true)
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