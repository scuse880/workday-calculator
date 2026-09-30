#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$WorkbookPath, [string]$OutDir, [int]$CodePage=949)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'vba-common.ps1')
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $PSScriptRoot) 'vba' }
$target=Resolve-VbaWorkbook $WorkbookPath
if (-not $PSCmdlet.ShouldProcess($OutDir, ($target + '의 bas/cls/frm 내보내기'))) { return }
Assert-VbaWorkbookClosed $target
$excel=$null; $book=$null; $project=$null; $stage=New-VbaStage
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($target,0,$true)
    $project=Get-VbaProject $book
    $count=0
    foreach ($c in $project.VBComponents) {
        $ext=switch ([int]$c.Type) { 1 {'.bas'} 2 {'.cls'} 3 {'.frm'} default {$null} }
        if ($null -eq $ext) { continue }
        $count++; $p=Join-Path $stage ($c.Name+$ext); $code=''
        if ($c.CodeModule.CountOfLines -gt 0) { $code=$c.CodeModule.Lines(1,$c.CodeModule.CountOfLines) }
        if ($c.Type -eq 3 -and $code -match '(?m)^\s*''\s*@RuntimeForm\s*$') {
            $header='VERSION 5.00' + "`r`n" + 'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ' + $c.Name + "`r`nEnd`r`n" + 'Attribute VB_Name = "' + $c.Name + '"' + "`r`n"
            [IO.File]::WriteAllText($p,$header+$code,(New-Object Text.UTF8Encoding($true)))
        } else {
            $c.Export($p)
            $text=[IO.File]::ReadAllText($p,(Get-VbaEncoding $CodePage))
            [IO.File]::WriteAllText($p,$text,(New-Object Text.UTF8Encoding($true)))
        }
        Write-Host ('내보내기 준비: '+$c.Name)
    }
    if ($count -eq 0) { throw '내보낼 구성요소가 없습니다. 기존 vba 폴더는 변경하지 않았습니다.' }
    $existing=@(Get-ChildItem -LiteralPath $OutDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.bas','.cls','.frm','.frx' })
    if ($existing.Count) {
        $backup=Join-Path (Split-Path $PSScriptRoot) ('backups/vba-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        New-Item -ItemType Directory -Force -Path $backup | Out-Null
        $existing | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $backup }
        Write-Host ('기존 소스 백업: '+$backup)
    }
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    Get-ChildItem -LiteralPath $stage -File | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $OutDir -Force }
    Write-Host ('내보내기 완료: {0}개. 통합문서는 저장하지 않았습니다.' -f $count)
    Write-Host '통합문서에 없는 기존 소스 파일은 삭제하지 않았습니다.'
} finally {
    Close-VbaSession $book $project $excel $stage
}
