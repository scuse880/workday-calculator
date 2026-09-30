#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess=$true)]
param([string]$WorkbookPath, [string]$SourceDir, [int]$CodePage=949, [switch]$ValidateOnly)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'vba-common.ps1')
if (-not $SourceDir) { $SourceDir = Join-Path (Split-Path $PSScriptRoot) 'vba' }
$target = Resolve-VbaWorkbook $WorkbookPath
$items = @(Get-VbaSources $SourceDir $CodePage)
if (-not $items.Count) { throw '가져올 bas/cls/frm 파일이 없습니다.' }
$items | ForEach-Object { Write-Host ('{0}: {1}' -f $_.File.Extension, $_.Name) }
Write-Host ('대상: {0} / {1}개' -f $target, $items.Count)
if ($ValidateOnly) { Write-Host '소스 검증 완료. Excel 실행 및 VBA 가져오기를 수행하지 않았습니다.'; return }
if (-not $PSCmdlet.ShouldProcess($target, 'VBA 모듈, 클래스, UserForm 가져오기')) { return }
Assert-VbaWorkbookClosed $target
$excel=$null; $book=$null; $project=$null; $stage=New-VbaStage
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=3
    $book = $excel.Workbooks.Open($target, 0, $false)
    Write-Host '대상 통합문서 열기 완료'
    if ($book.ReadOnly) { throw '통합문서가 읽기 전용입니다.' }
    $project = Get-VbaProject $book
    foreach ($item in $items) {
        foreach ($c in $project.VBComponents) {
            if ($c.Name -ieq $item.Name -and $c.Type -eq 100) { throw ('시트/통합문서 모듈 이름 충돌: ' + $item.Name) }
        }
    }
    $backupDir=Join-Path (Split-Path $target) 'backups'
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    $backup=Join-Path $backupDir (([IO.Path]::GetFileNameWithoutExtension($target)) + '.' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.xlsm')
    Copy-Item -LiteralPath $target -Destination $backup
    Write-Host ('백업 준비: '+$backup)
    $reusedForms=@{}
    foreach ($item in $items) {
        $old=$null
        foreach ($c in $project.VBComponents) { if ($c.Name -ieq $item.Name) { $old=$c; break } }
        if ($null -ne $old) {
            if ($item.Runtime -and $old.Type -eq 3) {
                $reusedForms[$item.Name]=$true
            } else {
                Write-Host ('기존 구성요소 제거: '+$item.Name)
                $project.VBComponents.Remove($old)
            }
        }
    }
    # UserForms first: establishes the Microsoft Forms reference before classes use MSForms.
    foreach ($item in $items | Where-Object { $_.Type -eq 3 }) {
        Write-Host ('폼 준비: '+$item.Name)
        if ($item.Runtime -and $reusedForms.ContainsKey($item.Name)) {
            $c=$project.VBComponents.Item($item.Name)
            $lines=$c.CodeModule.CountOfLines
            if ($lines -gt 0) { $c.CodeModule.DeleteLines(1,$lines) }
        }
        elseif ($item.Runtime) { $c=$project.VBComponents.Add(3); $c.Name=$item.Name }
        else { $c=Import-VbaNative $project $item $stage $CodePage }
        if ($c.Name -cne $item.Name) { throw ('폼 이름 오류: ' + $item.Name) }
    }
    foreach ($item in $items | Where-Object { $_.Type -ne 3 }) {
        Write-Host ('코드 가져오기: '+$item.Name)
        $c=Import-VbaNative $project $item $stage $CodePage
        if ($c.Name -cne $item.Name) { throw ('모듈 이름 오류: ' + $item.Name) }
    }
    foreach ($item in $items | Where-Object { $_.Runtime }) { $project.VBComponents.Item($item.Name).CodeModule.AddFromString($item.Code) }
    foreach ($item in $items) {
        $c=$project.VBComponents.Item($item.Name)
        $actualType=[int]$c.Type
        $lineCount=[int]$c.CodeModule.CountOfLines
        Write-Host ('검증: {0} type={1}/{2} lines={3}' -f $item.Name,$actualType,$item.Type,$lineCount)
        if ($actualType -ne $item.Type -or $lineCount -eq 0) { throw ('가져오기 검증 실패: ' + $item.Name) }
    }
    $book.Save()
    Write-Host ('가져오기 완료: ' + $target)
    Write-Host ('기존 파일 백업: ' + $backup)
    Write-Host 'VBA 편집기에서 [디버그 > VBAProject 컴파일] 후 사용하세요.'
} finally {
    Close-VbaSession $book $project $excel $stage
}
