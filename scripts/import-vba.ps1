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
    # 이벤트는 끄고, 가져온 로컬 코드를 명시적으로 초기화할 수 있게 연다.
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $book = $excel.Workbooks.Open($target, 0, $false)
    Write-Host '대상 통합문서 열기 완료'
    if ($book.ReadOnly) { throw '통합문서가 읽기 전용입니다.' }
    $project = Get-VbaProject $book
    $documentComponents=@{}
    foreach ($item in $items) {
        if ($item.Type -eq 100) {
            $codeName = if ($item.DocumentTarget -eq 'Workbook') { $book.CodeName } else { $book.Worksheets.Item($item.DocumentTarget).CodeName }
            $component = $project.VBComponents.Item($codeName)
            if ($component.Type -ne 100) { throw ('문서 모듈 대상 오류: ' + $item.Name) }
            $documentComponents[$item.Name]=$component
            continue
        }
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
        if ($item.Type -eq 100) { continue }
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
    foreach ($item in $items | Where-Object { $_.Type -ne 3 -and $_.Type -ne 100 }) {
        Write-Host ('코드 가져오기: '+$item.Name)
        $c=Import-VbaNative $project $item $stage $CodePage
        if ($c.Name -cne $item.Name) { throw ('모듈 이름 오류: ' + $item.Name) }
    }
    foreach ($item in $items | Where-Object { $_.Runtime }) { $project.VBComponents.Item($item.Name).CodeModule.AddFromString($item.Code) }
    foreach ($item in $items | Where-Object { $_.Type -eq 100 }) {
        $c=$documentComponents[$item.Name]
        $lines=$c.CodeModule.CountOfLines
        if ($lines -gt 0) { $c.CodeModule.DeleteLines(1,$lines) }
        $c.CodeModule.AddFromString($item.Code)
    }
    foreach ($item in $items) {
        $c=if ($item.Type -eq 100) { $documentComponents[$item.Name] } else { $project.VBComponents.Item($item.Name) }
        $actualType=[int]$c.Type
        $lineCount=[int]$c.CodeModule.CountOfLines
        Write-Host ('검증: {0} type={1}/{2} lines={3}' -f $item.Name,$actualType,$item.Type,$lineCount)
        if ($actualType -ne $item.Type -or $lineCount -eq 0) { throw ('가져오기 검증 실패: ' + $item.Name) }
    }
    $excel.VBE.ActiveVBProject=$project
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) { throw 'VBA 전체 컴파일 명령을 찾을 수 없습니다.' }
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) { throw 'VBA 전체 컴파일이 완료되지 않았습니다. 백업 파일은 유지했습니다.' }
    Write-Host 'VBA 전체 컴파일 완료'
    # 이벤트를 끄고 연 가져오기 세션에서도 보호와 선택 컨트롤을 저장한다.
    $excel.Run("'" + $book.Name.Replace("'", "''") + "'!InitializeWorkbookProtection")
    $book.Save()
    Write-Host ('가져오기 완료: ' + $target)
    Write-Host ('기존 파일 백업: ' + $backup)
    Write-Host '보호 및 선택 컨트롤 초기화 완료. 파일을 다시 열면 사용자 편집 보호가 복원됩니다.'
} finally {
    Close-VbaSession $book $project $excel $stage
}
