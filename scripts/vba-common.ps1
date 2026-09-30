#requires -Version 5.1
function Resolve-VbaWorkbook([string]$Path) {
    if ($Path) {
        $p = (Resolve-Path -LiteralPath $Path).Path
        if ([IO.Path]::GetExtension($p) -ine '.xlsm') { throw '대상은 xlsm 파일이어야 합니다.' }
        return $p
    }
    $dir = Join-Path (Split-Path $PSScriptRoot) 'workbook'
    $p = Join-Path $dir '근무일수계산.xlsm'
    if (Test-Path -LiteralPath $p) { return $p }
    $files = @(Get-ChildItem -LiteralPath $dir -Filter '*.xlsm' -File | Where-Object { -not $_.Name.StartsWith('~$') })
    if ($files.Count -ne 1) { throw '대상 xlsm이 없거나 여러 개입니다. -WorkbookPath로 지정하세요.' }
    return $files[0].FullName
}
function Get-VbaEncoding([int]$CodePage) {
    return [Text.Encoding]::GetEncoding($CodePage, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
}
function Get-VbaSources([string]$Folder, [int]$CodePage) {
    $ansi = Get-VbaEncoding $CodePage
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    $seen = @{}
    foreach ($f in Get-ChildItem -LiteralPath $Folder -File | Where-Object { $_.Extension -in '.bas','.cls','.frm' } | Sort-Object Name) {
        $bytes = [IO.File]::ReadAllBytes($f.FullName)
        try { $text = $utf8.GetString($bytes).TrimStart([char]0xFEFF) } catch { $text = $ansi.GetString($bytes) }
        $m = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\s*$')
        if (-not $m.Success) { throw ('VB_Name 누락: ' + $f.Name) }
        $name = $m.Groups[1].Value
        if ($name -cne $f.BaseName -or $name.Length -gt 31 -or $name -notmatch '^[\p{L}][\p{L}\p{N}_]*$') { throw ('파일명/VB_Name 오류: ' + $f.Name) }
        if ($seen.ContainsKey($name)) { throw ('중복 이름: ' + $name) }; $seen[$name] = $true
        if ($text -notmatch '(?m)^Option Explicit\s*$') { throw ('Option Explicit 누락: ' + $f.Name) }
        $null = $ansi.GetBytes($text) # Fail before opening Excel if encoding would lose data.
        $type = switch ($f.Extension.ToLowerInvariant()) { '.bas' {1} '.cls' {2} '.frm' {3} }
        $runtime = $type -eq 3 -and $text -match '(?m)^\s*''\s*@RuntimeForm\s*$'
        $code = ''
        if ($runtime) { $code = [regex]::Replace($text.Substring($text.IndexOf('Option Explicit')), '(?m)^Attribute [^\r\n]+\r?\n?', '') }
        if ($type -eq 3 -and -not $runtime -and $text -match '"([^"\r\n]+\.frx)"') {
            if (-not (Test-Path -LiteralPath (Join-Path $f.DirectoryName $Matches[1]))) { throw ('FRX 파일 누락: ' + $Matches[1]) }
        }
        [pscustomobject]@{ Name=$name; Type=$type; File=$f; Text=$text; Code=$code; Runtime=$runtime }
    }
}
function Assert-VbaWorkbookClosed([string]$Path) {
    $h = $null
    try { $h = [IO.File]::Open($Path, 'Open', 'Read', 'None') }
    catch { throw ('Excel에서 대상 파일을 닫고 다시 실행하세요: ' + $Path) }
    finally { if ($h) { $h.Dispose() } }
}
function Get-VbaProject($Book) {
    try {
        $p = $Book.VBProject
        if ($null -eq $p -or $p.Protection -ne 0) { throw '프로젝트 잠금' }
        $null = $p.VBComponents.Count
        return $p
    } catch { throw 'VBA 프로젝트 접근 실패. 보안 센터의 [VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음] 및 프로젝트 잠금을 확인하세요.' }
}
function Import-VbaNative($Project, $Item, [string]$Stage, [int]$CodePage) {
    $p = Join-Path $Stage $Item.File.Name
    $windowsText = [regex]::Replace($Item.Text, '\r\n|\r|\n', "`r`n")
    [IO.File]::WriteAllText($p, $windowsText, (Get-VbaEncoding $CodePage))
    if ($Item.Type -eq 3) {
        foreach ($match in [regex]::Matches($Item.Text, '"([^"\r\n]+\.frx)"')) {
            $name = $match.Groups[1].Value
            if ([IO.Path]::GetFileName($name) -cne $name) { throw ('폼 바이너리 경로 오류: ' + $name) }
            $frx = Join-Path $Item.File.DirectoryName $name
            if (-not (Test-Path -LiteralPath $frx)) { throw ('폼 바이너리가 없습니다: ' + $frx) }
            Copy-Item -LiteralPath $frx -Destination $Stage -Force
        }
    }
    return $Project.VBComponents.Import($p)
}
function New-VbaStage {
    $p = Join-Path ([IO.Path]::GetTempPath()) ('workdays-vba-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $p | Out-Null
    return $p
}
function Remove-VbaStage([string]$Path) {
    if (-not $Path) { return }
    $p = [IO.Path]::GetFullPath($Path)
    $base = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($p) -ine $base -or [IO.Path]::GetFileName($p) -notmatch '^workdays-vba-[a-f0-9]{32}$') { throw ('임시 경로 검증 실패: ' + $p) }
    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
}
function Release-VbaCom($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) }
}
function Close-VbaSession($Book, $Project, $Excel, [string]$Stage) {
    try { if ($null -ne $Book) { $Book.Close($false) } } catch { Write-Warning ('통합문서 닫기 실패: ' + $_.Exception.Message) }
    try { Release-VbaCom $Project } catch { Write-Warning ('VBA COM 해제 실패: ' + $_.Exception.Message) }
    try { Release-VbaCom $Book } catch { Write-Warning ('통합문서 COM 해제 실패: ' + $_.Exception.Message) }
    try { if ($null -ne $Excel) { $Excel.Quit() } } catch { Write-Warning ('Excel 종료 실패: ' + $_.Exception.Message) }
    try { Release-VbaCom $Excel } catch { Write-Warning ('Excel COM 해제 실패: ' + $_.Exception.Message) }
    try { Remove-VbaStage $Stage } catch { Write-Warning ('임시 폴더 정리 실패: ' + $_.Exception.Message) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
