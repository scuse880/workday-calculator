$ErrorActionPreference = 'Stop'
$path = (Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.xlsm' | Select-Object -First 1).FullName
Write-Host ('Path: [' + $path + ']')
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($path)
    $project = $book.VBProject
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'VBA compile command unavailable' }
    $compile.Execute()
    Write-Host ('Compile enabled after execute: ' + $compile.Enabled)
    if ($compile.Enabled) { throw 'VBA compile did not complete cleanly' }
    $smoke = @'
Option Explicit
Public Sub SmokePersonnelChange()
    Dim employeeA As CEmployee, employeeB As CEmployee
    Dim groupA As Collection, groupB As Collection, ordered As Collection
    Dim stage As Object, form As frmPersonnelChange
    Dim invalidNumber As Long
    Set gEmployeesByName = NewDictionary()
    Set groupA = New Collection
    Set groupB = New Collection
    Set employeeA = New CEmployee
    employeeA.Order = 1
    employeeA.Name = "Alpha"
    groupA.Add employeeA
    gEmployeesByName.Add employeeA.Name, groupA
    Set employeeB = New CEmployee
    employeeB.Order = 2
    employeeB.Name = "Beta"
    groupB.Add employeeB
    gEmployeesByName.Add employeeB.Name, groupB
    gEmployeesMonth = DateSerial(2026, 8, 1)
    Set ordered = GetEmployeesInOrder()
    Set form = New frmPersonnelChange
    form.Configure gEmployeesMonth, ordered
    form.SelectDay DateSerial(2026, 8, 3)
    form.SelectDay DateSerial(2026, 8, 5)
    form.Controls("cmdSave").Value = True
    If employeeA.PersonnelChangeStart <> DateSerial(2026, 8, 3) Then Err.Raise 5, , "first start not saved"
    If employeeA.PersonnelChangeEnd <> DateSerial(2026, 8, 5) Then Err.Raise 5, , "first end not saved"
    If Not IsPersonnelChangeDay(employeeA, DateSerial(2026, 8, 3)) Then Err.Raise 5, , "first day missing"
    If Not IsPersonnelChangeDay(employeeA, DateSerial(2026, 8, 5)) Then Err.Raise 5, , "last day missing"
    If IsPersonnelChangeDay(employeeA, DateSerial(2026, 8, 6)) Then Err.Raise 5, , "outside day included"
    On Error Resume Next
    SavePersonnelChange gEmployeesMonth, 1, DateSerial(2026, 8, 5), DateSerial(2026, 8, 3)
    invalidNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If invalidNumber = 0 Then Err.Raise 5, , "reverse range accepted"
    If employeeA.PersonnelChangeStart <> DateSerial(2026, 8, 3) Then Err.Raise 5, , "invalid save changed old range"
    form.SelectDay DateSerial(2026, 8, 7)
    If InStr(form.Controls("lblStatus").Caption, "2026-08-07") = 0 Then Err.Raise 5, , "third click did not start replacement"
    form.Controls("cmbEmployee").ListIndex = 1
    form.SelectDay DateSerial(2026, 8, 10)
    form.SelectDay DateSerial(2026, 8, 12)
    form.Controls("cmdSave").Value = True
    If employeeB.PersonnelChangeStart <> DateSerial(2026, 8, 10) Then Err.Raise 5, , "second employee start not saved"
    If employeeA.PersonnelChangeEnd <> DateSerial(2026, 8, 5) Then Err.Raise 5, , "first employee changed"
    form.Controls("cmdClear").Value = True
    If employeeB.PersonnelChangeStart <> 0 Then Err.Raise 5, , "clear failed"
    form.DisposeHandlers
    Unload form
    Set stage = NewRecordStage()
    CommitRecordStage stage, "Trip"
    Set ordered = GetEmployeesInOrder()
    Set employeeA = ordered(1)
    Set employeeB = ordered(2)
    If employeeA.PersonnelChangeStart <> DateSerial(2026, 8, 3) Then Err.Raise 5, , "range lost on reload"
    If employeeB.PersonnelChangeStart <> 0 Then Err.Raise 5, , "cleared range restored"
End Sub
'@
    $module = $project.VBComponents.Add(1)
    $module.Name = 'SmokePersonnel'
    $module.CodeModule.AddFromString($smoke)
    $compile.Execute()
    if ($compile.Enabled) { throw 'Smoke module compile did not complete cleanly' }
    $excel.Run("'" + $book.Name + "'!SmokePersonnelChange")
    Write-Host 'Personnel range and form-button smoke: PASS'
    $project.VBComponents.Remove($module)
    $book.Close($false)
    $book = $null
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
}
