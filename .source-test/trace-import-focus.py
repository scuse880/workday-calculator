from pathlib import Path
import shutil
import time
import json
import win32com.client
import pythoncom

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'.source-test/focus-ui'
OUT.mkdir(exist_ok=True)
target=OUT/'focus-ui.xlsm'
log=OUT/'focus-trace.txt'
log.write_text('',encoding='utf-8')
shutil.copy2(ROOT/'workbook/근무일수계산.xlsm',target)
excel=win32com.client.DispatchEx('Excel.Application')
excel.Visible=True
excel.DisplayAlerts=False
excel.EnableEvents=False
excel.AutomationSecurity=1
book=None
try:
    book=excel.Workbooks.Open(str(target),0,False)
    component=book.VBProject.VBComponents.Add(1)
    component.Name='UIFocusTrace'
    code='''Option Explicit
Public Sub UIDump(ByVal phase As String)
    Dim f As Integer, line As String
    On Error Resume Next
    line = phase & " | " & ActiveWorkbook.Name & " | " & ActiveSheet.Name
    line = line & " | " & TypeName(Selection) & " | " & ActiveCell.Address
    line = line & " | " & CStr(ActiveWindow.Zoom) & " | " & CStr(ActiveWindow.ScrollRow) & "," & CStr(ActiveWindow.ScrollColumn)
    line = line & " | events=" & CStr(Application.EnableEvents)
    f = FreeFile
    Open "LOGPATH" For Append As #f
    Print #f, line
    Close #f
End Sub
'''.replace('LOGPATH',str(log))
    component.CodeModule.AddFromString(code)
    module=book.VBProject.VBComponents('main_초과근무월별집계').CodeModule
    code=module.Lines(1,module.CountOfLines)
    code=code.replace('    ReadBaseSettings workMonth, startMinute, endMinute, False','    UIDump "MACRO_ENTRY"\r\n    ReadBaseSettings workMonth, startMinute, endMinute, False',1)
    code=code.replace('    selected = Application.GetOpenFilename','    UIDump "BEFORE_PICKER"\r\n    selected = Application.GetOpenFilename',1)
    code=code.replace('    If VarType(selected) = vbBoolean Then Exit Function','    UIDump "AFTER_PICKER"\r\n    If VarType(selected) = vbBoolean Then Exit Function',1)
    code=code.replace('    openedByCode = True','    UIDump "AFTER_OPEN"\r\n    openedByCode = True',1)
    code=code.replace('    Set SelectInputWorkbook = wb\r\n    Exit Function','    UIDump "AFTER_RESTORE"\r\n    Set SelectInputWorkbook = wb\r\n    Exit Function',1)
    code=code.replace('    wb.Close SaveChanges:=False','    UIDump "BEFORE_CLOSE"\r\n    wb.Close SaveChanges:=False\r\n    UIDump "AFTER_CLOSE"')
    code=code.replace('        Set dialog = New frmNeisPersonId','        UIDump "BEFORE_NEIS_NEW"\r\n        Set dialog = New frmNeisPersonId',1)
    code=code.replace('        dialog.Show vbModal','        UIDump "BEFORE_NEIS_SHOW"\r\n        dialog.Show vbModal\r\n        UIDump "AFTER_NEIS_SHOW"',1)
    code=code.replace('Cancelled:\r\n','Cancelled:\r\n    UIDump "CANCELLED"\r\n',1)
    code=code.replace('    Set gEmployeesByName = staged','    UIDump "COMMIT"\r\n    Set gEmployeesByName = staged',1)
    code=code.replace('vbInformation, "대상자 생성 완료"','vbInformation, "대상자 생성 완료"\r\n    UIDump "DONE"',1)
    module.DeleteLines(1,module.CountOfLines)
    module.AddFromString(code)
    module=book.VBProject.VBComponents('frmNeisPersonId').CodeModule
    module.AddFromString('Private Sub UserForm_Activate()\r\n    UIDump "NEIS_ACTIVATE"\r\nEnd Sub')
    module=book.VBProject.VBComponents(book.CodeName).CodeModule
    code=module.Lines(1,module.CountOfLines).replace('    ApplyWorkbookProtection','    UIDump "WB_ACTIVATE_ENTRY"\r\n    ApplyWorkbookProtection\r\n    UIDump "WB_ACTIVATE_EXIT"')
    module.DeleteLines(1,module.CountOfLines)
    module.AddFromString(code)
    excel.VBE.ActiveVBProject=book.VBProject
    compile=excel.VBE.CommandBars.FindControl(1,578)
    if compile.Enabled:compile.Execute()
    assert not compile.Enabled
    run=lambda name:excel.Run("'"+book.Name+"'!"+name)
    run('ApplyWorkbookProtection')
    work=book.Worksheets('작업')
    for a,v in [('B3','2026년'),('D3','08월'),('B5','8시'),('D5','30분'),('B6','16시'),('D6','30분')]:work.Range(a).Value2=v
    run('ApplyWorkbookProtection')
    book.Worksheets('작업결과').Activate()
    book.Worksheets('작업결과').Range('M105').Select()
    work.Activate()
    work.Range('B7').Select()
    excel.WindowState=-4137
    excel.ActiveWindow.Zoom=100
    excel.ActiveWindow.ScrollRow=1
    excel.ActiveWindow.ScrollColumn=1
    excel.EnableEvents=True
    print(json.dumps({'hwnd':excel.Hwnd,'ready':True,'target':str(target),'action':work.Shapes('btnEmployees').OnAction}),flush=True)
    # Root clicks the real worksheet button and handles the real native dialogs.
    started=time.monotonic()
    while time.monotonic()-started<240:
        pythoncom.PumpWaitingMessages()
        text=log.read_bytes().decode('cp949',errors='replace')
        if 'DONE |' in text or 'CANCELLED |' in text:
            time.sleep(1)
            print(text,flush=True)
            break
        time.sleep(.1)
finally:
    if book is not None:book.Close(False)
    excel.Quit()
