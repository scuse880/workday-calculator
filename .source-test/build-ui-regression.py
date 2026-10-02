from pathlib import Path
import json
import re
import win32com.client

ROOT=Path(__file__).resolve().parent.parent
target=ROOT/'.source-test/ui-regression/근무일수계산.xlsm'
excel=win32com.client.DispatchEx('Excel.Application')
excel.Visible=False
excel.DisplayAlerts=False
excel.EnableEvents=False
excel.AutomationSecurity=1
book=None
try:
    book=excel.Workbooks.Open(str(target),0,False)
    for filename in ['main_초과근무월별집계.bas','frmPersonnelChange.frm']:
        source=(ROOT/'vba'/filename).read_text(encoding='utf-8-sig')
        component=book.VBProject.VBComponents(Path(filename).stem)
        component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
        code=source[source.index('Option Explicit'):]
        component.CodeModule.AddFromString(code.replace('\n','\r\n'))
    excel.VBE.ActiveVBProject=book.VBProject
    compile=excel.VBE.CommandBars.FindControl(1,578)
    if compile.Enabled: compile.Execute()
    assert not compile.Enabled
    book.Save()
    print(json.dumps({'compiled':True,'changed_components':2,'saved':str(target)},ensure_ascii=False))
finally:
    if book is not None: book.Close(False)
    excel.Quit()
