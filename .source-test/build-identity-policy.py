from pathlib import Path
import json
import shutil
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / '.source-test/identity-policy'
OUT.mkdir(exist_ok=True)
original = ROOT / 'workbook/근무일수계산.xlsm'
baseline = OUT / 'before.xlsm'
target = OUT / '근무일수계산.xlsm'
live_copy = False
try:
    live = win32com.client.GetActiveObject('Excel.Application')
    for current in live.Workbooks:
        if Path(current.FullName) == original:
            current.SaveCopyAs(str(baseline))
            live_copy = True
            break
except Exception:
    if live_copy:
        raise
if not live_copy:
    shutil.copy2(original, baseline)
shutil.copy2(baseline, target)
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.AutomationSecurity = 1
book = None
try:
    book = excel.Workbooks.Open(str(target), 0, False)
    for filename in ['CEmployee.cls', 'main_초과근무월별집계.bas', 'main_근무상황목록.bas']:
        source = (ROOT / 'vba' / filename).read_text(encoding='utf-8-sig')
        module = book.VBProject.VBComponents(Path(filename).stem).CodeModule
        module.DeleteLines(1, module.CountOfLines)
        module.AddFromString(source[source.index('Option Explicit'):].replace('\n', '\r\n'))
    excel.VBE.ActiveVBProject = book.VBProject
    compile_control = excel.VBE.CommandBars.FindControl(1, 578)
    if compile_control.Enabled:
        compile_control.Execute()
    assert not compile_control.Enabled
    book.Save()
    print(json.dumps({'candidate': str(target), 'baseline_includes_unsaved_user_data': live_copy,
                      'compiled': True, 'changed_components': 3}, ensure_ascii=False))
finally:
    if book is not None:
        book.Close(False)
    excel.Quit()
