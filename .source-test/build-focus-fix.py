from pathlib import Path
import shutil
import json
import win32com.client

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'.source-test/focus-fix'
OUT.mkdir(exist_ok=True)
target=OUT/'근무일수계산.xlsm'
shutil.copy2(ROOT/'workbook/근무일수계산.xlsm',target)
excel=win32com.client.DispatchEx('Excel.Application')
excel.Visible=False
excel.DisplayAlerts=False
excel.EnableEvents=False
excel.AutomationSecurity=1
book=None
try:
    book=excel.Workbooks.Open(str(target),0,False)
    for filename in ['main_초과근무월별집계.bas','main_시트보호.bas','frmPersonnelChange.frm']:
        text=(ROOT/'vba'/filename).read_text(encoding='utf-8-sig')
        module=book.VBProject.VBComponents(Path(filename).stem).CodeModule
        module.DeleteLines(1,module.CountOfLines)
        module.AddFromString(text[text.index('Option Explicit'):].replace('\n','\r\n'))
    excel.VBE.ActiveVBProject=book.VBProject
    compile=excel.VBE.CommandBars.FindControl(1,578)
    if compile.Enabled:compile.Execute()
    assert not compile.Enabled
    book.Save()
    print(json.dumps({'saved':str(target),'compiled':True,'changed_components':3},ensure_ascii=False))
finally:
    if book is not None:book.Close(False)
    excel.Quit()
