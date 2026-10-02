from pathlib import Path
import importlib.util
import json
import win32com.client
ROOT=Path(__file__).resolve().parent.parent
spec=importlib.util.spec_from_file_location('layout',ROOT/'scripts/repair-work-layout.py')
layout=importlib.util.module_from_spec(spec)
spec.loader.exec_module(layout)
excel=win32com.client.DispatchEx('Excel.Application')
excel.Visible=False
excel.DisplayAlerts=False
excel.EnableEvents=True
excel.AutomationSecurity=1
book=None
try:
    book=excel.Workbooks.Open(str(ROOT/'workbook/근무일수계산.xlsm'),0,True)
    work=book.Worksheets('작업')
    steps=layout.verify(work)
    assert all(all(x[k] for k in ('aligned','fits_row','moves_with_note','note_present')) for x in steps)
    assert all(work.Range(a).Value2=='선택' for a in ('B3','D3','B5','D5','B6','D6'))
    assert work.ProtectContents and work.ProtectionMode
    for filename in ['main_초과근무월별집계.bas','frmPersonnelChange.frm']:
        source=(ROOT/'vba'/filename).read_text(encoding='utf-8-sig')
        source=source[source.index('Option Explicit'):]
        module=book.VBProject.VBComponents(Path(filename).stem).CodeModule
        actual=module.Lines(1,module.CountOfLines)
        normalize=lambda code:code.replace('\r\n','\n').strip().lower()
        assert normalize(source)==normalize(actual),filename
    assert all('Audit' not in c.Name and 'Preview' not in c.Name for c in book.VBProject.VBComponents)
    print(json.dumps({'actual_saved_workbook_reopen':True,'workbook_open_layout_preserved':True,'protection_and_defaults_restored':True,'two_source_modules_match':True,'test_code_not_persisted':True}))
finally:
    if book is not None:book.Close(False)
    excel.Quit()
