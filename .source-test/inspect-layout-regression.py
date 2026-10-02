from pathlib import Path
import zipfile
import xml.etree.ElementTree as E
import json
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for name in ['근무일수계산.xlsm', 'backups/근무일수계산.20261001-155921.b0962a05.xlsm', 'backups/근무일수계산.before-sheet-layout.20261001-070539.xlsm']:
    with zipfile.ZipFile(ROOT/'workbook'/name) as z:
        strings = [''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
        sheet = E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        cells = {}
        for c in sheet.findall('.//s:c', NS):
            v = c.find('s:v', NS)
            if v is not None:
                value = strings[int(v.text)] if c.get('t')=='s' else v.text
                cells[c.get('r')] = value
        print(name, json.dumps(cells, ensure_ascii=False))
        print('rows', [(x.get('r'), x.get('ht')) for x in sheet.findall('.//s:row', NS)])
        for n in z.namelist():
            if n.startswith('xl/drawings/drawing') and n.endswith('.xml'):
                print(n, z.read(n).decode()[:22000])
                break

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.AutomationSecurity = 1
book = None
try:
    book=excel.Workbooks.Open(str(ROOT/'workbook/근무일수계산.xlsm'), 0, True)
    ws=book.Worksheets('작업')
    print('shapes', json.dumps([dict(name=s.Name, left=s.Left, top=s.Top, width=s.Width, height=s.Height, placement=s.Placement, action=s.OnAction) for s in ws.Shapes],ensure_ascii=False))
    ws.Activate()
    print('before protection', excel.ActiveSheet.Name)
    excel.Run("'"+book.Name+"'!ApplyWorkbookProtection")
    print('after protection', excel.ActiveSheet.Name)
    excel.EnableEvents=True
    other=excel.Workbooks.Open(str(ROOT/'testsource/근무상황목록.xlsx'),0,True)
    print('after input open',book.ActiveSheet.Name)
    other.Close(False)
    print('after input close',excel.ActiveSheet.Name)
finally:
    if book is not None: book.Close(False)
    excel.Quit()
