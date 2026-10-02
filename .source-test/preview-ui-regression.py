from pathlib import Path
import importlib.util
import json
import time
import win32com.client
import win32gui
from PIL import ImageGrab

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'.source-test/ui-regression'
spec=importlib.util.spec_from_file_location('ui',ROOT/'scripts/test-personnel-ui.py')
ui=importlib.util.module_from_spec(spec)
spec.loader.exec_module(ui)

def capture(hwnd,path):
    # The native capture plugin is unavailable in this session. Capture only
    # this disposable Excel window through Pillow's window-specific API.
    win32gui.ShowWindow(hwnd,5)
    win32gui.SetWindowPos(hwnd,-1,0,0,0,0,0x53)
    time.sleep(0.4)
    ImageGrab.grab(bbox=win32gui.GetWindowRect(hwnd)).save(path)
    win32gui.SetWindowPos(hwnd,-2,0,0,0,0,0x13)

snapshots={}
for label,path in [('before',ROOT/'workbook/근무일수계산.xlsm'),('after',OUT/'근무일수계산.xlsm')]:
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=True
    excel.DisplayAlerts=False
    excel.EnableEvents=False
    excel.AutomationSecurity=1
    book=None
    try:
        book=excel.Workbooks.Open(str(path),0,True)
        work=book.Worksheets('작업')
        work.Activate()
        excel.WindowState=-4137
        book.Windows(1).Zoom=100
        book.Windows(1).ScrollRow=1
        book.Windows(1).ScrollColumn=1
        work.Range('A2').Select()
        capture(excel.Hwnd,OUT/f'work-{label}.png')
        snapshot={'cells':{},'widths':{},'shapes':{},'modules':{}}
        for ws in book.Worksheets:
            snapshot['cells'][ws.Name]=ws.UsedRange.Formula
            snapshot['widths'][ws.Name]=[ws.Columns(i).ColumnWidth for i in range(1,ws.UsedRange.Columns.Count+1)]
        snapshot['shapes']={s.Name:[s.Left,s.Top,s.Width,s.Height,s.Placement,s.OnAction] for s in work.Shapes}
        snapshot['modules']={c.Name:c.CodeModule.Lines(1,c.CodeModule.CountOfLines) if c.CodeModule.CountOfLines else '' for c in book.VBProject.VBComponents}
        c=book.VBProject.VBComponents.Add(1)
        c.Name='UIPreview'
        c.CodeModule.AddFromString(ui.HARNESS)
        run=lambda name:excel.Run("'"+book.Name+"'!"+name)
        assert run('UISetup')=='ok'
        form=run('UIGetForm')
        ui.pump()
        snapshot['form']={'width':form.Width,'height':form.Height,'inside_height':form.InsideHeight,
                          'controls':{c.Name:[c.Left,c.Top,c.Width,c.Height] for c in form.Controls}}
        capture(ui.find_form(excel.Hwnd),OUT/f'personnel-{label}.png')
        run('UIClose')
        snapshots[label]=snapshot
    finally:
        if book is not None:book.Close(False)
        excel.Quit()

before,after=snapshots['before'],snapshots['after']
changed_cells={}
for name,rows in before['cells'].items():
    changes=[]
    new=after['cells'][name]
    for r in range(max(len(rows),len(new))):
        oldrow=rows[r] if r<len(rows) else ()
        newrow=new[r] if r<len(new) else ()
        for col in range(max(len(oldrow),len(newrow))):
            a=oldrow[col] if col<len(oldrow) else None
            b=newrow[col] if col<len(newrow) else None
            if a!=b:changes.append([r+1,col+1])
    if changes:changed_cells[name]=changes
changed_modules=[name for name,code in before['modules'].items() if code.lower()!=after['modules'][name].lower()]
report={'changed_cells':changed_cells,'column_widths_preserved':before['widths']==after['widths'],
        'changed_modules':changed_modules,'form_before':{k:v for k,v in before['form'].items() if k!='controls'},
        'form_after':{k:v for k,v in after['form'].items() if k!='controls'},
        'form_controls_exactly_preserved':before['form']['controls']==after['form']['controls'],
        'button_sizes_and_actions_preserved':all(a[0]==after['shapes'][n][0] and a[2:4]==after['shapes'][n][2:4] and a[5]==after['shapes'][n][5] for n,a in before['shapes'].items()),
        'dropdown_geometry_preserved':all(a==after['shapes'][n] for n,a in before['shapes'].items() if n.startswith('select_'))}
(OUT/'preservation-report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
