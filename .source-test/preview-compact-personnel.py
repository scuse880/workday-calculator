from pathlib import Path
import importlib.util
import json
import tempfile
import time
import win32com.client
import win32gui
import win32process
from PIL import ImageGrab
ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'.source-test/focus-fix'
spec=importlib.util.spec_from_file_location('ui',ROOT/'scripts/test-personnel-ui.py')
ui=importlib.util.module_from_spec(spec)
spec.loader.exec_module(ui)
excel=win32com.client.DispatchEx('Excel.Application')
excel.Visible=True
excel.DisplayAlerts=False
excel.EnableEvents=False
excel.AutomationSecurity=1
book=None
try:
    book=excel.Workbooks.Open(str(OUT/'근무일수계산.xlsm'),0,True)
    ui.install_accessors(book,Path(tempfile.mkdtemp(prefix='personnel-preview-')))
    run=lambda n:excel.Run("'"+book.Name+"'!"+n)
    assert run('UISetup')=='ok'
    dialog=run('UIGetForm')
    hwnd=ui.find_form(excel.Hwnd)
    ui.focus_form(hwnd)
    ui.pump()
    ImageGrab.grab(bbox=win32gui.GetWindowRect(hwnd)).save(OUT/'compact-personnel.png')
    employee=dialog.Controls('cmbEmployee')
    employee.SetFocus()
    employee.Value=''
    ui.pump()
    pid=win32process.GetWindowThreadProcessId(hwnd)[1]
    def windows():
        output=[]
        win32gui.EnumWindows(lambda h,p:output.append((h,win32gui.GetClassName(h),win32gui.GetWindowRect(h),win32gui.GetWindow(h,4))) if win32process.GetWindowThreadProcessId(h)[1]==pid and win32gui.IsWindowVisible(h) else None,None)
        return output
    print('before',windows(),flush=True)
    ImageGrab.grab(bbox=win32gui.GetWindowRect(hwnd)).save(OUT/'compact-dropdown.png')
    print('wheel_start',employee.TopIndex,flush=True)
    ui.wheel_at(dialog,hwnd,employee,y_offset=54,x_offset=employee.Width+24)
    print('wheel_end',employee.TopIndex,'hook',run('PersonnelMouseWheel.AuditWheelInstalled'),flush=True)
    print('after',windows(),flush=True)
    ImageGrab.grab(bbox=win32gui.GetWindowRect(hwnd)).save(OUT/'compact-dropdown-after-wheel.png')
    print('direct',dialog.HandlePersonnelWheel(employee.Left+employee.Width+24,employee.Top+54,-120,True),employee.TopIndex,flush=True)
    run('UIClose')
finally:
    if book is not None:book.Close(False)
    excel.Quit()
