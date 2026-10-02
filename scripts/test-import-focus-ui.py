"""Verify import focus through the real worksheet button, picker and NEIS form.

Uses a visible, disposable Excel process and two synthetic employees. Production
file selection, workbook open/close, form input and commit/cancel paths stay intact.
Only message boxes are captured; observation hooks are inserted in the test copy.
Requires Windows, Excel, pywin32 and openpyxl. Moves the mouse during execution.
"""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import gc
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import tempfile
import time

import openpyxl
import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("real_sources", Path(__file__).with_name("test-real-sources.py"))
sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sources)
user32 = ctypes.windll.user32
user32.SetWindowPos.argtypes = (wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, ctypes.c_int, wintypes.UINT)
user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
user32.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
user32.SendMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.PostMessageW.argtypes = (wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.GetForegroundWindow.restype = wintypes.HWND


class KeyboardInput(ctypes.Structure):
    _fields_ = [("vk", wintypes.WORD), ("scan", wintypes.WORD), ("flags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("extra", ctypes.c_size_t)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("data", wintypes.DWORD),
                ("flags", wintypes.DWORD), ("time", wintypes.DWORD), ("extra", ctypes.c_size_t)]


class InputUnion(ctypes.Union):
    _fields_ = [("keyboard", KeyboardInput), ("mouse", MouseInput)]


class Input(ctypes.Structure):
    _fields_ = [("kind", wintypes.DWORD), ("data", InputUnion)]


user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(Input), ctypes.c_int)
user32.SendInput.restype = wintypes.UINT


def pump(seconds=0.15):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.02)


def focus(hwnd):
    current = ctypes.windll.kernel32.GetCurrentThreadId()
    foreground = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
    attached = foreground != current and bool(user32.AttachThreadInput(current, foreground, True))
    try:
        user32.SetWindowPos(hwnd, wintypes.HWND(-1), 0, 0, 0, 0, 0x43)
        if not user32.SetForegroundWindow(hwnd) and user32.GetForegroundWindow() != hwnd:
            raise AssertionError("Could not focus the disposable Excel window")
    finally:
        if attached:
            user32.AttachThreadInput(current, foreground, False)
    pump()


def click(x, y):
    user32.SetCursorPos(int(x), int(y))
    user32.mouse_event(2, 0, 0, 0, 0)
    user32.mouse_event(4, 0, 0, 0, 0)
    pump()


def key(code):
    user32.keybd_event(code, 0, 0, 0)
    user32.keybd_event(code, 0, 2, 0)
    pump()


def type_text(value):
    # Unicode input works regardless of the user's Korean/English keyboard mode.
    events = []
    for character in value:
        for flags in (4, 6):
            item = Input()
            item.kind = 1
            item.data.keyboard = KeyboardInput(0, ord(character), flags, 0, 0)
            events.append(item)
    packed = (Input * len(events))(*events)
    if user32.SendInput(len(events), packed, ctypes.sizeof(Input)) != len(events):
        raise AssertionError("Could not type synthetic NEIS identifiers")
    pump()


def windows(pid, children_of=None):
    result = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum(hwnd, unused):
        process = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process))
        if process.value == pid and user32.IsWindowVisible(hwnd):
            class_name = ctypes.create_unicode_buffer(128)
            caption = ctypes.create_unicode_buffer(512)
            user32.GetClassNameW(hwnd, class_name, len(class_name))
            user32.GetWindowTextW(hwnd, caption, len(caption))
            result.append({"hwnd": hwnd, "class": class_name.value,
                           "caption": caption.value, "id": user32.GetDlgCtrlID(hwnd)})
        return True

    if children_of is None:
        user32.EnumWindows(enum, 0)
    else:
        user32.EnumChildWindows(children_of, enum, 0)
    return result


def wait_for(predicate, description, timeout=25):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = predicate()
        if result:
            return result
        pump(0.05)
    raise AssertionError("Timed out waiting for " + description)


def wait_window(pid, caption):
    return wait_for(lambda: next((entry for entry in windows(pid)
                                  if entry["caption"] == caption), None), caption)["hwnd"]


def select_file(pid, hwnd, path):
    children = windows(pid, hwnd)
    edit = next((entry["hwnd"] for entry in children
                 if entry["class"] == "Edit" and entry["id"] == 1148), None)
    button = next((entry["hwnd"] for entry in children
                   if entry["class"] == "Button" and entry["id"] == 1), None)
    if not edit or not button:
        raise AssertionError("Native file picker controls were not found: " + repr(children))
    value = ctypes.create_unicode_buffer(str(path))
    user32.SendMessageW(edit, 0x000C, 0, ctypes.cast(value, ctypes.c_void_p).value)
    user32.PostMessageW(button, 0x00F5, 0, 0)
    pump()


def read_log(path):
    return path.read_text(encoding="utf-16").splitlines() if path.exists() else []


def records(lines):
    result = []
    for line in lines:
        fields = line.split("|")
        if fields[0] == "FOCUS" and len(fields) == 11:
            result.append({"phase": fields[1], "workbook": fields[2], "window": fields[3],
                           "sheet": fields[4], "selection": fields[5], "active_cell": fields[6],
                           "scroll_row": fields[7], "scroll_column": fields[8], "zoom": fields[9],
                           "events": fields[10]})
    return result


def state(record):
    return {name: value for name, value in record.items() if name != "phase"}


def control_point(hwnd, lines, name):
    fields = next(line.split("|") for line in reversed(lines)
                  if line.startswith("CONTROL|" + name + "|"))
    inside_width, inside_height, left, top, width, height = map(float, fields[2:])
    rect = wintypes.RECT()
    origin = wintypes.POINT(0, 0)
    user32.GetClientRect(hwnd, ctypes.byref(rect))
    user32.ClientToScreen(hwnd, ctypes.byref(origin))
    return (origin.x + (left + width / 2) * rect.right / inside_width,
            origin.y + (top + height / 2) * rect.bottom / inside_height)


AUDIT_CODE = r'''
Option Explicit
Public Sub UIWrite(ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").OpenTextFile(LOG_PATH, 8, True, -1)
    stream.WriteLine value
    stream.Close
End Sub

Public Sub UIFocusDump(ByVal phase As String)
    Dim selected As String
    selected = TypeName(Selection)
    If TypeOf Selection Is Range Then selected = Selection.Address
    UIWrite "FOCUS|" & phase & "|" & ActiveWorkbook.FullName & "|" & CStr(ActiveWindow.Hwnd) _
          & "|" & ActiveSheet.Name & "|" & selected & "|" & ActiveCell.Address _
          & "|" & CStr(ActiveWindow.ScrollRow) & "|" & CStr(ActiveWindow.ScrollColumn) _
          & "|" & CStr(ActiveWindow.Zoom) & "|" & CStr(Application.EnableEvents)
End Sub

Public Sub UIFormGeometry(ByVal form As Object)
    Dim name As Variant, control As Object, offsetLeft As Double, offsetTop As Double
    For Each name In Array("txtPersonId1", "txtPersonId2", "cmdConfirm", "cmdCancel")
        If Left$(CStr(name), 3) = "txt" Then
            Set control = form.Controls("fraRows").Controls(CStr(name))
            offsetLeft = form.Controls("fraRows").Left
            offsetTop = form.Controls("fraRows").Top
        Else
            Set control = form.Controls(CStr(name))
            offsetLeft = 0: offsetTop = 0
        End If
        UIWrite "CONTROL|" & CStr(name) & "|" & CStr(form.InsideWidth) & "|" & CStr(form.InsideHeight) _
              & "|" & CStr(control.Left + offsetLeft) & "|" & CStr(control.Top + offsetTop) _
              & "|" & CStr(control.Width) & "|" & CStr(control.Height)
    Next name
End Sub

Public Function UIMessage(ByVal prompt As Variant, Optional ByVal buttons As Long = 0, _
                          Optional ByVal title As String = "", Optional ByVal helpfile As Variant, _
                          Optional ByVal context As Variant) As Long
    UIWrite "MESSAGE|" & title & "|" & Replace(Replace(CStr(prompt), vbCr, " "), vbLf, " ")
    UIMessage = vbOK
End Function

Public Function UIEmployeeSignature() As String
    Dim employee As CEmployee
    If gEmployeesByName Is Nothing Then Exit Function
    For Each employee In GetEmployeesInOrder()
        UIEmployeeSignature = UIEmployeeSignature & employee.NeisPersonId & ";"
    Next employee
End Function
'''


def replace_module(module, code):
    if module.CountOfLines:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(code)


def instrument(book, log):
    component = book.VBProject.VBComponents.Add(1)
    component.Name = "UIFocusAudit"
    component.CodeModule.AddFromString(AUDIT_CODE.replace("LOG_PATH", sources.vba_string(str(log))))
    module = book.VBProject.VBComponents("main_초과근무월별집계").CodeModule
    code = module.Lines(1, module.CountOfLines)
    match = re.search(r"Public Sub LoadOvertimeEmployees\(\).*?\r?\nEnd Sub", code, re.S)
    if match is None:
        raise AssertionError("Could not find production employee import")
    macro = match.group()
    macro = macro.replace("    On Error GoTo Failed", '    UIFocusDump "MACRO_ENTRY"\r\n    On Error GoTo Failed', 1)
    macro = macro.replace("    If sourceBook Is Nothing Then Exit Sub",
                          '    If sourceBook Is Nothing Then\r\n        UIFocusDump "COMPLETE"\r\n        Exit Sub\r\n    End If', 1)
    macro = macro.replace("    Exit Sub\r\n", '    UIFocusDump "COMPLETE"\r\n    Exit Sub\r\n')
    macro = macro.replace("\r\nEnd Sub", '\r\n    UIFocusDump "COMPLETE"\r\nEnd Sub')
    # Keep the real picker and entire NEIS form path; only notifications are captured.
    code = code[:match.start()] + macro + code[match.end():]
    code = re.sub(r"\bMsgBox\b", "UIMessage", code)
    replace_module(module, code)

    module = book.VBProject.VBComponents("frmNeisPersonId").CodeModule
    code = re.sub(r"\bMsgBox\b", "UIMessage", module.Lines(1, module.CountOfLines))
    hook = '    UIFocusDump "NEIS_ACTIVATE"\r\n    UIFormGeometry Me\r\n'
    if re.search(r"Private Sub UserForm_Activate\(\)", code):
        code = re.sub(r"(Private Sub UserForm_Activate\(\)\r?\n)", lambda m: m.group(1) + hook, code, count=1)
    else:
        code += "\r\nPrivate Sub UserForm_Activate()\r\n" + hook + "End Sub\r\n"
    replace_module(module, code)

    module = book.VBProject.VBComponents(book.CodeName).CodeModule
    code = module.Lines(1, module.CountOfLines)
    code = re.sub(r"(?m)^([ \t]*ApplyWorkbookProtection[^\r\n]*)\r?$",
                  lambda m: '    UIFocusDump "WB_ACTIVATE_ENTRY"\r\n' + m.group(1)
                  + '\r\n    UIFocusDump "WB_ACTIVATE_EXIT"', code)
    replace_module(module, code)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=ROOT / "workbook" / "근무일수계산.xlsm")
    parser.add_argument("--source-dir", type=Path, default=ROOT / "vba")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    report = {"method": "visible Excel, real worksheet button, real picker, real NEIS form",
              "synthetic_employees": 2, "results": [], "traces": {}}
    original_hash = hashlib.sha256(args.workbook.read_bytes()).hexdigest()
    excel = book = work = window = button = run = None
    pid = None
    cursor_before = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(cursor_before))
    pythoncom.CoInitialize()

    def check(name, passed, **details):
        item = {"test": name, "passed": bool(passed), **details}
        report["results"].append(item)
        print(json.dumps(item, ensure_ascii=True), flush=True)
        if not passed:
            raise AssertionError(name)

    scratch = tempfile.TemporaryDirectory(prefix="workdays-focus-ui-", ignore_cleanup_errors=True)
    with scratch as temporary:
        stage = Path(temporary).resolve()
        temporary_root = Path(tempfile.gettempdir()).resolve()
        if stage.parent != temporary_root or not stage.name.startswith("workdays-focus-ui-"):
            raise AssertionError("Disposable test directory is outside the expected temporary root")
        target = stage / "focus-ui.xlsm"
        log = stage / "trace.txt"
        fixture = stage / "synthetic-overtime.xlsx"
        shutil.copy2(args.workbook, target)
        synthetic = openpyxl.Workbook()
        synthetic.active["C5"] = "검증교직원(800101)"
        synthetic.active["C6"] = "검증교직원(900202)"
        synthetic.save(fixture)
        synthetic.close()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = True
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 1
            book = excel.Workbooks.Open(str(target), 0, False)
            sources.install_sources(book, args.source_dir, stage, adapt_dialogs=False)
            instrument(book, log)
            check("production_with_observation_hooks_compiles", sources.compile_project(excel, book))
            run = lambda name: excel.Run("'" + book.Name + "'!" + name)
            run("ApplyWorkbookProtection")
            work = book.Worksheets("작업")
            work.Range("B3").Value2 = "2026년"
            work.Range("D3").Value2 = "08월"
            run("ApplyWorkbookProtection")
            excel.WindowState = -4137
            process = wintypes.DWORD()
            user32.GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(process))
            pid = process.value

            for scenario in ("success", "cancel", "close_x"):
                excel.EnableEvents = False
                # A visibly different inactive view catches unexpected reactivation.
                book.Worksheets("작업결과").Activate()
                book.Worksheets("작업결과").Range("M105").Select()
                excel.ActiveWindow.Zoom = 25
                excel.ActiveWindow.ScrollRow = 2
                excel.ActiveWindow.ScrollColumn = 2
                work.Activate()
                work.Range("B7").Select()
                window = excel.ActiveWindow
                window.Zoom = 100
                window.ScrollRow = 1
                window.ScrollColumn = 1
                excel.EnableEvents = True
                focus(excel.Hwnd)
                log.write_text("", encoding="utf-16")
                excel.Run("'" + book.Name + "'!UIFocusDump", "BEFORE_CLICK")
                employee_before = run("UIEmployeeSignature")
                button = work.Shapes("btnEmployees")
                check(scenario + "_real_button_action", "LoadOvertimeEmployees" in button.OnAction,
                      action=button.OnAction)
                x = window.PointsToScreenPixelsX((button.Left + button.Width / 2) * 96 / 72)
                y = window.PointsToScreenPixelsY((button.Top + button.Height / 2) * 96 / 72)
                click(x, y)
                picker = wait_window(pid, "입력할 xlsx 파일 선택")
                select_file(pid, picker, fixture)
                neis = wait_window(pid, "나이스 개인번호 등록")
                wait_for(lambda: any(line.startswith("CONTROL|cmdCancel|") for line in read_log(log)),
                         "real NEIS form observation")
                lines = read_log(log)
                observations = records(lines)
                before = next(item for item in observations if item["phase"] == "BEFORE_CLICK")
                during = next(item for item in observations if item["phase"] == "NEIS_ACTIVATE")
                before_matches = (before["sheet"] == "작업" and before["active_cell"] == "$B$7"
                                  and before["zoom"] == "100" and before["scroll_row"] == "1"
                                  and before["scroll_column"] == "1" and before["events"] == "True")
                during_matches = state(during) == state(before)
                # Close the modal form even when the focus assertion will fail.
                focus(neis)
                if scenario == "success":
                    click(*control_point(neis, lines, "txtPersonId1"))
                    type_text("C900000001")
                    key(0x09)
                    type_text("C900000002")
                    click(*control_point(neis, lines, "cmdConfirm"))
                elif scenario == "cancel":
                    click(*control_point(neis, lines, "cmdCancel"))
                else:
                    user32.PostMessageW(neis, 0x0010, 0, 0)
                wait_for(lambda: any(item["phase"] == "COMPLETE" for item in records(read_log(log))),
                         scenario + " import completion")
                pump(0.3)
                excel.Run("'" + book.Name + "'!UIFocusDump", "AFTER_MACRO")
                lines = read_log(log)
                report["traces"][scenario] = lines
                after = next(item for item in reversed(records(lines)) if item["phase"] == "AFTER_MACRO")
                check(scenario + "_caller_is_work_sheet", before_matches, before=before)
                check(scenario + "_focus_while_neis_is_open", during_matches, before=before, during=during)
                check(scenario + "_focus_after_completion", state(after) == state(before), after=after)
                signature = run("UIEmployeeSignature")
                expected = "C900000001;C900000002;" if scenario == "success" else employee_before
                check(scenario + "_production_commit_or_cancel", signature == expected)
                check(scenario + "_source_closed", excel.Workbooks.Count == 1)
        except Exception as error:
            report["error"] = repr(error)
            if log.exists():
                report["last_trace"] = read_log(log)
            if pid is not None:
                report["visible_windows_on_failure"] = windows(pid)
            raise
        finally:
            # Only windows in this disposable Excel process are dismissed.
            if pid is not None:
                for entry in windows(pid):
                    if entry["class"] == "#32770" or entry["class"].startswith("ThunderD"):
                        user32.PostMessageW(entry["hwnd"], 0x0010, 0, 0)
                pump(0.3)
            # Child COM proxies keep Excel alive even after Quit. Release the
            # shape, sheet and window before closing their parent workbook.
            button = window = work = run = None
            gc.collect()
            cleanup_errors = []
            if excel is not None:
                try:
                    excel.EnableEvents = False
                except pythoncom.com_error as error:
                    cleanup_errors.append("disable_events: " + repr(error))
            if book is not None:
                try:
                    book.Close(False)
                except pythoncom.com_error as error:
                    cleanup_errors.append("close_workbook: " + repr(error))
            book = None
            gc.collect()
            pump(0.15)
            if excel is not None:
                try:
                    excel.Quit()
                except pythoncom.com_error as error:
                    cleanup_errors.append("quit_excel: " + repr(error))
            excel = None
            gc.collect()
            pump(0.5)
            pythoncom.CoUninitialize()
            user32.SetCursorPos(cursor_before.x, cursor_before.y)
            # GetOpenFilename can leave Excel holding the temporary directory
            # briefly. A cleanup delay must not replace substantive test results.
            # The absolute target was checked against the temporary root above.
            scratch.cleanup()
            report["scratch_cleanup"] = {"removed": not stage.exists()}
            if stage.exists():
                report["scratch_cleanup"]["retained_directory"] = str(stage)
            if cleanup_errors:
                report["scratch_cleanup"]["com_cleanup_errors"] = cleanup_errors
            report["original_workbook_unchanged"] = hashlib.sha256(args.workbook.read_bytes()).hexdigest() == original_hash
            if args.report:
                args.report.parent.mkdir(parents=True, exist_ok=True)
                args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"checks": len(report["results"]), "passed": all(item["passed"] for item in report["results"])}))


if __name__ == "__main__":
    main()
