"""Exercise selection-only work inputs with native mouse and keyboard events."""
from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import importlib.util
import json
from pathlib import Path
import shutil
import tempfile
import time

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("real_sources", Path(__file__).with_name("test-real-sources.py"))
sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sources)
user32 = ctypes.windll.user32
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
user32.SetWindowPos.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]


def focus(hwnd):
    current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
    foreground_thread = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
    user32.AttachThreadInput(current_thread, foreground_thread, True)
    try:
        if not user32.SetWindowPos(hwnd, wintypes.HWND(-1), 0, 0, 0, 0, 0x43):
            raise AssertionError("Could not position the disposable Excel window")
        if not user32.SetForegroundWindow(hwnd):
            raise AssertionError("Could not focus the disposable Excel window")
    finally:
        user32.AttachThreadInput(current_thread, foreground_thread, False)
    pump()


def dismiss_dialog(hwnd):
    buttons = []
    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum(child, unused):
        class_name = ctypes.create_unicode_buffer(100)
        user32.GetClassNameW(child, class_name, 100)
        if class_name.value == "Button":
            buttons.append(child)
        return True
    user32.EnumChildWindows(hwnd, enum, 0)
    if buttons:
        user32.PostMessageW(buttons[0], 0xF5, 0, 0)
        pump()
    else:
        focus(hwnd)
        key(0x1B)


def pump(seconds=0.25):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.02)


def key(code):
    user32.keybd_event(code, 0, 0, 0)
    user32.keybd_event(code, 0, 2, 0)
    pump()


def click(x, y):
    user32.SetCursorPos(x, y)
    user32.mouse_event(2, 0, 0, 0, 0)
    user32.mouse_event(4, 0, 0, 0, 0)
    pump()


def excel_dialogs(pid):
    result = []
    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum(hwnd, unused):
        task_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(task_pid))
        if task_pid.value == pid and user32.IsWindowVisible(hwnd):
            class_name = ctypes.create_unicode_buffer(100)
            user32.GetClassNameW(hwnd, class_name, 100)
            if class_name.value == "#32770" or class_name.value.startswith(("bosa_sdm", "NUIDialog")):
                result.append(hwnd)
        return True
    user32.EnumWindows(enum, 0)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    report = {"results": []}
    excel = book = None
    cursor_before = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(cursor_before))
    pythoncom.CoInitialize()
    with tempfile.TemporaryDirectory(prefix="workdays-dropdown-") as folder:
        stage = Path(folder)
        copy = stage / "dropdown.xlsm"
        shutil.copy2(ROOT / "workbook" / "근무일수계산.xlsm", copy)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 1
            book = excel.Workbooks.Open(str(copy), 0, False)
            sources.install_sources(book, ROOT / "vba", stage, adapt_dialogs=False)
            if not sources.compile_project(excel, book):
                raise AssertionError("production full compile")
            excel.Run("'dropdown.xlsm'!ApplyWorkbookProtection")
            # OnAction and protection do not depend on worksheet event dispatch.
            excel.EnableEvents = False
            excel.DisplayAlerts = False
            excel.Visible = True
            excel.WindowState = -4143
            excel.Left = 12
            excel.Top = 12
            excel.Width = min(780, user32.GetSystemMetrics(0) * 72 / 96 - 24)
            excel.Height = min(580, user32.GetSystemMetrics(1) * 72 / 96 - 36)
            work = book.Worksheets("작업")
            work.Activate()
            window = excel.ActiveWindow
            window.Zoom = 100
            window.ScrollRow = 1
            window.ScrollColumn = 1
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
            focus(excel.Hwnd)

            def check(label, condition):
                result = {"test": label, "passed": bool(condition)}
                report["results"].append(result)
                print(json.dumps(result), flush=True)
                if not condition:
                    raise AssertionError(label)

            for address in ("B3", "B5", "B6", "D3", "D5", "D6"):
                selector = work.Shapes("select_" + address)
                before = work.Range(address).Value2
                selected_index = selector.ControlFormat.ListIndex
                # Shapes use 72pt/in; this Excel window API takes the document's 96px/in coordinates.
                x = window.PointsToScreenPixelsX((selector.Left + selector.Width - 8) * 96 / 72)
                y = window.PointsToScreenPixelsY((selector.Top + selector.Height / 2) * 96 / 72)
                focus(excel.Hwnd)
                click(x, y)
                key(0x26 if selected_index == selector.ControlFormat.ListCount else 0x28)
                key(0x0D)
                pump()
                check("dropdown_changes_cell_" + address, work.Range(address).Value2 != before)
                expected = work.Range(address).Value2
                grid = work.Range("A2")
                focus(excel.Hwnd)
                click(window.PointsToScreenPixelsX((grid.Left + 20) * 96 / 72),
                      window.PointsToScreenPixelsY((grid.Top + grid.Height / 2) * 96 / 72))
                work.Range(address).Select()
                key(0x32)
                dialogs = excel_dialogs(pid.value)
                for dialog in dialogs:
                    dismiss_dialog(dialog)
                focus(excel.Hwnd)
                key(0x1B)
                check("typing_blocked_" + address, work.Range(address).Locked and work.Range(address).Value2 == expected)
        finally:
            if book is not None:
                book.Close(False)
            if excel is not None:
                excel.Quit()
            book = excel = None
            pythoncom.CoUninitialize()
            user32.SetCursorPos(cursor_before.x, cursor_before.y)
            if args.report:
                args.report.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"checks": len(report["results"]), "passed": all(x["passed"] for x in report["results"])}))


if __name__ == "__main__":
    main()
