"""Verify the personnel form in a disposable Excel copy with 77 synthetic employees.

Runs production form code, full VBA compilation, filtering, duplicate identities,
save/clear, native mouse-wheel input, registered-row clicks, hook cleanup and layout.
Only test accessors are appended in the copy. Requires Windows, Excel and pywin32.
Use --layout-only for compilation, initial content and geometry checks without
native mouse input.
"""

import argparse
import ctypes
from ctypes import wintypes
import json
import importlib.util
from pathlib import Path
import re
import shutil
import tempfile
import time

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
source_spec = importlib.util.spec_from_file_location("real_sources", Path(__file__).with_name("test-real-sources.py"))
sources = importlib.util.module_from_spec(source_spec)
source_spec.loader.exec_module(sources)
user32 = ctypes.windll.user32


class MouseInput(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD), ("time", wintypes.DWORD), ("extra", ctypes.c_size_t),
    ]


class InputUnion(ctypes.Union):
    _fields_ = [("mouse", MouseInput)]


class Input(ctypes.Structure):
    _fields_ = [("kind", wintypes.DWORD), ("data", InputUnion)]


user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(Input), ctypes.c_int)
user32.SetWindowPos.argtypes = (
    wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
    ctypes.c_int, ctypes.c_int, wintypes.UINT,
)
user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
def pump(seconds=0.2):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.02)


def find_form(excel_hwnd):
    process = wintypes.DWORD()
    user32.GetWindowThreadProcessId(excel_hwnd, ctypes.byref(process))
    matches = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd, _):
        candidate = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(candidate))
        if candidate.value == process.value and user32.IsWindowVisible(hwnd):
            caption = ctypes.create_unicode_buffer(128)
            user32.GetWindowTextW(hwnd, caption, len(caption))
            if caption.value == "인사변동자 등록":
                matches.append(hwnd)
        return True

    user32.EnumWindows(callback, 0)
    if len(matches) != 1:
        raise AssertionError("Expected one personnel form in the test Excel process")
    return matches[0]


def control_point(dialog, hwnd, control, x_offset=None, y_offset=None):
    rect = wintypes.RECT()
    origin = wintypes.POINT(0, 0)
    user32.GetClientRect(hwnd, ctypes.byref(rect))
    user32.ClientToScreen(hwnd, ctypes.byref(origin))
    x_offset = control.Width / 2 if x_offset is None else x_offset
    y_offset = min(control.Height / 2, 12) if y_offset is None else y_offset
    return (
        int(origin.x + (control.Left + x_offset) * rect.right / dialog.InsideWidth),
        int(origin.y + (control.Top + y_offset) * rect.bottom / dialog.InsideHeight),
    )


def check_form_layout(dialog, check):
    # Freeze the calendar and vertical layout while allowing only the requested
    # widths, label reflow and horizontal action-button shift.
    expected = {
        "lblTitle": (17, 10, 500, 24),
        "lblEmployee": (17, 41, 72, 18),
        "cmbEmployee": (96, 39, 84, 22),
        "lblNotice": (17, 69, 500, 19),
        "lblInstructions": (17, 90, 500, 32),
        "lblMonth": (19, 130, 264, 19),
        "lblRegistered": (307, 130, 210, 20),
        "lstRegistered": (307, 154, 210, dialog.AuditRegisteredBaselineHeight()),
        "lblSelection": (307, 287, 210, 18),
        "lblStatus": (307, 307, 210, 35),
        "cmdSave": (235, 350, 90, 29),
        "cmdClear": (331, 350, 90, 29),
        "cmdClose": (427, 350, 90, 29),
    }
    for column in range(7):
        expected[f"lblWeek{column}"] = (19 + column * 38, 151, 36, 17)
    # August 2026 spans six calendar rows and exercises the lowest date buttons.
    for day in range(1, 32):
        row, column = divmod(6 + day - 1, 7)
        expected[f"day{day}"] = (19 + column * 38, 173 + row * 29, 36, 27)
    actual = {}
    clipped = []
    for control in dialog.Controls:
        left, top, width, height = [
            round(float(getattr(control, prop)), 2)
            for prop in ("Left", "Top", "Width", "Height")
        ]
        actual[control.Name] = (left, top, width, height)
        if (left < 0 or top < 0 or left + width > dialog.InsideWidth + 0.1
                or top + height > dialog.InsideHeight + 0.1):
            clipped.append(control.Name)
    check("form_width_compacted", dialog.Width, 550)
    # MSForms can quantize geometry when controls are first displayed.
    geometry_tolerance = 0.5
    mismatches = {
        name: {"actual": actual.get(name), "expected": expected.get(name)}
        for name in sorted(actual.keys() | expected.keys())
        if (name not in actual or name not in expected
            or any(abs(observed - baseline) > geometry_tolerance
                   for observed, baseline in zip(actual[name], expected[name])))
    }
    check("requested_widths_and_remaining_layout", not mismatches, True, mismatches=mismatches,
          controls_checked=len(expected), registered_baseline_height=expected["lstRegistered"][3],
          tolerance_points=geometry_tolerance)
    check("controls_inside_client_area", clipped, [],
          inside_width=dialog.InsideWidth, inside_height=dialog.InsideHeight)
    bottom_margin = min(
        dialog.InsideHeight - dialog.Controls(name).Top - dialog.Controls(name).Height
        for name in ("cmdSave", "cmdClear", "cmdClose")
    )
    check("action_button_bottom_margin", bottom_margin >= 11.9, True,
          margin_points=bottom_margin)
    check("notice_content_preserved", dialog.Controls("lblNotice").Caption,
          "등록한 기간의 날짜는 해당 교직원의 근무일수 계산에서 제외됩니다.")
    check("instructions_content_and_line_break", dialog.Controls("lblInstructions").Caption,
          "시작일과 종료일을 차례로 클릭하세요. 종료일은 시작일보다 늦어야 합니다.\r\n"
          "교직원 변경·닫기 시 저장하지 않은 선택은 취소됩니다.")
    check("instructions_wrap", dialog.Controls("lblInstructions").WordWrap, True)
    for name in ("lblNotice", "lblInstructions", "lblStatus"):
        required_height = dialog.AuditCaptionHeight(name)
        check(name + "_text_fits", required_height <= dialog.Controls(name).Height, True,
              required_height=required_height, available_height=dialog.Controls(name).Height)


def focus_form(hwnd):
    current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
    foreground_thread = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
    user32.AttachThreadInput(current_thread, foreground_thread, True)
    try:
        # HWND_TOPMOST must be pointer sized on Windows 64-bit.
        if not user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x43):
            raise AssertionError("Could not bring the disposable form to the front")
        if not user32.SetForegroundWindow(hwnd):
            raise AssertionError("Could not focus the disposable form")
    finally:
        user32.AttachThreadInput(current_thread, foreground_thread, False)


def wheel_at(dialog, hwnd, control, delta=-120, y_offset=None, x_offset=None):
    focus_form(hwnd)
    x, y = control_point(dialog, hwnd, control, x_offset=x_offset, y_offset=y_offset)
    user32.SetCursorPos(x, y)
    pump()
    event = Input(0, InputUnion(mouse=MouseInput(0, 0, delta & 0xFFFFFFFF, 0x800, 0, 0)))
    if user32.SendInput(1, ctypes.byref(event), ctypes.sizeof(event)) != 1:
        raise AssertionError("Native mouse-wheel input was not accepted")
    pump(0.5)


HARNESS = '''Option Explicit
Public UIForm As frmPersonnelChange
Public Function UISetup() As String
    Dim allEmployees As Collection, employee As CEmployee, group As Collection, i As Long
    On Error GoTo Failed
    Set gEmployeesByName = CreateObject("Scripting.Dictionary")
    Set allEmployees = New Collection
    gEmployeesMonth = DateSerial(2026, 8, 1)
    For i = 1 To 77
        Set employee = New CEmployee
        employee.Order = i
        Select Case i
            Case 1: employee.Name = "길누누"
            Case 2, 4: employee.Name = "김나나"
            Case 3: employee.Name = "김니니"
            Case Else: employee.Name = "검증대상" & Format$(i, "00")
        End Select
        employee.Birthdate = Format$(830100 + i, "000000")
        If i >= 5 Then
            employee.PersonnelChangeStart = DateSerial(2026, 8, 1)
            employee.PersonnelChangeEnd = DateSerial(2026, 8, 3)
        End If
        If Not gEmployeesByName.Exists(employee.Name) Then
            Set group = New Collection
            gEmployeesByName.Add employee.Name, group
        End If
        gEmployeesByName(employee.Name).Add employee
        allEmployees.Add employee
    Next i
    Set UIForm = New frmPersonnelChange
    UIForm.Configure gEmployeesMonth, allEmployees
    UIForm.Show vbModeless
    UISetup = "ok"
    Exit Function
Failed:
    UISetup = CStr(Err.Number) & ": " & Err.Description
End Function
Public Function UIGetForm() As Object
    Set UIGetForm = UIForm
End Function
Public Function UIClose() As Boolean
    On Error GoTo Failed
    If Not UIForm Is Nothing Then
        UIForm.DisposeHandlers
        Unload UIForm
        Set UIForm = Nothing
    End If
    UIClose = True
    Exit Function
Failed:
    UIClose = False
End Function
'''

FORM_AUDIT = '''
Public Function AuditCaptionHeight(ByVal controlName As String) As Single
    Dim original As MSForms.Label, measured As MSForms.Label
    Set original = Me.Controls(controlName)
    Set measured = Me.Controls.Add("Forms.Label.1", "auditCaption", False)
    With measured
        .Font.Name = original.Font.Name
        .Font.Size = original.Font.Size
        .Font.Bold = original.Font.Bold
        .Width = original.Width
        .WordWrap = True
        .Caption = original.Caption
        .AutoSize = True
        AuditCaptionHeight = .Height
    End With
    Me.Controls.Remove "auditCaption"
End Function
Public Function AuditTextWidth(ByVal textValue As String) As Single
    Dim measured As MSForms.Label
    Set measured = Me.Controls.Add("Forms.Label.1", "auditTextWidth", False)
    With measured
        .Font.Name = "맑은 고딕"
        .Font.Size = 9
        .WordWrap = False
        .Caption = textValue
        .AutoSize = True
        AuditTextWidth = .Width
    End With
    Me.Controls.Remove "auditTextWidth"
End Function
Public Function AuditRegisteredBaselineHeight() As Single
    Dim baseline As MSForms.ListBox
    ' Original creation order: IntegralHeight is True when Height is assigned.
    ' Compare the same Excel runtime's row rounding without changing the live list.
    Set baseline = Me.Controls.Add("Forms.ListBox.1", "auditRegisteredBaseline", False)
    With baseline
        .Left = 307
        .Top = 154
        .Width = 420
        .Height = 125
        .ColumnCount = 2
        .ColumnWidths = "165 pt;235 pt"
        .Font.name = "맑은 고딕"
        .Font.Size = 9
        .IntegralHeight = False
        AuditRegisteredBaselineHeight = .Height
    End With
    Me.Controls.Remove "auditRegisteredBaseline"
End Function
Public Function AuditSelectedIndex() As Long
    AuditSelectedIndex = mSelectedEmployeeIndex
End Function
Public Sub AuditSave()
    mSaveButton_Click
End Sub
Public Sub AuditSelectDay(ByVal dayNumber As Long)
    SelectDay DateSerial(2026, 8, dayNumber)
End Sub
Public Sub AuditClear()
    mClearButton_Click
End Sub
Public Sub AuditClose()
    mCloseButton_Click
End Sub
'''
HOOK_AUDIT = '''
Public Function AuditWheelInstalled() As Boolean
    AuditWheelInstalled = (mHook <> 0)
End Function
'''


def install_accessors(book, stage):
    project = book.VBProject
    for name, filename, audit in (
        ("frmPersonnelChange", "frmPersonnelChange.frm", FORM_AUDIT),
        ("PersonnelMouseWheel", "PersonnelMouseWheel.bas", HOOK_AUDIT),
    ):
        text = (ROOT / "vba" / filename).read_text(encoding="utf-8-sig")
        try:
            component = project.VBComponents(name)
        except pythoncom.com_error:
            component = None
        if filename.endswith(".bas"):
            if component is not None:
                project.VBComponents.Remove(component)
            native = stage / filename
            native.write_bytes((text + audit).replace("\r\n", "\n").replace("\n", "\r\n").encode("cp949"))
            project.VBComponents.Import(str(native))
        else:
            if component is None:
                component = project.VBComponents.Add(3)
                component.Name = name
            if component.CodeModule.CountOfLines:
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
            code = text[text.index("Option Explicit"):] + audit
            component.CodeModule.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    component = project.VBComponents.Add(1)
    component.Name = "PersonnelUIAudit"
    component.CodeModule.AddFromString(HARNESS.replace("\n", "\r\n"))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=ROOT / "workbook" / "근무일수계산.xlsm")
    parser.add_argument("--report", type=Path, default=ROOT / "scripts" / "test-personnel-ui-report.json")
    parser.add_argument("--layout-only", action="store_true",
                        help="Check compilation, initial content and layout through COM; skip native input")
    args = parser.parse_args()
    stage = Path(tempfile.mkdtemp(prefix="workday-personnel-ui-"))
    copy = stage / "personnel-ui.xlsm"
    shutil.copy2(args.workbook, copy)
    excel = book = dialog = None
    checks = []

    def run(name):
        return excel.Run("'personnel-ui.xlsm'!" + name)

    def check(name, actual, expected, **details):
        checks.append({"name": name, "actual": actual, "expected": expected, "passed": actual == expected, **details})
        print(json.dumps(checks[-1]), flush=True)
        if actual != expected:
            raise AssertionError(name)

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        book = excel.Workbooks.Open(str(copy), 0, False)
        sources.install_sources(book, ROOT / "vba", stage, adapt_dialogs=False)
        install_accessors(book, stage)
        book.Save()
        book.Close(False)
        book = excel.Workbooks.Open(str(copy), 0, False)
        excel.Visible = True
        excel.WindowState = -4140
        excel.VBE.MainWindow.Visible = True
        excel.VBE.ActiveVBProject = book.VBProject
        book.VBProject.VBComponents("PersonnelUIAudit").CodeModule.CodePane.Show()
        compile_control = excel.VBE.CommandBars.FindControl(1, 578)
        if compile_control.Enabled:
            compile_control.Execute()
        pump()
        check("full_compile", compile_control.Enabled, False)
        excel.VBE.MainWindow.Visible = False
        excel.WindowState = -4143
        # Workbook_Open is suppressed in this isolated UI test; enable VBA output writes.
        book.Worksheets("작업결과").Protect("workday-ui", True, True, True, True)
        check("setup", run("UISetup"), "ok")
        dialog = run("UIGetForm")
        pump()
        employee = dialog.Controls("cmbEmployee")
        registered = dialog.Controls("lstRegistered")
        if not args.layout_only:
            hwnd = find_form(excel.Hwnd)
        check("form_caption", dialog.Caption, "인사변동자 등록")
        check_form_layout(dialog, check)
        check("month_title", dialog.Controls("lblTitle").Caption, "2026년 8월")
        check("dropdown_rows", employee.ListRows, 25)
        popup_width = float(re.search(r"[\d.]+", str(employee.ListWidth)).group())
        check("dropdown_popup_width", popup_width, 210)
        check("registered_two_columns", registered.ColumnCount, 2)
        column_widths = [float(re.search(r"[\d.]+", width).group())
                         for width in registered.ColumnWidths.split(";")]
        check("registered_column_widths", column_widths, [78, 116])
        for label, sample, available_width in (
            ("duplicate_identity", "김나나(830104)", column_widths[0]),
            ("full_period", "2026-08-06~2026-08-08", column_widths[1]),
        ):
            required_width = dialog.AuditTextWidth(sample)
            check(label + "_text_fits_column", required_width <= available_width - 2, True,
                  required_width=required_width, available_width=available_width)
        check("initial_employees", employee.ListCount, 77)
        check("initial_registered", registered.ListCount, 73)
        if args.layout_only:
            return 0  # The shared finally block still writes the report and closes Excel.
        check("hook_installed", run("PersonnelMouseWheel.AuditWheelInstalled"), True)
        employee.SetFocus()
        employee.Value = "기"
        pump()
        check("open_syllable_filter", employee.ListCount, 4)
        check("filter_not_committed", dialog.AuditSelectedIndex(), 0)
        employee.ListIndex = 3
        pump()
        check("duplicate_index_identity", dialog.AuditSelectedIndex(), 4)
        check("selected_identity_tooltip", employee.ControlTipText, "김나나(830104)")
        dialog.AuditSelectDay(6)
        dialog.AuditSelectDay(8)
        dialog.AuditSave()
        check("saved_registered_count", registered.ListCount, 74)
        check("saved_period", registered.List[0][1], "2026-08-06~2026-08-08")
        check("selected_period_content", dialog.Controls("lblStatus").Caption,
              "2026-08-06 ~ 2026-08-08\r\n3일 · 저장됨")
        check("selected_period_text_fits", dialog.AuditCaptionHeight("lblStatus") <=
              dialog.Controls("lblStatus").Height, True)
        dialog.AuditClear()
        check("clear_registered_count", registered.ListCount, 73)
        employee.SetFocus()
        employee.Value = "ㄱ"
        pump()
        check("initial_consonant_filter", employee.ListCount, 77)
        employee.Value = "김나"
        pump()
        check("prefix_duplicate_filter", employee.ListCount, 2)
        employee.ListIndex = 0
        pump()
        check("first_duplicate_identity", dialog.AuditSelectedIndex(), 2)
        employee.Value = "존재하지않음"
        pump()
        check("empty_filter", employee.ListCount, 0)
        check("empty_filter_no_employee", dialog.AuditSelectedIndex(), 0)
        employee.Value = ""
        pump()
        check("restore_all_filter", employee.ListCount, 77)
        # Filtering can leave the restored list scrolled to its last page.
        # Start at row zero so a downward wheel event has room to move.
        employee.TopIndex = 0
        pump()
        old_top = employee.TopIndex
        wheel_at(dialog, hwnd, employee, y_offset=54, x_offset=employee.Width + 24)
        check("dropdown_actual_wheel_down", employee.TopIndex > old_top, True,
              old_top=old_top, new_top=employee.TopIndex,
              hook_installed=run("PersonnelMouseWheel.AuditWheelInstalled"))
        old_top = employee.TopIndex
        wheel_at(dialog, hwnd, employee, 120, y_offset=54, x_offset=employee.Width + 24)
        check("dropdown_actual_wheel_up", employee.TopIndex < old_top, True,
              old_top=old_top, new_top=employee.TopIndex,
              hook_installed=run("PersonnelMouseWheel.AuditWheelInstalled"))
        registered.SetFocus()
        pump()
        old_top = registered.TopIndex
        wheel_at(dialog, hwnd, registered)
        check("registered_actual_wheel_down", registered.TopIndex > old_top, True)
        old_top = registered.TopIndex
        wheel_at(dialog, hwnd, registered, 120)
        check("registered_actual_wheel_up", registered.TopIndex < old_top, True)
        registered.ListIndex = 4
        registered.SetFocus()
        x, y = control_point(dialog, hwnd, registered, x_offset=20, y_offset=8)
        user32.SetCursorPos(x, y)
        user32.mouse_event(0x0002, 0, 0, 0, 0)
        user32.mouse_event(0x0004, 0, 0, 0, 0)
        pump()
        check("registered_click_identity", dialog.AuditSelectedIndex(), 5)
        dialog.AuditClose()
        check("close_hook_cleanup", run("PersonnelMouseWheel.AuditWheelInstalled"), False)
        run("UIClose")
        dialog = None
        check("dispose_hook_cleanup", run("PersonnelMouseWheel.AuditWheelInstalled"), False)
    except Exception as exc:
        checks.append({"name": "execution_error", "error": str(exc), "passed": False})
    finally:
        args.report.write_text(json.dumps(checks, ensure_ascii=False, indent=2), encoding="utf-8")
        print("Report: " + str(args.report), flush=True)
        if excel is not None:
            try:
                run("UIClose")
            except Exception:
                pass
        if book is not None:
            book.Close(False)
        if excel is not None:
            excel.Quit()
    return 0 if checks and all(item["passed"] for item in checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
