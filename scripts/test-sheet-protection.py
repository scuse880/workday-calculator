"""Check production protection, worksheet events, and reopen in a disposable xlsm."""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import tempfile

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("real_sources", Path(__file__).with_name("test-real-sources.py"))
real_sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(real_sources)

VBA = r'''
Option Explicit
Public Function ProtectionSelectorText(ByVal address As String, ByVal index As Long) As String
    ProtectionSelectorText = ThisWorkbook.Worksheets("작업").Shapes("select_" & address).ControlFormat.List(index)
End Function

Public Function ProtectionVbaWrite() As Boolean
    Dim ws As Worksheet, original As Variant
    Set ws = ThisWorkbook.Worksheets("작업")
    original = ws.Range("B3").Value2
    ws.Range("B3").Value2 = "2027년"
    ProtectionVbaWrite = (ws.Range("B3").Value2 = "2027년")
    ws.Range("B3").Value2 = original
    Set ws = ThisWorkbook.Worksheets("방학중근무")
    ws.Range("E20").Interior.Color = RGB(1, 2, 3)
    ProtectionVbaWrite = ProtectionVbaWrite And (ws.Range("E20").Interior.Color = RGB(1, 2, 3))
    ws.Range("E20").Interior.Pattern = xlNone
End Function

Public Function ProtectionBorderCase(ByVal scenario As Long) As Boolean
    Dim ws As Worksheet, edge As Variant
    Set ws = ThisWorkbook.Worksheets("방학중근무")
    ws.Range("A13:D20").ClearContents
    ws.Range("A13:D20").Borders.LineStyle = xlLineStyleNone
    Select Case scenario
    Case 1
        ws.Range("A13").Value2 = "검증"
        ProtectionBorderCase = True
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical)
            ProtectionBorderCase = ProtectionBorderCase And (ws.Range("A13:D13").Borders(edge).LineStyle = xlContinuous)
        Next edge
        ProtectionBorderCase = ProtectionBorderCase And (ws.Range("A13").Borders(xlDiagonalDown).LineStyle = xlLineStyleNone)
    Case 2
        ws.Range("A13").Value2 = "검증"
        ws.Range("A13:D13").ClearContents
        ProtectionBorderCase = (ws.Range("A13:D13").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone And _
                                ws.Range("A13:D13").Borders(xlInsideVertical).LineStyle = xlLineStyleNone)
    Case 3
        ws.Range("A13").Value2 = "위"
        ws.Range("A14").Value2 = "삭제"
        ws.Range("A15").Value2 = "아래"
        ws.Range("A14:D14").ClearContents
        ProtectionBorderCase = (ws.Range("A13:D13").Borders(xlEdgeBottom).LineStyle = xlContinuous And _
                                ws.Range("A15:D15").Borders(xlEdgeTop).LineStyle = xlContinuous And _
                                ws.Range("A14:D14").Borders(xlInsideVertical).LineStyle = xlLineStyleNone)
    Case 4
        ws.Range("A13").Value2 = "검증"
        ws.Range("B13").Value2 = "남은 값"
        ws.Range("A13").ClearContents
        ProtectionBorderCase = (ws.Range("A13:D13").Borders(xlInsideVertical).LineStyle = xlContinuous)
    Case 5
        ws.Range("A13:A15").Value2 = "검증"
        ws.Range("A13:D15").ClearContents
        ProtectionBorderCase = (ws.Range("A13:D15").Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone And _
                                ws.Range("A13:D15").Borders(xlInsideVertical).LineStyle = xlLineStyleNone)
    Case 6
        ws.Range("A15:D15").Borders(xlEdgeTop).LineStyle = xlDouble
        ws.Range("A14").Value2 = "검증"
        ws.Range("A14:D14").ClearContents
        ProtectionBorderCase = (ws.Range("A15:D15").Borders(xlEdgeTop).LineStyle = xlDouble)
    Case 7
        ws.Range("A1000").Value2 = "끝"
        ws.Range("A3:D" & ws.Rows.Count).ClearContents
        ProtectionBorderCase = (ws.Range("A1000:D1000").Borders(xlInsideVertical).LineStyle = xlLineStyleNone)
    Case 8
        ws.Range("A13").Value2 = "위"
        ws.Range("A15").Value2 = "아래"
        ws.Range("A13:D13").ClearContents
        ws.Range("A15:D15").ClearContents
        ProtectionBorderCase = (ws.Range("A14:D14").Borders(xlEdgeTop).LineStyle = xlLineStyleNone And _
                                ws.Range("A14:D14").Borders(xlEdgeBottom).LineStyle = xlLineStyleNone)
    Case 9
        ws.Range("A13").Value2 = "위"
        ws.Range("A14").Value2 = "아래"
        ws.Range("A13:D13").ClearContents
        ws.Range("A14:D14").ClearContents
        ProtectionBorderCase = (ws.Range("A13:D14").Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone)
    End Select
End Function
'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    original = ROOT / "workbook" / "근무일수계산.xlsm"
    before = hashlib.sha256(original.read_bytes()).hexdigest()
    report = {"results": []}
    excel = book = None
    pythoncom.CoInitialize()
    with tempfile.TemporaryDirectory(prefix="workdays-protection-") as folder:
        stage = Path(folder)
        copy = stage / "protection.xlsm"
        shutil.copy2(original, copy)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 1
            book = excel.Workbooks.Open(str(copy), 0, False)
            real_sources.install_sources(book, ROOT / "vba", stage, adapt_dialogs=False)
            component = book.VBProject.VBComponents.Add(1)
            component.Name = "ProtectionAudit"
            component.CodeModule.AddFromString(VBA)

            def run(name, *values):
                return excel.Run("'protection.xlsm'!" + name, *values)

            def check(label, condition, **details):
                item = {"test": label, "passed": bool(condition), **details}
                report["results"].append(item)
                print(json.dumps(item, ensure_ascii=True), flush=True)
                if not condition:
                    raise AssertionError(label)

            check("full_compile", real_sources.compile_project(excel, book))
            settings = book.Worksheets("설정")
            settings_before = [list(settings.Range(settings.Cells(2, c), settings.Cells(settings.Rows.Count, c).End(-4162)).Value2)
                               for c in range(1, 5)]
            work = book.Worksheets("작업")
            original_values = {a: work.Range(a).Value2 for a in ("B3", "B5", "B6", "D3", "D5", "D6")}
            run("ApplyWorkbookProtection")
            check("activation_preserves_work_values", all(work.Range(a).Value2 == v for a, v in original_values.items()))
            for c in range(1, 5):
                after = list(settings.Range(settings.Cells(2, c), settings.Cells(settings.Rows.Count, c).End(-4162)).Value2)
                expected = settings_before[c - 1]
                if expected[0][0] != "선택":
                    expected = [("선택",)] + expected
                check(f"settings_placeholder_preserves_column_{c}", after == expected)
            settings_after = settings.UsedRange.Value2
            run("InitializeWorkbookProtection")
            check("initialization_defaults_all_six", all(work.Range(a).Value2 == "선택" for a in original_values))
            run("ApplyWorkbookProtection")
            check("settings_placeholder_is_idempotent", settings.UsedRange.Value2 == settings_after)
            for name in ("작업", "방학중근무", "사용방법 및 주의사항", "설정", "공휴일"):
                ws = book.Worksheets(name)
                check("protected_" + name, ws.ProtectContents and ws.ProtectionMode)
            check("structure_protected", book.ProtectStructure)
            for name in ("설정", "공휴일"):
                check("very_hidden_" + name, book.Worksheets(name).Visible == 2)
            for address in ("B3", "B5", "B6", "D3", "D5", "D6"):
                control = work.Shapes("select_" + address)
                cell = work.Range(address)
                check("selection_only_" + address, cell.Locked and control.Type == 8 and control.ControlFormat.ListIndex > 0
                      and abs(control.Left - cell.Left) < 0.1 and abs(control.Width - cell.Width) < 0.1)
                check("matching_row_and_dropdown_height_" + address,
                      abs(control.Height - cell.Height) < 0.1 and abs(control.Top - cell.Top) < 0.1
                      and abs(work.Cells(cell.Row, 1).Height - cell.Height) < 0.1,
                      row_height=cell.Height, dropdown_height=control.Height)
                check("placeholder_selected_" + address,
                      control.ControlFormat.ListIndex == 1 and run("ProtectionSelectorText", address, 1) == "선택"
                      and settings.Range(cell.Validation.Formula1.split("!")[1]).Row == 2)
                source_column = settings.Range(cell.Validation.Formula1.split("!")[1]).Column
                expected_items = [row[0] for row in settings.Range(settings.Cells(2, source_column),
                                  settings.Cells(settings.Rows.Count, source_column).End(-4162)).Value2 if row[0]]
                check("selector_preserves_all_options_" + address,
                      control.ControlFormat.ListCount == len(expected_items)
                      and all(run("ProtectionSelectorText", address, i + 1) == item
                              for i, item in enumerate(expected_items)), options=len(expected_items))
            check("other_work_cells_locked", work.Range("A1").Locked and work.Range("XFD1048576").Locked)
            vacation = book.Worksheets("방학중근무")
            check("vacation_editable_input", not vacation.Range("A3:D1048576").Locked)
            check("vacation_locked_header_and_outside", vacation.Range("A1:D2").Locked and vacation.Range("E1:XFD1048576").Locked)
            for address, text in (("B2", "날짜\n(yyyy-mm-dd)"), ("C2", "시작 시각\n(hh:mm)"), ("D2", "종료 시각\n(hh:mm)")):
                check("header_" + address, vacation.Range(address).Value2 == text and vacation.Range(address).WrapText)
            check("vba_value_and_format_write", run("ProtectionVbaWrite"))
            excel.EnableEvents = True
            for scenario in range(1, 10):
                check(f"vacation_border_{scenario:02}", run("ProtectionBorderCase", scenario))
            vacation.Range("A18").Value2 = "재열기"
            for address, value in (("B3", "2026년"), ("D3", "08월"), ("B5", "8시"), ("D5", "30분"),
                                   ("B6", "16시"), ("D6", "30분")):
                work.Range(address).Value2 = value
            book.Save()
            excel.EnableEvents = False
            book.Close(False)
            book = excel.Workbooks.Open(str(copy), 0, False)
            check("user_interface_only_is_not_persisted", not book.Worksheets("작업").ProtectionMode)
            book.Close(False)
            excel.EnableEvents = True
            book = excel.Workbooks.Open(str(copy), 0, False)
            work = book.Worksheets("작업")
            check("reopen_defaults_all_six", all(work.Range(a).Value2 == "선택" for a in original_values))
            check("reopen_defaults_all_six_controls", all(work.Shapes("select_" + a).ControlFormat.ListIndex == 1
                                                         for a in original_values))
            check("reopen_settings_remain_unchanged", book.Worksheets("설정").UsedRange.Value2 == settings_after)
            work.Range("B3").Value2 = "2026년"
            other_book = excel.Workbooks.Add()
            book.Activate()
            other_book.Close(False)
            check("actual_activate_preserves_selection", work.Range("B3").Value2 == "2026년"
                  and run("ProtectionSelectorText", "B3", work.Shapes("select_B3").ControlFormat.ListIndex) == "2026년")
            for name in ("작업", "방학중근무", "사용방법 및 주의사항", "설정", "공휴일"):
                check("reopen_restores_" + name, book.Worksheets(name).ProtectContents and book.Worksheets(name).ProtectionMode)
            check("reopen_vba_write", run("ProtectionVbaWrite"))
            book.Worksheets("방학중근무").Range("A18:D18").ClearContents()
            reopened = book.Worksheets("방학중근무")
            check("reopen_border_delete", reopened.Range("A18:D18").Borders(9).LineStyle == -4142,
                  bottom=reopened.Range("A18:D18").Borders(9).LineStyle,
                  row19_edges=[reopened.Range("A19").Borders(edge).LineStyle for edge in (7, 8, 9, 10)], events=excel.EnableEvents)
            check("original_workbook_unchanged", hashlib.sha256(original.read_bytes()).hexdigest() == before)
        finally:
            if book is not None:
                book.Close(False)
            if excel is not None:
                excel.Quit()
            book = excel = None
            pythoncom.CoUninitialize()
            if args.report:
                args.report.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"checks": len(report["results"]), "passed": all(x["passed"] for x in report["results"])}))


if __name__ == "__main__":
    main()
