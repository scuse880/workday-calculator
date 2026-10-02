"""Exercise work-status identity decisions in a disposable Excel workbook.

Production parsers, employee import, record staging/commit and calculation run
unchanged. Only dialog boundaries use test-real-sources.py's existing adapters.
Fixtures are synthetic except for the final four-source fixture regression; the
JSON report contains scenario labels and counts, never names or personal IDs.
Requires Windows Excel, pywin32 and openpyxl. No original workbook is saved.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import shutil
import tempfile

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("real_sources", Path(__file__).with_name("test-real-sources.py"))
real_sources = importlib.util.module_from_spec(spec)
spec.loader.exec_module(real_sources)

VBA = r'''
Public Sub IdentityInit()
    Set AuditIds = New Collection
    AuditIds.Add "C900000001"
    AuditIds.Add "C900000002"
    AuditSetBase
End Sub

Public Sub IdentityWriteOvertime(ByVal path As String)
    Dim source As Workbook, ws As Worksheet
    Set source = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = source.Worksheets(1)
    ws.Range("B5:B7").Value2 = "중고등학교교사"
    ws.Range("C5").Value2 = "검증동명(800101)"
    ws.Range("C6").Value2 = "검증동명(810202)"
    ws.Range("C7").Value2 = "검증단독(820303)"
    source.SaveAs path, xlOpenXMLWorkbook
    source.Close False
End Sub

Private Sub IdentityWriteRow(ByVal ws As Worksheet, ByVal rowNumber As Long, _
                             ByVal raw As String, ByVal role As String, ByVal reason As String, _
                             ByVal status As String, ByVal period As String)
    ws.Cells(rowNumber, 3).Value2 = raw
    If role = "#AUDIT_ROLE_ERROR#" Then
        ws.Cells(rowNumber, 4).Value = CVErr(xlErrValue)
    Else
        ws.Cells(rowNumber, 4).Value2 = role
    End If
    ws.Cells(rowNumber, 5).Value2 = period
    ws.Cells(rowNumber, 7).Value2 = reason
    ws.Cells(rowNumber, 11).Value2 = status
End Sub

Public Sub IdentityWriteWork(ByVal path As String, ByVal raw As String, ByVal role As String, _
                             Optional ByVal rawSecond As String = "", Optional ByVal reason As String = "", _
                             Optional ByVal status As String = "완결", Optional ByVal period As String = "", _
                             Optional ByVal repeat As Long = 1)
    Dim source As Workbook, ws As Worksheet, i As Long
    Set source = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = source.Worksheets(1)
    ' A valid preceding row proves later ambiguity rolls back staged work, too.
    IdentityWriteRow ws, 2, "검증단독", "중고등학교교사", "", "완결", _
                     "2026-08-07 14:00 ~ 2026-08-07 15:00"
    If Len(period) = 0 Then period = "2026-08-04 09:00 ~ 2026-08-04 10:00"
    For i = 1 To repeat
        IdentityWriteRow ws, i + 2, raw, role, reason, status, period
    Next i
    If Len(rawSecond) > 0 Then IdentityWriteRow ws, repeat + 3, rawSecond, role, reason, status, period
    source.SaveAs path, xlOpenXMLWorkbook
    source.Close False
End Sub

Public Sub IdentitySeedOtherData()
    Dim employee As CEmployee, item As Variant
    Dim trip As CTripRecord, vacation As CVacationWorkRecord
    For Each item In GetEmployeesInOrder()
        Set employee = item
        Set trip = New CTripRecord
        trip.StartAt = DateSerial(2026, 8, 6) + TimeSerial(9, 0, 0)
        trip.EndAt = DateSerial(2026, 8, 6) + TimeSerial(10, 0, 0)
        employee.TripRecords.Add trip
        Set vacation = New CVacationWorkRecord
        vacation.StartAt = DateSerial(2026, 8, 5) + TimeSerial(9, 0, 0)
        vacation.EndAt = DateSerial(2026, 8, 5) + TimeSerial(10, 0, 0)
        employee.VacationWorkRecords.Add vacation
        employee.PersonnelChangeStart = DateSerial(2026, 8, 10)
        employee.PersonnelChangeEnd = DateSerial(2026, 8, 11)
    Next item
    gTripLoaded = True
    gVacationWorkLoaded = True
End Sub

Public Sub IdentityRegistry(ByVal mode As String)
    Dim group As Collection
    Set group = gEmployeesByName("검증동명")
    Select Case mode
        Case "duplicate": group(2).NeisPersonId = group(1).NeisPersonId
        Case "missing": group(2).NeisPersonId = ""
        Case "invalid": group(2).NeisPersonId = "C123"
        Case "outside_role": group(2).JobTitle = "영양사"
        Case "missing_role": group(2).JobTitle = ""
        Case "work_flag_false": gWorkStatusLoaded = False
        Case "recommit_others": IdentityRecommitOthers
    End Select
End Sub

Private Sub IdentityRecommitOthers()
    Dim stage As Object, item As Variant
    Set stage = NewRecordStage()
    For Each item In GetEmployeesInOrder()
        Set stage(CStr(item.Order)) = item.TripRecords
    Next item
    CommitRecordStage stage, "Trip"
    Set stage = NewRecordStage()
    For Each item In GetEmployeesInOrder()
        Set stage(CStr(item.Order)) = item.VacationWorkRecords
    Next item
    CommitRecordStage stage, "VacationWork"
End Sub

Public Function IdentityCounts() As String
    Dim item As Variant
    For Each item In GetEmployeesInOrder()
        IdentityCounts = IdentityCounts & CStr(item.WorkStatusRecords.Count) & ":"
    Next item
End Function

Private Function IdentityRecords(ByVal records As Collection) As String
    Dim record As Variant
    IdentityRecords = CStr(records.Count) & "["
    For Each record In records
        IdentityRecords = IdentityRecords & CStr(CDbl(record.StartAt)) & "," & CStr(CDbl(record.EndAt)) & ";"
    Next record
    IdentityRecords = IdentityRecords & "]"
End Function

Public Function IdentityState(Optional ByVal includeWork As Boolean = True) As String
    Dim item As Variant, value As String
    value = CStr(CDbl(gEmployeesMonth)) & "|" & CStr(gTripLoaded) & "|" & CStr(gVacationWorkLoaded)
    If includeWork Then value = value & "|" & CStr(gWorkStatusLoaded)
    For Each item In GetEmployeesInOrder()
        value = value & "|" & CStr(item.Order) & ":" & item.Name & ":" & item.Birthdate & ":" & _
                item.NeisPersonId & ":" & item.JobTitle & ":" & CStr(CDbl(item.PersonnelChangeStart)) & ":" & _
                CStr(CDbl(item.PersonnelChangeEnd)) & ":" & IdentityRecords(item.TripRecords) & ":" & _
                IdentityRecords(item.VacationWorkRecords)
        If includeWork Then value = value & ":" & IdentityRecords(item.WorkStatusRecords)
    Next item
    IdentityState = value
End Function

Public Function IdentityResults() As String
    Dim ws As Worksheet, cell As Range, value As String, border As Variant
    Set ws = ThisWorkbook.Worksheets("작업결과")
    value = ws.UsedRange.Address
    For Each cell In ws.UsedRange.Cells
        value = value & "|" & cell.Address & ":" & IdentityValue(cell.Formula) & ":" & IdentityValue(cell.NumberFormat) & ":" & _
                IdentityValue(cell.Interior.Color) & ":" & IdentityValue(cell.Interior.Pattern) & ":" & IdentityValue(cell.Font.Color) & ":" & _
                IdentityValue(cell.Font.Bold) & ":" & IdentityValue(cell.Font.Size) & ":" & IdentityValue(cell.WrapText) & ":" & _
                IdentityValue(cell.HorizontalAlignment) & ":" & IdentityValue(cell.VerticalAlignment) & ":" & _
                IdentityValue(cell.RowHeight) & ":" & IdentityValue(cell.ColumnWidth)
        For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            value = value & ":" & IdentityValue(cell.Borders(border).LineStyle) & ":" & IdentityValue(cell.Borders(border).Weight)
        Next border
    Next cell
    IdentityResults = value
End Function

Private Function IdentityValue(ByVal value As Variant) As String
    If IsNull(value) Then
        IdentityValue = "<null>"
    ElseIf IsError(value) Then
        IdentityValue = "<error>"
    Else
        IdentityValue = CStr(value)
    End If
End Function
'''


def real_eligible_counts(overtime: Path, work_status: Path, target_ids: list[str]):
    """Independently count eligible rows after the shared fixture-ID adapter.

    The real work-status export omits birthdates. Tests assert target-ID sets,
    counts and identical before/after assignments, not the real birthdate pairing.
    """
    book, rows = real_sources.read_sheet(overtime)
    book.close()
    roster = {}
    for row in rows[4:]:
        if len(row) <= 2 or row[2] is None:
            continue
        match = re.fullmatch(r"([^()]+)\(([0-9]{6})\)", str(row[2]).strip())
        if match:
            roster.setdefault(match.group(1).strip(), []).append(str(row[1] or "").strip())
    book, rows = real_sources.read_sheet(work_status)
    book.close()
    identities = set(target_ids)
    eligible_targets = eligible_outside = 0
    for row in rows[1:]:
        if len(row) <= 10 or row[10] != "완결" or row[2] is None:
            continue
        reason = str(row[6] or "")
        if "육아시간" in reason or "모성보호시간" in reason:
            continue
        lines = str(row[2]).strip().splitlines()
        name = lines[0].strip()
        if name not in roster:
            continue
        if len(roster[name]) == 1:
            eligible_targets += 1
        elif len(lines) == 2 and lines[1].strip().strip("()").strip().upper() in identities:
            eligible_targets += 1
        elif str(row[3] or "").strip() == "영양사":
            eligible_outside += 1
        else:
            raise RuntimeError("The real fixture contains another unresolved duplicate identity")
    return eligible_targets, eligible_outside


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=ROOT / "workbook" / "근무일수계산.xlsm")
    parser.add_argument("--source-dir", type=Path, default=ROOT / "vba")
    parser.add_argument("--fixture-dir", type=Path, default=ROOT / "testsource")
    parser.add_argument("--baseline-workbook", type=Path, help="Read only: compare old work-status logic with identical target IDs")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    overtime = args.fixture_dir / "초과근무 월별집계(생년월일 표시).xlsx"
    work_status = args.fixture_dir / "근무상황목록.xlsx"
    originals = [args.workbook, overtime, work_status]
    if args.baseline_workbook:
        originals.append(args.baseline_workbook)
    hashes = {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in originals}
    report = {"production_dialog_adapters": ["file_picker", "messages", "duplicate_neis_ids"], "results": []}
    excel = book = audit = baseline = component = None
    pythoncom.CoInitialize()
    with tempfile.TemporaryDirectory(prefix="workdays-identity-") as temporary:
        stage = Path(temporary)
        copy = stage / "identity-audit.xlsm"
        shutil.copy2(args.workbook, copy)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 1
            baseline_code = None
            if args.baseline_workbook:
                baseline = excel.Workbooks.Open(str(args.baseline_workbook.resolve()), 0, True)
                component = baseline.VBProject.VBComponents("main_근무상황목록")
                baseline_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                component = None
                baseline.Close(False)
                baseline = None
            book = excel.Workbooks.Open(str(copy), 0, False)
            components = real_sources.install_sources(book, args.source_dir, stage)
            audit = book.VBProject.VBComponents.Add(1)
            audit.Name = "IdentityAudit"
            real_ids = real_sources.fixture_ids(overtime, work_status)
            real_expected_records, real_outside_records = real_eligible_counts(overtime, work_status, real_ids)
            configured = real_sources.AUDIT_VBA + VBA + "\nPublic Sub IdentityRealIds()\n    Set AuditIds = New Collection\n"
            configured += "\n".join("    AuditIds.Add " + real_sources.vba_string(value) for value in real_ids)
            configured += "\nEnd Sub\n"
            audit.CodeModule.AddFromString(configured)
            report["full_compile_passed"] = real_sources.compile_project(excel, book)
            if not report["full_compile_passed"]:
                raise RuntimeError("The complete VBA identity test project did not compile")
            print(json.dumps({"full_compile_passed": True}), flush=True)

            def run(name, *values):
                return excel.Run("'identity-audit.xlsm'!" + name, *values)

            def check(label, success, **details):
                result = {"test": label, "passed": bool(success), **details}
                report["results"].append(result)
                print(json.dumps(result, ensure_ascii=True), flush=True)
                if not success:
                    raise AssertionError(label)

            if "main_시트보호" in components:
                run("ApplyWorkbookProtection")
            run("IdentityInit")
            roster_path = stage / "synthetic-roster.xlsx"
            baseline_path = stage / "synthetic-baseline.xlsx"
            run("IdentityWriteOvertime", str(roster_path))
            run("IdentityWriteWork", str(baseline_path), "검증동명\n(C900000001)", "중고등학교교사", "검증동명\n(C900000002)")

            def import_work(path):
                run("AuditConfigure", str(path))
                run("LoadWorkStatusRecords")

            def seed(mode=""):
                run("IdentityInit")
                run("AuditConfigure", str(roster_path))
                run("LoadOvertimeEmployees")
                if run("AuditEmployeeCount") != 3:
                    raise AssertionError("synthetic roster failed")
                import_work(baseline_path)
                if run("IdentityCounts") != "1:1:1:":
                    raise AssertionError("baseline identity import failed")
                run("IdentitySeedOtherData")
                run("근무일수계산")
                if book.Worksheets("작업결과").Cells(1, 33).Value2 != "근무일수":
                    raise AssertionError("baseline calculation failed")
                if mode:
                    run("IdentityRegistry", mode)

            # outcome, expected per-target counts; every fixture starts with one
            # valid single-target row so skip/abort cannot accidentally look alike.
            cases = [
                ("outside_nutritionist_four_rows_quiet", "검증동명\n(C900000003)", "영양사", "quiet", "0:0:1:", {"repeat": 4}),
                ("occupation_survives_trip_and_vacation_commit", "검증동명\n(C900000003)", "영양사", "quiet", "0:0:1:", {"mode": "recommit_others"}),
                ("unmatched_target_role_warns", "검증동명\n(C900000003)", "중고등학교교사", "warn", "0:0:1:", {}),
                ("unmatched_unknown_role_warns", "검증동명\n(C900000003)", "검증미상직종", "warn", "0:0:1:", {}),
                ("unmatched_blank_role_warns", "검증동명\n(C900000003)", "", "warn", "0:0:1:", {}),
                ("unmatched_four_rows_one_warning", "검증동명\n(C900000003)", "중고등학교교사", "warn", "0:0:1:", {"repeat": 4}),
                ("matching_both_ids", "검증동명\n(C900000001)", "중고등학교교사", "quiet", "1:1:1:", {"second": "검증동명\n(C900000002)"}),
                ("matching_lowercase_crlf_spaces", "  검증동명  \r\n ( c900000001 ) ", "중고등학교교사", "quiet", "1:0:1:", {}),
                ("matching_bare_id_cr", "검증동명\rC900000002", "중고등학교교사", "quiet", "0:1:1:", {}),
                ("matching_id_overrides_outside_role", "검증동명\n(C900000001)", "영양사", "quiet", "1:0:1:", {}),
                ("same_outside_role_in_roster_is_uncertain", "검증동명\n(C900000003)", "영양사", "warn", "0:0:1:", {"mode": "outside_role"}),
                ("missing_roster_role_is_uncertain", "검증동명\n(C900000003)", "영양사", "warn", "0:0:1:", {"mode": "missing_role"}),
                ("unknown_person_bad_identity_quiet", "검증외부\n(C123)", "중고등학교교사", "quiet", "0:0:1:", {"period": "not an interval"}),
                ("single_target_missing_id_unchanged", "검증단독", "중고등학교교사", "quiet", "0:0:2:", {}),
                ("single_target_malformed_id_unchanged", "검증단독\n(C123)", "영양사", "quiet", "0:0:2:", {}),
                ("single_target_other_id_unchanged", "검증단독\n(C900000003)", "영양사", "quiet", "0:0:2:", {}),
                ("missing_source_id_aborts", "검증동명", "중고등학교교사", "abort", None, {}),
                ("malformed_source_id_aborts", "검증동명\n(C123)", "중고등학교교사", "abort", None, {}),
                ("extra_identity_line_aborts", "검증동명\n(C900000001)\nextra", "중고등학교교사", "abort", None, {}),
                ("missing_id_outside_role_still_uncertain", "검증동명", "영양사", "abort", None, {}),
                ("malformed_id_outside_role_still_uncertain", "검증동명\n(C123)", "영양사", "abort", None, {}),
                ("duplicate_registry_id_aborts", "검증동명\n(C900000001)", "중고등학교교사", "abort", None, {"mode": "duplicate"}),
                ("missing_registry_id_aborts", "검증동명\n(C900000003)", "영양사", "abort", None, {"mode": "missing"}),
                ("malformed_registry_id_aborts", "검증동명\n(C900000003)", "영양사", "abort", None, {"mode": "invalid"}),
                ("abort_preserves_false_loaded_flag", "검증동명", "중고등학교교사", "abort", None, {"mode": "work_flag_false"}),
                ("noncompleted_bad_id_bypasses_identity", "검증동명\n(C123)", "중고등학교교사", "quiet", "0:0:1:", {"status": "미완결", "period": "not an interval"}),
                ("parenting_bad_id_bypasses_identity", "검증동명\n(C123)", "중고등학교교사", "quiet", "0:0:1:", {"reason": "육아시간", "period": "not an interval"}),
                ("maternity_bad_id_bypasses_identity", "검증동명\n(C123)", "중고등학교교사", "quiet", "0:0:1:", {"reason": "모성보호시간", "period": "not an interval"}),
            ]
            for index, (label, raw, role, outcome, counts, options) in enumerate(cases):
                seed(options.get("mode", ""))
                before_state = run("IdentityState")
                before_other = run("IdentityState", False)
                before_results = run("IdentityResults")
                path = stage / f"synthetic-case-{index:02}.xlsx"
                run("IdentityWriteWork", str(path), raw, role, options.get("second", ""), options.get("reason", ""),
                    options.get("status", "완결"), options.get("period", ""), options.get("repeat", 1))
                import_work(path)
                stopped = int(run("AuditMessageCount", "근무상황목록 중단"))
                warned = int(run("AuditMessageCount", "제외한 근무상황 기록"))
                completed = int(run("AuditMessageCount", "근무상황목록 완료"))
                if outcome == "abort":
                    check(label, stopped == 1 and completed == 0 and run("AuditHasMessage", "기존 근무상황 기록은 유지")
                          and run("IdentityState") == before_state and run("IdentityResults") == before_results,
                          preserved_records_flags_results=run("IdentityState") == before_state and run("IdentityResults") == before_results)
                else:
                    expected_warnings = 1 if outcome == "warn" else 0
                    row_messages_complete = outcome != "warn" or all(run("AuditHasMessage", f"({row}행)") for row in range(3, 3 + options.get("repeat", 1)))
                    check(label, stopped == 0 and completed == 1 and warned == expected_warnings and
                          run("IdentityCounts") == counts and run("IdentityState", False) == before_other and row_messages_complete,
                          records=int(run("AuditRecordCount", "WorkStatus")), warning_popups=warned)
                check(label + "_input_closed", excel.Workbooks.Count == 1)

            # The real fixture includes four records for a third same-name
            # nutritionist. Derive target IDs from roster occupations: choosing
            # the two most frequent IDs would incorrectly register the outsider.
            run("IdentityRealIds")
            run("AuditConfigure", str(overtime))
            run("LoadOvertimeEmployees")
            check("real_roster_unchanged", run("AuditEmployeeCount") == 77, employees=int(run("AuditEmployeeCount")))
            import_work(work_status)
            count = int(run("AuditRecordCount", "WorkStatus"))
            check("real_outside_rows_quiet_target_records_unchanged", count == real_expected_records and real_outside_records == 4 and
                  run("AuditMessageCount", "제외한 근무상황 기록") == 0 and run("AuditMessageCount", "근무상황목록 중단") == 0 and
                  run("AuditMessageCount", "근무상황목록 완료") == 1, records=count, expected_target_records=real_expected_records,
                  outside_records=real_outside_records)
            check("real_source_closed", excel.Workbooks.Count == 1)
            if baseline_code is not None:
                before_replacement = run("IdentityState")
                component = book.VBProject.VBComponents("main_근무상황목록")
                component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                component.CodeModule.AddFromString(real_sources.adapt_source(baseline_code, "main_근무상황목록"))
                component = None
                check("baseline_work_status_module_compiles", real_sources.compile_project(excel, book))
                run("IdentityRealIds")
                run("AuditConfigure", str(overtime))
                run("LoadOvertimeEmployees")
                import_work(work_status)
                check("same_correct_ids_preserve_all_real_target_records", run("IdentityState") == before_replacement,
                      records=int(run("AuditRecordCount", "WorkStatus")))
                check("baseline_only_difference_is_outside_warning", run("AuditMessageCount", "제외한 근무상황 기록") == 1 and
                      run("AuditMessageCount", "근무상황목록 중단") == 0 and run("AuditMessageCount", "근무상황목록 완료") == 1,
                      baseline_warning_popups=int(run("AuditMessageCount", "제외한 근무상황 기록")))
            report["passed"] = True
        finally:
            audit = component = None
            if baseline is not None:
                try:
                    baseline.Close(False)
                except pythoncom.com_error:
                    pass
            baseline = None
            if book is not None:
                try:
                    book.Close(False)
                except pythoncom.com_error:
                    pass
            book = None
            if excel is not None:
                try:
                    excel.Quit()
                except pythoncom.com_error:
                    pass
            excel = None
            gc.collect()
            pythoncom.CoUninitialize()
            report["original_hashes_unchanged"] = all(hashlib.sha256(path.read_bytes()).hexdigest() == value for path, value in hashes.items())
            if args.report:
                args.report.parent.mkdir(parents=True, exist_ok=True)
                args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if not report["original_hashes_unchanged"]:
        raise AssertionError("An original workbook changed during testing")
    print(json.dumps({"passed": report.get("passed", False), "checks": len(report["results"]), "original_hashes_unchanged": report["original_hashes_unchanged"]}), flush=True)


if __name__ == "__main__":
    main()
