"""Run production VBA against the four unmasked fixtures in a disposable workbook.

Only dialog boundaries (file selection, message boxes, NEIS input, trip selection)
are adapted. Parsers, staging, commits and calculations execute production code.
The original workbook and input files are never saved. Reports contain counts and
test labels only. Requires Windows, Excel, pywin32 and openpyxl.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import shutil
import tempfile
import warnings

import openpyxl
import pythoncom
import win32com.client


ROOT = Path(__file__).resolve().parent.parent


def read_sheet(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = book.worksheets[0]
    # NEIS exports can advertise A1 dimensions despite containing hundreds of rows.
    sheet.reset_dimensions()
    return book, list(sheet.values)


def vba_string(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def fixture_ids(overtime: Path, work_status: Path) -> list[str]:
    book, rows = read_sheet(overtime)
    book.close()
    names = Counter()
    roster_jobs: dict[str, set[str]] = {}
    for row in rows[4:]:
        if len(row) > 2 and row[2] is not None:
            text = str(row[2]).strip()
            match = re.fullmatch(r"([^()]+)\(([0-9]{6})\)", text)
            if match:
                name = match.group(1).strip()
                names[name] += 1
                if len(row) > 1 and row[1]:
                    roster_jobs.setdefault(name, set()).add(str(row[1]).strip())
    duplicates = {name: count for name, count in names.items() if count > 1}
    book, rows = read_sheet(work_status)
    book.close()
    identities: dict[str, Counter] = {name: Counter() for name in duplicates}
    for row in rows[1:]:
        if len(row) <= 2 or not row[2]:
            continue
        lines = str(row[2]).strip().splitlines()
        if len(lines) != 2 or lines[0].strip() not in identities:
            continue
        identity = lines[1].strip().strip("()").strip().upper()
        # A same-name non-target can have more records than either target.
        # Frequency is not evidence of target identity; use the fixture roster's
        # job titles and reject ambiguity instead of guessing the top N IDs.
        if (re.fullmatch(r"C[0-9]{9}", identity) and len(row) > 3
                and str(row[3]).strip() in roster_jobs.get(lines[0].strip(), set())):
            identities[lines[0].strip()][identity] += 1
    result = []
    for name, count in duplicates.items():
        choices = sorted(identities[name].items())
        if len(choices) != count:
            raise RuntimeError("Fixture target identities cannot be resolved from roster job titles")
        result.extend(identity for identity, _ in choices)
    return result


def adapt_source(text: str, name: str) -> str:
    # Capture message boxes in memory, so expected validation errors cannot block COM.
    text = re.sub(r"\bMsgBox\b", "AuditMsgBox", text)
    if name == "main_초과근무월별집계":
        text, count = re.subn(
            r"selected = Application\.GetOpenFilename\([^\r\n]+\)",
            "selected = AuditPickInputPath()",
            text, count=1,
        )
        if count != 1:
            raise RuntimeError("File dialog adapter could not find its production boundary")
        text, count = re.subn(
            r"        Set dialog = New frmNeisPersonId\r?\n.*?        Set dialog = Nothing",
            "        AuditAssignNeis duplicates\n        If AuditCancelNeis Then GoTo Cancelled",
            text, count=1, flags=re.S,
        )
        if count != 1:
            raise RuntimeError("NEIS dialog adapter could not find its production boundary")
    if name == "main_출장목록개인별":
        text, count = re.subn(
            r"                    Set selector = New frmTripEmployee\r?\n.*?                    Set selector = Nothing",
            "                    If AuditCancelTrip Then GoTo Cancelled\n"
            "                    Set employee = candidates(1)\n"
            "                    skipRecord = False",
            text, count=1, flags=re.S,
        )
        if count != 1:
            raise RuntimeError("Trip dialog adapter could not find its production boundary")
    if name == "main_근무상황목록":
        # Long person-level warnings use a scrollable form. Initialize and configure
        # the real form, then capture its message at the same modal boundary.
        text = text.replace("    dialog.Show vbModal", "    AuditMsgBox text, vbExclamation, title")
    return text


def install_sources(book, source_dir: Path, stage: Path, adapt_dialogs: bool = True) -> list[str]:
    project = book.VBProject
    items = []
    for path in sorted(source_dir.iterdir()):
        if path.suffix.lower() not in {".bas", ".cls", ".frm"}:
            continue
        text = path.read_text(encoding="utf-8-sig")
        match = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
        if not match:
            raise RuntimeError("VBA source has no VB_Name")
        name = match.group(1)
        doc = re.search(r"^\s*'\s*@DocumentModule:(.+)\s*$", text, re.M)
        kind = 100 if doc else {".bas": 1, ".cls": 2, ".frm": 3}[path.suffix.lower()]
        items.append((path, name, text, kind, doc.group(1).strip() if doc else None))
    for path, name, text, kind, document in sorted(items, key=lambda x: x[3] != 3):
        code = adapt_source(text, name) if adapt_dialogs else text
        if kind == 100:
            document_code_name = book.CodeName if document == "Workbook" else book.Worksheets(document).CodeName
            component = project.VBComponents(document_code_name)
        else:
            try:
                old = project.VBComponents(name)
            except pythoncom.com_error:
                old = None
            if old is not None and old.Type == 100:
                raise RuntimeError("Production component conflicts with a document module")
            if kind == 3 and old is not None and old.Type == 3:
                component = old
            else:
                if old is not None:
                    project.VBComponents.Remove(old)
                if kind in {1, 2}:
                    # Native import preserves API declarations and class attributes.
                    # AddFromString can insert an extra "()" after a continued Declare.
                    native = stage / path.name
                    windows_text = re.sub(r"\r\n|\r|\n", "\r\n", code)
                    native.write_bytes(windows_text.encode("cp949", errors="strict"))
                    component = project.VBComponents.Import(str(native))
                    if component.Name != name or component.Type != kind:
                        raise RuntimeError("Native test import changed a component name or type")
                    continue
                component = project.VBComponents.Add(kind)
                component.Name = name
        code = code[code.index("Option Explicit"):]
        code = re.sub(r"^Attribute [^\r\n]+\r?\n?", "", code, flags=re.M)
        if component.CodeModule.CountOfLines:
            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
        component.CodeModule.AddFromString(code)
    return [item[1] for item in items]


def compile_project(excel, book) -> bool:
    """Compile the active whole project, including code outside exercised macros."""
    excel.VBE.ActiveVBProject = book.VBProject
    compile_control = excel.VBE.CommandBars.FindControl(1, 578)
    if compile_control is None:
        raise RuntimeError("Could not locate the VBA project compile command")
    if compile_control.Enabled:
        compile_control.Execute()
    return not bool(compile_control.Enabled)


AUDIT_VBA = r'''
Option Explicit
Public AuditInputPath As String
Public AuditMessages As Collection
Public AuditIds As Collection
Public AuditCancelNeis As Boolean
Public AuditCancelTrip As Boolean
Private AuditExpectedError As Boolean

Public Function AuditMsgBox(ByVal prompt As Variant, Optional ByVal buttons As Variant = 0, _
                           Optional ByVal title As Variant = Empty, Optional ByVal helpfile As Variant = Empty, _
                           Optional ByVal context As Variant = Empty) As VbMsgBoxResult
    If AuditMessages Is Nothing Then Set AuditMessages = New Collection
    AuditMessages.Add Array(CStr(prompt), CStr(title))
    AuditMsgBox = vbOK
End Function

Public Sub AuditClearMessages()
    Set AuditMessages = New Collection
End Sub

Public Function AuditPickInputPath() As Variant
    AuditPickInputPath = False
    If Len(AuditInputPath) = 0 Then Exit Function
    AuditPickInputPath = AuditInputPath
End Function

Public Sub AuditAssignNeis(ByVal employees As Collection)
    Dim i As Long
    If AuditCancelNeis Then Exit Sub
    For i = 1 To employees.Count
        employees(i).NeisPersonId = CStr(AuditIds(i))
    Next i
End Sub

Public Sub AuditConfigure(ByVal path As String, Optional ByVal cancelNeis As Boolean = False, _
                          Optional ByVal cancelTrip As Boolean = False)
    AuditInputPath = path
    AuditCancelNeis = cancelNeis
    AuditCancelTrip = cancelTrip
    AuditClearMessages
End Sub

Public Function AuditMessageCount(Optional ByVal titleFragment As String = "") As Long
    Dim message As Variant
    If AuditMessages Is Nothing Then Exit Function
    For Each message In AuditMessages
        If Len(titleFragment) = 0 Or InStr(CStr(message(1)), titleFragment) > 0 Then AuditMessageCount = AuditMessageCount + 1
    Next message
End Function

Public Function AuditHasMessage(ByVal fragment As String) As Boolean
    Dim message As Variant
    If AuditMessages Is Nothing Then Exit Function
    For Each message In AuditMessages
        If InStr(CStr(message(0)), fragment) > 0 Then AuditHasMessage = True: Exit Function
    Next message
End Function

Public Function AuditEmployeeCount() As Long
    Dim employees As Collection
    Set employees = GetEmployeesInOrder()
    AuditEmployeeCount = employees.Count
End Function

Public Function AuditRecordCount(ByVal kind As String) As Long
    Dim item As Variant
    For Each item In GetEmployeesInOrder()
        Select Case kind
            Case "WorkStatus": AuditRecordCount = AuditRecordCount + item.WorkStatusRecords.Count
            Case "Trip": AuditRecordCount = AuditRecordCount + item.TripRecords.Count
            Case "VacationWork": AuditRecordCount = AuditRecordCount + item.VacationWorkRecords.Count
        End Select
    Next item
End Function

Public Sub AuditSetBase()
    With ThisWorkbook.Worksheets("작업")
        .Range("B3").Value2 = "2026년"
        .Range("D3").Value2 = "8월"
        .Range("B5").Value2 = "8시"
        .Range("D5").Value2 = "30분"
        .Range("B6").Value2 = "16시"
        .Range("D6").Value2 = "30분"
    End With
End Sub

Public Function AuditBaseSettingsRegression(ByVal scenario As Long) As Boolean
    Dim workMonth As Date, a As Long, b As Long, expectedError As Boolean
    Dim addresses As Variant, i As Long, description As String
    AuditSetBase
    On Error GoTo Failed
    With ThisWorkbook.Worksheets("작업")
        Select Case scenario
            Case 2
                .Range("B3").Value2 = 2026: .Range("D3").Value2 = 8
                .Range("B5").Value2 = 8: .Range("D5").Value2 = 30
                .Range("B6").Value2 = 16: .Range("D6").Value2 = 30
            Case 3
                .Range("B3").Value2 = " 2026년 ": .Range("D3").Value2 = " 8월 "
                .Range("B5").Value2 = " 8시 ": .Range("D5").Value2 = " 30분 "
                .Range("B6").Value2 = " 16시 ": .Range("D6").Value2 = " 30분 "
            Case 4
                .Range("B5").Value2 = "0시": .Range("D5").Value2 = "0분"
                .Range("B6").Value2 = "23시": .Range("D6").Value2 = "59분"
            Case 5 To 10
                expectedError = True
                addresses = Array("B3", "D3", "B5", "D5", "B6", "D6")
                .Range(CStr(addresses(scenario - 5))).Value2 = "선택"
            Case 11
                expectedError = True: .Range("B5").Value2 = "8분"
            Case 12
                expectedError = True: .Range("D5").Value2 = "30시"
            Case 13
                expectedError = True: .Range("B5").Value2 = "8 시"
        End Select
    End With
    ReadBaseSettings workMonth, a, b
    If scenario = 4 Then
        AuditBaseSettingsRegression = workMonth = DateSerial(2026, 8, 1) And a = 0 And b = 1439
    Else
        AuditBaseSettingsRegression = Not expectedError And workMonth = DateSerial(2026, 8, 1) And a = 510 And b = 990
    End If
    GoTo CleanUp
Failed:
    description = Err.Description
    AuditBaseSettingsRegression = expectedError
    If scenario >= 5 And scenario <= 10 Then AuditBaseSettingsRegression = expectedError And InStr(description, "드롭다운") > 0
    Resume CleanUp
CleanUp:
    AuditSetBase
End Function

Public Function AuditBoundaryCase(ByVal scenario As Long) As Boolean
    Dim parsed As Date, month As Date, a As Long, b As Long, name As String, birth As String, normalized As String
    Dim ws As Worksheet, holidays As Object, saved As Boolean, decision As Variant, span As Variant
    Dim oldMap As Object, oldMonth As Date, employee As CEmployee, emptyValue As Variant
    Dim raised As Boolean
    AuditExpectedError = False
    On Error GoTo ExpectedFailure
    Select Case scenario
        Case 1: AuditExpectedError = True: parsed = ParseDateTimeValue("2026-02-30 09:00", "-", "audit")
        Case 2: parsed = ParseDateTimeValue("2028-02-29 09:00", "-", "audit"): AuditBoundaryCase = (Day(parsed) = 29)
        Case 3: AuditExpectedError = True: parsed = ParseDateTimeValue("2026-08-01 24:00", "-", "audit")
        Case 4: AuditExpectedError = True: parsed = ParseDateTimeValue("2026-08-01 09:60", "-", "audit")
        Case 5: AuditExpectedError = True: parsed = ParseDateTimeValue("2026-08-01 09:00:01", "-", "audit")
        Case 6: AuditExpectedError = True: parsed = ParseDateTimeValue(60, "-", "audit")
        Case 7: parsed = ParseDateTimeValue(0, "-", "audit", True): AuditBoundaryCase = (parsed = DateSerial(1904, 1, 1))
        Case 8: AuditExpectedError = True: parsed = ParseDateTimeValue(0.5, "-", "audit")
        Case 9: AuditExpectedError = True: parsed = ParseDateTimeValue(DateSerial(2026, 8, 1) + TimeSerial(9, 0, 1), "-", "audit")
        Case 10: AuditExpectedError = True: parsed = ParseDateTimeValue(CVErr(xlErrValue), "-", "audit")
        Case 11: AuditExpectedError = True: ValidateInterval DateSerial(2026, 8, 1), DateSerial(2026, 8, 1), "audit"
        Case 12: AuditExpectedError = True: ValidateInterval DateSerial(2026, 8, 2), DateSerial(2026, 8, 1), "audit"
        Case 13: ParseNameBirth "  audit(001001)  ", name, birth, "audit": AuditBoundaryCase = (name = "audit" And birth = "001001")
        Case 14: AuditExpectedError = True: ParseNameBirth "audit(12345)", name, birth, "audit"
        Case 15: AuditExpectedError = True: ParseNameBirth "(001001)", name, birth, "audit"
        Case 16: AuditBoundaryCase = TryNormalizeNeisId(" c123456789 ", normalized) And normalized = "C123456789"
        Case 17: AuditBoundaryCase = Not TryNormalizeNeisId("C12345678", normalized)
        Case 18: AuditBoundaryCase = Not TryNormalizeNeisId("abc****", normalized)
        Case 19: AuditExpectedError = True: ThisWorkbook.Worksheets("작업").Range("B3").Value2 = "26년": ReadBaseSettings month, a, b
        Case 20: AuditExpectedError = True: ThisWorkbook.Worksheets("작업").Range("D3").Value2 = 13: ReadBaseSettings month, a, b
        Case 21: AuditExpectedError = True: ThisWorkbook.Worksheets("작업").Range("B5").Value2 = 24: ReadBaseSettings month, a, b
        Case 22: AuditExpectedError = True: ThisWorkbook.Worksheets("작업").Range("D6").Value2 = 60: ReadBaseSettings month, a, b
        Case 23: AuditExpectedError = True: ThisWorkbook.Worksheets("작업").Range("B6").Value2 = 8: ReadBaseSettings month, a, b
        Case 24: AuditExpectedError = True: RequireEmployeeState DateSerial(2026, 9, 1)
        Case 25: AuditExpectedError = True: SavePersonnelChange DateSerial(2026, 8, 1), 1, DateSerial(2026, 8, 3), DateSerial(2026, 8, 3)
        Case 26: AuditExpectedError = True: SavePersonnelChange DateSerial(2026, 8, 1), 1, DateSerial(2026, 8, 31), DateSerial(2026, 9, 1)
        Case 27: AuditExpectedError = True: SavePersonnelChange DateSerial(2026, 8, 1), 1, 0, DateSerial(2026, 8, 4)
        Case 28: AuditExpectedError = True: SavePersonnelChange DateSerial(2026, 8, 1), 999, DateSerial(2026, 8, 3), DateSerial(2026, 8, 4)
        Case 29: AuditExpectedError = True: SavePersonnelChange DateSerial(2026, 8, 1), 1, DateSerial(2026, 8, 3) + 0.5, DateSerial(2026, 8, 4)
        Case 30: span = RecordDaySpan(DateSerial(2026, 7, 31) + TimeSerial(9, 0, 0), DateSerial(2026, 8, 1), "Y", DateSerial(2026, 8, 1), 510, 990): AuditBoundaryCase = IsEmpty(span)
        Case 31: span = RecordDaySpan(DateSerial(2026, 8, 3) + TimeSerial(18, 0, 0), DateSerial(2026, 8, 4) + TimeSerial(7, 0, 0), "X", DateSerial(2026, 8, 3), 510, 990): AuditBoundaryCase = IsEmpty(span)
        Case 32: span = RecordDaySpan(DateSerial(2026, 8, 3) + TimeSerial(18, 0, 0), DateSerial(2026, 8, 4) + TimeSerial(7, 0, 0), "Y", DateSerial(2026, 8, 3), 510, 990): AuditBoundaryCase = (span(0) = 1080 And span(1) = 1440)
        Case 33: span = RecordDaySpan(DateSerial(2026, 8, 3) + TimeSerial(18, 0, 0), DateSerial(2026, 8, 4) + TimeSerial(7, 0, 0), "Y", DateSerial(2026, 8, 4), 510, 990): AuditBoundaryCase = (span(0) = 0 And span(1) = 420)
        Case 34: AuditBoundaryCase = IntervalsContain(Array(Array(540, 600), Array(780, 900)), Array(Array(840, 900)))
        Case 35: AuditBoundaryCase = Not IntervalsContain(Array(Array(540, 660)), Array(Array(840, 900)))
        Case 36: AuditBoundaryCase = Not IntervalsOutsideWork(Array(Array(510, 990)), 510, 990)
        Case 37: AuditBoundaryCase = IntervalsOutsideWork(Array(Array(509, 990)), 510, 990)
        Case 38: span = MergeIntervals(Array(Array(600, 720), Array(540, 600), Array(660, 780))): AuditBoundaryCase = (UBound(span) = 0 And span(0)(0) = 540 And span(0)(1) = 780)
        Case 39: decision = JudgeIntervals(Empty, Empty, Empty, 510, 990): AuditBoundaryCase = (decision(0) = 1 And decision(1) = 1)
        Case 40: decision = JudgeIntervals(Array(Array(540, 600)), Empty, Empty, 510, 990): AuditBoundaryCase = (decision(0) = 0 And decision(1) = 2)
        Case 41: decision = JudgeIntervals(Empty, Array(Array(540, 600)), Empty, 510, 990): AuditBoundaryCase = (decision(0) = 1 And decision(1) = 3)
        Case 42: decision = JudgeIntervals(Empty, Empty, Array(Array(540, 600)), 510, 990): AuditBoundaryCase = (decision(0) = 1 And decision(1) = 4)
        Case 43: decision = JudgeIntervals(Array(Array(540, 600)), Array(Array(540, 600)), Empty, 510, 990): AuditBoundaryCase = (decision(0) = 1 And decision(1) = 5)
        Case 44: decision = JudgeIntervals(Array(Array(840, 900)), Array(Array(540, 660)), Empty, 510, 990): AuditBoundaryCase = (decision(0) = 0 And decision(1) = 5)
        Case 45: decision = JudgeIntervals(Array(Array(840, 900)), Array(Array(500, 660)), Empty, 510, 990): AuditBoundaryCase = (decision(0) = -1 And decision(1) = 5)
        Case 46: decision = JudgeIntervals(Array(Array(540, 600)), Empty, Array(Array(540, 600)), 510, 990): AuditBoundaryCase = (decision(0) = 1 And decision(1) = 6)
        Case 47: decision = JudgeIntervals(Array(Array(840, 900)), Empty, Array(Array(540, 660)), 510, 990): AuditBoundaryCase = (decision(0) = 0 And decision(1) = 6)
        Case 48: decision = JudgeIntervals(Empty, Array(Array(540, 600)), Array(Array(540, 600)), 510, 990): AuditBoundaryCase = (decision(0) = -1)
        Case 49: decision = JudgeIntervals(Array(Array(540, 600)), Array(Array(540, 600)), Array(Array(540, 600)), 510, 990): AuditBoundaryCase = (decision(0) = -1)
        Case 50: AuditBoundaryCase = (MinuteText(1440) = "24:00")
    End Select
    GoTo CleanUp
ExpectedFailure:
    raised = True
    AuditBoundaryCase = AuditExpectedError And Err.Number <> 0
    Resume CleanUp
CleanUp:
    If AuditExpectedError And Not raised Then AuditBoundaryCase = False
    AuditSetBase
End Function

Public Sub AuditSeedVacation()
    Dim employee As CEmployee, employees As Collection
    Set employees = GetEmployeesInOrder()
    Set employee = employees(1)
    With ThisWorkbook.Worksheets("방학중근무")
        .Range("A3:D100").ClearContents
        .Range("A3").Value2 = employee.Name & "(" & employee.Birthdate & ")"
        .Range("B3").Value2 = "2026-08-03"
        .Range("C3").Value2 = "09:00"
        .Range("D3").Value2 = "12:00"
    End With
End Sub

Public Function AuditVacationError(ByVal scenario As Long) As Boolean
    Dim before As Long, employee As CEmployee, group As Collection, key As Variant
    AuditSeedVacation
    before = AuditRecordCount("VacationWork")
    With ThisWorkbook.Worksheets("방학중근무")
        Select Case scenario
            Case 1: .Range("A3").Value2 = "__AUDIT_UNKNOWN_EMPLOYEE__"
            Case 2: .Range("B3").Value2 = "2026-09-01"
            Case 3: .Range("B3").Value2 = "2026-02-30"
            Case 4: .Range("C3").Value2 = "24:00"
            Case 5: .Range("D3").Value2 = "09:00"
            Case 6: .Range("D3").Value2 = "08:00"
            Case 7: .Range("C3").ClearContents
            Case 8: .Range("C3").Value2 = CDbl(TimeSerial(9, 0, 1))
            Case 9: .Range("B3").Value2 = CDbl(DateSerial(2026, 8, 3)) + 0.5
            Case 10: .Range("A3").Value2 = .Range("A3").Value2 & "()"
            Case 11, 12
                For Each key In gEmployeesByName.Keys
                    Set group = gEmployeesByName(key)
                    If group.Count > 1 Then Exit For
                Next key
                .Range("A3").Value2 = CStr(key)
                If scenario = 12 Then .Range("A3").Value2 = CStr(key) & "(000000)"
        End Select
    End With
    AuditClearMessages
    LoadVacationWorkRecords
    AuditVacationError = AuditHasMessage("기존 방학중근무 자료는 유지") And AuditRecordCount("VacationWork") = before
End Function

Public Function AuditHolidayCase(ByVal scenario As Long) As Boolean
    Dim ws As Worksheet, holidays As Object, saved As Boolean, raised As Boolean
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("공휴일")
    ws.Range("A2:AF100").ClearContents
    Select Case scenario
        Case 1
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
            AuditHolidayCase = Not saved And holidays.Count = 0
        Case 2
            Set holidays = NewDictionary()
            SaveHolidays DateSerial(2026, 8, 1), holidays
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
            AuditHolidayCase = saved And holidays.Count = 0
        Case 3
            ws.Range("A2").Value2 = "2026-08-01"
            ws.Range("B2").Value2 = "2026-09-01"
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
        Case 4
            ws.Range("A2:A3").Value2 = "2026-08-01"
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
        Case 5
            ws.Range("A2").Value2 = "2026-08-02"
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
        Case 6
            Set holidays = NewDictionary()
            holidays.Add HolidayKey(DateSerial(2026, 8, 15)), DateSerial(2026, 8, 15)
            SaveHolidays DateSerial(2026, 8, 1), holidays
            Set holidays = LoadHolidays(DateSerial(2026, 8, 1), saved)
            AuditHolidayCase = saved And holidays.Count = 1
    End Select
    GoTo CleanUp
Failed:
    raised = True
    AuditHolidayCase = scenario >= 3 And scenario <= 5
    Resume CleanUp
CleanUp:
    ws.Range("A2:AF100").ClearContents
End Function

Public Function AuditTotalDays() As Long
    Dim ws As Worksheet, lastColumn As Long, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets("작업결과")
    lastColumn = 33
    lastRow = AuditEmployeeCount() + 1
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, lastColumn).Value2) Then AuditTotalDays = AuditTotalDays + CLng(ws.Cells(r, lastColumn).Value2)
    Next r
End Function

Public Function AuditPersonnelEffect() As Boolean
    Dim before As Long, after As Long, employee As CEmployee, i As Long, ws As Worksheet, employees As Collection
    AuditClearMessages
    근무일수계산
    Set ws = ThisWorkbook.Worksheets("작업결과")
    Set employees = GetEmployeesInOrder()
    For i = 1 To AuditEmployeeCount()
        Set employee = employees(i)
        If employee.WorkStatusRecords.Count = 0 And employee.TripRecords.Count = 0 And employee.VacationWorkRecords.Count = 0 Then Exit For
    Next i
    If i > AuditEmployeeCount() Then i = 1
    before = CLng(ws.Cells(i + 1, 33).Value2)
    SavePersonnelChange DateSerial(2026, 8, 1), i, DateSerial(2026, 8, 3), DateSerial(2026, 8, 4)
    근무일수계산
    after = CLng(ws.Cells(i + 1, 33).Value2)
    AuditPersonnelEffect = (Len(CStr(ws.Cells(i + 1, 4).Value2)) = 0 And Len(CStr(ws.Cells(i + 1, 5).Value2)) = 0 And after <= before)
    ClearPersonnelChange DateSerial(2026, 8, 1), i
End Function

Public Sub AuditWriteWorkFixture(ByVal path As String, ByVal scenario As Long)
    Dim source As Workbook, ws As Worksheet, employees As Collection, employee As CEmployee
    Dim group As Collection, key As Variant, statuses As Variant, i As Long
    Set source = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = source.Worksheets(1)
    Set employees = GetEmployeesInOrder()
    Set employee = employees(1)
    If scenario = 1 Then
        statuses = Array("완결취소", "미완결", "완결 처리중", "", "완결")
        For i = 0 To UBound(statuses)
            ws.Cells(i + 2, 3).Value2 = employee.Name
            ws.Cells(i + 2, 5).Value2 = "malformed interval must be ignored"
            ws.Cells(i + 2, 11).Value2 = statuses(i)
        Next i
        ' 미완결 행은 이름·사유·상태가 Excel 오류값이어도 검증하지 않는다.
        ws.Cells(2, 3).Value = CVErr(xlErrNA)
        ws.Cells(3, 7).Value = CVErr(xlErrValue)
        ws.Cells(4, 11).Value = CVErr(xlErrValue)
        ws.Cells(6, 5).Value2 = "2026-08-03 09:00 ~ 2026-08-03 10:00"
    ElseIf scenario = 2 Then
        For Each key In gEmployeesByName.Keys
            Set group = gEmployeesByName(key)
            If group.Count > 1 Then Exit For
        Next key
        For i = 2 To 5
            ws.Cells(i, 3).Value2 = CStr(key) & vbLf & "(C999999999)"
            ws.Cells(i, 4).Value2 = "audit"
            ws.Cells(i, 5).Value2 = "2026-08-03 09:00 ~ 2026-08-03 10:00"
            ws.Cells(i, 11).Value2 = "완결"
        Next i
    ElseIf scenario = 4 Then
        For Each key In gEmployeesByName.Keys
            Set group = gEmployeesByName(key)
            If group.Count > 1 Then Exit For
        Next key
        For i = 2 To 41
            ws.Cells(i, 3).Value2 = CStr(key) & vbLf & "(C999999999)"
            ws.Cells(i, 4).Value2 = "audit"
            ws.Cells(i, 5).Value2 = "2026-08-03 09:00 ~ 2026-08-03 10:00"
            ws.Cells(i, 11).Value2 = "완결"
        Next i
    ElseIf scenario = 3 Then
        ws.Cells(2, 3).Value2 = employee.Name
        ws.Cells(2, 5).Value2 = "2026-02-30 09:00 ~ 2026-08-03 10:00"
        ws.Cells(2, 11).Value2 = "완결"
    End If
    source.SaveAs path, xlOpenXMLWorkbook
    source.Close SaveChanges:=False
End Sub
'''


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, default=ROOT / "workbook" / "근무일수계산.xlsm")
    parser.add_argument("--source-dir", type=Path, default=ROOT / "vba")
    parser.add_argument("--fixture-dir", type=Path, default=ROOT / "testsource")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    fixtures = sorted(p for p in args.fixture_dir.glob("*.xlsx") if ".masked." not in p.name and not p.name.startswith("~$"))
    if len(fixtures) != 4:
        raise RuntimeError("Expected the four unmasked fixture workbooks")
    overtime = args.fixture_dir / "초과근무 월별집계(생년월일 표시).xlsx"
    missing_birth = args.fixture_dir / "초과근무 월별집계(생년월일 미표시).xlsx"
    work_status = args.fixture_dir / "근무상황목록.xlsx"
    trip = args.fixture_dir / "출장 근무상황부 개인별.xlsx"
    originals = [args.workbook, *fixtures]
    hashes = {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in originals}
    ids = fixture_ids(overtime, work_status)
    report = {"fixture_count": len(fixtures), "production_dialog_adapters": ["file_picker", "messages", "duplicate_neis_ids", "first_trip_candidate"], "results": []}
    excel = book = None
    pythoncom.CoInitialize()
    with tempfile.TemporaryDirectory(prefix="workdays-audit-") as temporary:
        stage = Path(temporary)
        copy = stage / "audit.xlsm"
        shutil.copy2(args.workbook, copy)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            # Events stay disabled, so document startup never races the source import.
            excel.AutomationSecurity = 1
            book = excel.Workbooks.Open(str(copy), 0, False)
            components = install_sources(book, args.source_dir, stage)
            audit = book.VBProject.VBComponents.Add(1)
            audit.Name = "AuditRuntime"
            configured = AUDIT_VBA + '\nPublic Sub AuditInit()\n    Set AuditIds = New Collection\n'
            configured += "\n".join("    AuditIds.Add " + vba_string(identity) for identity in ids)
            configured += "\n    AuditSetBase\nEnd Sub\n"
            audit.CodeModule.AddFromString(configured)
            excel.AutomationSecurity = 1
            report["full_compile_passed"] = compile_project(excel, book)
            if not report["full_compile_passed"]:
                raise RuntimeError("The complete VBA test project did not compile")
            print(json.dumps({"full_compile_passed": True}, ensure_ascii=True), flush=True)

            def run(name, *values):
                return excel.Run("'audit.xlsm'!" + name, *values)

            def check(label: str, success, **details):
                result = {"test": label, "passed": bool(success), **details}
                report["results"].append(result)
                print(json.dumps(result, ensure_ascii=True), flush=True)
                if not success:
                    raise AssertionError(label)

            def capture_focus():
                window = excel.ActiveWindow
                return {
                    "workbook": excel.ActiveWorkbook.FullName,
                    "sheet": excel.ActiveSheet.Name,
                    "selection": excel.Selection.Address,
                    "window": window.Hwnd,
                    "scroll_row": window.ScrollRow,
                    "scroll_column": window.ScrollColumn,
                    "zoom": window.Zoom,
                }

            def run_import_with_focus(label, macro, sheet_name="작업"):
                # Exercise the real open/close implementation with activation events
                # enabled, while adapting only the file-picker and modal dialogs.
                excel.EnableEvents = True
                book.Activate()
                book.Worksheets(sheet_name).Activate()
                book.Worksheets(sheet_name).Range("A1").Select()
                excel.ActiveWindow.ScrollRow = 2
                excel.ActiveWindow.ScrollColumn = 1
                before = capture_focus()
                run(macro)
                after = capture_focus()
                check(label, after == before and excel.EnableEvents,
                      active_sheet=after["sheet"], view_unchanged=after == before)
                excel.EnableEvents = False

            if "main_시트보호" in components:
                run("ApplyWorkbookProtection")
            run("AuditInit")
            for scenario in range(1, 14):
                check(f"display_units_settings_regression_{scenario:02}", run("AuditBaseSettingsRegression", scenario))
            run("AuditConfigure", str(missing_birth))
            run_import_with_focus("validation_failure_preserves_import_focus", "LoadOvertimeEmployees")
            check("missing_duplicate_birthdates_rejected", run("AuditEmployeeCount") == 0 and run("AuditHasMessage", "동명이인"))
            run("AuditConfigure", str(overtime))
            run_import_with_focus("overtime_import_preserves_focus", "LoadOvertimeEmployees")
            employee_count = int(run("AuditEmployeeCount"))
            check("overtime_actual_fixture", employee_count == 77, employees=employee_count)
            run("AuditConfigure", str(overtime), True)
            run_import_with_focus("cancelled_neis_dialog_preserves_focus", "LoadOvertimeEmployees")
            check("cancelled_neis_dialog_preserves_employees", run("AuditEmployeeCount") == employee_count)
            run("AuditConfigure", str(work_status))
            run_import_with_focus("work_status_import_preserves_focus", "LoadWorkStatusRecords")
            work_count = int(run("AuditRecordCount", "WorkStatus"))
            check("work_status_actual_fixture", work_count == 203, records=work_count)
            run("AuditConfigure", str(trip))
            run_import_with_focus("trip_import_preserves_focus", "LoadTripRecords")
            trip_count = int(run("AuditRecordCount", "Trip"))
            check("trip_actual_fixture", trip_count == 110, records=trip_count)
            run("AuditConfigure", str(trip), False, True)
            run_import_with_focus("cancelled_trip_dialog_preserves_focus", "LoadTripRecords")
            check("cancelled_trip_dialog_preserves_records", run("AuditRecordCount", "Trip") == trip_count)
            for macro in ("LoadOvertimeEmployees", "LoadWorkStatusRecords", "LoadTripRecords"):
                run("AuditConfigure", "")
                run_import_with_focus(f"cancelled_picker_preserves_focus_{macro}", macro)
            run("AuditConfigure", str(stage / "missing-input.xlsx"))
            run_import_with_focus("failed_file_open_preserves_focus", "LoadWorkStatusRecords")
            check("failed_file_open_preserves_records", run("AuditRecordCount", "WorkStatus") == work_count)
            preopened = excel.Workbooks.Open(str(work_status), 0, True)
            try:
                run("AuditConfigure", str(work_status))
                run_import_with_focus("already_open_input_preserves_focus", "LoadWorkStatusRecords")
                check("already_open_input_remains_open", preopened.Name == work_status.name)
            finally:
                preopened.Close(False)
            run("AuditConfigure", str(work_status))
            run_import_with_focus("import_preserves_non_work_caller_sheet", "LoadWorkStatusRecords", "방학중근무")

            # Pure and validation cases use the same production public functions.
            for scenario in range(1, 51):
                check(f"boundary_{scenario:02}", run("AuditBoundaryCase", scenario))
            for scenario in range(1, 7):
                check(f"holiday_{scenario:02}", run("AuditHolidayCase", scenario))
            run("AuditClearMessages")
            run("근무일수계산")
            check("calculation_without_holiday_setup", book.Worksheets("작업결과").Cells(1, 33).Value2 == "근무일수" and not run("AuditHasMessage", "공휴일을 확인"), total_days=int(run("AuditTotalDays")))
            check("personnel_period_excluded", run("AuditPersonnelEffect"))

            run("AuditSeedVacation")
            run("AuditClearMessages")
            run("LoadVacationWorkRecords")
            check("vacation_generated_valid_row", run("AuditRecordCount", "VacationWork") == 1)
            for scenario in range(1, 13):
                check(f"vacation_error_preserves_previous_{scenario:02}", run("AuditVacationError", scenario))
            run("AuditSeedVacation")
            run("AuditClearMessages")
            run("LoadVacationWorkRecords")
            run("근무일수계산")
            check("calculation_with_vacation", not run("AuditHasMessage", "기존 방학중근무 자료는 유지"))
            run("AuditConfigure", "")
            run("LoadWorkStatusRecords")
            check("cancelled_file_preserves_work_status", run("AuditRecordCount", "WorkStatus") == work_count)

            # A completed row alone reaches interval parsing; near matches are ignored.
            test_work = stage / "work-boundaries-1.xlsx"
            run("AuditWriteWorkFixture", str(test_work), 1)
            run("AuditConfigure", str(test_work))
            run("LoadWorkStatusRecords")
            check("non_exact_complete_rows_skipped_before_parsing", run("AuditRecordCount", "WorkStatus") == 1)
            test_work = stage / "work-boundaries-3.xlsx"
            run("AuditWriteWorkFixture", str(test_work), 3)
            run("AuditConfigure", str(test_work))
            run("LoadWorkStatusRecords")
            check("malformed_completed_period_preserves_previous", run("AuditRecordCount", "WorkStatus") == 1 and run("AuditHasMessage", "기존 근무상황 기록은 유지"))
            test_work = stage / "work-boundaries-2.xlsx"
            run("AuditWriteWorkFixture", str(test_work), 2)
            run("AuditConfigure", str(test_work))
            run("LoadWorkStatusRecords")
            check("four_identity_mismatches_one_person_popup", run("AuditRecordCount", "WorkStatus") == 0 and run("AuditMessageCount") == 2 and all(run("AuditHasMessage", f"({row}행)") for row in range(2, 6)), popups=int(run("AuditMessageCount")))
            check("identity_warning_has_neis_label_without_candidate_roster", run("AuditHasMessage", "나이스 개인번호:") and not run("AuditHasMessage", "원본 개인번호") and not run("AuditHasMessage", "등록된 동명이인:"))
            test_work = stage / "work-boundaries-4.xlsx"
            run("AuditWriteWorkFixture", str(test_work), 4)
            run("AuditConfigure", str(test_work))
            run("LoadWorkStatusRecords")
            check("long_warning_one_person_popup_keeps_all_rows", run("AuditRecordCount", "WorkStatus") == 0 and run("AuditMessageCount") == 2 and all(run("AuditHasMessage", f"({row}행)") for row in range(2, 42)), popups=int(run("AuditMessageCount")))

            # Reimport succeeds and clears all old records and personnel registrations.
            run("AuditConfigure", str(overtime))
            run("LoadOvertimeEmployees")
            check("employee_reimport_resets_records", run("AuditEmployeeCount") == employee_count and run("AuditRecordCount", "Trip") == 0 and run("AuditRecordCount", "VacationWork") == 0)
            report["passed"] = True
        finally:
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
            pythoncom.CoUninitialize()
            report["original_hashes_unchanged"] = all(hashlib.sha256(path.read_bytes()).hexdigest() == value for path, value in hashes.items())
            if args.report:
                args.report.parent.mkdir(parents=True, exist_ok=True)
                args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"passed": report.get("passed", False), "checks": len(report["results"]), "original_hashes_unchanged": report["original_hashes_unchanged"]}, ensure_ascii=True))


if __name__ == "__main__":
    main()
