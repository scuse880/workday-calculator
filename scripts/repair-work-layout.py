"""Repair only work-step alignment in a separate Excel workbook copy.

Preserves the input rows, column widths, button dimensions and formatting. The
seven buttons and their notes share cell anchors; no buttons are recreated.
Requires Windows, Excel and pywin32.
"""
from pathlib import Path
import argparse
import json
import shutil
import win32com.client

STEPS = [
    ("btnHolidays", 9), ("btnEmployees", 11), ("btnPersonnel", 13),
    ("btnWorkStatus", 15), ("btnTrips", 17), ("btnVacation", 19),
    ("btnCalculate", 21),
]


def repair(work):
    work.Unprotect("workday-ui")
    # Capture before writing: the old sheet has six notes for seven buttons.
    personnel_note = " 인사변동자가 있는 경우 근무일수에서 제외할 기간을 등록합니다."
    if work.Range("F13").Value2 == personnel_note:
        notes = [work.Range(f"F{row}").Value2 for _, row in STEPS]
    else:
        old_notes = [work.Range(f"F{row}").Value2 for row in (9, 11, 13, 15, 17, 19)]
        notes = old_notes[:2] + [personnel_note] + old_notes[2:]
    # The final step used a short trailing row and needs the same 30pt as 1–6.
    work.Rows(21).RowHeight = work.Rows(19).RowHeight
    work.Range("E19:F19").Copy(work.Range("E21:F21"))
    for (name, row), note in zip(STEPS, notes):
        work.Range(f"E{row}").Value2 = "<-"
        work.Range(f"F{row}").Value2 = note
        work.Range(f"E{row}:F{row}").VerticalAlignment = -4108  # xlCenter
        button = work.Shapes(name)
        # Set the cell attachment after moving; future input-row height changes
        # must move the button and note together without resizing the button.
        button.Top = work.Cells(row, 1).Top
        button.Placement = 2  # xlMove
    work.Protect("workday-ui", True, True, True, True)


def verify(work):
    return [{"button": name, "row": row,
             "aligned": abs(work.Shapes(name).Top - work.Cells(row, 1).Top) <= 0.1,
             "fits_row": abs(work.Shapes(name).Height - work.Rows(row).RowHeight) <= 0.1,
             "moves_with_note": work.Shapes(name).Placement == 2,
             "note_present": bool(work.Range(f"F{row}").Value2)}
            for name, row in STEPS]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.source.resolve() == args.output.resolve():
        raise ValueError("Use a separate output copy")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.source, args.output)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3
    book = None
    try:
        book = excel.Workbooks.Open(str(args.output.resolve()), 0, False)
        repair(book.Worksheets("작업"))
        report = verify(book.Worksheets("작업"))
        assert all(all(x[k] for k in ("aligned", "fits_row", "moves_with_note", "note_present")) for x in report)
        book.Save()
        print(json.dumps(report, ensure_ascii=False, indent=2))
    finally:
        if book is not None:
            book.Close(False)
        excel.Quit()


if __name__ == "__main__":
    main()
