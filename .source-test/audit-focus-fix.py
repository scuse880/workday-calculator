from pathlib import Path
import json
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / '.source-test/focus-fix'
target = ROOT / 'workbook/근무일수계산.xlsm'
deployment = json.loads((OUT / 'deployment.json').read_text(encoding='utf-8-sig'))
backup = Path(deployment['backup'])
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.AutomationSecurity = 1
book = None

def code(text):
    return '\n'.join(line.rstrip() for line in text.strip().splitlines()).lower()

def snapshot(path, events):
    global book
    excel.EnableEvents = events
    book = excel.Workbooks.Open(str(path), 0, True)
    modules = {item.Name: code(item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
                              if item.CodeModule.CountOfLines else '')
               for item in book.VBProject.VBComponents}
    sheets = [(s.Name, s.Visible) for s in book.Worksheets]
    def action(value):
        # Excel rewrites self-qualified macro paths when a backup is renamed.
        for prefix in (book.Name + '!', "'" + book.Name + "'!"):
            if value.startswith(prefix):
                return '<this-workbook>!' + value[len(prefix):]
        return value
    shapes = {s.Name: [(a.Name, a.Left, a.Top, a.Width, a.Height, action(a.OnAction))
                      for a in s.Shapes] for s in book.Worksheets}
    protection = bool(book.ProtectStructure and book.Worksheets('작업').ProtectContents)
    book.Close(False)
    book = None
    return modules, sheets, shapes, protection

try:
    old, old_sheets, old_shapes, _ = snapshot(backup, False)
    new, new_sheets, new_shapes, protection = snapshot(target, True)
    expected = {'main_초과근무월별집계', 'main_시트보호', 'frmPersonnelChange'}
    changed = {name for name in old.keys() | new.keys() if old.get(name) != new.get(name)}
    matched = {}
    for name in expected:
        suffix = '.frm' if name.startswith('frm') else '.bas'
        text = (ROOT / 'vba' / (name + suffix)).read_text(encoding='utf-8-sig')
        matched[name] = new[name] == code(text[text.index('Option Explicit'):])
    report = {'changed_components': sorted(changed), 'expected_components_only': changed == expected,
              'source_matches': matched, 'sheet_names_order_visibility_preserved': old_sheets == new_sheets,
              'all_sheet_shapes_preserved': old_shapes == new_shapes, 'reopen_protection_active': protection}
    if old_shapes != new_shapes:
        report['shape_differences'] = {name: {'before': old_shapes[name], 'after': new_shapes[name]}
                                       for name in old_shapes if old_shapes[name] != new_shapes[name]}
    report['passed'] = (report['expected_components_only'] and all(matched.values())
                        and old_sheets == new_sheets and old_shapes == new_shapes and protection)
    (OUT / 'deployed-audit.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
    assert report['passed']
finally:
    if book is not None:
        book.Close(False)
    excel.Quit()
