from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1] / "testsource"
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def redact(value: str) -> str:
    return "".join("K" if "가" <= c <= "힣" else "0" if c.isdigit() else "x" if c.isalpha() else c for c in value)[:70]


def read(path: Path):
    with zipfile.ZipFile(path) as z:
        sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        strings = ["".join(t.text or "" for t in si.iter(M + "t")) for si in sroot.findall(M + "si")]
        w = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        cells = {}
        for c in w.findall(".//" + M + "c"):
            v = c.find(M + "v")
            if c.get("t") == "s" and v is not None:
                cells[c.get("r")] = strings[int(v.text)]
            elif c.get("t") == "inlineStr":
                cells[c.get("r")] = "".join(t.text or "" for t in c.iter(M + "t"))
            elif v is not None:
                cells[c.get("r")] = v.text or ""
        return cells, strings, sroot


files = sorted(ROOT.glob("*.xlsx"))
orig = [p for p in files if ".masked." not in p.name]
books = {p.name: read(p) for p in orig}
overtime = next(c for n, (c, _, _) in books.items() if "생년월일 표시" in n)
no_birth = next(c for n, (c, _, _) in books.items() if "생년월일 미표시" in n)
workstatus = next(c for n, (c, _, _) in books.items() if n.startswith("근무상황"))
trip = next(c for n, (c, _, _) in books.items() if n.startswith("출장"))

employees = []
for cell, value in overtime.items():
    match = re.fullmatch(r"C(\d+)", cell)
    if match and int(match.group(1)) >= 5:
        parsed = re.fullmatch(r"([^()]+)\((\d{6})\)", value.strip())
        if parsed:
            employees.append((parsed.group(1), parsed.group(2), cell))
names = {name for name, _, _ in employees}
counts = Counter(name for name, _, _ in employees)
print("EMPLOYEES", len(employees), "unique_names", len(counts), "duplicate_groups", Counter(counts.values()))
print("SPECIAL_DUPLICATE", counts.get("이은경", 0))
print("NO_BIRTH_NAME_COUNT", sum(1 for c, v in no_birth.items() if re.fullmatch(r"C\d+", c) and v in names))

for label, cells in [("workstatus", workstatus), ("trip", trip), ("overtime", overtime), ("no_birth", no_birth)]:
    long_digits = Counter()
    emails = Counter()
    all_name_hits = Counter()
    for addr, value in cells.items():
        col = re.match(r"[A-Z]+", addr).group()
        long_digits[col] += len(re.findall(r"(?<!\d)\d{9,}(?!\d)", value))
        emails[col] += value.count("@")
        all_name_hits[col] += sum(value.count(name) for name in names)
    print(label, "long_digits_by_col", dict(long_digits), "emails_by_col", dict(emails), "known_name_hits_by_col", dict(all_name_hits))

trip_a = [v for c, v in trip.items() if re.fullmatch(r"A\d+", c) and v]
trip_person = [v for v in trip_a if "(" in v and "*" in v]
print("TRIP_PERSON_ROWS", len(trip_person), "unique", len(set(trip_person)), "samples_redacted", [redact(v) for v in list(dict.fromkeys(trip_person))[:25]])
print("TRIP_A_OTHER_PATTERNS", Counter(redact(v) for v in trip_a if v not in trip_person).most_common(12))

work_c = [v for c, v in workstatus.items() if re.fullmatch(r"C\d+", c) and "(" in v]
print("WORK_C_ROWS", len(work_c), "unique", len(set(work_c)), "patterns", Counter(redact(v) for v in work_c).most_common(6))
special_ids = Counter()
for value in work_c:
    if value.startswith("이은경"):
        for neis in re.findall(r"[Cc]\d{9}", value):
            special_ids[neis] += 1
print("SPECIAL_IDS", len(special_ids), "frequencies", sorted(special_ids.values()))

for label, cells, columns, rows in [
    ("TRIP_SAMPLE", trip, "ABFGHIJ", [5, 12, 19, 26, 33, 40, 47, 54, 100, 300, 500]),
    ("WORK_SAMPLE", workstatus, "CJI", [2, 16, 41, 196, 240, 280, 281]),
]:
    print(label)
    for row in rows:
        print(" ", row, " ".join(f"{column}:{redact(cells.get(f'{column}{row}', ''))}" for column in columns))
for label, cells, col in [("TRIP_B", trip, "B"), ("TRIP_H", trip, "H"), ("TRIP_I", trip, "I"), ("TRIP_J", trip, "J"), ("WORK_J", workstatus, "J")]:
    values = [v for c, v in cells.items() if re.fullmatch(col + r"\d+", c) and v]
    print(label, "unique", len(set(values)), "patterns", Counter(redact(v) for v in values).most_common(12))

for name, (cells, strings, sroot) in books.items():
    print("PACKAGE", name, "shared_rich", sum(len(si.findall(M + "r")) > 0 for si in sroot.findall(M + "si")), "formula_cells", sum(1 for x in cells.values() if x.startswith("=")))
