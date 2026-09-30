from collections import defaultdict
import runpy
import zipfile
from xml.etree import ElementTree as ET

m = runpy.run_path("scripts/mask-testsource.py")
paths = m["source_paths"]()
books = {role: m["read_book"](path) for role, path in paths.items()}
overtime = m["overtime_people"](books["overtime"][0])
work = m["work_people"](books["workstatus"][0])
trip = m["trip_people"](books["trip"][0])
names = m["new_names"](m["gather_names"](books, overtime, work, trip))
births = m["new_birthdates"](overtime)
ids = m["new_person_ids"](work, [v for _, strings in books.values() for v in strings])
prefixes = m["new_prefixes"](trip)
phones = m["new_phones"](books["trip"][0])
transform = m["make_transform"](names, births, ids, prefixes, phones)
cell_transform = m["make_cell_transform"](transform, names, overtime, work, trip)
role = "trip"
with zipfile.ZipFile(paths[role]) as archive:
    tree = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
data = defaultdict(lambda: [[], []])
for cell in tree.findall(".//" + m["MAIN"] + "c"):
    if cell.get("t") != "s":
        continue
    idx = int(cell.find(m["MAIN"] + "v").text)
    addr = cell.get("r")
    old = books[role][1][idx]
    changed = cell_transform(role, addr, old) != old
    data[idx][int(changed)].append(addr)
for idx, (same, changed) in data.items():
    if same and changed:
        print("CONFLICT", idx, "changed", changed[:14], "unchanged", same[:14])
