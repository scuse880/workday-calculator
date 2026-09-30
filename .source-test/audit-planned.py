from collections import Counter
import re
import runpy

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
print("alias_count", len(names))
for role, (cells, _) in books.items():
    changed = Counter()
    for address, value in cells.items():
        if cell_transform(role, address, value) != value:
            changed[re.match(r"[A-Z]+", address).group()] += 1
    print(role, dict(changed))
    try:
        desired = m["planned_shared_strings"](paths[role], role, books[role][1], cell_transform, transform)
        print(role, "changed_shared_strings", sum(a != b for a, b in zip(books[role][1], desired)))
    except RuntimeError as error:
        print(role, "conflict", error)
