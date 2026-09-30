"""Create privacy-masked copies of the four testsource XLSX files.

The originals are read only. Only OOXML shared strings and document-author fields
are changed, so source layout, cell types, styles, and page breaks stay intact.
No identity mapping is written to disk.
"""

from __future__ import annotations

from collections import Counter
from datetime import date, timedelta
from hashlib import sha256
from itertools import product
from pathlib import Path
import re
import secrets
from xml.etree import ElementTree as ET
import zipfile


SOURCE_DIR = Path(__file__).resolve().parents[1] / "testsource"
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DC = "{http://purl.org/dc/elements/1.1/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
ET.register_namespace("", MAIN[1:-1])
ET.register_namespace("cp", CP[1:-1])
ET.register_namespace("dc", DC[1:-1])
ET.register_namespace("dcterms", "http://purl.org/dc/terms/")
ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")

PERSON_ID_RE = re.compile(r"(?<![A-Za-z0-9])[Cc]\d{9}(?!\d)")
ID_PREFIX_RE = re.compile(r"\(([A-Za-z0-9]{3})(\*+)\)")
BIRTH_RE = re.compile(r"(?<!\d)\d{6}(?!\d)")
PHONE_RE = re.compile(r"(?<!\d)(?:\d{10,11}|\d{2,3}-\d{3,4}-\d{4})(?!\d)")
KOREAN_NAME_RE = re.compile(r"[가-힣]{2,4}")
DATE_TIME_RE = re.compile(r"\d{4}\.\d{2}\.\d{2}")
WORK_PERSON_RE = re.compile(r"^([^\r\n]+)[\r\n]+\(([Cc]\d{9})\)$")
OVERTIME_PERSON_RE = re.compile(r"^([^()]+)\((\d{6})\)$")


def fail(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def digest(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def source_paths() -> dict[str, Path]:
    paths = [p for p in SOURCE_DIR.glob("*.xlsx") if not p.stem.endswith(".masked")]
    fail(len(paths) == 4, "Expected exactly four original XLSX source files")
    roles = {
        "overtime": next((p for p in paths if "생년월일 표시" in p.name), None),
        "overtime_no_birth": next((p for p in paths if "생년월일 미표시" in p.name), None),
        "workstatus": next((p for p in paths if p.name.startswith("근무상황")), None),
        "trip": next((p for p in paths if p.name.startswith("출장")), None),
    }
    fail(all(roles.values()) and len(set(roles.values())) == 4, "Source-file roles are ambiguous")
    return roles


def read_book(path: Path) -> tuple[dict[str, str], list[str]]:
    with zipfile.ZipFile(path) as archive:
        fail("xl/sharedStrings.xml" in archive.namelist(), "Source uses unexpected string storage")
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings = ["".join(t.text or "" for t in si.iter(MAIN + "t")) for si in root.findall(MAIN + "si")]
        fail(all(not si.findall(MAIN + "r") for si in root.findall(MAIN + "si")), "Unexpected rich-text source cell")
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        cells: dict[str, str] = {}
        for cell in sheet.findall(".//" + MAIN + "c"):
            value = cell.find(MAIN + "v")
            if cell.get("t") == "s" and value is not None:
                cells[cell.get("r", "")] = strings[int(value.text or "0")]
            elif cell.get("t") == "inlineStr":
                cells[cell.get("r", "")] = "".join(t.text or "" for t in cell.iter(MAIN + "t"))
            elif value is not None:
                cells[cell.get("r", "")] = value.text or ""
        return cells, strings


def overtime_people(cells: dict[str, str]) -> dict[int, tuple[str, str]]:
    people = {}
    for address, value in cells.items():
        match = re.fullmatch(r"C(\d+)", address)
        if match and int(match.group(1)) >= 5:
            parsed = OVERTIME_PERSON_RE.fullmatch(value.strip())
            if parsed:
                people[int(match.group(1))] = parsed.groups()
    return people


def work_people(cells: dict[str, str]) -> dict[int, tuple[str, str]]:
    people = {}
    for address, value in cells.items():
        match = re.fullmatch(r"C(\d+)", address)
        if match and int(match.group(1)) >= 2:
            parsed = WORK_PERSON_RE.fullmatch(value.strip())
            if parsed:
                people[int(match.group(1))] = parsed.groups()
    return people


def trip_people(cells: dict[str, str]) -> dict[int, tuple[str, str, str]]:
    people = {}
    for address, value in cells.items():
        match = re.fullmatch(r"B(\d+)", address)
        if not match or int(match.group(1)) < 5:
            continue
        row = int(match.group(1))
        if not DATE_TIME_RE.match(cells.get(f"C{row}", "")) or not DATE_TIME_RE.match(cells.get(f"D{row}", "")):
            continue
        display = value.strip()
        masked_id = ID_PREFIX_RE.search(display)
        if masked_id:
            fail(display.endswith(masked_id.group(0)), "Unexpected travel ID placement")
            front = display[: masked_id.start()].rstrip()
            fail(len(front) > 2 and front[-2:].isdigit(), "Unexpected travel duplicate-name suffix")
            people[row] = (front[:-2].strip(), front[-2:], masked_id.group(1))
        else:
            people[row] = (display, "", "")
    return people


def gather_names(
    books: dict[str, tuple[dict[str, str], list[str]]],
    overtime: dict[int, tuple[str, str]],
    work: dict[int, tuple[str, str]],
    trip: dict[int, tuple[str, str, str]],
) -> set[str]:
    names = {name for name, _ in overtime.values()}
    names.update(name for name, _ in work.values())
    names.update(name for name, _, _ in trip.values())
    for role in ("overtime", "overtime_no_birth"):
        sender = books[role][0].get("D2", "").strip()
        if KOREAN_NAME_RE.fullmatch(sender):
            names.add(sender)
    work_cells = books["workstatus"][0]
    for row in work:
        value = work_cells.get(f"J{row}", "").strip()
        if KOREAN_NAME_RE.fullmatch(value):
            names.add(value)
    fail(all(KOREAN_NAME_RE.fullmatch(name) for name in names), "Unexpected person-name format")
    return names


def new_names(original_names: set[str]) -> dict[str, str]:
    syllables = "가나다라마바사아자차카타파하거너더러머버서어저처커터퍼허"
    candidates = ("가" + a + b for a, b in product(syllables, repeat=2))
    selected = []
    for candidate in candidates:
        if candidate in original_names or any(name in candidate for name in original_names):
            continue
        selected.append(candidate)
        if len(selected) == len(original_names):
            break
    fail(len(selected) == len(original_names), "Could not generate enough distinct alias names")
    return dict(zip(sorted(original_names), selected))


def new_birthdates(overtime: dict[int, tuple[str, str]]) -> dict[str, str]:
    originals = {birth for _, birth in overtime.values()}
    mapped = {}
    day = date(1980, 1, 1)
    for old in sorted(originals):
        while day.strftime("%y%m%d") in originals or day.strftime("%y%m%d") in mapped.values():
            day += timedelta(days=1)
        mapped[old] = day.strftime("%y%m%d")
        day += timedelta(days=1)
    return mapped


def new_person_ids(work: dict[int, tuple[str, str]], all_strings: list[str]) -> dict[str, str]:
    originals = {m.group().lower() for value in all_strings for m in PERSON_ID_RE.finditer(value)}
    special = Counter(person_id.lower() for name, person_id in work.values() if name == "이은경")
    fail(len(special) >= 2, "The specified duplicate-name ID case is missing")
    frequent = [person_id for person_id, _ in special.most_common(2)]
    fail(special[frequent[1]] > max((v for k, v in special.items() if k not in frequent), default=0),
         "Cannot distinguish the two matching IDs from warning-only IDs")
    mapped = {frequent[0]: "c111111111", frequent[1]: "c222222222"}
    used = set(originals) | set(mapped.values())
    for original in sorted(originals - set(mapped)):
        while True:
            candidate = f"c{secrets.randbelow(10**9):09d}"
            if candidate not in used:
                mapped[original] = candidate
                used.add(candidate)
                break
    return mapped


def new_prefixes(trip: dict[int, tuple[str, str, str]]) -> dict[str, str]:
    originals = {prefix for _, _, prefix in trip.values() if prefix}
    candidates = (char * 3 for char in "abcdefghijklmnopqrstuvwxyz")
    mapped = {}
    for original in sorted(originals):
        candidate = next((c for c in candidates if c not in originals), None)
        fail(candidate is not None, "Too many travel ID prefixes")
        mapped[original] = candidate
    return mapped


def new_phones(cells: dict[str, str]) -> dict[str, str]:
    originals = {}
    for address, value in cells.items():
        if re.fullmatch(r"G\d+", address) and PHONE_RE.fullmatch(value):
            originals[re.sub(r"\D", "", value)] = None
    mapped = {}
    for index, original in enumerate(sorted(originals), 1):
        candidate = f"{index:0{len(original)}d}"
        fail(candidate not in originals, "Synthetic contact number collides with source")
        mapped[original] = candidate
    return mapped


def make_transform(
    names: dict[str, str], births: dict[str, str], person_ids: dict[str, str],
    prefixes: dict[str, str], phones: dict[str, str],
):
    name_re = re.compile("|".join(re.escape(name) for name in sorted(names, key=len, reverse=True)))

    def replace_phone(match: re.Match[str]) -> str:
        original = match.group()
        digits = re.sub(r"\D", "", original)
        if digits not in phones:
            return original
        replacement_digits = iter(phones[digits])
        return "".join(next(replacement_digits) if char.isdigit() else char for char in original)

    def transform(value: str) -> str:
        value = name_re.sub(lambda match: names[match.group()], value)
        value = PERSON_ID_RE.sub(lambda match: person_ids[match.group().lower()], value)
        value = ID_PREFIX_RE.sub(
            lambda match: f"({prefixes.get(match.group(1), match.group(1))}{match.group(2)})", value
        )
        value = BIRTH_RE.sub(lambda match: births.get(match.group(), match.group()), value)
        return PHONE_RE.sub(replace_phone, value)

    return transform


def make_cell_transform(transform, names: dict[str, str], overtime_rows, work_rows, trip_rows):
    name_re = re.compile("|".join(re.escape(name) for name in sorted(names, key=len, reverse=True)))
    free_text_name_re = re.compile(r"(?<![가-힣])(?:" + name_re.pattern + r")(?![가-힣])")

    def cell_transform(role: str, address: str, value: str) -> str:
        match = re.fullmatch(r"([A-Z]+)(\d+)", address)
        if not match:
            return value
        column, row_text = match.groups()
        row = int(row_text)
        if role in {"overtime", "overtime_no_birth"}:
            return transform(value) if (column == "C" and row in overtime_rows) or address == "D2" else value
        if role == "workstatus":
            return transform(value) if row in work_rows and column in {"C", "J"} else value
        if role != "trip":
            return value
        if column == "B" and row in trip_rows:
            return transform(value)
        if column == "F" and (free_text_name_re.search(value) or ID_PREFIX_RE.search(value)):
            return transform(value)
        if column == "G" and PHONE_RE.fullmatch(value):
            return transform(value)
        if column in {"H", "I", "J"}:
            if value in names:
                return names[value]
            for original in names:
                if value.startswith(original + "(") and value.endswith(")"):
                    return names[original] + value[len(original):]
        return value

    return cell_transform


def planned_shared_strings(path: Path, role: str, strings: list[str], cell_transform, transform):
    with zipfile.ZipFile(path) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    desires: dict[int, set[str]] = {}
    references: dict[int, list[str]] = {}
    for cell in sheet.findall(".//" + MAIN + "c"):
        if cell.get("t") != "s":
            continue
        index = int(cell.find(MAIN + "v").text or "0")
        value = cell_transform(role, cell.get("r", ""), strings[index])
        desires.setdefault(index, set()).add(value)
        references.setdefault(index, []).append(cell.get("r", ""))
    conflicts = [(index, references[index][:8]) for index, values in desires.items() if len(values) != 1]
    fail(not conflicts,
         f"A shared string is used by conflicting personal and non-personal cells: {conflicts[:4]}")
    return [next(iter(desires[index])) if index in desires else transform(value)
            for index, value in enumerate(strings)]


def write_copy(path: Path, output: Path, desired_strings: list[str]) -> None:
    if output.exists():
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(output) as previous:
            fail(source.namelist() == previous.namelist(),
                 "Existing masked destination is not a script-generated copy")
            fail(all(source.read(part) == previous.read(part)
                     for part in source.namelist()
                     if part not in {"xl/sharedStrings.xml", "docProps/core.xml"}),
                 "Existing masked destination has unrelated edits")
            fail(source.read("xl/sharedStrings.xml") != previous.read("xl/sharedStrings.xml"),
                 "Existing masked destination has not been anonymized")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(output, "w") as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "xl/sharedStrings.xml":
                root = ET.fromstring(payload)
                for index, string_item in enumerate(root.findall(MAIN + "si")):
                    texts = list(string_item.iter(MAIN + "t"))
                    fail(len(texts) == 1, "Unexpected rich-text item")
                    texts[0].text = desired_strings[index]
                # XML parsers normalize literal CRLF to LF. Preserve Excel's
                # original in-cell line breaks as character references.
                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True).replace(b"\r", b"&#13;")
            elif item.filename == "docProps/core.xml":
                root = ET.fromstring(payload)
                for field in (root.find(DC + "creator"), root.find(CP + "lastModifiedBy")):
                    if field is not None:
                        field.text = "Masked test source"
                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(item, payload)


def assert_no_original_identifiers(
    output: Path, names: dict[str, str], births: dict[str, str], person_ids: dict[str, str],
    prefixes: dict[str, str], phones: dict[str, str],
) -> None:
    with zipfile.ZipFile(output) as archive:
        fail(archive.testzip() is None, "Masked archive failed CRC check")
        for item in archive.namelist():
            if not (item.endswith(".xml") or item.endswith(".rels")):
                continue
            root = ET.fromstring(archive.read(item))
            values = [element.text or "" for element in root.iter()]
            values.extend(value for element in root.iter() for value in element.attrib.values())
            for value in values:
                fail(not any(re.search(r"(?<![가-힣])" + re.escape(name) + r"(?![가-힣])", value)
                             for name in names), "Source person-name token remains in masked OOXML")
                fail(not any(re.search(r"(?<!\d)" + re.escape(birth) + r"(?!\d)", value) for birth in births),
                     "Source birthdate remains in masked OOXML")
                fail(not any(old in value.lower() for old in person_ids), "Source person ID remains in masked OOXML")
                fail(not any(f"({prefix}" in value for prefix in prefixes), "Source user-ID prefix remains in masked OOXML")
                fail(not any(phone in re.sub(r"\D", "", value) for phone in phones),
                     "Source contact number remains in masked OOXML")


def validate_copy(
    original: Path, output: Path, old_cells: dict[str, str], role: str, cell_transform,
) -> dict[str, str]:
    with zipfile.ZipFile(original) as old_zip, zipfile.ZipFile(output) as new_zip:
        fail(old_zip.namelist() == new_zip.namelist(), "OOXML package structure changed")
        sheet = ET.fromstring(old_zip.read("xl/worksheets/sheet1.xml"))
        text_addresses = {
            cell.get("r") for cell in sheet.findall(".//" + MAIN + "c")
            if cell.get("t") in {"s", "inlineStr"}
        }
        for part in old_zip.namelist():
            if part not in {"xl/sharedStrings.xml", "docProps/core.xml"}:
                fail(old_zip.read(part) == new_zip.read(part), "A non-PII workbook part changed")
    new_cells, new_strings = read_book(output)
    fail(old_cells.keys() == new_cells.keys(), "Cell locations changed")
    mismatches = [
        (address, "text" if address in text_addresses else "other", len(value), len(new_cells[address]))
        for address, value in old_cells.items()
        if new_cells[address] != (cell_transform(role, address, value) if address in text_addresses else value)
    ]
    fail(not mismatches, f"Cell masking mismatch at {mismatches[:8]}")
    return new_cells


def main() -> None:
    paths = source_paths()
    old_hashes = {role: digest(path) for role, path in paths.items()}
    books = {role: read_book(path) for role, path in paths.items()}
    overtime = overtime_people(books["overtime"][0])
    work = work_people(books["workstatus"][0])
    trip = trip_people(books["trip"][0])
    fail(len(overtime) == 77 and len(work) == 294 and len(trip) == 134,
         "Expected source record counts changed")
    fail(Counter(name for name, _ in overtime.values())["이은경"] == 2,
         "Expected duplicate-name pair changed")
    names = new_names(gather_names(books, overtime, work, trip))
    births = new_birthdates(overtime)
    all_strings = [value for _, strings in books.values() for value in strings]
    person_ids = new_person_ids(work, all_strings)
    prefixes = new_prefixes(trip)
    phones = new_phones(books["trip"][0])
    transform = make_transform(names, births, person_ids, prefixes, phones)
    cell_transform = make_cell_transform(transform, names, overtime, work, trip)
    masked = {}
    for role, path in paths.items():
        output = path.with_name(path.stem + ".masked.xlsx")
        desired_strings = planned_shared_strings(path, role, books[role][1], cell_transform, transform)
        write_copy(path, output, desired_strings)
        masked[role] = validate_copy(path, output, books[role][0], role, cell_transform)
        assert_no_original_identifiers(output, names, births, person_ids, prefixes, phones)
    fail(all(digest(path) == old_hashes[role] for role, path in paths.items()),
         "An original source file changed")

    masked_overtime = overtime_people(masked["overtime"])
    masked_work = work_people(masked["workstatus"])
    masked_trip = trip_people(masked["trip"])
    fail(len(masked_overtime) == 77 and len(masked_work) == 294 and len(masked_trip) == 134,
         "Masked source parser counts changed")
    fail(Counter(name for name, _ in masked_overtime.values()).most_common(1)[0][1] == 2,
         "Duplicate-name scenario was lost")
    fail(all(masked_overtime[row] == (names[name], births[birth]) for row, (name, birth) in overtime.items()),
         "Overtime employee correspondence changed")
    fail(all(masked_work[row] == (names[name], person_ids[person_id.lower()])
             for row, (name, person_id) in work.items()),
         "Work-status employee correspondence changed")
    fail(all(masked_trip[row] == (names[name], suffix, prefixes.get(prefix, ""))
             for row, (name, suffix, prefix) in trip.items()),
         "Trip employee correspondence changed")
    fail({new_id for new_name, new_id in masked_work.values() if new_name == names["이은경"]}
         >= {"c111111111", "c222222222"}, "Specified duplicate-name IDs are missing")
    no_birth_old = books["overtime_no_birth"][0]
    no_birth_new = masked["overtime_no_birth"]
    fail(all(no_birth_new.get(f"C{row}", "") == names[name]
             for row, (name, _) in overtime.items()
             if no_birth_old.get(f"C{row}", "") == name),
         "No-birthdate overtime counterpart changed")
    target_names = {name for name, _ in overtime.values()}
    old_target_trips = sum(name in target_names for name, _, _ in trip.values())
    new_target_trips = sum(name in {names[n] for n in target_names} for name, _, _ in masked_trip.values())
    fail(old_target_trips == new_target_trips == 110, "Target trip count changed")
    print("Masked source copies: 4")
    print("Original source hashes: unchanged")
    print("Parsed records: overtime 77, work status 294, trip 134 (target 110)")
    print("Duplicate-name, ID mismatch, birthdate, user-ID prefix, and contact patterns: preserved")
    print("Original personal identifiers in masked OOXML: none detected")


if __name__ == "__main__":
    main()
