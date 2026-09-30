from __future__ import annotations

from collections import Counter
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1] / "testsource"
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def redacted(value: str) -> str:
    out = []
    for char in value:
        if "가" <= char <= "힣":
            out.append("K")
        elif char.isdigit():
            out.append("0")
        elif char.isalpha():
            out.append("x")
        else:
            out.append(char)
    return "".join(out)[:70]


for path in sorted(ROOT.glob("*.xlsx")):
    if ".masked." in path.name:
        continue
    with zipfile.ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in si.iter(M + "t")) for si in shared.findall(M + "si")]
        print("FILE", path.name, "shared", len(strings), "parts", len(archive.namelist()))
        for name in sorted(n for n in archive.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)):
            tree = ET.fromstring(archive.read(name))
            rows = tree.findall(".//" + M + "row")
            types = Counter()
            cells = []
            for cell in tree.findall(".//" + M + "c"):
                kind = cell.attrib.get("t", "n")
                types[kind] += 1
                v = cell.find(M + "v")
                if kind == "s" and v is not None:
                    value = strings[int(v.text)]
                elif kind == "inlineStr":
                    value = "".join(t.text or "" for t in cell.iter(M + "t"))
                else:
                    value = v.text if v is not None else ""
                cells.append((cell.attrib.get("r", "?"), value, kind))
            print("SHEET", name, "rows", len(rows), "cells", len(cells), "types", dict(types))
            for address, value, kind in cells[:200]:
                print(" ", address, kind, redacted(value))
