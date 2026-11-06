"""Module 01: an .xlsx file is a zip of XML. Build one by hand, read it by hand, edit it by hand.

No Excel, no openpyxl until the very last check. Everything here is the standard library:
zipfile to open the container, xml.etree to read the parts.

    python main.py
"""

from __future__ import annotations

import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile

# Every XML tag in a spreadsheet lives in this namespace; ElementTree writes it as {uri}tag.
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# The five parts that make a minimal workbook. Real files add styles, themes, doc properties...
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Satışlar" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""

# Strings are stored once, in sharedStrings.xml; cells with t="s" hold an index into that list.
SHARED_STRINGS = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
<si><t>Ürün</t></si><si><t>Adet</t></si><si><t>Çay</t></si></sst>"""

# A cell is <c r="B2" t="..."><v>value</v></c>. No t means number. A formula adds <f>text</f>.
SHEET = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>5</v></c></row>
<row r="3"><c r="B3"><f>B2*2</f><v>10</v></c></row>
</sheetData>
</worksheet>"""

PARTS = {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS,
    "xl/workbook.xml": WORKBOOK,
    "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
    "xl/sharedStrings.xml": SHARED_STRINGS,
    "xl/worksheets/sheet1.xml": SHEET,
}


def write_xlsx(path: str, parts: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            z.writestr(name, text.encode("utf-8"))


def read_cells(path: str) -> dict[str, tuple[object, str | None]]:
    """Return {address: (value, formula)} for the first sheet, using only zipfile + ElementTree."""
    with zipfile.ZipFile(path) as z:
        # 1. workbook.xml lists the sheets; each <sheet> points at a relationship id...
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        first = wb.find(f"{NS}sheets")[0]  # type: ignore[index]
        rid = first.get(f"{NS_REL}id")
        # 2. ...and the .rels file maps that id to a part name.
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == rid)
        # 3. shared strings: a plain list, index -> text
        sst = [
            "".join(t.text or "" for t in si.iter(f"{NS}t"))
            for si in ET.fromstring(z.read("xl/sharedStrings.xml"))
        ]
        sheet = ET.fromstring(z.read(f"xl/{target}"))
    cells: dict[str, tuple[object, str | None]] = {}
    for c in sheet.iter(f"{NS}c"):
        v = c.find(f"{NS}v")
        f = c.find(f"{NS}f")
        value: object = None
        if v is not None and v.text is not None:
            value = sst[int(v.text)] if c.get("t") == "s" else float(v.text)
        cells[c.get("r", "")] = (value, f.text if f is not None else None)
    return cells


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "hand-made.xlsx")
        write_xlsx(path, PARTS)

        print("Parts inside the zip:")
        with zipfile.ZipFile(path) as z:
            for info in z.infolist():
                print(f"  {info.filename:32} {info.file_size:5} bytes")

        cells = read_cells(path)
        print("Cells:", cells)
        assert cells["A1"] == ("Ürün", None)
        assert cells["A2"] == ("Çay", None)
        assert cells["B2"] == (5.0, None)
        # a formula cell has BOTH the formula text and the value Excel last calculated
        assert cells["B3"] == (10.0, "B2*2")

        # Editing by hand: change the number 5 to 7 by rewriting one part. Notice we did NOT
        # touch B3's cached <v>10</v> - it is now stale until something recalculates it.
        edited = os.path.join(tmp, "edited.xlsx")
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(edited, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b'<c r="B2"><v>5</v></c>', b'<c r="B2"><v>7</v></c>')
                zout.writestr(item, data)
        after = read_cells(edited)
        assert after["B2"] == (7.0, None)
        assert after["B3"] == (10.0, "B2*2"), "stale cached value: this is why formulas need care"
        print("After hand edit: B2 =", after["B2"][0], "but B3 still caches", after["B3"][0])

        # Does a real library agree that this is a valid workbook? (the only non-stdlib line)
        import openpyxl

        wb = openpyxl.load_workbook(edited)
        ws = wb["Satışlar"]
        assert ws["A2"].value == "Çay" and ws["B2"].value == 7 and ws["B3"].value == "=B2*2"
        print("openpyxl reads the hand-made file:", ws["A2"].value, ws["B2"].value, ws["B3"].value)

    print("MODULE OK")


if __name__ == "__main__":
    main()
