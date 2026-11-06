"""Module 03: formulas vs values - what is cached, what is not, and how to tell.

A formula cell stores TWO things: the formula text (<f>) and the result the last program
calculated (<v>). openpyxl can read either, but it cannot calculate. So a tool must never show
None silently; it must say "this file was never calculated".

    python main.py
"""

from __future__ import annotations

import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile

import openpyxl
from openpyxl.formula.translate import Translator

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def cache_state(path: str, part: str = "xl/worksheets/sheet1.xml") -> dict[str, int]:
    """Count formula cells, and how many have no cached result. Streams the XML, one cell at a time."""
    formulas = no_cache = 0
    with zipfile.ZipFile(path) as z, z.open(part) as fh:
        for _event, el in ET.iterparse(fh, events=("end",)):
            if el.tag != f"{NS}c":
                continue
            f, v = el.find(f"{NS}f"), el.find(f"{NS}v")
            if f is not None:
                formulas += 1
                # <v/> or a missing <v> means "never calculated". The exception: t="str" with an
                # empty <v> is a formula that really returned the empty string "".
                if (v is None or not v.text) and el.get("t") != "str":
                    no_cache += 1
            el.clear()
    return {"formulas": formulas, "no_cache": no_cache}


# A sheet as Excel writes it: shared formula in D2:D4 (children carry only si="0"), cached values.
EXCEL_SHEET = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="2"><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2"><f t="shared" ref="D2:D4" si="0">B2*C2</f><v>6</v></c></row>
<row r="3"><c r="B3"><v>4</v></c><c r="C3"><v>5</v></c><c r="D3"><f t="shared" si="0"/><v>20</v></c></row>
<row r="4"><c r="B4"><v>1</v></c><c r="C4"><v>1</v></c><c r="D4"><f t="shared" si="0"/><v>1</v></c></row>
<row r="5"><c r="D5"><f>SUM(D2:D4)</f><v>27</v></c></row>
</sheetData></worksheet>"""


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        # --- 1. A script-written workbook: formulas, no cached values ------------------------
        script = os.path.join(tmp, "script.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws["A1"], ws["A2"], ws["A3"] = 10, 20, "=SUM(A1:A2)"
        wb.save(script)
        with zipfile.ZipFile(script) as z:
            xml = z.read("xl/worksheets/sheet1.xml").decode()
        assert "<f>SUM(A1:A2)</f><v />" in xml or "<f>SUM(A1:A2)</f><v/>" in xml
        print("openpyxl wrote:", re.search(r'<c r="A3".*?</c>', xml).group(0))  # type: ignore[union-attr]
        assert openpyxl.load_workbook(script, data_only=True).active["A3"].value is None
        assert cache_state(script) == {"formulas": 1, "no_cache": 1}
        print("cache_state :", cache_state(script), "-> tell the user, never show a blank")

        # --- 2. An Excel-written workbook: cached values present -----------------------------
        excel = os.path.join(tmp, "excel.xlsx")
        with zipfile.ZipFile(script) as zin, zipfile.ZipFile(excel, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = EXCEL_SHEET.encode()
                zout.writestr(item, data)
        assert cache_state(excel) == {"formulas": 4, "no_cache": 0}
        values = openpyxl.load_workbook(excel, data_only=True).active
        formulas = openpyxl.load_workbook(excel).active
        assert [values[f"D{r}"].value for r in (2, 3, 4, 5)] == [6, 20, 1, 27]
        # openpyxl expands the shared formula for every child cell for us:
        assert [formulas[f"D{r}"].value for r in (2, 3, 4)] == ["=B2*C2", "=B3*C3", "=B4*C4"]
        print("excel file  : values", [values[f"D{r}"].value for r in (2, 3, 4, 5)],
              "formulas", [formulas[f"D{r}"].value for r in (2, 3, 4)])

        # What does expanding a shared formula mean? Shift every relative reference.
        assert Translator("=B2*C2", origin="D2").translate_formula("D4") == "=B4*C4"
        assert Translator("=$A$1+B2", origin="D2").translate_formula("E5") == "=$A$1+C5"

        # --- 3. The trap: load + save drops EVERY cached value --------------------------------
        resaved = os.path.join(tmp, "resaved.xlsx")
        openpyxl.load_workbook(excel).save(resaved)
        assert cache_state(resaved) == {"formulas": 4, "no_cache": 4}
        print("after openpyxl load+save:", cache_state(resaved), "(formulas kept, results gone)")

        # --- 4. Do not invent a calculator ----------------------------------------------------
        # It is tempting to evaluate =SUM(A1:A2) yourself. Do not: Excel has ~500 functions with
        # edge cases (text-to-number coercion, 1900 vs 1904 dates, locale list separators).
        # A tool that is occasionally wrong about numbers is worse than one that says
        # "not calculated - open it in Excel". Report the state; do not paper over it.

    print("MODULE OK")


if __name__ == "__main__":
    main()
