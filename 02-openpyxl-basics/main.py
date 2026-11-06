"""Module 02: openpyxl basics - the three ways to open a workbook and what each one gives you.

    python main.py
"""

from __future__ import annotations

import datetime as dt
import os
import tempfile

import openpyxl
from openpyxl.styles import Font, PatternFill


def build(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sales"
    ws.append(["Item", "Qty", "Price", "Total", "Date"])
    for i in range(2, 6):
        ws.append([f"Item {i}", i, 2.5 * i, f"=B{i}*C{i}", dt.date(2024, 3, i)])
        ws[f"E{i}"].number_format = "dd/mm/yyyy"      # how it is DISPLAYED; the stored value is a number
    ws["A1"].font = Font(bold=True)
    ws["A1"].fill = PatternFill("solid", fgColor="FFFF00")
    ws.merge_cells("G1:H1")
    ws["G1"] = "merged"
    wb.save(path)


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "sales.xlsx")
        build(path)

        # 1. Normal mode: the whole workbook in memory as objects. Styles, merges, validations
        #    are all there, and so are formulas - as strings starting with "=".
        wb = openpyxl.load_workbook(path)
        ws = wb["Sales"]
        assert ws["D2"].value == "=B2*C2" and ws["D2"].data_type == "f"
        assert ws["E2"].value == dt.datetime(2024, 3, 2)          # a number + a date format = datetime
        assert ws["E2"].number_format == "dd/mm/yyyy"
        assert ws["A1"].font.bold is True
        assert [str(r) for r in ws.merged_cells.ranges] == ["G1:H1"]
        print("normal mode : formula", ws["D2"].value, "| date", ws["E2"].value.date(), "| bold", ws["A1"].font.bold)

        # 2. data_only=True: you get the values Excel CACHED last time it calculated, not formulas.
        #    This file was written by openpyxl, which cannot calculate, so nothing is cached.
        wbv = openpyxl.load_workbook(path, data_only=True)
        assert wbv["Sales"]["D2"].value is None
        print("data_only   : D2 =", wbv["Sales"]["D2"].value, "(never calculated - module 03 explains)")

        # 3. read_only=True: streams rows without building the object tree. Much less memory and
        #    faster to start, but no merged cells, no row/column dimensions, one pass only.
        wbr = openpyxl.load_workbook(path, read_only=True)
        rows = list(wbr["Sales"].iter_rows(min_row=1, max_row=2, values_only=True))
        assert rows[0][:3] == ("Item", "Qty", "Price")
        assert not hasattr(wbr["Sales"], "merged_cells")
        wbr.close()   # read-only workbooks hold the zip open; always close them
        print("read_only   :", rows[1][:4])

        # Cells are created on first access - looking is not free in normal mode.
        before = len(ws._cells)
        _ = ws["Z99"]
        assert len(ws._cells) == before + 1
        print("Touching ws['Z99'] created a cell:", before, "->", len(ws._cells))

        # Saving and reloading keeps what openpyxl understands...
        again = os.path.join(tmp, "again.xlsx")
        wb.save(again)
        wb2 = openpyxl.load_workbook(again)["Sales"]
        assert wb2["A1"].fill.fgColor.rgb == "00FFFF00" and wb2["D2"].value == "=B2*C2"
        print("round trip keeps style + formulas")

    print("MODULE OK")


if __name__ == "__main__":
    main()
