"""Module 05: a safe set_cells - explicit coercion, dry run, guards, new file, atomic write.

    python main.py
"""

from __future__ import annotations

import datetime as dt
import math
import os
import re
import tempfile
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.formula import ArrayFormula


class EditError(Exception):
    """An expected failure with a message that tells the caller what to do."""


# -------------------------------------------------------------------------------------------
# 1. Coercion: the agent sends JSON; the rules for what each JSON value MEANS must be explicit.
# -------------------------------------------------------------------------------------------
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def coerce(raw: Any) -> tuple[str, Any]:
    if raw is None:
        return "empty", None
    if isinstance(raw, bool):                       # bool before int: True is an int in Python
        return "bool", raw
    if isinstance(raw, (int, float)):
        if isinstance(raw, float) and not math.isfinite(raw):
            raise EditError("NaN/infinity cannot be stored in a cell")
        return "number", raw
    if not isinstance(raw, str):
        raise EditError(f"unsupported value {raw!r}; use string, number, boolean or null")
    if raw.startswith("="):
        return "formula", raw                       # '=SUM(A1:A3)' is a formula
    if raw.startswith("'"):
        return "text", raw[1:]                      # "'=x" is the literal text =x
    if ISO_DATE.match(raw):
        try:
            return "date", dt.date.fromisoformat(raw)
        except ValueError:
            pass
    return "text", raw                              # everything else stays text - never auto-number


def write(cell: Any, kind: str, value: Any) -> None:
    cell.value = value
    if kind == "text":
        cell.data_type = "s"                        # openpyxl would turn "=x" / "#REF!" into formula/error
    if kind == "date" and cell.number_format == "General":
        cell.number_format = "yyyy-mm-dd"           # keep an existing date format if there is one


# -------------------------------------------------------------------------------------------
# 2. Guards: places where a write would corrupt the sheet.
# -------------------------------------------------------------------------------------------
def guard(ws: Any, row: int, col: int) -> None:
    for m in ws.merged_cells.ranges:
        if (row, col) != (m.min_row, m.min_col) and m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            raise EditError(f"{ws.title}!R{row}C{col} is inside merged range {m.coord}; write to its top-left cell")
    for ref in ws.array_formulae.values():
        r = CellRange(ref)
        if r.min_row <= row <= r.max_row and r.min_col <= col <= r.max_col:
            raise EditError(f"{ws.title}!R{row}C{col} is part of array formula {ref}")


# -------------------------------------------------------------------------------------------
# 3. The tool: plan first, apply only when dry_run is False, write a NEW file atomically.
# -------------------------------------------------------------------------------------------
def count_formulas(wb: Any) -> int:
    return sum(1 for ws in wb.worksheets for c in ws._cells.values() if c.data_type == "f")


def set_cells(path: str, sheet: str, changes: list[dict[str, Any]], dry_run: bool = True,
              output: str | None = None, overwrite: bool = False) -> str:
    out = output or path.replace(".xlsx", ".edited.xlsx")
    if os.path.exists(out) and not overwrite:
        raise EditError(f"{os.path.basename(out)} already exists; pass overwrite=True")
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet]
    plan = []
    for ch in changes:
        cell = ws[ch["cell"]]
        guard(ws, cell.row, cell.column)
        kind, value = coerce(ch["value"])
        plan.append((cell, kind, value, cell.value))
    lines = [f"{c.coordinate}: {before!r} -> {kind} {value!r}" for c, kind, value, before in plan]
    if dry_run:
        return "DRY RUN, nothing written\n" + "\n".join(lines)
    for c, kind, value, _ in plan:
        write(c, kind, value)
    expected = count_formulas(wb)
    tmp = out[:-5] + ".tmp.xlsx"                    # openpyxl insists on the .xlsx extension
    wb.save(tmp)
    # Verify what we wrote before we let it replace anything: formulas must not have been lost.
    if count_formulas(openpyxl.load_workbook(tmp)) != expected:
        os.remove(tmp)
        raise EditError("saved file lost formulas; nothing was written")
    os.replace(tmp, out)                            # atomic on one filesystem: all or nothing
    return "Saved " + os.path.basename(out) + "\n" + "\n".join(lines)


def snapshot(path: str, sheet: str) -> dict[str, tuple[Any, ...]]:
    ws = openpyxl.load_workbook(path)[sheet]
    def plain(v: Any) -> Any:                        # ArrayFormula objects compare by identity
        return ("array", v.ref, v.text) if isinstance(v, ArrayFormula) else v

    return {c.coordinate: (plain(c.value), c.data_type, c.number_format, repr(c.font), repr(c.alignment))
            for row in ws.iter_rows() for c in row if c.value is not None or c.has_style}


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "book.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "S"
        ws.append(["name", "qty", "price", "total", "date"])
        ws.append(["tea", 2, 3.5, "=B2*C2", dt.date(2024, 1, 2)])
        ws.append(["çay", 4, 1.0, "=B3*C3", dt.date(2024, 1, 3)])
        ws["A1"].font = Font(bold=True)
        ws["B2"].alignment = Alignment(horizontal="center")
        ws["E2"].number_format = "dd/mm/yyyy"
        ws.merge_cells("G1:H1")
        ws["J1"] = ArrayFormula("J1:J2", "=B2:B3*2")
        wb.save(src)
        original = open(src, "rb").read()
        before = snapshot(src, "S")

        # dry run: shows the plan, writes nothing
        plan = set_cells(src, "S", [{"cell": "B2", "value": 9}, {"cell": "D3", "value": 5}])
        assert plan.startswith("DRY RUN") and "'=B3*C3' -> number 5" in plan
        assert os.listdir(tmp) == ["book.xlsx"]
        print(plan)

        # real run: new file, original untouched, formulas still formulas, untouched cells identical
        changes = [
            {"cell": "B2", "value": 9},
            {"cell": "A3", "value": "'=literal"},     # literal text that starts with '='
            {"cell": "F2", "value": "2030-01-31"},    # ISO date -> real date
            {"cell": "F3", "value": "007"},           # strings stay strings
        ]
        print(set_cells(src, "S", changes, dry_run=False))
        assert open(src, "rb").read() == original
        out = os.path.join(tmp, "book.edited.xlsx")
        after = snapshot(out, "S")
        changed = {k for k in set(before) | set(after) if before.get(k) != after.get(k)}
        assert changed == {"B2", "A3", "F2", "F3"}, changed
        new = openpyxl.load_workbook(out)["S"]
        assert new["D2"].value == "=B2*C2" and new["A3"].value == "=literal" and new["A3"].data_type == "s"
        assert new["F2"].value == dt.datetime(2030, 1, 31) and new["F3"].value == "007"
        assert new["B2"].alignment.horizontal == "center", "style of the edited cell survives"
        print("changed cells:", sorted(changed), "- everything else identical")

        # refusals
        for bad, why in [
            ({"cell": "H1", "value": 1}, "merged"),
            ({"cell": "J2", "value": 1}, "array"),
            ({"cell": "A1", "value": {"x": 1}}, "unsupported"),
            ({"cell": "A1", "value": float("nan")}, "NaN"),
        ]:
            try:
                set_cells(src, "S", [bad], output=os.path.join(tmp, "probe.xlsx"))
            except EditError as e:
                assert why in str(e), (why, str(e))
            else:
                raise AssertionError(why)
        try:
            set_cells(src, "S", changes, dry_run=False)           # output already exists
        except EditError as e:
            assert "already exists" in str(e)
        assert sorted(os.listdir(tmp)) == ["book.edited.xlsx", "book.xlsx"], "no temp files left behind"
        print("guards and overwrite protection work")

    print("MODULE OK")


if __name__ == "__main__":
    main()
