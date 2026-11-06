"""Module 06: checking formulas without calculating them.

Four checks, each built on the same trick: tokenize the formula instead of regex-ing it.
  1. broken references (#REF!, missing sheet)
  2. R1C1 patterns: a formula that differs from both neighbours is probably a mistake
  3. hardcoded numbers inside arithmetic
  4. circular references (cycle detection on the dependency graph)

    python main.py
"""

from __future__ import annotations

import re
import tempfile
import os
from collections import defaultdict

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter, column_index_from_string


def tokens(formula: str) -> list[tuple[str, str, str]]:
    return [(t.value, t.type, t.subtype) for t in Tokenizer(formula).items]


# ---- 1. why a tokenizer: text inside quotes must not look like a reference --------------------
def find_refs(formula: str) -> list[str]:
    return [v for v, typ, sub in tokens(formula) if typ == "OPERAND" and sub == "RANGE"]


# ---- 2. R1C1: the same formula filled down is the same pattern ---------------------------------
CELL = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")


def to_r1c1(formula: str, row: int, col: int) -> str:
    out = []
    for value, typ, sub in tokens(formula):
        if typ == "OPERAND" and sub == "RANGE":
            prefix, _, rest = value.rpartition("!")
            pieces = []
            for piece in rest.split(":"):
                m = CELL.match(piece)
                if not m:
                    pieces.append(piece)
                    continue
                c_abs, letters, r_abs, digits = m.groups()
                r, c = int(digits), column_index_from_string(letters)
                rr = f"R{r}" if r_abs else ("R" if r == row else f"R[{r - row}]")
                cc = f"C{c}" if c_abs else ("C" if c == col else f"C[{c - col}]")
                pieces.append(rr + cc)
            value = (prefix + "!" if prefix else "") + ":".join(pieces)
        if typ != "WHITE-SPACE":
            out.append(value)
    return "=" + "".join(out)


def pattern_outliers(column: dict[int, str]) -> list[int]:
    """Rows whose pattern differs from the row above AND the row below, which share one pattern."""
    return [r for r, p in column.items()
            if column.get(r - 1) is not None and column.get(r - 1) == column.get(r + 1) != p]


# ---- 3. hardcoded numbers ------------------------------------------------------------------------
COMMON = {0.0, 1.0, 2.0, 12.0, 100.0}


def hardcoded_numbers(formula: str) -> list[str]:
    sig = [t for t in tokens(formula) if t[1] != "WHITE-SPACE"]
    found = []
    for i, (value, typ, sub) in enumerate(sig):
        if typ == "OPERAND" and sub == "NUMBER" and float(value) not in COMMON:
            neighbours = [sig[j] for j in (i - 1, i + 1) if 0 <= j < len(sig)]
            if any(n[1] == "OPERATOR-INFIX" and n[0] in "+-*/^" for n in neighbours):
                found.append(value)          # 1.18 in =A1*1.18 yes; the 2 in ROUND(x,2) no (a SEP)
    return found


# ---- 4. circular references: strongly connected components ----------------------------------------
def cell_deps(formula: str) -> list[tuple[int, int]]:
    deps = []
    for ref in find_refs(formula):
        if "!" in ref:
            continue                              # this teaching version checks one sheet
        for piece in ref.replace("$", "").split(":"):
            m = CELL.match(piece)
            if m:
                deps.append((int(m.group(4)), column_index_from_string(m.group(2))))
    return deps                                   # ranges: only the corners; the real tool expands them


def find_cycles(graph: dict[tuple[int, int], list[tuple[int, int]]]) -> list[list[tuple[int, int]]]:
    """Tarjan's algorithm, iterative so a 100k-long chain cannot hit the recursion limit."""
    index: dict = {}
    low: dict = {}
    on: set = set()
    stack: list = []
    cycles: list = []
    counter = 0
    for start in graph:
        if start in index:
            continue
        work = [(start, 0)]
        while work:
            node, i = work.pop()
            if i == 0:
                index[node] = low[node] = counter
                counter += 1
                stack.append(node)
                on.add(node)
            nbrs = graph.get(node, [])
            pushed = False
            while i < len(nbrs):
                w = nbrs[i]
                i += 1
                if w not in index:
                    work.append((node, i))
                    work.append((w, 0))
                    pushed = True
                    break
                if w in on:
                    low[node] = min(low[node], index[w])
            if pushed:
                continue
            if low[node] == index[node]:
                comp = []
                while True:
                    w = stack.pop()
                    on.discard(w)
                    comp.append(w)
                    if w == node:
                        break
                if len(comp) > 1 or node in graph.get(node, []):
                    cycles.append(sorted(comp))
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
    return cycles


def main() -> None:
    # 1 ----------------------------------------------------------------------------------------
    assert find_refs('=SUM(A1:B2)+"A9 and Sheet2!C3"+Sheet2!D4') == ["A1:B2", "Sheet2!D4"]
    print("refs ignore text inside quotes:", find_refs('=SUM(A1:B2)+"A9"+Sheet2!D4'))

    # 2 ----------------------------------------------------------------------------------------
    assert to_r1c1("=A2*2", 2, 2) == to_r1c1("=A3*2", 3, 2) == "=RC[-1]*2"
    assert to_r1c1("=$A$1+B2", 5, 3) == "=R1C1+R[-3]C[-1]"
    col = {r: to_r1c1(f, r, 2) for r, f in {2: "=A2*2", 3: "=A3*2", 4: "=A4*3", 5: "=A5*2", 6: "=A6*2"}.items()}
    assert pattern_outliers(col) == [4]
    print("pattern outlier rows:", pattern_outliers(col))

    # 3 ----------------------------------------------------------------------------------------
    assert hardcoded_numbers("=A1*1.18") == ["1.18"]
    assert hardcoded_numbers("=ROUND(A1*2,2)+1") == []
    assert hardcoded_numbers("=A1*100/12") == []
    print("hardcoded:", hardcoded_numbers("=A1*1.18+B1/7"))

    # 4 ----------------------------------------------------------------------------------------
    sheet = {"A1": "=B1+1", "B1": "=A1+1", "C1": "=C1*2", "D1": "=A1+E1", "E1": "=5"}
    graph = {}
    for addr, f in sheet.items():
        m = CELL.match(addr)
        assert m
        graph[(int(m.group(4)), column_index_from_string(m.group(2)))] = cell_deps(f)
    cycles = find_cycles(graph)
    names = [[f"{get_column_letter(c)}{r}" for r, c in comp] for comp in cycles]
    assert sorted(names) == [["A1", "B1"], ["C1"]], names
    print("cycles:", names, "(D1 depends on a cycle but is not in one)")

    # a long chain must not blow the recursion limit
    chain = {(i, 1): [(i + 1, 1)] for i in range(1, 50_000)}
    assert find_cycles(chain) == []

    # Finally, against a real workbook ----------------------------------------------------------
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "m.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        for r in range(1, 7):
            ws[f"A{r}"] = r
            ws[f"B{r}"] = f"=A{r}*2" if r != 4 else "=A4*3"
        wb.save(path)
        ws = openpyxl.load_workbook(path).active
        column = defaultdict(str)
        for r in range(1, 7):
            column[r] = to_r1c1(ws[f"B{r}"].value, r, 2)
        assert pattern_outliers(column) == [4]
        print("real workbook: B4 is the odd one out")

    print("MODULE OK")


if __name__ == "__main__":
    main()
