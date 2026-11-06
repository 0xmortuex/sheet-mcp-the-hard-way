"""Module 07: wrap it as an MCP server and test it over real stdio.

Two roles in one file (same pattern as the finished sheet-mcp):

    python main.py --serve --dir DIR   run the MCP server (a client starts this as a subprocess)
    python main.py                     demo/test: builds a workbook, launches itself with --serve,
                                       connects as an MCP client and checks everything end to end

The tools are deliberately tiny versions of sheet-mcp's: they reuse the ideas of modules 01-06
(zip inventory, cached-value state, dry-run set_cells with a new output file) in ~100 lines.
"""

from __future__ import annotations

import argparse
import asyncio
import functools
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

mcp = MCPServer(
    "mini-sheet",
    instructions=(
        "Read and edit .xlsx files. Start with workbook_overview. set_cells is a dry run unless "
        "dry_run=false and always writes a NEW file."
    ),
)

ROOT = Path.cwd()   # set from --dir in main()
LOSSY = {"xl/charts/": "charts", "xl/pivottables/": "pivot tables", "xl/vbaproject.bin": "macros",
         "xl/media/": "images", "xl/externallinks/": "external links"}


# ------------------------------------------------------------------ error handling (Module 05)
class SheetError(Exception):
    """Expected, explainable failure."""


def tool(fn: Any) -> Any:
    """Register fn as an MCP tool, turning SheetError into ToolError.

    mcp 2.x shows the client ONLY the message of a ToolError. Any other exception becomes a
    bare 'Error executing tool X', so expected failures must be re-raised as ToolError.
    """
    @functools.wraps(fn)    # keeps the signature + docstring: the SDK builds the tool schema from them
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except SheetError as e:
            raise ToolError(str(e)) from e

    mcp.tool()(wrapper)
    return fn


def resolve(path: str) -> Path:
    """The sandbox: every path must stay inside ROOT."""
    p = (ROOT / path).resolve()
    if not p.is_relative_to(ROOT):
        raise SheetError(f"{path!r} is outside the allowed directory {ROOT}")
    if not p.is_file():
        raise SheetError(f"{path!r} not found under {ROOT}")
    return p


def inventory(p: Path) -> dict[str, str]:
    with zipfile.ZipFile(p) as z:
        return {label: name for name in z.namelist() for prefix, label in LOSSY.items() if name.lower().startswith(prefix)}


# ------------------------------------------------------------------ the tools
@tool
def workbook_overview(path: str) -> str:
    """Sheets, formula counts, formulas without cached values, and what a save would lose."""
    p = resolve(path)
    lost = inventory(p)
    wb = openpyxl.load_workbook(p)
    lines = [f"{p.name}:"]
    for ws in wb.worksheets:
        n = sum(1 for c in ws._cells.values() if c.data_type == "f")
        lines.append(f"  sheet {ws.title!r}: {ws.max_row} rows x {ws.max_column} cols, {n} formulas")
    with zipfile.ZipFile(p) as z:
        xml = "".join(z.read(n).decode("utf-8") for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
    uncached = len(re.findall(r"</f><v\s*/>|</f><v></v>", xml))
    if uncached:
        lines.append(f"  {uncached} formula(s) have no cached value (never calculated)")
    lines.append("  would LOSE on save: " + (", ".join(sorted(lost)) if lost else "nothing"))
    return "\n".join(lines)


@tool
def read_range(path: str, sheet: str, range: str, mode: str = "values") -> str:
    """Read a range as 'A1: value' lines. mode: values (cached results) or formulas."""
    if mode not in ("values", "formulas"):
        raise SheetError("mode must be 'values' or 'formulas'")
    wb = openpyxl.load_workbook(resolve(path), data_only=mode == "values")
    if sheet not in wb.sheetnames:
        raise SheetError(f"no sheet {sheet!r}. Sheets: {', '.join(wb.sheetnames)}")
    out = []
    cells = wb[sheet][range]
    if not isinstance(cells, tuple):          # ws["A3"] is one Cell; ws["A1:A3"] is a tuple of rows
        cells = ((cells,),)
    for row in cells:
        for c in row:
            out.append(f"{c.coordinate}: {c.value if c.value is not None else '(empty or not calculated)'}")
    return "\n".join(out)


@tool
def set_cells(path: str, sheet: str, changes: list[dict[str, Any]], dry_run: bool = True) -> str:
    """Set cells ({"cell": "B2", "value": ...}; '=...' is a formula). Dry run by default;
    a real run writes <name>.edited.xlsx and never touches the original."""
    p = resolve(path)
    if inventory(p):
        raise SheetError("Refusing to edit: saving would lose " + ", ".join(sorted(inventory(p))))
    out = p.with_name(p.stem + ".edited.xlsx")
    if out.exists() and not dry_run:
        raise SheetError(f"{out.name} already exists; remove it first")
    wb = openpyxl.load_workbook(p)
    if sheet not in wb.sheetnames:
        raise SheetError(f"no sheet {sheet!r}. Sheets: {', '.join(wb.sheetnames)}")
    ws = wb[sheet]
    lines = []
    for ch in changes:
        cell = ws[ch["cell"]]
        lines.append(f"{cell.coordinate}: {cell.value!r} -> {ch['value']!r}")
        if not dry_run:
            cell.value = ch["value"]
    if dry_run:
        return "DRY RUN, nothing written\n" + "\n".join(lines)
    wb.save(out)
    return f"Saved {out.name}\n" + "\n".join(lines)


# ------------------------------------------------------------------ the demo / test client
async def demo() -> None:
    from mcp import ClientSession
    from mcp.client.stdio import StdioServerParameters, stdio_client

    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp).resolve()
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "S"
        ws["A1"], ws["A2"], ws["A3"] = 10, 20, "=SUM(A1:A2)"
        wb.save(root / "book.xlsx")
        # a workbook that must be protected: fake chart part
        with zipfile.ZipFile(root / "book.xlsx") as zin, zipfile.ZipFile(root / "chart.xlsx", "w") as zout:
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
            zout.writestr("xl/charts/chart1.xml", "<c/>")

        params = StdioServerParameters(
            command=sys.executable, args=[os.path.abspath(__file__), "--serve", "--dir", str(root)],
            env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        )
        async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
            init = await session.initialize()
            assert init.server_info.name == "mini-sheet"
            names = {t.name for t in (await session.list_tools()).tools}
            assert names == {"workbook_overview", "read_range", "set_cells"}, names
            print("tools:", sorted(names))

            async def call(name: str, args: dict[str, Any]) -> tuple[bool, str]:
                res = await session.call_tool(name, args)
                return bool(res.is_error), res.content[0].text  # type: ignore[union-attr]

            err, text = await call("workbook_overview", {"path": "book.xlsx"})
            assert not err and "1 formulas" in text and "never calculated" in text and "LOSE on save: nothing" in text, text
            print(text)

            err, text = await call("read_range", {"path": "book.xlsx", "sheet": "S", "range": "A1:A3"})
            assert "A3: (empty or not calculated)" in text      # honest about the missing cache
            err, text = await call("read_range", {"path": "book.xlsx", "sheet": "S", "range": "A3", "mode": "formulas"})
            assert "=SUM(A1:A2)" in text

            err, text = await call("set_cells", {"path": "book.xlsx", "sheet": "S", "changes": [{"cell": "A1", "value": 99}]})
            assert not err and text.startswith("DRY RUN") and not (root / "book.edited.xlsx").exists()
            err, text = await call("set_cells", {"path": "book.xlsx", "sheet": "S", "changes": [{"cell": "A1", "value": 99}], "dry_run": False})
            assert not err and openpyxl.load_workbook(root / "book.edited.xlsx")["S"]["A1"].value == 99
            assert openpyxl.load_workbook(root / "book.xlsx")["S"]["A1"].value == 10      # original untouched
            print("edit wrote book.edited.xlsx; book.xlsx unchanged")

            # errors must reach the agent WITH their message
            err, text = await call("set_cells", {"path": "chart.xlsx", "sheet": "S", "changes": [{"cell": "A1", "value": 1}], "dry_run": False})
            assert err and "would lose charts" in text, text
            err, text = await call("read_range", {"path": "../etc/passwd", "sheet": "S", "range": "A1"})
            assert err and "outside the allowed directory" in text, text
            err, text = await call("read_range", {"path": "book.xlsx", "sheet": "Nope", "range": "A1"})
            assert err and "Sheets: S" in text, text
            print("errors carry their messages:", text)
    print("MODULE OK")


def main() -> None:
    global ROOT
    ap = argparse.ArgumentParser()
    ap.add_argument("--serve", action="store_true")
    ap.add_argument("--dir", default=".")
    args = ap.parse_args()
    if args.serve:
        ROOT = Path(args.dir).resolve()
        mcp.run()          # stdio: stdout is the protocol channel, never print() here
    else:
        asyncio.run(demo())


if __name__ == "__main__":
    main()
