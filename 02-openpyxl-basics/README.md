# Module 02: openpyxl Basics

## The Concept

openpyxl is a pure-Python library that reads and writes the parts from Module 01. It is the
foundation of the finished sheet-mcp, so it pays to know exactly what it does and what it costs.

There are three ways to open a workbook.

| Mode | What you get | What you give up |
|---|---|---|
| `load_workbook(p)` | Everything in memory as objects: cells, styles, merges, validations, comments, tables | Memory (a 200,000-row sheet is hundreds of MB) and startup time |
| `load_workbook(p, data_only=True)` | The values Excel **cached**, instead of formulas | Formulas. And if the file was never calculated, you get `None` |
| `load_workbook(p, read_only=True)` | A row stream: fast, constant memory | Merged cells, row/column properties, random access, editing |

Three things in `main.py` are worth pausing on.

**Dates are numbers plus a format.** Excel stores 45353 and a number format `dd/mm/yyyy`;
openpyxl turns that into a `datetime`. The format is display only, so a tool that shows dates
should say what the format is (the finished server shows ISO 8601 and notes the format). Which
date system the file uses (1900 or 1904) shifts every date by four years and a day; openpyxl
reads that from `workbook.xml` for you.

**Cells appear when you look.** `ws["Z99"]` creates a cell. In normal mode, merely reading a
far-away address grows the sheet. The finished server therefore never probes cells it does not
have to, and reads big sheets through a streaming reader.

**A round trip keeps what openpyxl understands.** Styles, number formats, merges and formulas
survive `load` then `save`. The question for the next two modules is: *what about the rest?*

## Run It

```bash
python main.py
```

You will see the same workbook read three ways, plus a round trip that keeps the bold header
fill and the formula text.

## Design Decisions For The Server

- **Read with a streaming reader, edit with the full object model.** Reading a 200k-row sheet in
  normal mode is slow and big. Editing needs the object model because it has to write the file
  back. The finished server uses its own streaming reader for reads, and openpyxl normal mode for
  edits (with a size limit on what it will edit).
- **Always close read-only workbooks.** They hold the zip file open; on Windows that blocks
  deleting or replacing the file afterwards.

## Next

[Module 03](../03-formulas-vs-values/): the `D2 = None` you saw above is not a bug. It is the
central difficulty of working with spreadsheets from code.
