# Module 01: An xlsx Is a Zip of XML

## The Concept

Every tool that edits spreadsheets for an AI has one job that matters more than any other: not
to damage the file. To protect a file you have to know what it is made of, and `.xlsx` is not a
mystery format. Rename `book.xlsx` to `book.zip`, open it, and you see this:

```
[Content_Types].xml          what kind of data each part is
_rels/.rels                  "the workbook is xl/workbook.xml"
xl/workbook.xml              the list of sheets (names, order, hidden or not)
xl/_rels/workbook.xml.rels   which part holds each sheet
xl/worksheets/sheet1.xml     the cells of one sheet
xl/sharedStrings.xml         every distinct text string, stored once
xl/styles.xml                fonts, fills, number formats
xl/charts/chart1.xml         a chart (if there is one)
xl/vbaProject.bin            macros (if there are any)
```

Three things about this layout drive the rest of the course.

**Parts are separate files, linked by name.** `workbook.xml` does not contain sheet 1; it says
"a sheet called Satışlar lives at relationship `rId1`", and `workbook.xml.rels` says `rId1` is
`worksheets/sheet1.xml`. A library that does not understand a part simply does not carry it along.
That is how charts vanish (Module 04).

**A cell holds a value and, optionally, a formula.** `<c r="B3"><f>B2*2</f><v>10</v></c>` means
"B3 contains the formula `B2*2`, and the last program to calculate it got 10". The `<v>` is a
cache. Nothing keeps it fresh (Module 03).

**Text is not in the cell.** `<c r="A2" t="s"><v>2</v></c>` means "the text is item number 2 of
the shared strings table". Number cells have no `t` attribute.

## Run It

```bash
python main.py
```

It builds a six-part workbook by hand with `zipfile`, reads it back with `xml.etree` by following
the chain `workbook.xml` -> `.rels` -> `sheet1.xml`, edits a number by rewriting one part of the
zip, and finally asks openpyxl whether it considers the hand-made file valid.

Look at the second-to-last output line: after the hand edit `B2` is 7, yet `B3` still says 10.
The formula was not recalculated, because nothing here can calculate. Keep that picture in your
head: **a formula's cached value can be stale**, and a tool that edits cells has to be honest
about it.

## Why This Matters For An MCP Server

- You can inspect a file without opening it with a spreadsheet library: list the parts, find a
  `vbaProject.bin`, count formulas by streaming `sheet1.xml`. That is cheap and does not depend on
  what any library chooses to load.
- A file you did not create is untrusted input. A zip can be a bomb, and XML can carry entity
  tricks. The container is where you check for both (Module 04).
- When the server says "this edit would lose your pivot table", it is reading part names out of this
  zip. There is no magic.

## Next

[Module 02](../02-openpyxl-basics/) uses openpyxl to read and write these parts for you, and shows
what each way of opening a workbook costs.
