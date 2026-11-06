# Break It: openpyxl

## Challenge 1: Leave a Read-Only Workbook Open

Remove `wbr.close()`, then try to `os.remove` the file at the end of the `with` block (on Windows).

**Hint:** The zip handle is still open. On Windows the delete fails; on Linux it succeeds and
leaks a file descriptor.

## Challenge 2: Probe Many Cells

In normal mode run `for r in range(1, 100000): ws.cell(r, 1)` and check `ws.max_row` before and
after, then save.

**Hint:** `ws.cell(...)` creates the cell. The sheet now "uses" 100,000 rows, which changes what a
tool reports as the used range. This is why the finished server computes the used range from the
XML, not from openpyxl.

## Challenge 3: Write a Date as a String

Put `"2024-03-02"` (a string) in a cell and save, then reload. Compare with `dt.date(2024, 3, 2)`.

**Hint:** The string stays text; Excel will not sort or compute with it. Module 05 turns ISO
date strings into real dates, deliberately and only for strings that match a strict pattern.
