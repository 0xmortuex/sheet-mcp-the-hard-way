# Break It: The Wrapper

## Challenge 1: Print to Stdout

Add `print("reading", path)` at the top of `read_range`, run the demo, then change it to
`print(..., file=sys.stderr)`.

**Hint:** stdout carries the protocol. The first form injects a non-JSON line into the stream.

## Challenge 2: Raise a Plain Exception

Change one `raise SheetError(...)` inside `set_cells` to `raise ValueError(...)`. What text does the
client receive? Compare with the `SheetError` version.

**Hint:** The client gets only `Error executing tool set_cells`. This is the reason the tool
wrapper exists.

## Challenge 3: Escape the Sandbox With a Symlink

Create a symlink inside the directory that points outside it and call `workbook_overview` on it
(on Windows this needs developer mode or an elevated shell).

**Hint:** `resolve()` follows the link, and the resulting path fails the `is_relative_to` check.
Try the same with `os.path.abspath`, which does not follow links, to see why the order matters.
