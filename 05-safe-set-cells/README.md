# Module 05: A Safe set_cells

## The Concept

`set_cells` looks trivial: `ws["B2"] = 9`. A tool an AI can use unsupervised needs five more
things around that line.

### 1. Explicit coercion

The agent sends JSON. What does `"007"` mean? `"=SUM(A1:A3)"`? `"2024-12-31"`? If the tool
guesses, it will guess wrong for someone. So the rules are written down and printed back in the
dry run:

| Input | Stored as |
|---|---|
| `"=SUM(A1:A3)"` | formula |
| `"'=literal"` | the text `=literal` (one leading apostrophe removed, like typing it in Excel) |
| `5`, `2.5` | number; `true` is a boolean (checked **before** int, because `True` is an `int` in Python) |
| `"2030-01-31"` | a real date, only if it matches a strict ISO pattern |
| `"007"`, `"123"` | text. Never silently converted to a number |
| `null` | clears the cell |

Text needs one extra step: openpyxl treats a string starting with `=` as a formula and `#REF!` as
an error. `cell.data_type = "s"` after assignment forces it to stay text.

### 2. A dry run by default

The first call shows `B2: 2 -> number 9` and writes nothing. For a formula cell it also shows
`'=B3*C3' -> number 5`: the agent (and the human reading its transcript) sees that a formula is
about to become a constant. Making the safe action the default means a confused agent can only
do harm by asking explicitly.

### 3. Guards

Some cells must not be written: the interior of a merged range (only its top-left cell holds a
value) and any cell in an array formula's range (writing one cell breaks the array). The guard
raises with a message saying which cell to use instead.

### 4. A new file, written atomically

The result goes to `book.edited.xlsx`; the original's bytes are never touched (the demo asserts
this). The write goes to a temporary file and `os.replace` moves it into place, so a crash cannot
leave half a workbook. Before replacing anything, the tool re-opens the temp file and checks that
the number of formulas matches what it expects. A silent loss of formulas aborts the save.

### 5. Proof that nothing else changed

The demo takes a snapshot of every cell (value, type, number format, font, alignment) before and
after, and asserts that exactly the four edited cells differ. That is the test that justifies the
claim "untouched cells stay identical". The finished sheet-mcp does this for validations,
conditional formats, merged cells, column widths and defined names too.

## Run It

```bash
python main.py
```

## Next

[Module 06](../06-formula-checks/): finding the formulas that are already wrong.
