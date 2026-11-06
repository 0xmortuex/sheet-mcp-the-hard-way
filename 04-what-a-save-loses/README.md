# Module 04: What a Save Loses

## The Concept

openpyxl's contract is: read the parts I understand, write those back. Parts it does not model
are not copied. Nobody tells you. A user asks an agent to "fix the total in B7" and gets back a
workbook with the same numbers and no chart, no pivot table and no macros.

The defence has three layers, in this order.

**1. Inventory before you touch anything.** The zip's part names prove what is inside:
`xl/charts/` means charts, `xl/pivotCache/` means pivot tables, `xl/vbaProject.bin` means macros,
`xl/slicers/`, `xl/externalLinks/`, `xl/threadedComments/`, `xl/embeddings/`, `xl/metadata.xml`
(dynamic-array formulas)... `inventory()` is a dictionary lookup over `namelist()`. It does not
depend on what a library chooses to load.

Some features hide *inside* a part instead. Sparklines and some conditional formats live in a
sheet's `<extLst>` block, recognisable by a GUID. `main.py` finds sparklines by that GUID, then
proves the loss is real: it round-trips the file through openpyxl and the sparklines are gone.
openpyxl does emit a warning for some of these ("...will be removed"), which is worth catching as
a second net, but it does not warn for all of them.

**2. Refuse, with a way out.** `refuse_if_lossy` raises unless the caller passes `allow_loss=True`.
The message names what would be lost, so the agent can tell the user. The default is safe; the
override is explicit and per call. Never make "lose things" a silent default, and never make
"allow loss" a server-wide setting an agent can flip without a human noticing.

**3. Never write over the original.** Even with `allow_loss`, the result goes to a new file. The
user keeps the intact original.

## The Container Is Untrusted Too

An xlsx comes from email and downloads. Before parsing, `safe_check` enforces:

- **Zip bombs.** 64 MB of zeros compresses to about 64 KB. The check compares each entry's
  claimed size to its compressed size (a ratio over 500:1 is not a spreadsheet), and caps entry
  count and total size.
- **XML entity attacks.** A `<!DOCTYPE>` with nested `<!ENTITY>` definitions can expand a few
  hundred bytes into gigabytes (the "billion laughs"). Excel never writes a DOCTYPE, so a file
  that has one is refused outright, before any parser sees it.

The finished server does these checks while streaming each part, counts the bytes that really
come out (a header can lie about sizes), and also refuses non-UTF-8 XML, duplicate entries and
`..` in entry names.

## Run It

```bash
python main.py
```

Every assertion in it is a claim made above, checked for real.

## Next

[Module 05](../05-safe-set-cells/): with a safe file in hand, write cells without damaging
anything else.
