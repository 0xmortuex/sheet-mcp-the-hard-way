# Module 03: Formulas vs Values

## The Concept

A formula cell is two facts glued together:

```xml
<c r="D5"><f>SUM(D2:D4)</f><v>27</v></c>
        └ formula text  └ the result the last calculator got
```

Excel and LibreOffice write both. **openpyxl writes only the formula** and leaves `<v/>` empty,
because it cannot calculate. So a workbook produced by a script (or saved by an openpyxl-based
tool) has formulas and no results.

This creates three situations a tool has to tell apart:

1. **The cell is empty.** `data_only=True` gives `None`.
2. **The cell is a formula that was calculated.** `data_only=True` gives the number.
3. **The cell is a formula that was never calculated.** `data_only=True` also gives `None`.

Cases 1 and 3 look identical to a naive reader. If an AI is told "D5 is None", it will say "D5 is
empty", which is wrong and can lead to a wrong edit. The honest answer is "D5 is a formula and has
no cached value, because this file was never calculated". `cache_state` in `main.py` shows how to
know: stream the sheet XML and count `<f>` elements whose `<v>` is missing or empty.

One exception hides in there: a formula cell with `t="str"` and an empty `<v>` is a formula that
really did return the empty string. The finished server tests for it.

## Shared Formulas

Excel saves a column of identical formulas compactly: the first cell holds the text and `si="0"`;
the others hold only `si="0"` and a position. openpyxl expands them for you (`Translator` shifts
relative references), so `D3` reads as `=B3*C3`. If you read the XML yourself, you have to do that
expansion, or you will report formulas as empty. `main.py` shows both.

## The Trap: A Save Drops Every Cached Value

Load an Excel file with `load_workbook(p)` and save it: every `<v>` is gone (see the
"after openpyxl load+save" line). The formulas are intact and Excel recalculates when it opens
the file, but **everything that reads the file without calculating** (pandas, a preview pane, the
next call to your own tool) now sees nothing. The finished server says so in the reply to every
save: "Formulas are kept, but their cached results were dropped".

## Do Not Write A Calculator

It is tempting: `=SUM(A1:A2)` looks easy. But Excel has hundreds of functions with coercion
rules, locale-dependent separators, 1900 leap-year bugs and implicit intersection. A calculator
that is right 95% of the time produces confident wrong numbers, which is worse than "not
calculated". The finished sheet-mcp reports the cache state and points the user at Excel or
LibreOffice to refresh it.

## Run It

```bash
python main.py
```

## Next

[Module 04](../04-what-a-save-loses/): the cached values are only the first thing a save loses.
