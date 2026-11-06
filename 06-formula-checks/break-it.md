# Break It: Formula Checks

## Challenge 1: The Total Row

Add `"=SUM(B2:B6)"` as the last row of the pattern column and run `pattern_outliers`. Why is it not
reported? Change the rule to "differs from the row above" and see how many false alarms a normal
sheet produces.

**Hint:** Totals, subtotals and headers are legitimate breaks. Requiring both neighbours to agree
is a cheap way to keep only the suspicious middle-of-the-column case.

## Challenge 2: A Reference In A String

Run `find_refs('=IF(A1>0,"see B2","")')`. Then replace the tokenizer with
`re.findall(r"[A-Z]+\d+", formula)` and compare.

**Hint:** The regex reports `B2`, a piece of text.

## Challenge 3: Cross-Sheet Cycles

Extend `cell_deps` to keep the sheet name (`Sheet2!A1`) in the node key and detect a cycle between
two sheets.

**Hint:** Nodes become `(sheet, row, col)`. The hard part is the prefix: `'My Sheet'!A1` is
quoted, and `Sheet1:Sheet3!A1` is a 3-D reference spanning several sheets.
