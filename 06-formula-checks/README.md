# Module 06: Formula Checks

## The Concept

An agent asked "is this budget right?" cannot recalculate, but it can find formulas that are
suspicious by shape. Four checks cover most real mistakes, and all four need the same ingredient:
a **tokenizer**, not a regular expression.

openpyxl ships one: `Tokenizer("=SUM(A1:B2)+\"A9\"").items` gives typed tokens (function open,
range operand, number operand, text operand, infix operator...). Text inside quotes is a TEXT
token, so `"A9"` is never mistaken for a reference. That alone removes the biggest class of
false positives a regex approach has.

### 1. Broken references

A token `#REF!`, or a sheet prefix naming a sheet that does not exist, is a formula that will show
an error the moment Excel opens it.

### 2. R1C1 patterns: the formula that does not match its neighbours

Fill-down makes `=A2*2`, `=A3*2`, `=A4*2` look different in A1 notation but identical in R1C1
("the cell one to my left, times two"). Converting each formula to R1C1 (`=RC[-1]*2`) turns "are
these the same formula?" into a string comparison. A cell whose pattern differs from the cell
above and the cell below, which agree with each other, is the classic overwritten-formula bug.
The check is deliberately narrow (both neighbours must agree) so a `SUM` total row at the end of a
column is not reported.

### 3. Hardcoded numbers

`=A1*1.18` hides a VAT rate inside a formula. The rule: a number is suspicious when it sits next to
an arithmetic operator and is not a common constant. The `2` in `ROUND(x,2)` sits next to a
separator, not an operator, so it passes. The list of exemptions (0, 1, 12, 100...) is a judgement
call; report these as hints, not errors.

### 4. Circular references

Build a graph: each formula cell points at the cells it reads. A cycle in that graph is a circular
reference. `find_cycles` is Tarjan's strongly-connected-components algorithm written iteratively,
because a 50,000-long dependency chain would exceed Python's recursion limit. A cell that depends
on a cycle without being in one (D1 in the demo) is correctly not reported.

Be honest about what you cannot see. This graph follows direct references only. `INDIRECT`,
`OFFSET` and defined names compute addresses at run time, so cycles through them are invisible.
The finished tool says so in every report.

## Run It

```bash
python main.py
```

## Next

[Module 07](../07-mcp-wrapper-and-testing/): put the pieces behind an MCP server and test it the
way a real client would use it.
