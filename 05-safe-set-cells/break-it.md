# Break It: set_cells

## Challenge 1: The Bool Trap

Swap the order of the `bool` and `int` checks in `coerce` and call it with `True`.

**Hint:** `isinstance(True, int)` is `True`, so the boolean is stored as the number 1.

## Challenge 2: Two Changes, One Cell

Send `[{"cell": "A1", ...}, {"cell": "a1", ...}]`. What does the plan show, and which wins?

**Hint:** Last write wins silently. The finished server rejects the same cell twice in one call,
after normalising `a1` to `A1`.

## Challenge 3: Skip the Verification

Comment out the formula-count check before `os.replace`, then arrange for `wb.save` to lose a
formula (for example by editing a workbook with an array formula in a way that drops it).

**Hint:** You want a check that can only fail when something is wrong, and that runs on the
bytes you are about to publish, not on the in-memory object you built them from.
