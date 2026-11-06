# Break It: Formulas and Values

## Challenge 1: The Empty-String Formula

Add `<c r="E9" t="str"><f>""</f><v></v></c>` to the sheet XML in `main.py` and run `cache_state`.
Is it counted as "no cache"? Remove the `t != "str"` condition and run again.

**Hint:** Without the exception, a formula that correctly evaluates to `""` is reported as
uncalculated. With it, a genuinely uncalculated cell with `t="str"` would be missed. Which error
is worse for an agent?

## Challenge 2: Expand a Shared Formula by Hand

Read the Excel-style sheet with `ElementTree` only and print the formula of `D4`. You will get
`None`. Implement the expansion with `Translator`, keeping the master's address in a dict by `si`.

**Hint:** Children come after the master in document order, so one pass is enough.

## Challenge 3: Stale Caches

Take the hand-edited file from Module 01 (B2 changed, B3's cache untouched). Write a check that
could notice the cache is stale without calculating anything.

**Hint:** You cannot know in general. This is why the finished server never claims a cached value is
*current*, only that it is *cached*.
