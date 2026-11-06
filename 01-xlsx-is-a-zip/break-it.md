# Break It: The Zip

## Challenge 1: Rename a Shared String

Change `<si><t>Çay</t></si>` in `sharedStrings.xml` to `Kahve`. Which cells change? Now change the
cell `A2` to `t="s"` with `<v>7</v>`. What does `read_cells` do, and what would Excel do?

**Hint:** Index 7 does not exist. Your reader raises `IndexError`; Excel reports the file as
corrupt. A reader that must survive other people's files has to turn that into a clear error.

## Challenge 2: Delete `[Content_Types].xml`

Remove that part and open the result with openpyxl.

**Hint:** The part list is not decorative. Without the content types, a consumer cannot tell a
worksheet from a chart. Check what error you get, then ask what your server should say.

## Challenge 3: A Cell Without a Value

Add `<c r="C5"/>` and `<c r="C6" s="3"/>` to the sheet. Does `read_cells` report them? Should a
styled but empty cell count towards "the used range"?

**Hint:** Excel writes empty styled cells all the time (a formatted column). The finished server
only counts cells with a value, a formula or an inline string.
