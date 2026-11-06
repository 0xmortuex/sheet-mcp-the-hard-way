# Break It: The Inventory

## Challenge 1: A Drawing That Is Only a Picture

Add `xl/drawings/drawing1.xml` containing a text-box shape (`<xdr:sp>`) to a workbook. Does
`inventory` notice? Part names alone cannot tell a text box from a picture-only drawing.

**Hint:** The finished server reads small drawing parts and looks for shape elements with a regex.
Decide which is worse: a false alarm or a missed text box.

## Challenge 2: A Bomb That Lies

Build a zip whose header claims `file_size = 100` but whose data inflates to 100 MB. Python's
`zipfile` stops at the claimed size. What does `safe_check` conclude, and what would a *streaming*
check that counts real bytes do differently?

**Hint:** Python raises or truncates depending on version; the robust rule is "verify the number of
bytes you actually read against the header, and refuse on mismatch".

## Challenge 3: UTF-16 XML

Re-encode `sharedStrings.xml` as UTF-16 and keep the DOCTYPE. Does the byte search for
`<!doctype` still find it?

**Hint:** No: in UTF-16 each character has a zero byte after it. That is why the finished server
refuses XML that is not UTF-8 instead of trying to decode every encoding.
