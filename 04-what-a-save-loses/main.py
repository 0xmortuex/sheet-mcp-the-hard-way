"""Module 04: what a save loses, and how to detect it BEFORE saving.

openpyxl reads the parts of a workbook it understands and writes those back. Everything else
vanishes silently. The defence is an inventory of the zip's part names, done up front, plus a
refusal. This module also makes the file itself safe to open (zip bombs, XML entities).

    python main.py
"""

from __future__ import annotations

import os
import tempfile
import warnings
import zipfile

import openpyxl

# ---------------------------------------------------------------------------------------------
# Part-name rules: a prefix of a part inside the zip -> the feature it proves is present.
# ---------------------------------------------------------------------------------------------
LOSSY_PREFIXES = {
    "xl/charts/": "charts",
    "xl/media/": "images",
    "xl/pivottables/": "pivot tables",
    "xl/pivotcache/": "pivot tables",
    "xl/slicers/": "slicers",
    "xl/externallinks/": "external links",
    "xl/threadedcomments/": "threaded comments",
    "xl/embeddings/": "embedded objects",
    "xl/vbaproject.bin": "macros (VBA)",
    "xl/metadata.xml": "dynamic-array metadata",
}
# Sheet XML can hide features in <extLst> blocks, identified by a GUID, not by a part name.
SPARKLINE_GUID = b"{05c60535-1f16-4fd2-b633-f4f36f0b64e0}"


def inventory(path: str) -> dict[str, list[str]]:
    found: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            low = name.lower()
            for prefix, label in LOSSY_PREFIXES.items():
                if low.startswith(prefix):
                    found.setdefault(label, []).append(name)
            if low.startswith("xl/worksheets/") and low.endswith(".xml"):
                if SPARKLINE_GUID in z.read(name).lower():
                    found.setdefault("sparklines", []).append(name)
    return found


def refuse_if_lossy(path: str, allow_loss: bool = False) -> None:
    lost = inventory(path)
    if lost and not allow_loss:
        raise RuntimeError(
            "Refusing to edit: saving would lose " + ", ".join(sorted(lost)) + ". Pass allow_loss=True to override."
        )


# ---------------------------------------------------------------------------------------------
# Making the container safe: an xlsx is attacker-controlled input.
# ---------------------------------------------------------------------------------------------
MAX_ENTRIES, MAX_TOTAL, MAX_RATIO = 10_000, 1 << 30, 500


def safe_check(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        if len(infos) > MAX_ENTRIES:
            raise ValueError("too many zip entries")
        if sum(i.file_size for i in infos) > MAX_TOTAL:
            raise ValueError("expands to more than 1 GiB")
        for i in infos:
            # file_size is what the header CLAIMS; ratio catches 64 MB of zeros stored in 64 KB.
            if i.file_size > 1 << 20 and i.file_size > MAX_RATIO * max(i.compress_size, 1):
                raise ValueError(f"{i.filename}: compression ratio too high (zip bomb?)")
            if i.filename.endswith((".xml", ".rels")):
                head = z.read(i.filename)
                # Excel never writes a DOCTYPE; its only use in a workbook is an entity attack.
                if b"<!doctype" in head.lower() or b"<!entity" in head.lower():
                    raise ValueError(f"{i.filename}: DOCTYPE/ENTITY declaration (billion laughs?)")


def make_base(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "data"
    ws["B1"] = "=LEN(A1)"
    wb.save(path)


def add_parts(src: str, dst: str, extra: dict[str, bytes], replace: dict[str, bytes] | None = None) -> None:
    replace = replace or {}
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item.filename, replace.get(item.filename, zin.read(item.filename)))
        for name, data in extra.items():
            zout.writestr(name, data)


def main() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        base = os.path.join(tmp, "base.xlsx")
        make_base(base)
        assert inventory(base) == {}
        refuse_if_lossy(base)                     # a plain workbook is fine
        print("plain workbook: nothing would be lost")

        # --- hand-zip parts that openpyxl never writes (we only need the NAMES to be there) ---
        fancy = os.path.join(tmp, "fancy.xlsx")
        add_parts(base, fancy, {
            "xl/charts/chart1.xml": b"<c/>",
            "xl/pivotCache/pivotCacheDefinition1.xml": b"<p/>",
            "xl/vbaProject.bin": b"\xd0\xcf\x11\xe0",
            "xl/externalLinks/externalLink1.xml": b"<e/>",
        })
        found = inventory(fancy)
        print("fancy workbook would lose:", sorted(found))
        assert set(found) == {"charts", "pivot tables", "macros (VBA)", "external links"}
        try:
            refuse_if_lossy(fancy)
        except RuntimeError as e:
            print("refused:", e)
        else:
            raise AssertionError("must refuse")
        refuse_if_lossy(fancy, allow_loss=True)    # the explicit opt-in

        # --- prove the loss is REAL for something hiding inside the sheet XML ------------------
        with zipfile.ZipFile(base) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode()
        spark = sheet.replace(
            "</worksheet>",
            '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" '
            'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
            "<x14:sparklineGroups/></ext></extLst></worksheet>",
        )
        sparkfile = os.path.join(tmp, "spark.xlsx")
        add_parts(base, sparkfile, {}, replace={"xl/worksheets/sheet1.xml": spark.encode()})
        assert "sparklines" in inventory(sparkfile)
        resaved = os.path.join(tmp, "resaved.xlsx")
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            openpyxl.load_workbook(sparkfile).save(resaved)   # the "harmless" round trip
        assert "sparklines" not in inventory(resaved), "openpyxl dropped the sparklines"
        # openpyxl DOES emit a warning for some losses - a second safety net, but only some:
        print("round trip through openpyxl: sparklines gone; warnings:", [str(w.message) for w in caught])
        assert any("will be removed" in str(w.message) for w in caught)

        # --- the container itself ---------------------------------------------------------------
        safe_check(base)
        bomb = os.path.join(tmp, "bomb.xlsx")
        with zipfile.ZipFile(bomb, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/junk.bin", b"\0" * (64 << 20))
        assert os.path.getsize(bomb) < 200_000
        laughs = os.path.join(tmp, "laughs.xlsx")
        add_parts(base, laughs, {"xl/sharedStrings.xml": (
            b'<?xml version="1.0"?><!DOCTYPE l [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;">]>'
            b"<sst>&b;</sst>")})
        for name in (bomb, laughs):
            try:
                safe_check(name)
            except ValueError as e:
                print("rejected", os.path.basename(name), "->", e)
            else:
                raise AssertionError(f"{name} must be rejected")

    print("MODULE OK")


if __name__ == "__main__":
    main()
