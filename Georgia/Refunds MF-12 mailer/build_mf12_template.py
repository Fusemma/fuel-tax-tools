"""Build the MF-12 template the refund macro fills.

The state's blank MF-12 is locked with an empty open password: filling the
form is allowed, changing the page is not. Excel's VBA can't unlock it, so
this script makes a filled-in-able copy once, at home:

- same page, same 146 fields, same scripts and buttons
- every box the macro fills gets a reserved slot in two places: the field's
  value, and a small drawing (the "appearance") that shows the value on the
  page. Slots are runs of underscores after a marker like MF12V:R05.GAL:
- the macro swaps each slot for real text padded with spaces to the exact
  same length, so nothing else in the file moves and the file stays valid

If Georgia ever revises the form: download the new blank, point BLANK at it,
run this again, and check the proof page it renders.

Usage: python build_mf12_template.py
"""
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject,
                           FloatObject, NameObject, PdfObject)


class RawString(PdfObject):
    """A value written into the file byte for byte. pypdf's own strings
    escape colons and underscores, which would hide the markers."""

    def __init__(self, text):
        self.data = text.encode("ascii")

    def write_to_stream(self, stream, encryption_key=None):
        stream.write(self.data)

HERE = Path(__file__).parent
BLANK = HERE / "MF-12 blank (state 2017 rev).pdf"
OUT = HERE / "mf12_template.pdf"

HEADER_KEYS = {
    "NAME OF SELLER": "SELLER",
    "SELLER'S LICENSE NUMBER": "LICENSE",
    "NAME OF PURCHASER": "BUYER",
    "ADDRESS": "ADDRESS",
    "CITY OR TOWN": "CITY",
    "STATE": "STATE",
    "ZIP CODE": "ZIP",
}
ROW_COLS = {  # first part of the field name -> column key
    "DATE": "DATE",
    "INVOICE NUMBER": "INV",
    "GALLONS": "GAL",
    "TAX PAID TO SELLER BY PURCHASER": "TAX",
    "EXEMPT PORTION": "EXEMPT",
    "TYPE OF FUEL": "FUEL",
}
# slot sizes in bytes: (value slot, drawing slot). Generous on purpose;
# the macro refuses a value that doesn't fit rather than cut it.
SLOT = {
    "SELLER": (150, 360), "BUYER": (150, 360), "ADDRESS": (150, 360),
    "CITY": (80, 200), "LICENSE": (60, 150), "STATE": (30, 100),
    "ZIP": (30, 100), "FUEL": (100, 220),
}
ROW_SLOT = (50, 160)


def full_name(w):
    parts, o = [], w
    while o is not None:
        if "/T" in o:
            parts.append(str(o["/T"]))
        o = o["/Parent"].get_object() if "/Parent" in o else None
    return ".".join(reversed(parts))


def main():
    reader = PdfReader(str(BLANK))
    if reader.is_encrypted:
        reader.decrypt("")
    info = dict(reader.metadata or {})
    writer = PdfWriter(clone_from=reader)
    page = writer.pages[0]
    font_cobo = writer._root_object["/AcroForm"]["/DR"]["/Font"].raw_get("/CoBo")

    widgets = [a.get_object() for a in page["/Annots"]]
    named = [(full_name(w), w) for w in widgets]

    # rows come from position, top row first, so field-name quirks can't
    # put a value on the wrong line
    date_ys = sorted({round(float(w["/Rect"][1]), 1)
                      for n, w in named if n.startswith("DATE.")}, reverse=True)
    assert len(date_ys) == 21, date_ys

    def row_of(w):
        y = float(w["/Rect"][1])
        best = min(range(21), key=lambda i: abs(date_ys[i] - y))
        assert abs(date_ys[best] - y) < 5, (y, date_ys[best])
        return best + 1

    done = {}
    for name, w in named:
        base = name.split(".")[0]
        if base in HEADER_KEYS:
            key = HEADER_KEYS[base]
            vslot, aslot = SLOT[key]
        elif base in ROW_COLS:
            col = ROW_COLS[base]
            key = "R%02d.%s" % (row_of(w), col)
            vslot, aslot = SLOT["FUEL"] if col == "FUEL" else ROW_SLOT
            # sanity: the number at the end of the field name should agree
            idx = int(name.split(".")[-1])
            if name == "TAX PAID TO SELLER BY PURCHASER.0.0.0":
                idx = 0
            assert idx + 1 == int(key[1:3]), (name, key)
        else:
            continue  # the Print and Clear buttons stay exactly as the state made them
        assert key not in done, key
        rect = [float(x) for x in w["/Rect"]]
        wd, ht = rect[2] - rect[0], rect[3] - rect[1]
        q = int(w.get("/Q", 0))

        vmark = "MF12V:%s:" % key
        w[NameObject("/V")] = RawString("(" + vmark + "_" * (vslot - 2 - len(vmark)) + ")")

        amark = "%%MF12A:%s:W=%06.2f:H=%06.2f:Q=%d:" % (key, wd, ht, q)
        body = amark + "_" * (aslot - len(amark))
        data = ("/Tx BMC\nq 1 1 %.2f %.2f re W n\n%s\nQ\nEMC\n"
                % (wd - 2, ht - 2, body)).encode("ascii")
        if key == "SELLER":
            # a comment line the macro stamps with a fingerprint of the rows on
            # this form, so a later run can tell if the report changed since
            data += b"%MF12ID:" + b"_" * 120 + b"\n"
        ap = DecodedStreamObject()
        ap.set_data(data)
        ap.update({
            NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Form"),
            NameObject("/BBox"): ArrayObject([FloatObject(0), FloatObject(0),
                                              FloatObject(round(wd, 3)), FloatObject(round(ht, 3))]),
            NameObject("/Resources"): DictionaryObject({
                NameObject("/Font"): DictionaryObject({NameObject("/CoBo"): font_cobo})}),
        })
        w[NameObject("/AP")] = DictionaryObject({NameObject("/N"): writer._add_object(ap)})
        done[key] = (name, round(wd, 2), round(ht, 2), q)

    assert len(done) == 7 + 21 * 6, len(done)
    af = writer._root_object["/AcroForm"]
    if "/NeedAppearances" in af:
        del af["/NeedAppearances"]
    writer.add_metadata({k: v for k, v in info.items() if k != "/Producer"})
    writer.add_metadata({"/Producer": info.get("/Producer", "")})
    with open(OUT, "wb") as fh:
        writer.write(fh)

    # every slot must appear exactly once in the file, or the macro can't trust it
    raw = OUT.read_bytes()
    for key in done:
        assert raw.count(("(MF12V:%s:" % key).encode()) == 1, key
        assert raw.count(("%%MF12A:%s:" % key).encode()) == 1, key
    assert raw.count(b"%MF12ID:") == 1
    print("wrote", OUT.name, len(raw), "bytes,", len(done), "fields")


if __name__ == "__main__":
    sys.exit(main())
