"""Python twin of the macro's MF-12 fill, step for step.

The VBA does exactly this: find each slot marker in mf12_template.pdf, write
the value and a small drawing of it, pad with spaces to the slot's length.
This twin exists so a change to the fill can be proven here first, and so the
macro's output can be compared byte for byte against it.

Usage: python fill_twin.py <values.json> <out.pdf>
values.json maps slot keys (SELLER, R01.GAL, ...) to [value, shown_text].
"""
import json
import sys
from pathlib import Path

TEMPLATE = Path(__file__).parent / "mf12_template.pdf"
CHAR_W = 0.6          # Courier: every character is 0.6 of the font size wide
CAP_H = 0.562         # Courier-Bold capital height, to center text up and down


def pdf_escape(text):
    out = bytearray()
    for ch in text:
        code = ord(ch)
        if ch in "()\\":
            out += b"\\" + ch.encode("ascii")
        elif 32 <= code < 127:
            out.append(code)
        else:
            try:
                out += ch.encode("cp1252")
            except UnicodeEncodeError:
                out += b"?"
    return bytes(out)


def drawing(shown, w, h, q):
    if shown == "":
        return b""
    size = 10.0
    width = len(shown) * CHAR_W * size
    if width > w - 4:
        size = int((w - 4) / (len(shown) * CHAR_W) * 100) / 100.0
        width = len(shown) * CHAR_W * size
    if q == 1:
        x = (w - width) / 2
    elif q == 2:
        x = w - 2 - width
    else:
        x = 2.0
    y = (h - CAP_H * size) / 2
    return b"BT /CoBo %.2f Tf 0 g %.2f %.2f Td (%s) Tj ET" % (size, x, y, pdf_escape(shown))


def fill(template, values):
    s = bytearray(template)
    pos = 0
    while True:                                   # value slots
        p = s.find(b"(MF12V:", pos)
        if p < 0:
            break
        k = s.find(b":", p + 7)
        key = s[p + 7:k].decode("ascii")
        e = k + 1
        while s[e] == 0x5F:
            e += 1
        assert s[e] == 0x29, key
        n = e - p + 1
        value = values.get(key, ["", ""])[0]
        rep = b"(" + pdf_escape(value) + b")"
        if len(rep) > n:
            raise ValueError("too long for the form: " + key)
        s[p:p + n] = rep + b" " * (n - len(rep))
        pos = p + n
    pos = 0
    while True:                                   # drawing slots
        p = s.find(b"%MF12A:", pos)
        if p < 0:
            break
        k = s.find(b":", p + 7)
        key = s[p + 7:k].decode("ascii")
        w = float(s[k + 3:k + 9])
        h = float(s[k + 12:k + 18])
        q = int(s[k + 21:k + 22])
        e = k + 23
        while e < len(s) and s[e] == 0x5F:
            e += 1
        n = e - p
        rep = drawing(values.get(key, ["", ""])[1], w, h, q)
        if len(rep) > n:
            raise ValueError("too long for the form: " + key)
        s[p:p + n] = rep + b" " * (n - len(rep))
        pos = p + n
    p = s.find(b"%MF12ID:")                       # fingerprint slot
    e = p + 8
    while s[e] == 0x5F:
        e += 1
    rep = b"%MF12ID:" + values.get("_ID", "").encode("ascii")
    s[p:e] = rep + b" " * (e - p - len(rep))
    return bytes(s)


if __name__ == "__main__":
    vals = json.loads(Path(sys.argv[1]).read_text())
    Path(sys.argv[2]).write_bytes(fill(TEMPLATE.read_bytes(), vals))
