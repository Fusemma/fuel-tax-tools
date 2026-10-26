"""Tests the content check on throwaway folders, and on this repository.

    python .github/scripts/test_content_check.py

Each test builds a folder, runs the check on it with its own blocked
words, and looks at the exit code and what it printed. Prints PASS or
FAIL for each test and a count at the end; exit 1 on any FAIL.
"""
import gzip
import hashlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
CHECK = os.path.join(HERE, "content_check.py")
STATE_PAGE = os.path.join(REPO, "General", "State info", "preview", "State page.pdf")
results = []


def run(files, blocked, allowed=""):
    """Writes files ({path: bytes}) to a new folder and runs the check on it."""
    folder = tempfile.mkdtemp(prefix="check-test-")
    try:
        for path, data in files.items():
            full = os.path.join(folder, *path.split("/"))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "wb") as f:
                f.write(data)
        env = dict(os.environ, BLOCKED_WORDS=blocked, ALLOWED_PHRASES=allowed)
        r = subprocess.run([sys.executable, CHECK, folder], env=env, capture_output=True)
        return r.returncode, r.stdout.decode("utf-8", "replace")
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def check(name, ok, out=""):
    results.append(ok)
    print(("PASS " if ok else "FAIL ") + name)
    if not ok and out:
        print("    " + out.strip().replace("\n", "\n    "))


def fails(name, files, blocked, why, allowed=""):
    """The check must fail, say why, and never print a blocked word."""
    code, out = run(files, blocked, allowed)
    quiet = all(w.strip().lower() not in out.lower() for w in blocked.split("\n") if w.strip())
    check(name, code == 1 and why in out and quiet, out)


def passes(name, files, blocked, allowed=""):
    code, out = run(files, blocked, allowed)
    check(name, code == 0, out)


CLEAN = {"clean.txt": b"Nothing to see here.\n"}


# A PNG: one gray dot, with optional extra chunks.
def chunk(kind, body):
    return len(body).to_bytes(4, "big") + kind + body + zlib.crc32(kind + body).to_bytes(4, "big")


def png(extra=b"", shade=128):
    head = (1).to_bytes(4, "big") * 2 + bytes([8, 0, 0, 0, 0])
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", head) + extra
            + chunk(b"IDAT", zlib.compress(bytes([0, shade]))) + chunk(b"IEND", b""))


def office(members):
    """A zip shaped like an Office file."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        for name, data in members.items():
            z.writestr(name, data)
    return buf.getvalue()


# A one-page PDF, every stream packed with Flate the way most PDFs are.
def stream(head, data):
    packed = zlib.compress(data)
    return (b"<< %s /Filter /FlateDecode /Length %d >>\nstream\n" % (head, len(packed))
            + packed + b"\nendstream")


def pdf(ops, font, *more, trailer=b"", fonts=b"/F1 5 0 R"):
    """Objects 1 to 5 are the catalog, the page list, the page, its drawing
    and font F1; more objects follow as 6, 7 and on."""
    # The page takes its fonts and pictures from the page list above it.
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1"
            b" /Resources << /Font << %s >> /XObject << /Im1 6 0 R >> >> >>" % fonts,
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R >>",
            stream(b"", ops), font, *more]
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    spots = []
    for n, body in enumerate(objs, 1):
        spots.append(len(out))
        out += b"%d 0 obj\n" % n + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % s for s in spots)
    out += b"trailer\n<< /Size %d /Root 1 0 R %s >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objs) + 1, trailer, xref)
    return bytes(out)


HELVETICA = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
# "Quokka" split across a kern, then "Island" after a gap with no space in it.
QUOKKA_ISLAND = b"BT /F1 12 Tf 72 700 Td [(Quo) -30 (kka) -400 (Island)] TJ ET"

# A font whose codes are glyph numbers, not letters, as Word and Excel
# often write them: only its character map (ToUnicode) says what they are.
GLYPH_FONT = (b"<< /Type /Font /Subtype /Type0 /BaseFont /Arial /Encoding /Identity-H"
              b" /DescendantFonts [7 0 R] %s >>")
CID_FONT = (b"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Arial"
            b" /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>")
CHARACTER_MAP = b"""/CIDInit /ProcSet findresource begin 12 dict begin begincmap
/CMapName /Test def
1 begincodespacerange <0000> <FFFF> endcodespacerange
2 beginbfchar <0101> <0051> <0102> <0075> endbfchar
2 beginbfrange <0103> <0105> [<006F> <006B> <0061>] <0003> <0003> <0020> endbfrange
endcmap CMapName currentdict /CMap defineresource pop end end"""
# "Quok" and "ka", two pieces with a move between them.
GLYPH_OPS = b"BT /F1 12 Tf 72 700 Td <0101010201030104> Tj 0 -14 Td <01040105> Tj ET"

# Two fonts in one word: "Quok" in F1, then two glyph codes in F2.
TWO_FONTS = b"BT /F1 12 Tf 72 700 Td (Quok) Tj /F2 12 Tf <%s> Tj ET"
TWO_FONT_OBJECTS = (stream(b"", CHARACTER_MAP), CID_FONT, GLYPH_FONT % b"/ToUnicode 6 0 R")
# A font whose character map gives code 1 as a ligature: fi or fl in one letter.
LIGATURE_FONT = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>"
LIGATURE_MAP = (b"1 begincodespacerange <00> <FF> endcodespacerange\n"
                b"1 beginbfchar <01> <%s> endbfchar\n1 beginbfrange <61> <7A> <0061> endbfrange")

PICTURE_OPS = b"q 10 0 0 10 72 600 cm /Im1 Do Q " + QUOKKA_ISLAND
PDF_PICTURE = stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1"
                     b" /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x80")
INLINE_OPS = b"q 10 0 0 10 72 600 cm BI /W 1 /H 1 /CS /G /BPC 8 ID \x80 EI Q " + QUOKKA_ISLAND


def sha(data):
    return hashlib.sha256(data).hexdigest()


def main():
    # Gap 1: words on a PDF page, packed where raw bytes never show them.
    fails("PDF: a blocked word on the page, packed and split by a kern",
          {"page.pdf": pdf(QUOKKA_ISLAND, HELVETICA)}, "quokka", "blocked word")
    fails("PDF: a blocked word in glyph codes, read through the font's character map",
          {"page.pdf": pdf(GLYPH_OPS, GLYPH_FONT % b"/ToUnicode 6 0 R", stream(b"", CHARACTER_MAP), CID_FONT)},
          "quokka", "blocked word")
    fails("PDF: a two-word entry on the page with no space between, only a gap",
          {"page.pdf": pdf(QUOKKA_ISLAND, HELVETICA)}, "quokka island", "blocked word")
    fails("PDF: a blocked word as the author, in hex",
          {"page.pdf": pdf(b"", HELVETICA, b"<< /Author <FEFF00510075006F006B006B0061> >>",
                           trailer=b"/Info 6 0 R")},
          "quokka", "blocked word")
    if os.path.exists(STATE_PAGE):
        with open(STATE_PAGE, "rb") as f:
            state_page = f.read()
        fails("PDF: State page.pdf with a word printed on its page set as blocked",
              {"State page.pdf": state_page}, "disbursements", "blocked word")
        fails("PDF: State page.pdf with a two-word name on its page set as blocked",
              {"State page.pdf": state_page}, "fulton county", "blocked word")
        passes("PDF: State page.pdf with a word not on it", {"State page.pdf": state_page}, "wombat")
    else:
        check("PDF: State page.pdf is in the repository", False, STATE_PAGE)
    passes("PDF: a clean page passes", {"page.pdf": pdf(QUOKKA_ISLAND, HELVETICA)}, "wombat")
    passes("PDF: a clean page through a character map passes",
           {"page.pdf": pdf(GLYPH_OPS, GLYPH_FONT % b"/ToUnicode 6 0 R", stream(b"", CHARACTER_MAP), CID_FONT)},
           "wombat")
    passes("PDF: the blocked word inside an allowed phrase passes",
           {"page.pdf": pdf(QUOKKA_ISLAND, HELVETICA)}, "quokka", "quokka island")
    fails("PDF: a blocked word drawn half in one font, half in another",
          {"page.pdf": pdf(TWO_FONTS % b"01040105", HELVETICA, *TWO_FONT_OBJECTS,
                           fonts=b"/F1 5 0 R /F2 8 0 R")}, "quokka", "blocked word")
    passes("PDF: the second font's codes read as its own letters, not the first font's",
           {"page.pdf": pdf(TWO_FONTS % b"01030102", HELVETICA, *TWO_FONT_OBJECTS,
                            fonts=b"/F1 5 0 R /F2 8 0 R")}, "quokka")
    fails("PDF: a blocked word with the fi ligature in it",
          {"page.pdf": pdf(b"BT /F1 12 Tf 72 700 Td (\\001ord) Tj ET", LIGATURE_FONT,
                           stream(b"", LIGATURE_MAP % b"FB01"))}, "fiord", "blocked word")
    passes("PDF: the fl ligature is not fi",
           {"page.pdf": pdf(b"BT /F1 12 Tf 72 700 Td (\\001ord) Tj ET", LIGATURE_FONT,
                            stream(b"", LIGATURE_MAP % b"FB02"))}, "fiord")
    fails("PDF: a stream packed some other way can't be read",
          {"page.pdf": pdf(b"", HELVETICA, b"<< /Filter /LZWDecode /Length 2 >>\nstream\n\x80\x0b\nendstream")},
          "wombat", "can't be read")
    fails("PDF: a stream packed with a predictor can't be read",
          {"page.pdf": pdf(b"", HELVETICA, stream(b"/DecodeParms << /Predictor 12 /Columns 4 >>", b"\x00abcd"))},
          "wombat", "can't be read")
    fails("PDF: glyph codes with no character map can't be read",
          {"page.pdf": pdf(GLYPH_OPS, GLYPH_FONT % b"", b"null", CID_FONT)}, "wombat", "can't be read")
    fails("PDF: a locked PDF can't be read",
          {"page.pdf": pdf(QUOKKA_ISLAND, HELVETICA, b"<< /Filter /Standard /V 1 /R 2 /P -4 >>",
                           trailer=b"/Encrypt 6 0 R")},
          "wombat", "can't be read")

    # Gap 2: words in a picture are never text.
    fails("Picture: a PNG fails", {"chart.png": png()}, "wombat", "can't be read")
    fails("Picture: a PNG under another name fails", {"chart.dat": png()}, "wombat", "can't be read")
    fails("Picture: a picture inside an Office file fails",
          {"report.docx": office({"word/document.xml": "<w/>", "word/media/image1.png": png()})},
          "wombat", "can't be read")
    fails("Picture: a picture in a PDF fails",
          {"page.pdf": pdf(PICTURE_OPS, HELVETICA, PDF_PICTURE)}, "wombat", "can't be read")
    fails("Picture: a picture drawn inside a PDF page fails",
          {"page.pdf": pdf(INLINE_OPS, HELVETICA)}, "wombat", "can't be read")
    looked = b"%s  chart.png\n" % sha(png()).encode()
    passes("Picture: a picture on .github/looked-at.txt passes",
           {"chart.png": png(), ".github/looked-at.txt": looked}, "wombat")
    fails("Picture: a changed picture with only the old one listed fails",
          {"chart.png": png(shade=129), ".github/looked-at.txt": looked}, "wombat", "can't be read")
    page = pdf(PICTURE_OPS, HELVETICA, PDF_PICTURE)
    passes("Picture: a PDF listed whole passes with its picture",
           {"page.pdf": page, ".github/looked-at.txt": b"%s  page.pdf\n" % sha(page).encode()}, "wombat")
    marked = png(chunk(b"zTXt", b"Author\x00\x00" + zlib.compress(b"Quokka")))
    fails("Picture: a blocked word packed in a listed picture's text still fails",
          {"chart.png": marked, ".github/looked-at.txt": b"%s  chart.png\n" % sha(marked).encode()},
          "quokka", "blocked word")
    passes("Picture: a text file that starts like a BMP is not a picture",
           {"BMW fleet.txt": b"BMW 2024 fleet list\n"}, "wombat")

    # Packed files: opened where the standard library can, else they fail.
    fails("Packed: a blocked word inside a .gz file",
          {"notes.txt.gz": gzip.compress(b"Quokka\n")}, "quokka", "blocked word")
    passes("Packed: a clean .gz file passes", {"notes.txt.gz": gzip.compress(b"Quokka\n")}, "wombat")
    fails("Packed: a 7z file can't be read",
          {"notes.7z": b"7z\xbc\xaf\x27\x1c\x00\x04" + bytes(24)}, "wombat", "can't be read")
    fails("Packed: a zip that will not open can't be read",
          {"broken.xlsx": b"PK\x03\x04" + bytes(40)}, "wombat", "can't be read")
    broken = b"PK\x03\x04" + bytes(40)
    passes("Packed: a broken zip on .github/looked-at.txt passes",
           {"broken.xlsx": broken, ".github/looked-at.txt": sha(broken).encode() + b"  broken.xlsx\n"}, "wombat")

    # Gap 3: file and folder names.
    fails("Name: a file named with a blocked word fails",
          {"Quokka notes.txt": b"Nothing to see here.\n"}, "quokka", "in its name")
    fails("Name: a folder named with a blocked word fails",
          {"Quokka/clean.txt": b"Nothing to see here.\n"}, "quokka", "in its name")
    fails("Name: a two-word entry run together fails", {"HazelNut.txt": b"x\n"}, "hazel nut", "in its name")
    fails("Name: a two-word entry joined by an underscore fails",
          {"hazel_nut.txt": b"x\n"}, "hazel nut", "in its name")
    fails("Name: a blocked word in wide letters fails",
          {"Ｑｕｏｋｋａ.txt": b"x\n"}, "quokka", "in its name")
    fails("Name: a file inside an Office file named with a blocked word fails",
          {"report.docx": office({"word/quokka.xml": "<w/>"})}, "quokka", "in its name")
    passes("Name: clean names pass", dict(CLEAN, **{"Rates check/state files/a.csv": b"1,2\n"}), "quokka")
    passes("Name: a blocked word inside an allowed phrase passes",
           {"Quokka Island/clean.txt": b"x\n"}, "quokka", "quokka island")

    # This repository, with a word that is nowhere in it (built here, so
    # not even this file holds it): everything reads.
    env = dict(os.environ, BLOCKED_WORDS="qzx" + "wvjk", ALLOWED_PHRASES="")
    r = subprocess.run([sys.executable, CHECK, REPO], env=env, capture_output=True)
    out = r.stdout.decode("utf-8", "replace")
    check("This repository: every file reads or is on .github/looked-at.txt", r.returncode == 0, out)

    print(f"{results.count(True)} passed, {results.count(False)} failed")
    return 1 if False in results else 0


if __name__ == "__main__":
    sys.exit(main())
