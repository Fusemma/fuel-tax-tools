"""The content check: no blocked word in anything here, and nothing it
can't read passes unless a person has looked at it.

    BLOCKED_WORDS=... ALLOWED_PHRASES=... python .github/scripts/content_check.py [folder]

Both lists come from repository secrets, one entry per line. A word
matches anywhere, any capitals, as part of a longer word too. An allowed
phrase (ALLOWED_PHRASES) is blanked out first, so a blocked word inside
an allowed phrase passes and the bare word fails.

What it reads, under the folder (.git aside):
- Every file and folder name, spelled plain (wide letters and ligatures
  made ordinary, NFKC) with its spaces, dots, dashes and underscores
  taken out, so HazelNut.txt and hazel_nut.txt both match "hazel nut".
- Every file's bytes, as UTF-8, Latin-1 and UTF-16.
- A zip, which is what an Office file (.xlsx, .xlsm, .docx and kin) is:
  each file in it, its name and its bytes, read the same way.
  vbaProject.bin is read as bytes; its packed macro source is not
  unpacked. A .gz, .bz2 or .xz file is unpacked and read the same way.
- A PDF: every stream but a picture, unpacked (Flate, ASCII hex,
  ASCII85), read the same way; every text string in its objects, plain
  or hex, such as the author; and the words on its pages, forms and
  patterns. Each string there is turned into letters by the font it is
  drawn in (its character map, else its encoding), then again by every
  font in the file in case one was not found, and read like a name.
- A PNG: its packed text (zTXt, iTXt), unpacked.

What it can't read fails, unless a person has looked at it and found no
blocked word: a picture, a sound or a video, known by its first bytes,
whether a file, inside a zip or in a PDF; a 7z, RAR, Zstandard or cabinet
file; a zip, .gz, .bz2 or .xz that will not open; and a PDF that is
locked, packs a stream any other way, has a font with no way to turn its
codes into letters, or can't be taken apart. Each prints the line to add
to .github/looked-at.txt after the look: its SHA-256 and where it is.
Every .github/looked-at.txt under the folder counts, and only the SHA-256
on each line, so a changed file fails again until it is looked at again.
A listed file passes everything in it.

A hit prints the file and the word's line number in the list, never the
word; a name that holds one prints as [blocked name]. Exit 1 on any hit,
on anything it can't read, if the blocked list is empty, if it reads no
files at all, or if a file or folder it finds will not open (a path too
long for Windows does both).
"""
import base64
import bz2
import gzip
import hashlib
import io
import lzma
import os
import pathlib
import re
import sys
import unicodedata
import zipfile
import zlib

SKIP = {".git"}
LOOKED_AT = (".github", "looked-at.txt")

# First bytes of a picture, a sound or a video: words in one are never text.
MEDIA = re.compile(
    rb"\x89PNG|\xff\xd8\xff|GIF8[79]a|BM.{4}\x00{4}|II\*\x00|MM\x00\*|8BPS"
    rb"|\x00\x00[\x01\x02]\x00|\x00\x00\x00\x0cjP  |\xffO\xffQ|RIFF.{4}(?:WEBP|WAVE|AVI )"
    rb"|.{4}ftyp|ID3|fLaC|OggS|\x1aE\xdf\xa3", re.S)
# First bytes of a packed file the check can't open: 7z, RAR, Zstandard, cabinet.
SEALED = re.compile(rb"7z\xbc\xaf\x27\x1c|Rar!\x1a\x07|\x28\xb5\x2f\xfd|MSCF\x00{4}")
UNPACK = [(re.compile(rb"\x1f\x8b\x08"), gzip.decompress),
          (re.compile(rb"BZh[1-9]1AY&SY"), bz2.decompress),
          (re.compile(rb"\xfd7zXZ\x00"), lzma.decompress)]


def entries(name):
    raw = os.environ.get(name, "").replace("\r", "")
    return [w.strip().lower() for w in raw.split("\n") if w.strip()]


def readings(data):
    """Every way a file's bytes could carry a word."""
    out = [data.decode("utf-8", errors="replace"), data.decode("latin-1")]
    for start in (0, 1):
        out.append(data[start:].decode("utf-16-le", errors="replace"))
    return out


def squeeze(text):
    """Text with its spaces, dots, dashes and underscores taken out."""
    return re.sub(r"[\s._-]+", "", text)


def hits(text, words, allow):
    low = text.lower()
    for a in allow:
        if a:
            low = low.replace(a, " ")
    return [n for n, w in enumerate(words, 1) if w and w in low]


def sha(data):
    return hashlib.sha256(data).hexdigest()


# Reading a PDF. Its page words sit in packed streams as codes that each
# font turns into letters, so the check takes it apart.

class Name(str):
    """A PDF name, like /Font."""


class Ref(int):
    """A pointer to another object, like 5 0 R."""


class Unread(Exception):
    """Something in a PDF the check can't read, and why."""


TOKEN = re.compile(
    rb"\s+|%[^\r\n]*|(<<|>>|\[|\]|\{|\})|<([0-9A-Fa-f\s]*)>"
    rb"|/([^\s()<>\[\]{}/%]*)|([^\s()<>\[\]{}/%]+)|.", re.S)
LITERAL = re.compile(rb"[^()\\]+|\\([0-7]{1,3}|\r\n|[\s\S])|\\|\(|\)")
ESCAPES = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b", b"f": b"\f",
           b"\r\n": b"", b"\r": b"", b"\n": b""}
INLINE_END = re.compile(rb"\sEI(?=[\s/<\[(%]|\Z)")


def unhex(digits):
    digits = re.sub(rb"[^0-9A-Fa-f]", b"", digits)
    return bytes.fromhex((digits + b"0" * (len(digits) % 2)).decode())


def literal(data, pos):
    """A (string) from just past its opening bracket: its bytes, and where it ends."""
    out, depth = bytearray(), 1
    while pos < len(data):
        m = LITERAL.match(data, pos)
        pos = m.end()
        if m[1] is not None:
            e = m[1]
            out += bytes([int(e, 8) & 255]) if e[0] in b"01234567" else ESCAPES.get(e, e)
        elif m[0] == b"(":
            depth += 1
            out += m[0]
        elif m[0] == b")":
            depth -= 1
            if depth == 0:
                return bytes(out), pos
            out += m[0]
        elif m[0] != b"\\":
            out += m[0]
    return bytes(out), pos


def tokens(data):
    """A PDF's tokens as (kind, value). The kind is "d" for a bracket, "s"
    for a string's bytes, "n" for a name, "w" for a number or keyword, and
    "picture" for the bytes of a picture drawn inline (BI ... ID ... EI)."""
    pos = 0
    while pos < len(data):
        if data[pos] == 0x28:
            s, pos = literal(data, pos + 1)
            yield "s", s
            continue
        m = TOKEN.match(data, pos)
        pos = m.end()
        if m[1]:
            yield "d", m[1]
        elif m[2] is not None:
            yield "s", unhex(m[2])
        elif m[3] is not None:
            name = re.sub(rb"#([0-9A-Fa-f]{2})", lambda h: bytes([int(h[1], 16)]), m[3])
            yield "n", Name(name.decode("latin-1"))
        elif m[4]:
            yield "w", m[4]
            if m[4] == b"ID":
                end = INLINE_END.search(data, pos + 1)
                yield "picture", data[pos + 1:end.start() if end else len(data)]
                pos = end.end() if end else len(data)


def value(toks, tok):
    """One PDF object from its first token on: a dict, list, string bytes,
    Name, Ref, number, or a keyword or operator as plain text."""
    kind, v = tok
    if kind == "d" and v in (b"<<", b"["):
        items, close = [], (b">>" if v == b"<<" else b"]")
        for t in toks:
            if t == ("d", close):
                break
            if (t == ("w", b"R") and len(items) > 1
                    and type(items[-1]) is int and type(items[-2]) is int):
                items[-2:] = [Ref(items[-2])]
            else:
                items.append(value(toks, t))
        if v == b"[":
            return items
        return {k: x for k, x in zip(items[::2], items[1::2]) if isinstance(k, Name)}
    if kind == "w":
        for number in (int, float):
            try:
                return number(v)
            except ValueError:
                pass
        return v.decode("latin-1")
    return v if kind in ("s", "n") else None


def parse(data):
    toks = tokens(data)
    tok = next(toks, None)
    return value(toks, tok) if tok else None


def strings_in(v):
    if isinstance(v, bytes):
        yield v
    elif isinstance(v, (dict, list)):
        for x in (v.values() if isinstance(v, dict) else v):
            yield from strings_in(x)


def pdf_text(s):
    """A PDF text string: UTF-16 after its byte-order mark, else a letter a byte."""
    if s[:2] in (b"\xfe\xff", b"\xff\xfe"):
        return s[2:].decode("utf-16-be" if s[0] == 0xFE else "utf-16-le", "replace")
    return s.decode("cp1252", "replace")


def shown(data, font=None):
    """What a page's drawing shows: each string it draws as text, in order,
    with the name of the font it is drawn in; the strings in its
    marked-content tags (like ActualText); and the pictures it draws inline."""
    strings, marked, pictures, operands, saved = [], [], [], [], []
    toks = tokens(data)
    for tok in toks:
        if tok[0] == "picture":
            pictures.append(tok[1])
            continue
        v = value(toks, tok)
        if type(v) is not str:
            operands.append(v)
            continue
        if v in ("Tj", "TJ", "'", '"') and operands:
            strings += [(font, s) for s in strings_in(operands[-1])]
        elif v == "Tf" and len(operands) > 1:
            font = operands[-2]
        elif v == "q":
            saved.append(font)
        elif v == "Q" and saved:
            font = saved.pop()
        elif v in ("BDC", "DP"):
            marked += strings_in(operands)
        operands = []
    return strings, marked, pictures


def inflate(data):
    try:
        return zlib.decompressobj().decompress(data)
    except zlib.error:
        try:
            return zlib.decompressobj(-15).decompress(data[2:])
        except zlib.error:
            raise Unread("a stream that will not unpack") from None


def a85(data):
    data = re.sub(rb"\s", b"", data).split(b"~>")[0]
    return base64.a85decode(data[2:] if data.startswith(b"<~") else data)


UNPACKERS = {"FlateDecode": inflate, "Fl": inflate,
             "ASCIIHexDecode": lambda d: unhex(d.split(b">")[0]),
             "AHx": lambda d: unhex(d.split(b">")[0]),
             "ASCII85Decode": a85, "A85": a85}


def character_map(text):
    """A font's character map (ToUnicode): its table from codes to letters,
    and the code lengths in bytes, longest first."""
    pairs = rb"<([^>]*)>\s*<([^>]*)>"
    table, sizes = {}, set()
    for block in re.findall(rb"begincodespacerange(.*?)endcodespacerange", text, re.S):
        sizes |= {len(unhex(lo)) for lo, _ in re.findall(pairs, block)}
    for block in re.findall(rb"beginbfchar(.*?)endbfchar", text, re.S):
        for code, letters in re.findall(pairs, block):
            table[unhex(code)] = unhex(letters).decode("utf-16-be", "replace")
    for block in re.findall(rb"beginbfrange(.*?)endbfrange", text, re.S):
        for lo, hi, to in re.findall(rb"<([^>]*)>\s*<([^>]*)>\s*(<[^>]*>|\[[^\]]*\])", block):
            lo, hi = unhex(lo), unhex(hi)
            first = int.from_bytes(lo, "big")
            count = max(0, min(int.from_bytes(hi, "big") - first + 1, 65536))
            if to[:1] == b"[":
                letters = [unhex(t).decode("utf-16-be", "replace")
                           for t in re.findall(rb"<([^>]*)>", to)]
            else:
                start = unhex(to)
                base, top = int.from_bytes(start, "big"), 1 << (8 * len(start))
                letters = [((base + i) % top).to_bytes(len(start), "big").decode("utf-16-be", "replace")
                           for i in range(count)] if start else []
            for i, s in enumerate(letters[:count]):
                table[(first + i).to_bytes(len(lo), "big")] = s
    # Some maps claim two-byte codes and list one-byte ones, so both count.
    return table, sorted(sizes | {len(k) for k in table} or {2}, reverse=True)


def by_codes(table, sizes):
    """Letters for a string of codes from a character map, longest code first."""
    def letters(s):
        out, i = [], 0
        while i < len(s):
            for n in sizes:
                if s[i:i + n] in table:
                    out.append(table[s[i:i + n]])
                    i += n
                    break
            else:
                i += sizes[-1]
        return "".join(out)
    return letters


GLYPHS = {"space": " ", "hyphen": "-", "period": ".", "comma": ",", "ampersand": "&",
          "at": "@", "underscore": "_", "quoteright": "'", "quotesingle": "'",
          "slash": "/", "fi": "fi", "fl": "fl", "ff": "ff", "ffi": "ffi", "ffl": "ffl",
          "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4",
          "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9"}


def glyph(name):
    """The letters a glyph name stands for: "" for one that is no letter,
    None for a numbered name like g12, which doesn't say."""
    out = ""
    for part in name.split(".")[0].split("_"):
        if len(part) == 1 or part in GLYPHS:
            out += GLYPHS.get(part, part)
        elif re.fullmatch(r"uni(?:[0-9A-F]{4})+", part):
            out += "".join(chr(int(part[i:i + 4], 16)) for i in range(3, len(part), 4))
        elif re.fullmatch(r"u[0-9A-F]{4,6}", part):
            out += chr(min(int(part[1:], 16), 0x10FFFF))
        elif re.search(r"\d", part):
            return None
    return out


def one_byte(table):
    return lambda s: "".join(table[b] for b in s)


class Pdf:
    """A PDF taken apart: its objects, each stream unpacked when asked."""
    OBJECT = re.compile(rb"(\d+)\s+\d+\s+obj\b")
    STREAM = re.compile(rb"\bstream(?:\r\n|\n|\r)")
    FONTS = {"Type0", "Type1", "MMType1", "TrueType", "Type3"}

    def __init__(self, data):
        self.objects, self.values, self.problems, self.ways = {}, {}, [], {}
        pos = 0
        while m := self.OBJECT.search(data, pos):
            end = data.find(b"endobj", m.end())
            end = len(data) if end < 0 else end
            s = self.STREAM.search(data, m.end(), end)
            if s:
                stop = data.find(b"endstream", s.end())
                stop = len(data) if stop < 0 else stop
                self.objects[int(m[1])] = (data[m.end():s.start()], data[s.end():stop])
                pos = stop
            else:
                self.objects[int(m[1])] = (data[m.end():end], None)
                pos = end
        trailers = [parse(data[m.end() - 2:m.end() + 4096])
                    for m in re.finditer(rb"trailer\s*<<", data)]
        trailers += [self.dict(n) for n in self.objects if self.dict(n).get("Type") == "XRef"]
        self.locked = any(isinstance(t, dict) and "Encrypt" in t for t in trailers)
        if self.locked:
            return
        for num in list(self.objects):
            if self.dict(num).get("Type") == "ObjStm":
                try:
                    body, first = self.body(num), self.dict(num).get("First", 0)
                    nums = [int(x) for x in body[:first].split()]
                except (Unread, ValueError) as e:
                    self.problems.append((num, str(e) if isinstance(e, Unread) else
                                          "an object stream that will not open"))
                    continue
                for i in range(0, len(nums) - 1, 2):
                    stop = first + nums[i + 3] if i + 3 < len(nums) else len(body)
                    self.objects.setdefault(nums[i], (body[first + nums[i + 1]:stop], None))

    def get(self, num):
        if num not in self.values:
            self.values[num] = parse(self.objects.get(num, (b"", None))[0])
        return self.values[num]

    def dict(self, num):
        d = self.get(num)
        return d if isinstance(d, dict) else {}

    def resolve(self, v):
        return self.get(v) if isinstance(v, Ref) else v

    def has_stream(self, num):
        return self.objects.get(num, (b"", None))[1] is not None

    def body(self, num):
        """A stream, unpacked."""
        d, data = self.dict(num), self.objects.get(num, (b"", None))[1]
        if data is None:
            raise Unread("a missing stream")
        filters = self.resolve(d.get("Filter", []))
        filters = filters if isinstance(filters, list) else [filters]
        parms = self.resolve(d.get("DecodeParms"))
        parms = parms if isinstance(parms, list) else [parms] * len(filters)
        for i, f in enumerate(filters):
            f = self.resolve(f)
            if f not in UNPACKERS:
                raise Unread(f"a stream packed as {f}")
            data = UNPACKERS[f](data)
            p = self.resolve(parms[i]) if i < len(parms) else None
            if isinstance(p, dict) and p.get("Predictor", 1) != 1:
                raise Unread("a stream packed with a predictor")
        return data

    def way(self, key, make):
        """One function per way of reading, so fonts that read alike share one."""
        if key not in self.ways:
            self.ways[key] = make()
        return self.ways[key]

    def font(self, num):
        """A way to turn a font's codes into letters, or None if it has none.
        A Type0 font's codes are one or more bytes, read through its character
        map. Any other font's codes are one byte, read through its encoding,
        with its character map laid over it."""
        d = self.dict(num)
        to_unicode, mapped, sizes = d.get("ToUnicode"), {}, []
        if isinstance(to_unicode, Ref):
            try:
                mapped, sizes = character_map(self.body(to_unicode))
            except Unread:
                pass
        encoding = self.resolve(d.get("Encoding"))
        if d.get("Subtype") == "Type0":
            if mapped:
                return self.way(("map", int(to_unicode)), lambda: by_codes(mapped, sizes))
            if isinstance(encoding, Name) and re.search("UCS2|UTF16", encoding):
                return self.way("utf-16", lambda: lambda s: s.decode("utf-16-be", "replace"))
            return None
        base, differences = encoding, []
        if isinstance(encoding, dict):
            base, differences = encoding.get("BaseEncoding"), self.resolve(encoding.get("Differences"))
        codec = "mac_roman" if base == "MacRomanEncoding" else "cp1252"
        table = [bytes([b]).decode(codec, "replace") for b in range(256)]
        code, numbered = 0, False
        for item in differences if isinstance(differences, list) else []:
            if type(item) is int:
                code = item
            elif isinstance(item, Name):
                letters = glyph(item)
                numbered = numbered or letters is None
                table[code % 256] = letters or ""
                code += 1
        for c, letters in mapped.items():
            if len(c) == 1 or (len(c) == 2 and c[0] == 0):
                table[c[-1]] = letters
        if numbered and not mapped:
            return None
        return self.way(tuple(table), lambda: one_byte(table))

    def fonts(self):
        """Each font, by object number, with its way to turn codes into letters."""
        return {num: self.font(num) for num in list(self.objects)
                if self.dict(num).get("Subtype") in self.FONTS
                and self.dict(num).get("Type", "Font") == "Font"}

    def inherited(self, d, key):
        """A page's own entry, else the nearest one up its page tree."""
        for _ in range(64):
            if key in d:
                return self.resolve(d[key])
            d = self.resolve(d.get("Parent"))
            if not isinstance(d, dict):
                return None

    def pages(self):
        """For each page, form and pattern: the strings it draws as text, each
        with the font object it is drawn in (None if not found). Also the
        strings in their marked-content tags, and the pictures they draw inline."""
        groups, marked, pictures = [], [], []
        for num in list(self.objects):
            d = self.dict(num)
            if d.get("Type") == "Page":
                refs = d.get("Contents")
                if isinstance(refs, Ref) and not self.has_stream(refs):
                    refs = self.get(refs)
                streams = [r for r in (refs if isinstance(refs, list) else [refs])
                           if isinstance(r, Ref) and self.has_stream(r)]
                resources = self.inherited(d, "Resources")
            elif self.has_stream(num) and (d.get("Subtype") == "Form" or d.get("PatternType") == 1):
                streams, resources = [num], self.resolve(d.get("Resources"))
            else:
                continue
            names = self.resolve(resources.get("Font")) if isinstance(resources, dict) else None
            names = names if isinstance(names, dict) else {}
            # A page's streams are one drawing; a command may start in one
            # and end in the next.
            bodies = []
            for n in streams:
                try:
                    bodies.append(self.body(n))
                except Unread:
                    pass
            strings, tags, drawn = shown(b"\n".join(bodies))
            groups.append([(names.get(f) if isinstance(f, str) else None, s) for f, s in strings])
            marked += tags
            pictures += drawn
        return groups, marked, pictures


def png_text(data):
    """The packed text in a PNG's zTXt and iTXt pieces, unpacked."""
    out, pos = [], 8
    while pos + 8 <= len(data):
        size, kind = int.from_bytes(data[pos:pos + 4], "big"), data[pos + 4:pos + 8]
        body, pos = data[pos + 8:pos + 8 + size], pos + 12 + size
        if kind == b"zTXt":
            out.append(zlib.decompress(body.partition(b"\x00")[2][1:]))
        elif kind == b"iTXt":
            rest = body.partition(b"\x00")[2]
            if rest[:1] == b"\x01":
                out.append(zlib.decompress(rest[2:].split(b"\x00", 2)[-1]))
    return out


class Check:
    def __init__(self, words, allow, looked):
        self.words, self.allow, self.looked = words, allow, looked
        self.tight = [squeeze(w) for w in words], [squeeze(a) for a in allow]
        self.found, self.unseen, self.passed = set(), set(), set()

    def read(self, label, data):
        for text in readings(data):
            for n in hits(text, self.words, self.allow):
                self.found.add((label, "", n))

    def tight_hits(self, text):
        # NFKC spells out ligatures (the fi in "Pacific") and wide letters.
        return hits(squeeze(unicodedata.normalize("NFKC", text)), *self.tight)

    def read_tight(self, label, text, where=""):
        for n in self.tight_hits(text):
            self.found.add((label, where, n))

    def name(self, label, name):
        self.read_tight(label, name, " in its name")

    def cannot(self, label, why, mark, within):
        """Something it can't read: passes only if it, or a file it is in, was looked at."""
        if mark in self.looked or within & self.looked:
            self.passed.add(mark)
        else:
            self.unseen.add((label, why, mark))

    def scan(self, label, data, within=frozenset()):
        mark = sha(data)
        within = within | {mark}
        self.read(label, data)
        if MEDIA.match(data):
            self.cannot(label, "a picture, a sound or a video", mark, within)
        if SEALED.match(data):
            self.cannot(label, "a packed file the check can't open", mark, within)
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            try:
                for text in png_text(data):
                    self.read(f"{label} > its text", text)
            except zlib.error:
                self.cannot(label, "a picture whose text will not unpack", mark, within)
        for first, unpack in UNPACK:
            if first.match(data):
                try:
                    inside = unpack(data)
                except (OSError, EOFError, ValueError, zlib.error, lzma.LZMAError):
                    self.cannot(label, "a packed file that will not unpack", mark, within)
                else:
                    self.scan(f"{label} > unpacked", inside, within)
        if data[:4] == b"PK\x03\x04":
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for info in z.infolist():
                        inner = f"{label} > {info.filename}"
                        self.name(inner, info.filename)
                        if not info.is_dir():
                            self.scan(inner, z.read(info), within)
            except zipfile.BadZipFile:
                self.cannot(label, "a zip that will not open", mark, within)
        if data[:5] == b"%PDF-" or re.search(rb"[\r\n]%PDF-", data[:1024]):
            try:
                self.pdf(label, data, mark, within)
            except Exception:  # built some way this reader doesn't follow
                self.cannot(label, "a PDF the check could not take apart", mark, within)

    def pdf(self, label, data, mark, within):
        doc = Pdf(data)
        if doc.locked or not doc.objects:
            why = "a locked PDF" if doc.locked else "a PDF the check could not take apart"
            self.cannot(label, why, mark, within)
            return
        for num, why in doc.problems:
            self.cannot(f"{label} > object {num}", why, mark, within)
        strings = [s for head, _ in doc.objects.values() for kind, s in tokens(head) if kind == "s"]
        for n in hits("\n".join(map(pdf_text, strings)), self.words, self.allow):
            self.found.add((f"{label} > text in its objects", "", n))
        for num, (_, raw) in list(doc.objects.items()):
            d = doc.dict(num)
            if raw is None or d.get("Type") == "XRef":
                continue
            if d.get("Subtype") == "Image":
                self.cannot(f"{label} > picture, object {num}", "a picture", sha(raw), within)
                continue
            try:
                body = doc.body(num)
            except Unread as e:
                self.cannot(f"{label} > object {num}", str(e), mark, within)
            else:
                self.scan(f"{label} > object {num}", body, within)
        fonts = doc.fonts()
        for num, letters in fonts.items():
            if letters is None:
                self.cannot(f"{label} > font, object {num}",
                            "a font with no way to turn its codes into letters", mark, within)
        groups, marked, pictures = doc.pages()
        if fonts and not groups:
            self.cannot(label, "a PDF whose pages the check could not find", mark, within)
        for picture in pictures:
            self.cannot(f"{label} > picture drawn in a page", "a picture", sha(picture), within)
        for n in hits("\n".join(map(pdf_text, marked)), self.words, self.allow):
            self.found.add((f"{label} > page text", "", n))
        # Each string through the font it is drawn in; then, in case a font
        # was not found, every string through each font in turn and through
        # plain one-byte and two-byte text.
        plain = lambda s: s.decode("cp1252", "replace")
        self.read_tight(f"{label} > page text", "|".join(
            "".join((fonts.get(f) or plain)(s) for f, s in group) for group in groups))
        ways = {plain, lambda s: s.decode("utf-16-be", "replace")} | set(fonts.values()) - {None}
        for letters in ways:
            self.read_tight(f"{label} > page text", "|".join(
                "".join(letters(s) for _, s in group) for group in groups))

    def shown_as(self, label):
        """The label with each name that holds a blocked word blanked, so the
        log never shows one."""
        return "".join(p if p in ("/", " > ") or not self.tight_hits(p)
                       else "[blocked name]" for p in re.split(r"(/| > )", label))


def main():
    # A name in any script prints, rather than stopping the check on Windows.
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    words, allow = entries("BLOCKED_WORDS"), entries("ALLOWED_PHRASES")
    if not words:
        print("BLOCKED_WORDS is empty or missing; the check cannot run")
        return 1
    root = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    folders, files, unread, count = [], [], [], 0

    def cannot_list(e):
        # A folder that will not list (a path too long for Windows) would
        # otherwise be passed over without a word.
        folder = pathlib.Path(e.filename or "")
        if folder != root:
            unread.append(folder.relative_to(root).as_posix() + "/")

    for top, dirs, names in os.walk(root, onerror=cannot_list):
        dirs[:] = sorted(d for d in dirs if d not in SKIP)
        folders += [pathlib.Path(top) / d for d in dirs]
        files += [pathlib.Path(top) / n for n in sorted(names) if n not in SKIP]
    looked = set()
    for p in files:
        if p.relative_to(root).parts[-2:] == LOOKED_AT:
            try:
                text = p.read_text("utf-8", "replace").lower()
            except OSError:
                continue  # it is listed below as a file that will not open
            looked |= set(re.findall(r"^\s*([0-9a-f]{64})\b", text, re.M))
    check = Check(words, allow, looked)
    for p in folders:
        check.name(p.relative_to(root).as_posix() + "/", p.name)
    for p in files:
        label = p.relative_to(root).as_posix()
        check.name(label, p.name)
        try:
            data = p.read_bytes()
        except OSError:
            unread.append(label)
            continue
        count += 1
        check.scan(label, data)
    show = check.shown_as
    for label, where, n in sorted(check.found):
        print(f"{show(label)}: blocked word{where}, line {n} of the list")
    for label, why, mark in sorted(check.unseen):
        print(f"{show(label)}: {why}, so its words can't be read. Once a person has looked "
              f"and found no blocked word, add this line to .github/looked-at.txt:\n"
              f"    {mark}  {show(label)}")
    for label in unread:
        print(f"{show(label)}: will not open, so it was not checked")
    print(f"{count} files read, {len(check.found)} hit(s), {len(check.unseen)} that can't be "
          f"read, {len(check.passed)} passed as looked at by a person")
    if count == 0:
        print(f"No files read in {root}: the folder is empty, missing, or its "
              "paths are too long to open. Nothing was checked, so this fails.")
    return 1 if check.found or check.unseen or unread or count == 0 else 0


if __name__ == "__main__":
    sys.exit(main())
