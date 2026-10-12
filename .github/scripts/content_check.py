"""The content check: no blocked word in any file here.

    BLOCKED_WORDS=... ALLOWED_PHRASES=... python .github/scripts/content_check.py [folder]

Both lists come from repository secrets, one entry per line. A word
matches anywhere, any capitals, as part of a longer word too. An allowed
phrase (ALLOWED_PHRASES) is blanked out first, so a blocked word inside
an allowed phrase passes and the bare word fails.

Every file is read: text as UTF-8, everything else as bytes in both
ASCII and UTF-16. An Office file (.xlsx, .xlsm, .docx and kin) is a zip,
so each file inside it is read the same way, vbaProject.bin included.
A hit prints the file and the word's line number in the list, never the
word, so the log stays clean. Exit 1 on any hit, or if the blocked list
is empty.
"""
import io
import os
import pathlib
import sys
import zipfile

SKIP = {".git"}


def entries(name):
    raw = os.environ.get(name, "").replace("\r", "")
    return [w.strip().lower() for w in raw.split("\n") if w.strip()]


def readings(data):
    """Every way a file's bytes could carry a word."""
    out = [data.decode("utf-8", errors="replace"), data.decode("latin-1")]
    for start in (0, 1):
        out.append(data[start:].decode("utf-16-le", errors="replace"))
    return out


def hits(text, words, allow):
    low = text.lower()
    for a in allow:
        low = low.replace(a, " ")
    return [n for n, w in enumerate(words, 1) if w in low]


def scan(label, data, words, allow, found):
    for text in readings(data):
        for n in hits(text, words, allow):
            found.add((label, n))
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for info in z.infolist():
                    if not info.is_dir():
                        scan(f"{label} > {info.filename}", z.read(info), words, allow, found)
        except zipfile.BadZipFile:
            found.add((f"{label} (looks like a zip but will not open)", 0))


def main():
    words, allow = entries("BLOCKED_WORDS"), entries("ALLOWED_PHRASES")
    if not words:
        print("BLOCKED_WORDS is empty or missing; the check cannot run")
        return 1
    root = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    found, count = set(), 0
    for p in sorted(root.rglob("*")):
        if not p.is_file() or SKIP & set(p.relative_to(root).parts):
            continue
        count += 1
        scan(p.relative_to(root).as_posix(), p.read_bytes(), words, allow, found)
    for label, n in sorted(found):
        print(f"{label}: blocked word, line {n} of the list")
    print(f"{count} files read, {len(found)} hit(s)")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
