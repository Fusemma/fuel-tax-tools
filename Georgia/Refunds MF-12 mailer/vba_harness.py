"""Run a VBA module in real Excel on this PC, to test it before it ships.

Excel here won't let a program paste code into a workbook: "Trust access to
the VBA project object model" is off, and it stays off (security setting).
So this builds the macro workbook by hand instead. It copies a known-good
.xlsm, swaps the source text of its Module1 for the module under test, and
throws away the saved compiled copy, so Excel compiles the new source when
the file opens. The file format is Microsoft's published MS-OVBA spec:
a compressed "dir" stream lists the modules and where each one's source
starts; each module stream holds compressed source text.

Works for any macro, not only this build. Base workbook used so far:
Work Related\\Excel Related\\GL Sort Test.xlsm (read only, never changed).

    from vba_harness import build_xlsm
    build_xlsm(base_xlsm, {"Module1": source_text}, out_xlsm)
"""
import os
import re
import shutil
import struct
import tempfile
import zipfile

import pythoncom

STGM_READ, STGM_READWRITE = 0x0, 0x2
STGM_SHARE_EXCLUSIVE, STGM_CREATE = 0x10, 0x1000


# ---- MS-OVBA 2.4.1 compression ------------------------------------------

def _bit_count(diff):
    return max((diff - 1).bit_length(), 4)


def decompress(data):
    assert data[0] == 1, "not a compressed container"
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        flag = header >> 15
        chunk_end = min(pos + size, len(data))
        pos += 2
        start = len(out)
        if not flag:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    bc = _bit_count(len(out) - start)
                    length = (token & (0xFFFF >> bc)) + 3
                    offset = (token >> (16 - bc)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
    return bytes(out)


def _compress_chunk(chunk):
    out = bytearray()
    heads = {}
    pos = 0
    n = len(chunk)
    while pos < n:
        flag_at = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos > 0:
                bc = _bit_count(pos)
                max_len = (0xFFFF >> bc) + 3
                max_off = 1 << bc
                for cand in reversed(heads.get(chunk[pos:pos + 3], [])):
                    off = pos - cand
                    if off > max_off:
                        break
                    ln = 0
                    while pos + ln < n and ln < max_len and chunk[cand + ln] == chunk[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, off
                        if ln == max_len:
                            break
            if best_len >= 3:
                bc = _bit_count(pos)
                token = ((best_off - 1) << (16 - bc)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                step = best_len
            else:
                out.append(chunk[pos])
                step = 1
            for i in range(pos, min(pos + step, n - 2)):
                heads.setdefault(chunk[i:i + 3], []).append(i)
            pos += step
        out[flag_at] = flags
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for i in range(0, max(len(data), 1), 4096):
        chunk = data[i:i + 4096]
        body = _compress_chunk(chunk)
        if len(body) + 2 > 4098:
            raise ValueError("chunk would not compress; real VBA text always does")
        out += struct.pack("<H", 0xB000 | (len(body) + 2 - 3))
        out += body
    return bytes(out)


# ---- the dir stream -------------------------------------------------------

def parse_dir(raw):
    recs, pos = [], 0
    while pos < len(raw):
        rid, size = struct.unpack_from("<HI", raw, pos)
        pos += 6
        if rid == 0x0009:          # PROJECTVERSION: "size" is a fixed 4, data is 6
            size = 6
        recs.append([rid, raw[pos:pos + size]])
        pos += size
    return recs


def write_dir(recs):
    out = bytearray()
    for rid, data in recs:
        size = 4 if rid == 0x0009 else len(data)
        out += struct.pack("<HI", rid, size) + data
    return bytes(out)


# ---- the compound file ------------------------------------------------------

def _read_tree(stg):
    tree = {}
    for st in stg.EnumElements():
        name, typ = st[0], st[1]
        if typ == 1:
            tree[name] = _read_tree(stg.OpenStorage(name, None, STGM_READ | STGM_SHARE_EXCLUSIVE, None, 0))
        else:
            s = stg.OpenStream(name, None, STGM_READ | STGM_SHARE_EXCLUSIVE, 0)
            tree[name] = s.Read(st[2])
    return tree


def _write_tree(stg, tree):
    mode = STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE
    for name, val in tree.items():
        if isinstance(val, dict):
            sub = stg.CreateStorage(name, mode, 0, 0)
            _write_tree(sub, val)
            sub.Commit(0)
        else:
            s = stg.CreateStream(name, mode, 0, 0)
            s.Write(val)


def rebuild_project(bin_bytes, new_sources):
    """Return a vbaProject.bin with the given modules' source replaced and
    every module stored as source only (Excel recompiles on open)."""
    tmp = tempfile.mkdtemp()
    try:
        src = os.path.join(tmp, "in.bin")
        with open(src, "wb") as fh:
            fh.write(bin_bytes)
        tree = _read_tree(pythoncom.StgOpenStorage(src, None, STGM_READ | STGM_SHARE_EXCLUSIVE, None, 0))
        vba = tree["VBA"]
        recs = parse_dir(decompress(vba["dir"]))
        codepage = next(struct.unpack("<H", d)[0] for r, d in recs if r == 0x0003)
        # walk the module records: name, stream name, offset
        modules, cur = [], None
        for rec in recs:
            rid, data = rec
            if rid == 0x0019:
                cur = {"name": data.decode("cp%d" % codepage)}
                modules.append(cur)
            elif rid == 0x001A and cur is not None:
                cur["stream"] = data.decode("cp%d" % codepage)
            elif rid == 0x0031 and cur is not None:
                cur["offset_rec"] = rec
        for m in modules:
            off = struct.unpack("<I", m["offset_rec"][1])[0]
            text = decompress(vba[m["stream"]][off:])
            if m["name"] in new_sources:
                body = new_sources[m["name"]].replace("\r\n", "\n").replace("\n", "\r\n")
                if not body.startswith("Attribute VB_Name"):
                    body = 'Attribute VB_Name = "%s"\r\n' % m["name"] + body
                text = body.encode("cp%d" % codepage)
            vba[m["stream"]] = compress(text)
            m["offset_rec"][1] = struct.pack("<I", 0)
        missing = set(new_sources) - {m["name"] for m in modules}
        assert not missing, "base workbook has no module named %s" % missing
        vba["dir"] = compress(write_dir(recs))
        vba["_VBA_PROJECT"] = b"\xCC\x61\xFF\xFF\x00\x00\x00"   # no compiled cache
        for k in [k for k in vba if k.startswith("__SRP_")]:
            del vba[k]
        dst = os.path.join(tmp, "out.bin")
        stg = pythoncom.StgCreateDocfile(dst, STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0)
        _write_tree(stg, tree)
        stg.Commit(0)
        del stg
        with open(dst, "rb") as fh:
            return fh.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


_DUMMY = {"string": "zS", "long": "zL", "double": "zD", "boolean": "zB", "date": "zDt", "variant": "zV",
          "object": "zO", "worksheet": "zWs", "workbook": "zWb", "range": "zR", "vbmsgboxstyle": "zMsg"}


def compile_hook(source, safe_dir, first=(), setup="", overrides=None, skip=()):
    """Return the source plus a Public Function zzCompileAll that steps into
    every procedure once. Excel compiles a procedure the first time it runs,
    so a typo in a part a test never reaches would otherwise ship. Each call
    runs under On Error Resume Next with blank arguments; run it only in a
    throwaway workbook. `first` names run before `setup` (VBA lines that
    point module state at safe_dir); `overrides` gives argument text by name."""
    overrides = overrides or {}
    text = source.replace(" _\r\n", " ").replace(" _\n", " ")
    procs = re.findall(r"^(?:Public|Private)\s+(Sub|Function)\s+(\w+)\s*\(([^)]*)\)", text, re.M)
    calls_first, calls = [], []
    for kind, name, params in procs:
        if name in skip:
            continue
        if name in overrides:
            args = overrides[name]
        else:
            args = []
            for prm in [x.strip() for x in params.split(",") if x.strip()]:
                m = re.search(r"\bAs\s+(\w+)", prm)
                typ = (m.group(1) if m else "Variant").lower()
                args.append(_DUMMY.get(typ, "zV"))
            args = ", ".join(args)
        line = "    Call %s(%s): Err.Clear" % (name, args)
        (calls_first if name in first else calls).append((first.index(name) if name in first else 0, line))
    calls_first.sort()
    setup_lines = ["    " + l.strip() for l in setup.splitlines() if l.strip()]
    body = ["", "Public Function zzCompileAll() As String",
            "    Dim zS As String, zL As Long, zD As Double, zB As Boolean, zDt As Date, zV As Variant",
            "    Dim zO As Object, zWs As Worksheet, zWb As Workbook, zR As Range, zMsg As VbMsgBoxStyle",
            "    Dim zSafe As String",
            "    zSafe = \"%s\"" % safe_dir.replace('"', '""'),
            "    On Error Resume Next"]
    body += setup_lines
    body += [l for _, l in calls_first]
    body += setup_lines
    body += [l for _, l in calls]
    body += ['    zzCompileAll = "stepped into %d procedures"' % (len(calls) + len(calls_first)), "End Function", ""]
    return source.rstrip() + "\r\n" + "\r\n".join(body)


def build_xlsm(base_xlsm, new_sources, out_xlsm):
    with zipfile.ZipFile(base_xlsm) as zin:
        new_bin = rebuild_project(zin.read("xl/vbaProject.bin"), new_sources)
        with zipfile.ZipFile(out_xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_bin if item.filename == "xl/vbaProject.bin" else zin.read(item.filename)
                zout.writestr(item, data)
    return out_xlsm


if __name__ == "__main__":
    import sys
    # self-test: round-trip every module of the base workbook
    base = sys.argv[1]
    with zipfile.ZipFile(base) as z:
        raw = z.read("xl/vbaProject.bin")
    tmpd = tempfile.mkdtemp()
    p = os.path.join(tmpd, "b.bin")
    open(p, "wb").write(raw)
    tree = _read_tree(pythoncom.StgOpenStorage(p, None, STGM_READ | STGM_SHARE_EXCLUSIVE, None, 0))
    d = decompress(tree["VBA"]["dir"])
    assert decompress(compress(d)) == d
    assert write_dir(parse_dir(d)) == d
    print("dir round-trips,", len(d), "bytes")
