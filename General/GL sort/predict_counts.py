"""Python copy of SortGL's grouping and ORD pairing, to predict the popup numbers."""
import re, sys, openpyxl

VENDOR_LETTERS = 4
BOL_DIGITS = 0
SURE, EXC, MAYBE = 0.4, 0.5, 0.6


def squash(s):
    return re.sub(r"[^a-z0-9]", "", str(s if s is not None else "").lower())


def bolkey(s):
    s = str(s if s is not None else "").upper()
    p = s.find("BL")
    q = -1
    while p >= 0:
        q = p + 2
        while q < len(s) and s[q] in " -:#":
            q += 1
        if q < len(s) and s[q] in "0123456789":
            break
        p = s.find("BL", p + 1)
    if p < 0:
        return ""
    d = ""
    while q < len(s) and s[q] in "0123456789":
        d += s[q]
        q += 1
    if BOL_DIGITS:
        d = d[:BOL_DIGITS]
    return "BL" + d


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def run(rows):
    n = len(rows)
    rank, pair = [0.0] * n, [0] * n
    counts = [0, 0, 0]
    for i, r in enumerate(rows):
        g, h = str(r[6] or "").upper(), str(r[7] or "").upper()
        if "MISC" in h: rank[i] = 1
        elif "ORD" in g: rank[i] = 2
        elif "BANK" in h: rank[i] = 3
        if rank[i]: counts[int(rank[i]) - 1] += 1
    net = lambda i: num(rows[i][2]) - num(rows[i][3])
    key = lambda i, a: squash(rows[i][5])[:VENDOR_LETTERS] + "~" + f"{a:.2f}"

    ords = {}
    for i in range(n):
        if rank[i] == 2 and net(i) < 0:
            ords.setdefault(key(i, -net(i)), []).append(i)

    seq = sure = amb = paired = 0
    for i in range(n):
        if rank[i] != 0 or net(i) <= 0:
            continue
        k = key(i, net(i))
        cands = ords.get(k)
        if not cands:
            continue
        bol, hits, pick = bolkey(rows[i][7]), 0, None
        if bol:
            for c in cands:
                if bolkey(rows[c][7]) == bol:
                    hits += 1
                    pick = c
            if hits != 1:
                pick = None
        if pick is None and len(cands) == 1 and hits == 0:
            if not bol or not bolkey(rows[cands[0]][7]):
                pick = cands[0]
        if pick is not None:
            seq += 1
            paired += 1
            if hits == 1:
                sure += 1
                rank[i], pair[i] = SURE, seq
            else:
                rank[i], pair[i] = EXC, n + seq
            pair[pick] = pair[i]
            rest = [c for c in cands if c != pick]
            if rest: ords[k] = rest
            else: del ords[k]
        else:
            rank[i] = MAYBE
            amb += 1

    order = sorted(range(n), key=lambda i: (rank[i], pair[i], i))
    top = sum(1 for x in rank if x < 1)
    top_sum_d = sum(num(rows[i][2]) for i in range(n) if rank[i] < 1)
    top_sum_c = sum(num(rows[i][3]) for i in range(n) if rank[i] < 1)
    first = 5
    left = top - paired - amb
    res = {
        "left on top": left, "MISC": counts[0], "ORD": counts[1], "BANK": counts[2],
        "green": sure, "yellow": paired - sure, "orange": amb, "exceptions total": paired,
        "green rows": (first + left, first + left + sure - 1),
        "yellow rows": (first + left + sure, first + left + paired - 1),
        "orange rows": (first + left + paired, first + top - 1),
        "sum row": first + top, "net row": first + top + 2,
        "AP net": round(top_sum_d - top_sum_c, 2),
        "moved start": first + top + 10,
        "ORD still open": counts[1] - paired,
    }
    ord_end = first + top + 10 + counts[0] + counts[1] - 1
    res["blue rows"] = (ord_end - paired + 1, ord_end)
    return res, rank, pair, order


if __name__ == "__main__":
    wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
    rows = [r for r in wb.active.iter_rows(min_row=5, values_only=True)]
    res = run(rows)[0]
    for k, v in res.items():
        print(f"{k:18} {v:,}" if isinstance(v, (int, float)) else f"{k:18} {v[0]:,} to {v[1]:,}")
