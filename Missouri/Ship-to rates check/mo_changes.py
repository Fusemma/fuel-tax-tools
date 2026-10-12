"""Step one reference: where the Missouri USE rate changed between two quarters.

Reads last quarter's and this quarter's DOR files (the boundary file, zipped
or not, and the jurisdiction total rate xlsx) and works out:

1. Codes whose use rate changed (the code is in both rate files).
2. Streets where some house numbers now sit in a different code with a
   different use rate. Compared house number by house number on each street,
   not row by row, because the state re-splits ranges between quarters (a new
   single-house row inside an old range is a move the row compare misses).

This Python copy is the test oracle; the Access copy that ships must give
the same rows (see test_access.py).

Run: python mo_changes.py LAST_BOUNDARY LAST_RATES NOW_BOUNDARY NOW_RATES [OUT_DIR]
"""
import csv, io, sys, zipfile, collections
from pathlib import Path
import openpyxl
from mo_match import tokens, canon, cut_unit, one

C_TYPE, C_BEGIN, C_END, C_LOW, C_HIGH, C_PARITY = 0, 1, 2, 3, 4, 5
C_PRE, C_NAME, C_SUF, C_POST, C_CITY, C_ZIP = 6, 7, 8, 9, 14, 15
C_JCODE = 69


def open_csv(path):
    path = str(path)
    if path.lower().endswith(".zip"):
        zf = zipfile.ZipFile(path)
        return io.TextIOWrapper(zf.open(zf.namelist()[0]), encoding="latin-1", newline="")
    return open(path, newline="", encoding="latin-1")


def full_key(pre, name, suf, post):
    """The street, standardized: 'N', 'MAIN', 'ST', '' -> 'N MAIN ST'."""
    p, s, q = one(pre), one(suf), one(post)
    nm = canon(cut_unit(tokens(name)))
    return " ".join(canon([x for x in [p] + nm + [s, q] if x]))


def city_key(city):
    """Postal city, standardized so the state's own spellings meet:
    SAINT LOUIS = ST. LOUIS, O'FALLON = O FALLON."""
    t = []
    for c in city or "":
        if c in ".'":
            continue
        if "a" <= c <= "z":
            c = chr(ord(c) - 32)
        t.append(c if ("A" <= c <= "Z" or "0" <= c <= "9") else " ")
    words = {"SAINT": "ST", "SAINTE": "STE", "MOUNT": "MT", "FORT": "FT"}
    return "".join(words.get(w, w) for w in "".join(t).split())


def load_boundary(path):
    """A rows in force on the file's own date (its newest begin date)."""
    rows = []
    with open_csv(path) as f:
        r = csv.reader(f)
        next(r)
        for row in r:
            rows.append(row)
    asof = max(row[C_BEGIN] for row in rows)
    out = []
    for row in rows:
        if row[C_TYPE] != "A" or not (row[C_BEGIN] <= asof <= row[C_END]):
            continue
        out.append(row)
    return asof, out


def rate_value(v):
    """'6.9750%' or 0.06975 -> 0.06975 (rounded to 6 places)."""
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 6)
    s = str(v).strip()
    if s.endswith("%"):
        return round(float(s[:-1]) / 100, 6)
    return round(float(s), 6)


def load_rates(path):
    wb = openpyxl.load_workbook(str(path), read_only=True)
    ws = wb.worksheets[0]
    out = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or row[8] is None:
            continue
        out[str(row[8]).strip()] = (rate_value(row[9]), rate_value(row[10]))
    return out


def side_of(par):
    p = (par or "").strip()
    return p if p in ("O", "E") else "B"


def effective(rows, a, b, par):
    """Codes in force for house numbers a..b of one parity: the narrowest
    rows covering them win (the state's carve-outs)."""
    cov = [r for r in rows if r[0] <= a and r[1] >= b and (r[2] == "B" or r[2] == par)]
    if not cov:
        return None
    w = min(r[1] - r[0] for r in cov)
    return tuple(sorted(set(r[3] for r in cov if r[1] - r[0] == w)))


def first_of(a, b, par):
    """First and last house number of parity par in a..b, or None."""
    f = a if (a % 2 == 1) == (par == "O") else a + 1
    l = b if (b % 2 == 1) == (par == "O") else b - 1
    return (f, l) if f <= l else None


def compare(last_rows, now_rows, last_rates, now_rates):
    def use_last(c):
        v = last_rates.get(c); return v[1] if v else None
    def use_now(c):
        v = now_rates.get(c); return v[1] if v else None

    # 1. codes whose use rate changed
    changes = []
    for c in sorted(set(last_rates) & set(now_rates)):
        if last_rates[c][1] != now_rates[c][1]:
            changes.append((c, last_rates[c][1], now_rates[c][1]))

    # mail cities for each changed code, from this quarter's rows
    want = set(c for c, _, _ in changes)
    cities = collections.defaultdict(lambda: collections.defaultdict(list))
    for row in now_rows:
        if row[C_JCODE] in want:
            cities[row[C_JCODE]][city_key(row[C_CITY])].append(row[C_CITY].strip())
    change_rows = []
    for c, old, new in changes:
        cs = cities.get(c, {})
        order = sorted(cs.items(), key=lambda kv: (-len(kv[1]), kv[0]))
        names = [min(v) for k, v in order]
        change_rows.append({"code": c, "county": c[6:9], "old": old, "new": new,
                            "cities": names, "rows": sum(len(v) for v in cs.values())})

    # 2. streets
    def streets(rows):
        s = collections.defaultdict(list)
        for row in rows:
            key = row[C_ZIP] + " " + full_key(row[C_PRE], row[C_NAME], row[C_SUF], row[C_POST])
            s[key].append(row)
        return s
    S_last, S_now = streets(last_rows), streets(now_rows)

    runs = []   # (street key, side, first, last, old codes, new codes)
    texts = []  # lettered house numbers, compared as written
    for key in sorted(set(S_last) & set(S_now)):
        L, N = S_last[key], S_now[key]
        def numeric(rows):
            out = []
            for row in rows:
                lo, hi = row[C_LOW].strip(), row[C_HIGH].strip()
                if lo.isdigit() and hi.isdigit():
                    a, b = int(lo), int(hi)
                    if a > b:
                        a, b = b, a
                    out.append((a, b, side_of(row[C_PARITY]), row[C_JCODE]))
            return out
        nl, nn = numeric(L), numeric(N)
        if sorted(nl) != sorted(nn):
            pts = sorted(set([r[0] for r in nl + nn] + [r[1] + 1 for r in nl + nn]))
            for par in ("O", "E"):
                run = None
                for a, b1 in zip(pts, pts[1:]):
                    b = b1 - 1
                    fl = first_of(a, b, par)
                    if not fl:
                        continue
                    e_last, e_now = effective(nl, a, b, par), effective(nn, a, b, par)
                    if e_last is None and e_now is None:
                        continue  # no such house in either quarter; a run carries over the gap
                    moved = (e_last is not None and e_now is not None and e_last != e_now and
                             set(use_last(c) for c in e_last) != set(use_now(c) for c in e_now))
                    if moved and run and run[4] == e_last and run[5] == e_now:
                        run[3] = fl[1]
                        continue
                    if run:
                        runs.append(tuple(run))
                    run = [key, par, fl[0], fl[1], e_last, e_now] if moved else None
                if run:
                    runs.append(tuple(run))
        # lettered numbers (1616B): same text range, compared as written
        def lettered(rows):
            d = collections.defaultdict(set)
            for row in rows:
                lo, hi = row[C_LOW].strip(), row[C_HIGH].strip()
                if not (lo.isdigit() and hi.isdigit()):
                    d[(lo, hi, side_of(row[C_PARITY]))].add(row[C_JCODE])
            return d
        tl, tn = lettered(L), lettered(N)
        for k in sorted(set(tl) & set(tn)):
            e_last, e_now = tuple(sorted(tl[k])), tuple(sorted(tn[k]))
            if e_last != e_now and set(use_last(c) for c in e_last) != set(use_now(c) for c in e_now):
                texts.append((key, k[2], k[0], k[1], e_last, e_now))

    # an odd run and an even run over the same houses, same codes -> both sides
    lines = []
    evens = [r for r in runs if r[1] == "E"]
    for key, par, f, l, el, en in [r for r in runs if r[1] == "O"]:
        pair = None
        for e in evens:
            if e[0] == key and e[4] == el and e[5] == en and abs(e[2] - f) <= 1 and abs(e[3] - l) <= 1:
                pair = e
                break
        if pair:
            evens.remove(pair)
            lines.append((key, "B", str(min(f, pair[2])), str(max(l, pair[3])), el, en))
        else:
            lines.append((key, "O", str(f), str(l), el, en))
    for key, par, f, l, el, en in evens:
        lines.append((key, "E", str(f), str(l), el, en))
    for key, par, lo, hi, el, en in texts:
        lines.append((key, par, lo, hi, el, en))

    # street display and mail city from this quarter's rows
    move_rows = []
    for key, par, lo, hi, el, en in lines:
        rows = S_now[key]
        names = collections.Counter(
            " ".join(x.strip() for x in (r[C_PRE], r[C_NAME], r[C_SUF], r[C_POST]) if x.strip()) for r in rows)
        street = sorted(names.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        cs = collections.Counter(r[C_CITY].strip() for r in rows)
        city = sorted(cs.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        ul = sorted(set(use_last(c) for c in el), key=lambda x: (x is None, x))
        un = sorted(set(use_now(c) for c in en), key=lambda x: (x is None, x))
        move_rows.append({"zip": key[:5], "street": street, "city": city, "side": par,
                          "lo": lo, "hi": hi, "old_codes": ", ".join(el), "new_codes": ", ".join(en),
                          "old": ul[0] if len(ul) == 1 else None, "new": un[0] if len(un) == 1 else None,
                          "county": en[0][6:9]})
    return change_rows, move_rows


def fmt(v):
    return "" if v is None else f"{v:.6f}"


def main():
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    lb, lr, nb, nr = sys.argv[1:5]
    out = Path(sys.argv[5]) if len(sys.argv) > 5 else Path(".")
    a_last, last_rows = load_boundary(lb)
    a_now, now_rows = load_boundary(nb)
    changes, moves = compare(last_rows, now_rows, load_rates(lr), load_rates(nr))
    print(f"{a_last} to {a_now}: {len(changes)} codes changed use rate, {len(moves)} street lines moved")
    with open(out / "expected_changes.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["code", "old", "new", "cities", "rows"])
        for c in changes:
            w.writerow([c["code"], fmt(c["old"]), fmt(c["new"]), "|".join(c["cities"]), c["rows"]])
    with open(out / "expected_moves.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["zip", "street", "city", "side", "lo", "hi", "old_codes", "new_codes", "old", "new"])
        for m in sorted(moves, key=lambda m: (m["zip"], m["street"], m["side"], m["lo"], m["hi"], m["old_codes"], m["new_codes"])):
            w.writerow([m["zip"], m["street"], m["city"], m["side"], m["lo"], m["hi"],
                        m["old_codes"], m["new_codes"], fmt(m["old"]), fmt(m["new"])])


if __name__ == "__main__":
    main()
