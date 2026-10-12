"""Builds MO Rates Check in real Access from MORatesCheck.bas and checks it.

1. A fresh blank database gets the module (plus test_hooks.bas) and runs SetUp.
2. The street spelling rules: every distinct street and city in both real
   quarters goes through the Access copy and must come out the same as the
   Python reference (mo_match.py / mo_changes.py).
3. Step one: Q3 2026 then Q4 2026 are loaded from the state's real files into
   their own files in a data folder (the database stays small, no compacting),
   and the use tax changes must match mo_changes.py line for line.
4. The report is printed to PDF and its first pages saved as pictures.
5. Step two: a made-up ship-to list, checked by Access and by mo_shiptos.py,
   then typed map codes, a changed address, and the 100879/4 copy.
6. A later quarter replaces the oldest file; an older one and a wrong file are
   turned away with nothing changed.

Run from this folder: python test_access.py   (about ten minutes)
Access must not be open. Nothing is installed or changed outside the temp
folder; Access's macro trust is lowered only inside the copy this script
starts, and goes away when it quits.
"""
import csv, io, os, shutil, sys, tempfile, threading, time, zipfile, collections
from pathlib import Path

import pythoncom
import win32com.client
import win32con
import win32gui
import win32process

import mo_changes
from mo_match import tokens, canon, cut_unit, one

HERE = Path(__file__).parent
STATE = HERE / "state files"
WORK = Path(tempfile.gettempdir()) / "mo_rates_test"
DATA = WORK / "data"  # the test database's data folder
fails = 0


def check(name, ok, detail=""):
    global fails
    if not ok:
        fails += 1
    print(("PASS " if ok else "FAIL ") + name + ("" if ok else "  " + str(detail)))


def crlf_copy(src, dst):
    text = Path(src).read_text(encoding="ascii")
    Path(dst).write_bytes(text.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))


def watchdog(pid, stop, seen):
    """Any message box from that Access gets its text logged and is closed, so
    a compile error can't hang the run with nobody there to click."""
    def on_window(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
                if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
                    texts = []
                    win32gui.EnumChildWindows(hwnd, lambda c, _: texts.append(win32gui.GetWindowText(c)) or True, None)
                    seen.append((win32gui.GetWindowText(hwnd), [t for t in texts if t]))
                    print("   [dialog closed]", seen[-1])
                    # a run-time error box: End stops the code (Continue would loop)
                    buttons = []
                    win32gui.EnumChildWindows(hwnd, lambda c, _: buttons.append(c) or True, None)
                    for b in buttons:
                        if win32gui.GetWindowText(b) == "&End":
                            win32gui.SendMessage(b, win32con.BM_CLICK, 0, 0)
                            return True
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, win32con.IDOK, 0)
                    time.sleep(0.3)
                    if win32gui.IsWindow(hwnd):
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
        return True
    while not stop.is_set():
        try:
            win32gui.EnumWindows(on_window, None)
        except Exception:
            pass
        stop.wait(1.0)


def prepare():
    WORK.mkdir(exist_ok=True)
    for q in ("q3", "q4"):
        csv_path = WORK / f"dor-boundary-database-{q}-2026.csv"
        if not csv_path.exists():
            with zipfile.ZipFile(STATE / f"dor-boundary-database-{q}-2026.zip") as zf:
                with zf.open(zf.namelist()[0]) as src, open(csv_path, "wb") as dst:
                    shutil.copyfileobj(src, dst)
        shutil.copy(STATE / f"dor-jurisdiction-total-rate-{q}-2026.xlsx", WORK)
    crlf_copy(HERE / "MORatesCheck.bas", WORK / "MORatesCode.txt")
    crlf_copy(HERE / "test_hooks.bas", WORK / "TestHooks.txt")


def expected_keys():
    """Every distinct street (as the state writes it) and city in both quarters."""
    streets, cities = set(), set()
    for q in ("q3", "q4"):
        with mo_changes.open_csv(WORK / f"dor-boundary-database-{q}-2026.csv") as f:
            r = csv.reader(f)
            next(r)
            for row in r:
                if row[0] == "A":
                    streets.add((row[6], row[7], row[8], row[9]))
                    cities.add(row[14])
    streets = sorted(s for s in streets if not any("|" in x for x in s))
    cities = sorted(cities)
    with open(WORK / "keys_in.txt", "w", encoding="latin-1", newline="\r\n") as f:
        for s in streets:
            f.write("|".join(s) + "\n")
    with open(WORK / "cities_in.txt", "w", encoding="latin-1", newline="\r\n") as f:
        for c in cities:
            f.write(c + "\n")
    want_keys = []
    for pre, nm, suf, post in streets:
        p, s, q = one(pre), one(suf), one(post)
        n = canon(cut_unit(tokens(nm)))
        full = mo_changes.full_key(pre, nm, suf, post)
        want_keys.append("|".join([" ".join(n), p, s, q, full]))
    want_cities = [mo_changes.city_key(c) for c in cities]
    return want_keys, want_cities


def run(app, *args):
    """Access's Run, called plainly: pywin32 312's typed call sends all 31
    arguments and Access refuses them. Keep the result, not the arguments."""
    ole = app._oleobj_
    r = ole.Invoke(ole.GetIDsOfNames("Run"), 0, pythoncom.DISPATCH_METHOD, True, *args)
    return r[0] if isinstance(r, tuple) else r


def read_rows(path):
    with open(path, newline="", encoding="latin-1") as f:
        return list(csv.DictReader(f))


# ------------------------------------------------------------------ step two

FULL_WORDS = {"N": "North", "S": "South", "E": "East", "W": "West", "ST": "Street", "AVE": "Avenue", "RD": "Road",
              "DR": "Drive", "LN": "Lane", "CT": "Court", "BLVD": "Boulevard", "PL": "Place", "TER": "Terrace",
              "CIR": "Circle", "PKWY": "Parkway", "HWY": "Highway"}


def make_ship_to_list(b4, rates3, rates4, n=2000):
    """A made-up ship-to list: fake names and account numbers on real public
    address ranges, written the messy ways people type them. Returns (rows, truth)."""
    import random
    from mo_match import COUNTIES, city_key
    rng = random.Random(20261002)
    pool = [(z, r) for z, rows in b4.by_zip.items() for r in rows if r["low"].isdigit() and r["high"].isdigit()]
    zip_cities = collections.defaultdict(set)
    for z, rows in b4.by_zip.items():
        for r in rows:
            zip_cities[z].add(r["city"].strip())
    out, truth = [], {}
    while len(out) < n:
        z, r = rng.choice(pool)
        lo, hi = sorted((int(r["low"]), int(r["high"])))
        num = rng.randint(lo, hi)
        if r["parity"] == "O" and num % 2 == 0 or r["parity"] == "E" and num % 2 == 1:
            num = num + 1 if num < hi else num - 1
            if not (lo <= num <= hi):
                continue
        pre, name, suf, post = r["raw"]
        exact = " ".join(x for x in (str(num), pre, name, suf, post) if x)
        how, code, _ = b4.match(exact, z)
        if how not in ("EXACT", "CLOSE"):
            continue
        words = [FULL_WORDS.get(w, w) if rng.random() < 0.4 else w for w in exact.split()]
        street = " ".join(words)
        street = street.title() if rng.random() < 0.6 else street
        if rng.random() < 0.15:
            street += rng.choice([" Ste 200", " Unit 4", " #12", " Bldg B"])
        city = r["city"].strip()
        if rng.random() < 0.03:
            others = sorted(c for c in zip_cities[z] if city_key(c) != city_key(city))
            if others:
                city = rng.choice(others)
        city = city.title() if rng.random() < 0.5 else city
        if city.upper().startswith("SAINT ") and rng.random() < 0.5:
            city = "St. " + city[6:]
        cname = COUNTIES.get(code[6:9], "")
        county = rng.choice([cname, cname.upper() + " COUNTY", cname + " Co", "NULL", cname])
        junk = rng.choice(["Attn Receiving", "NULL", "", "Suite 200", "Bldg 3", "C/O Plant Office"])
        if rng.random() < 0.02:
            street = rng.choice(["PO Box " + str(rng.randint(1, 999)), "Rural Route 2", "Main Office"])
        place = rng.random()
        if place < 0.6:
            a1, a2 = junk, street
        elif place < 0.9:
            a1, a2 = street, rng.choice(["NULL", "", junk])
        else:
            a1, a2 = street, street
        use4 = rates4.get(code, (None, None))[1]
        sales4 = rates4.get(code, (None, None))[0]
        use3 = rates3.get(code, (None, None))[1]
        t = rng.random()
        if t < 0.70:
            tax = use4
        elif t < 0.80:
            tax = use3
        elif t < 0.85:
            tax = sales4
        elif t < 0.90:
            tax = round(rng.choice([0.04225, 0.06975, 0.08475, 0.09113]), 6)
        elif t < 0.95:
            tax = "NULL"
        else:
            tax = f"{use4 * 100:.3f}%" if use4 is not None else "NULL"
        k = len(out) + 1
        acct, ship = str(100000 + k), str(rng.randint(1, 4))
        status_c = "Inactive" if rng.random() < 0.03 else "Active"
        status_s = "Inactive" if rng.random() < 0.03 else "Active"
        out.append({"Acct #": acct, "Ship-To #": ship, "Customer Name": f"Test Customer {k:04d}",
                    "Ship-To Name": f"Test Site {k:04d}", "Address 1": a1, "Address 2": a2, "City": city,
                    "State": "MO", "County": county, "Tax": tax,
                    "TaxExemption": "LICENSED" if rng.random() < 0.05 else "NULL",
                    "ProductTypes": rng.choice(["2,4,G", "2,G", "2,4"]), "Customer Status": status_c,
                    "Ship-To Status": status_s})
        truth[(acct, ship)] = code
    return out, truth


def write_ship_to_list(rows, xlsx_path, csv_path):
    import openpyxl
    head = list(rows[0].keys())
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(head)
    for r in rows:
        ws.append([float(r[h]) if h == "Tax" and isinstance(r[h], float) else
                   (int(r[h]) if h in ("Acct #", "Ship-To #") else r[h]) for h in head])
    wb.save(xlsx_path)
    with open(csv_path, "w", newline="", encoding="latin-1") as f:
        w = csv.writer(f)
        w.writerow(head)
        for r in rows:
            w.writerow([r[h] for h in head])


def python_results(recs, b4, b3, rates4, rates3, pins=None):
    import mo_shiptos
    out = {}
    for rec in recs:
        rec = dict(rec)
        rec["pinned"] = (pins or {}).get((rec["acct"], rec["ship_to"]), "")
        out[(rec["acct"], rec["ship_to"])] = mo_shiptos.check(rec, b4, b3, rates4, rates3, "Oct 1")
    return out


def compare_ship_tos(name, got_rows, want):
    got = {(r["acct"], r["ship_to"]): r for r in got_rows}
    check(f"{name}: same {len(want)} ship-tos", set(got) == set(want), (len(got), len(want)))
    diffs = []
    for k, w in want.items():
        g = got.get(k)
        if not g:
            continue
        exp = (w["result"], w["code_now"], w["how_now"], w["where_now"], w["line_now"], mo_changes.fmt(w["use_now"]),
               w["code_last"], mo_changes.fmt(w["use_last"]), "Yes" if w["exempt"] else "")
        act = (g["result"], g["code_now"], g["how_now"], g["where_now"], g["line_now"], g["use_now"], g["code_last"],
               g["use_last"], g["exempt"])
        if exp != act or (w["result"] != "Map" and w["why"] != g["why"]):
            diffs.append((k, act, g["why"], exp, w["why"]))
    check(f"{name}: result, codes, how sure, line used, rates and why all match", not diffs, diffs[:4])
    return got


def clipboard_text():
    """The clipboard's text, or None."""
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        return None
    finally:
        win32clipboard.CloseClipboard()


def set_clipboard_text(text):
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def make_later_quarter():
    """A made-up Q1 2027: Q4 2026's first 3,000 rows dated Jan 1, 2027, and its
    rate file under the Q1 2027 name. Only for the file swap, not the answers."""
    src = WORK / "dor-boundary-database-q4-2026.csv"
    dst = WORK / "dor-boundary-database-q1-2027.csv"
    with open(src, encoding="latin-1", newline="") as f, open(dst, "w", encoding="latin-1", newline="") as g:
        g.write(f.readline())
        for i, line in enumerate(f):
            if i >= 3000:
                break
            a = line.split(",", 3)
            g.write(",".join([a[0], "20270101", a[2], a[3]]))
    shutil.copy(WORK / "dor-jurisdiction-total-rate-q4-2026.xlsx", WORK / "dor-jurisdiction-total-rate-q1-2027.xlsx")
    return dst


def main():
    os.chdir(HERE)
    print("preparing files in", WORK)
    prepare()
    db_path = WORK / "MO Rates Check test.accdb"
    for p in (db_path, db_path.with_suffix(".laccdb")):
        if p.exists():
            p.unlink()
    if DATA.exists():
        shutil.rmtree(DATA)

    t0 = time.time()
    want_keys, want_cities = expected_keys()
    a3, rows3 = mo_changes.load_boundary(WORK / "dor-boundary-database-q3-2026.csv")
    a4, rows4 = mo_changes.load_boundary(WORK / "dor-boundary-database-q4-2026.csv")
    changes, moves = mo_changes.compare(rows3, rows4, mo_changes.load_rates(WORK / "dor-jurisdiction-total-rate-q3-2026.xlsx"),
                                        mo_changes.load_rates(WORK / "dor-jurisdiction-total-rate-q4-2026.xlsx"))
    print(f"python reference ready in {time.time() - t0:.0f}s: {len(changes)} changes, {len(moves)} moves, "
          f"{len(want_keys)} streets, {len(want_cities)} cities")

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Access.Application")
    stop, seen = threading.Event(), []
    try:
        pid = win32process.GetWindowThreadProcessId(app.hWndAccessApp())[1]
        threading.Thread(target=watchdog, args=(pid, stop, seen), daemon=True).start()
        app.AutomationSecurity = 1  # this copy of Access only; gone when it quits
        app.NewCurrentDatabase(str(db_path))
        app.LoadFromText(5, "MORatesCode", str(WORK / "MORatesCode.txt"))
        app.LoadFromText(5, "TestHooks", str(WORK / "TestHooks.txt"))
        print("module loaded")

        msg = run(app, "TestSetUp", str(DATA))
        check("SetUp runs", msg.startswith("Set up is done"), msg)
        check("compact on close is off", run(app, "TestOption", "Auto Compact") in ("0", "False"))

        # 2. spelling rules on every real street and city
        n = run(app, "TestKeys", str(WORK / "keys_in.txt"), str(WORK / "keys_out.txt"))
        got = (WORK / "keys_out.txt").read_text(encoding="latin-1").splitlines()
        bad = [(i, want_keys[i], got[i]) for i in range(min(len(got), len(want_keys))) if got[i] != want_keys[i]]
        check(f"street pieces match the Python rules on all {len(want_keys)} real streets",
              len(got) == len(want_keys) and not bad, bad[:5])
        run(app, "TestCityKeys", str(WORK / "cities_in.txt"), str(WORK / "cities_out.txt"))
        got = (WORK / "cities_out.txt").read_text(encoding="latin-1").splitlines()
        bad = [(want_cities[i], got[i]) for i in range(min(len(got), len(want_cities))) if got[i] != want_cities[i]]
        check(f"city keys match on all {len(want_cities)} real cities", len(got) == len(want_cities) and not bad, bad[:5])

        # 3. step one on the real files
        for q in ("q3", "q4"):
            msg = run(app, "TestLoad", str(WORK / f"dor-boundary-database-{q}-2026.csv"),
                          str(WORK / f"dor-jurisdiction-total-rate-{q}-2026.xlsx"))
            print("  load", q, "->", msg)
            check(f"{q} loads", msg.startswith("OK"), msg)
        print("  ", run(app, "TestInfo"))
        print("  ", run(app, "TestExport", str(WORK)))
        files = sorted(p.name for p in DATA.iterdir())
        check("each quarter is a file of its own in the data folder",
              files == ["Quarter 2026-07-01.accdb", "Quarter 2026-10-01.accdb"], files)
        mb = db_path.stat().st_size / 1048576
        check(f"the database itself stays small ({mb:.1f} MB)", mb < 10, mb)

        got = read_rows(WORK / "actual_changes.csv")
        want = {c["code"]: c for c in changes}
        check(f"same {len(changes)} codes with a new use rate", sorted(r["code"] for r in got) == sorted(want),
              (len(got), len(want)))
        diffs = []
        for r in got:
            c = want.get(r["code"])
            if not c:
                continue
            exp = (mo_changes.fmt(c["old"]), mo_changes.fmt(c["new"]), "|".join(c["cities"]), str(c["rows"]))
            if (r["old"], r["new"], r["cities"], r["rows"]) != exp:
                diffs.append((r["code"], (r["old"], r["new"], r["cities"][:60], r["rows"]), exp[:2] + (exp[2][:60], exp[3])))
        check("old rate, new rate, mail cities and address count all match", not diffs, diffs[:5])

        def mkey(m):
            return (m["zip"], m["street"], m["city"], m["side"], m["lo"], m["hi"], m["old_codes"], m["new_codes"],
                    m["old"], m["new"])
        got_m = collections.Counter(mkey(r) for r in read_rows(WORK / "actual_moves.csv"))
        want_m = collections.Counter((m["zip"], m["street"], m["city"], m["side"], m["lo"], m["hi"], m["old_codes"],
                                      m["new_codes"], mo_changes.fmt(m["old"]), mo_changes.fmt(m["new"])) for m in moves)
        only_got = list((got_m - want_m).elements())
        only_want = list((want_m - got_m).elements())
        check(f"same {len(moves)} street lines (house numbers, side, codes, rates)", not only_got and not only_want,
              f"only Access: {only_got[:4]}  only Python: {only_want[:4]}")

        # 4. the report
        pdf = WORK / "Use tax changes.pdf"
        if pdf.exists():
            pdf.unlink()
        run(app, "TestReportPdf", str(pdf))
        check("report prints to PDF", pdf.exists())
        if pdf.exists():
            import pymupdf as fitz
            doc = fitz.open(str(pdf))
            print(f"   report: {doc.page_count} pages")
            for i in range(min(3, doc.page_count)):
                doc[i].get_pixmap(dpi=70).save(str(WORK / f"report_page{i + 1}.png"))

        # 5. step two: a made-up ship-to list, checked by Access and by mo_shiptos.py
        import mo_shiptos
        from mo_match import Boundary
        t0 = time.time()
        b4 = Boundary(str(WORK / "dor-boundary-database-q4-2026.csv"), a4)
        b3 = Boundary(str(WORK / "dor-boundary-database-q3-2026.csv"), a3)
        rates4 = mo_changes.load_rates(WORK / "dor-jurisdiction-total-rate-q4-2026.xlsx")
        rates3 = mo_changes.load_rates(WORK / "dor-jurisdiction-total-rate-q3-2026.xlsx")
        rows, truth = make_ship_to_list(b4, rates3, rates4)
        xlsx, csv_path = WORK / "Ship-to list test.xlsx", WORK / "Ship-to list test.csv"
        write_ship_to_list(rows, xlsx, csv_path)
        recs = mo_shiptos.read_list(xlsx)
        want = python_results(recs, b4, b3, rates4, rates3)
        print(f"   python step two ready in {time.time() - t0:.0f}s: " +
              str(collections.Counter(w["result"] for w in want.values())))
        check("the CSV copy of the list reads the same as the xlsx", mo_shiptos.read_list(csv_path) == recs)

        msg = run(app, "TestLoadList", str(xlsx))
        print("   ", msg)
        run(app, "TestExportShipTos", str(WORK / "actual_shiptos.csv"))
        got = compare_ship_tos("ship-to list (xlsx)", read_rows(WORK / "actual_shiptos.csv"), want)
        placed_right = sum(1 for k, w in want.items() if w["code_now"] and w["how_now"] != "PINNED" and
                           w["code_now"] == truth[k])
        placed_wrong = [(k, w["code_now"], truth[k]) for k, w in want.items() if w["code_now"] and
                        w["code_now"] != truth[k] and w["result"] != "Map"]
        print(f"   placed on the right code: {placed_right} of {len(want)}; wrong code and not sent to the map: "
              f"{len(placed_wrong)} {placed_wrong[:3]}")

        msg = run(app, "TestLoadList", str(csv_path))
        print("    csv again:", msg)
        check("loading the same list as CSV changes nothing", "0 new ship-tos, 0 with a changed address" in msg, msg)
        run(app, "TestExportShipTos", str(WORK / "actual_shiptos2.csv"))
        compare_ship_tos("same list from CSV", read_rows(WORK / "actual_shiptos2.csv"), want)

        # typed map codes: the true code for up to 20 the matcher couldn't place,
        # and one wrong code on a placed ship-to (must come back for a look)
        pins = {}
        for k, w in want.items():
            if w["result"] == "Map" and not w["code_now"] and len(pins) < 20:
                pins[k] = truth[k]
        placed_ok = [k for k, w in want.items() if w["result"] == "OK" and w["how_now"] == "EXACT"]
        if placed_ok:
            pins[placed_ok[0]] = "37000-051-000"
        for (acct, ship), code in pins.items():
            run(app, "TestPin", acct, ship, code)
        print("    pinned:", run(app, "TestCheck"))
        want2 = python_results(recs, b4, b3, rates4, rates3, pins)
        run(app, "TestExportShipTos", str(WORK / "actual_shiptos3.csv"))
        got3 = compare_ship_tos("after typing map codes", read_rows(WORK / "actual_shiptos3.csv"), want2)
        if placed_ok:
            g = got3[placed_ok[0]]
            check("a typed code the state's file disagrees with goes back to the map",
                  g["result"] == "Map" and "you typed" in g["why"], (g["result"], g["why"]))

        # a later list with a changed address clears that ship-to's typed code
        k = next(iter(pins))
        moved = [dict(r) for r in rows if (r["Acct #"], r["Ship-To #"]) == k]
        moved[0]["Address 2"] = "999 Nowhere Lane"
        write_ship_to_list(moved, WORK / "Ship-to list moved.xlsx", WORK / "Ship-to list moved.csv")
        msg = run(app, "TestLoadList", str(WORK / "Ship-to list moved.csv"))
        print("    moved:", msg)
        run(app, "TestExportShipTos", str(WORK / "actual_shiptos4.csv"))
        g = {(r["acct"], r["ship_to"]): r for r in read_rows(WORK / "actual_shiptos4.csv")}[k]
        check("a changed address clears its typed map code", g["pinned"] == "" and "1 with a changed address" in msg,
              (g["pinned"], msg))

        for frm in ("MO Rates Check", "Ship-tos to fix", "Look up on the map"):
            pdf_name = frm + ".pdf"
            pdf = WORK / pdf_name
            if pdf.exists():
                pdf.unlink()
            run(app, "TestFormPdf", str(pdf), frm)
            check(f"{pdf_name} prints", pdf.exists())
            if pdf.exists():
                import pymupdf as fitz
                doc = fitz.open(str(pdf))
                print(f"   {pdf_name}: {doc.page_count} pages")
                for i in range(min(2, doc.page_count)):
                    doc[i].get_pixmap(dpi=70).save(str(WORK / f"{pdf.stem}_{i + 1}.png"))

        # one click on a 100879/4 box puts exactly that on the clipboard
        acct, ship = rows[0]["Acct #"], rows[0]["Ship-To #"]
        before = clipboard_text()
        msg = run(app, "TestCopy", acct, ship)
        got = clipboard_text()
        check(f"copying {acct}/{ship} puts exactly that on the clipboard", msg == "copied" and got == f"{acct}/{ship}",
              (msg, got))
        if before is not None:
            set_clipboard_text(before)

        # 6. a later quarter replaces the oldest file; nothing else changes
        later = make_later_quarter()
        msg = run(app, "TestLoad", str(later), str(WORK / "dor-jurisdiction-total-rate-q1-2027.xlsx"))
        print("   later quarter:", msg)
        files = sorted(p.name for p in DATA.iterdir())
        check("a later quarter's file replaces the oldest one", msg.startswith("OK") and
              files == ["Quarter 2026-10-01.accdb", "Quarter 2027-01-01.accdb"], (msg, files))
        info = run(app, "TestInfo")
        print("  ", info)
        check("the ship-tos were checked again against it", "checked against Q1 2027" in info, info)
        msg = run(app, "TestLoad", str(WORK / "dor-boundary-database-q3-2026.csv"),
                  str(WORK / "dor-jurisdiction-total-rate-q3-2026.xlsx"))
        check("a quarter older than both on hand is turned away", msg.startswith("FAILED") and
              "older than both" in msg and sorted(p.name for p in DATA.iterdir()) == files, msg)
        wrong = WORK / "not a boundary file.csv"
        wrong.write_text("Acct #,Ship-To #,City\n1,2,X\n")
        msg = run(app, "TestLoad", str(wrong), str(WORK / "dor-jurisdiction-total-rate-q4-2026.xlsx"))
        check("a wrong file is turned away with nothing changed", msg.startswith("FAILED") and
              sorted(p.name for p in DATA.iterdir()) == files, msg)
        mb = db_path.stat().st_size / 1048576
        check(f"the database is still small after all that ({mb:.1f} MB)", mb < 10, mb)
        check("no message boxes popped up", not seen, seen)
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit(2)
        except Exception:
            pass
        stop.set()
        del app
        pythoncom.CoUninitialize()
    print(f"\n{fails} failures")


if __name__ == "__main__":
    main()
