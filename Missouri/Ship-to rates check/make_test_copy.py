"""Builds a test copy to try before it goes out: MO Rates Check test 4.accdb
on the desktop, made from MORatesCheck.bas the way it is made at work
(SetUp), with Q3 and Q4 2026 loaded and a made-up ship-to list of 2,000 fake
ship-tos. Its state data goes in the folder beside it, "MO Rates Check test 4
data" (at work it's %LOCALAPPDATA%\\MO Rates Check). The test-only module is
removed before it's handed over. Each look-over gets the next number.

Run from this folder after test_access.py passes: python make_test_copy.py
Access must be closed.
"""
import os, threading, time
from pathlib import Path

import pythoncom
import win32com.client
import win32process

import mo_changes
import test_access as t
from mo_match import Boundary

DESKTOP = Path(os.path.expanduser("~")) / "Desktop"
NAME = "MO Rates Check test 4"


def main():
    os.chdir(t.HERE)
    t.prepare()
    db_path = DESKTOP / f"{NAME}.accdb"
    data = DESKTOP / f"{NAME} data"
    for p in (db_path, data):
        if p.exists():
            raise SystemExit(f"{p} already exists; move it to the Recycle Bin first")
    a4, _ = mo_changes.load_boundary(t.WORK / "dor-boundary-database-q4-2026.csv")
    b4 = Boundary(str(t.WORK / "dor-boundary-database-q4-2026.csv"), a4)
    rates4 = mo_changes.load_rates(t.WORK / "dor-jurisdiction-total-rate-q4-2026.xlsx")
    rates3 = mo_changes.load_rates(t.WORK / "dor-jurisdiction-total-rate-q3-2026.xlsx")
    rows, _ = t.make_ship_to_list(b4, rates3, rates4)
    xlsx = t.WORK / "Ship-to list test.xlsx"
    t.write_ship_to_list(rows, xlsx, t.WORK / "Ship-to list test.csv")

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Access.Application")
    stop, seen = threading.Event(), []
    try:
        pid = win32process.GetWindowThreadProcessId(app.hWndAccessApp())[1]
        threading.Thread(target=t.watchdog, args=(pid, stop, seen), daemon=True).start()
        app.AutomationSecurity = 1  # this copy of Access only; gone when it quits
        app.NewCurrentDatabase(str(db_path))
        app.LoadFromText(5, "MORatesCode", str(t.WORK / "MORatesCode.txt"))
        app.LoadFromText(5, "TestHooks", str(t.WORK / "TestHooks.txt"))
        print(t.run(app, "TestSetUp", str(data)))
        for q in ("q3", "q4"):
            print(t.run(app, "TestLoad", str(t.WORK / f"dor-boundary-database-{q}-2026.csv"),
                        str(t.WORK / f"dor-jurisdiction-total-rate-{q}-2026.xlsx")))
        print(t.run(app, "TestLoadList", str(xlsx)))
        print(t.run(app, "TestInfo"))
        app.DoCmd.DeleteObject(5, "TestHooks")
        print("test-only module removed; dialogs seen:", seen)
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
    time.sleep(2)
    print(f"{db_path}: {db_path.stat().st_size / 1048576:.1f} MB")
    for p in sorted(data.iterdir()):
        print(f"{p}: {p.stat().st_size / 1048576:.0f} MB")


if __name__ == "__main__":
    main()
