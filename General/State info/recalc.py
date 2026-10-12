"""Opens State Info.xlsx in Excel, recalculates, strips author details,
saves. Run after build_workbook.py so the saved file carries values.

GA Sales\\recalc.ps1 does the same job, but on this workbook Excel turns
away the call right after the recalc while it's still busy, so this one
waits for the recalc to finish and retries a busy call.
"""
import os
import time

import pywintypes
import win32com.client

PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "State Info.xlsx")
BUSY = -2147418111  # RPC_E_CALL_REJECTED: Excel is busy, try again shortly


def patiently(action):
    for _ in range(60):
        try:
            return action()
        except pywintypes.com_error as e:
            if e.hresult != BUSY:
                raise
            time.sleep(0.5)
    raise RuntimeError("Excel stayed busy for 30 seconds")


xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = patiently(lambda: xl.Workbooks.Open(PATH))
    patiently(xl.CalculateFull)
    while patiently(lambda: xl.CalculationState) != 0:  # 0 means done
        time.sleep(0.2)
    patiently(lambda: setattr(wb, "RemovePersonalInformation", True))
    patiently(wb.Save)
    patiently(lambda: wb.Close(False))
    print("recalculated", PATH)
finally:
    patiently(xl.Quit)
