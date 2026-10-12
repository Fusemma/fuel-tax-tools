"""Prints the State page of State Info.xlsx to PDF through real Excel and
turns each printed page into a picture, for look-overs. Output lands in
preview\\.
"""
import os
import shutil
import time

import pymupdf
import pywintypes
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "preview")
COPY = os.path.join(HERE, "preview copy.xlsx")
PDF = os.path.join(OUT, "State page.pdf")
BUSY = -2147418111


def patiently(action):
    for _ in range(60):
        try:
            return action()
        except pywintypes.com_error as e:
            if e.hresult != BUSY:
                raise
            time.sleep(0.5)
    raise RuntimeError("Excel stayed busy for 30 seconds")


os.makedirs(OUT, exist_ok=True)
shutil.copyfile(os.path.join(HERE, "State Info.xlsx"), COPY)
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = patiently(lambda: xl.Workbooks.Open(COPY))
    patiently(lambda: wb.Worksheets("State").ExportAsFixedFormat(0, PDF))
    wb._FlagAsMethod("Close")  # after the export, late binding mistakes Close for a property
    patiently(lambda: wb.Close(False))
finally:
    patiently(xl.Quit)
    os.remove(COPY)

doc = pymupdf.open(PDF)
for i, pdf_page in enumerate(doc, 1):
    picture = os.path.join(OUT, f"State page {i}.png")
    pdf_page.get_pixmap(dpi=110).save(picture)
    print(picture)
