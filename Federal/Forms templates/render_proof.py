"""Print a built template to a PDF proof through a private, hidden Word.

    python render_proof.py "<template>.docx" ["<proof>.pdf"]

Word only ever renders here; the build scripts make the .docx (see
HANDOFF.md for why). Prints the page count and each click-and-type field
and box. Only the Word this script started is ever closed, and a watchdog
kills it if a dialog hangs it. test_exemption_template.py also uses
render() to type into a copy in memory first; that copy is never saved.
"""
import os
import subprocess
import sys
import threading
import time

import pythoncom
import win32com.client

KINDS = {1: 'field', 8: 'box'}


def word_pids():
    out = subprocess.run(
        ['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV', '/NH'],
        capture_output=True, text=True).stdout
    return {int(line.split('","')[1]) for line in out.splitlines()
            if line.startswith('"WINWORD')}


def kill(pids):
    for pid in pids:
        subprocess.run(['taskkill', '/F', '/PID', str(pid)], capture_output=True)


def render(src, pdf, fill=None, tick=(), cleared=()):
    """Print src to pdf. First types fill {field title: text}, ticks the
    boxes titled in tick, and types into then clears the fields titled in
    cleared, if given. Returns (pages, [(kind, title, state)])."""
    before = word_pids()
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    mine = word_pids() - before
    if not mine:
        # Handed a Word someone already has open; hiding or quitting it
        # would close their documents, so stop without touching it.
        raise RuntimeError('Word gave back a copy already open on screen. '
                           'Close Word and run again.')
    watchdog = threading.Timer(120, kill, [mine])
    watchdog.start()
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(src, False, True, False)  # read-only, not in recent files
        try:
            for cc in doc.ContentControls:
                if cc.Type == 8 and cc.Title in tick:
                    cc.Checked = True
                elif cc.Type == 1 and fill and fill.get(cc.Title):
                    cc.Range.Text = fill[cc.Title]
                elif cc.Type == 1 and cc.Title in cleared:
                    cc.Range.Text = 'typed, then cleared'
                    cc.Range.Text = ''
            doc.ExportAsFixedFormat(pdf, 17)  # 17 = PDF
            pages = doc.ComputeStatistics(2)
            controls = []
            for cc in doc.ContentControls:
                if cc.Type == 8:
                    state = 'checked' if cc.Checked else 'unchecked'
                elif cc.ShowingPlaceholderText:
                    state = 'hint: ' + cc.Range.Text
                else:
                    state = 'typed: ' + cc.Range.Text
                controls.append((KINDS.get(cc.Type, cc.Type), cc.Title, state))
        finally:
            doc.Close(0)
    finally:
        watchdog.cancel()
        try:
            word.Quit(0)
        except Exception:  # noqa: BLE001
            pass
        time.sleep(1)
        kill(mine & word_pids())
    return pages, controls


def main():
    src = os.path.abspath(sys.argv[1])
    pdf = (os.path.abspath(sys.argv[2]) if len(sys.argv) > 2
           else os.path.splitext(src)[0] + ' - proof.pdf')
    pages, controls = render(src, pdf)
    print('pages:', pages)
    for kind, title, state in controls:
        print(kind, '|', title, '|', state)
    print('proof:', pdf)


if __name__ == '__main__':
    main()
