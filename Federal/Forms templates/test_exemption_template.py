"""Checks the exemption certificate template the way it was proofed before shipping.

    python build_exemption_template.py
    python test_exemption_template.py

Prints the template through Word (render_proof.py) and checks the proof:
one page; the wording word for word against "Combined certificate text
(Pub 510 Dec 2025).txt"; every line of that file not marked * found word
for word in IRS Model Certificate M or P (p510.pdf pages 53 and 56);
hints gray and everything else black; room left under the box. Then types
made-up details into every field, ticks two boxes, and checks that copy
still prints on one page. Nothing is saved; the proof PDFs are deleted at
the end.
"""
import os

import pymupdf

from render_proof import render

HERE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(HERE, 'Exemption Certificate - TEMPLATE.docx')
TEXT = os.path.join(HERE, 'Combined certificate text (Pub 510 Dec 2025).txt')
IRS_PDF = os.path.join(HERE, 'p510.pdf')
MODEL_PAGES = {'M': 53, 'P': 56}
JOINED_LINES = 7   # lines marked * in the text file

HINTS = ['Vendor name', 'Vendor address', 'Vendor FEIN', 'Invoice no.',
         'Gallons', 'MM/DD/YYYY', 'MM/DD/YYYY', 'Account or order no.',
         'Name of person signing', 'Title of person signing', 'Buyer name',
         'Buyer FEIN', 'Buyer address']
BOXES = ['State or local government', 'Nonprofit educational organization',
         'Single purchase', 'All purchases on an account']
# Made-up details, long ones on purpose, to prove a filled copy still fits
SAMPLE = {
    'Vendor name': 'Sample Fuel Company, LLC',
    'Vendor address': '1000 Example Parkway, Suite 200, Anytown, IN 46000',
    'Vendor FEIN': '12-3456789',
    'Effective date': '10/01/2026',
    'Expiration date': '09/30/2027',
    'Buyer’s account or order number': 'A-100572',
    'Name of person signing': 'Pat Example',
    'Title of person signing': 'Director of Transportation and Facilities',
    'Buyer name': 'Anytown County Highway Department',
    'Buyer FEIN': '98-7654321',
    'Buyer address': '2200 North County Road 500 West, Building C, Anytown, IN 46000',
}
TICK = ('State or local government', 'All purchases on an account')
EMPTY_BOX, TICKED_BOX = '☐', '☒'

failures = []
checks = [0]


def check(ok, what):
    checks[0] += 1
    print('ok  ' if ok else 'FAIL', what)
    if not ok:
        failures.append(what)


def words(text):
    """Text as a list of words, quote marks and the section sign made plain."""
    text = (text.replace('’', "'").replace('“', '"').replace('”', '"')
            .replace('§', 'section'))
    return text.split()


def text_lines():
    """The form's lines from the text file, in order."""
    with open(TEXT, encoding='utf-8') as f:
        body = f.read().split('=== Certificate text ===')[1]
    return [line.strip() for line in body.splitlines() if line.strip()]


def plain(line):
    """A line without its * mark, blanks and box marks."""
    if line.startswith('* '):
        line = line[2:]
    return line.replace('________', ' ').replace('[ ]', ' ')


def read_pdf(path):
    """First page's text, its text pieces, and the lowest line drawn."""
    with pymupdf.open(path) as pdf:
        page = pdf[0]
        pieces = [s for block in page.get_text('dict')['blocks']
                  for line in block.get('lines', []) for s in line['spans']
                  if s['text'].strip()]
        bottom = max(d['rect'].y1 for d in page.get_drawings())
        return page.get_text(), pieces, bottom, page.rect.height


def main():
    blank_pdf = os.path.join(HERE, 'test proof blank.pdf')
    filled_pdf = os.path.join(HERE, 'test proof filled.pdf')
    try:
        lines = text_lines()
        with pymupdf.open(IRS_PDF) as pdf:
            models = [' '.join(words(pdf[n - 1].get_text()))
                      for n in MODEL_PAGES.values()]
        joined = [l for l in lines if l.startswith('* ')]
        untraced = [l for l in lines if not l.startswith('* ')
                    and words(plain(l))
                    and not any(' '.join(words(plain(l))) in m for m in models)]
        check(not untraced, 'every unmarked line is word for word from Model M or P'
              + ''.join('\n       not found: ' + l for l in untraced))
        check(len(joined) == JOINED_LINES,
              '%d lines join M and P wording (marked * in the text file)' % len(joined))

        pages, controls = render(DOCX, blank_pdf)
        check(pages == 1, 'blank prints on one page (Word says %d)' % pages)
        check([t for k, t, s in controls if k == 'box'] == BOXES,
              'four boxes, in order')
        check(all(s == 'unchecked' for k, t, s in controls if k == 'box'),
              'boxes start empty')
        hints = [s[len('hint: '):] for k, t, s in controls
                 if k == 'field' and s.startswith('hint: ')]
        check(hints == HINTS, 'all 13 fields show their hint')

        text, pieces, bottom, height = read_pdf(blank_pdf)
        text = text.replace(EMPTY_BOX, ' ')
        for hint in HINTS:
            text = text.replace(hint, ' ', 1)
        check(words(text) == words(' '.join(plain(l) for l in lines)),
              'wording matches the text file word for word')

        gray = sorted(s['text'].strip() for s in pieces if s['color'] == 0x808080)
        check(gray == sorted(HINTS), 'hints print gray')
        check(all(s['color'] in (0, 0x808080) for s in pieces),
              'everything else prints black')
        room = height - 36 - bottom   # 36 pt is the half-inch bottom margin
        check(room > 18, 'room under the box: %.0f pt' % room)

        pages, controls = render(DOCX, filled_pdf, SAMPLE, TICK,
                                 cleared=('Number of gallons',))
        check(pages == 1, 'filled copy prints on one page (Word says %d)' % pages)
        typed = {t: s[len('typed: '):] for k, t, s in controls
                 if s.startswith('typed: ')}
        check(typed == SAMPLE, 'every typed detail landed in its field')
        state = dict((t, s) for k, t, s in controls)['Number of gallons']
        check(state == 'hint: Gallons',
              'a field typed in and then cleared shows its hint again (%s)' % state)
        ticked = tuple(t for k, t, s in controls if s == 'checked')
        check(ticked == TICK, 'the two clicked boxes are checked')
        text, pieces, bottom, height = read_pdf(filled_pdf)
        sizes = {round(s['size'], 1) for s in pieces
                 if s['text'].strip() in (EMPTY_BOX, TICKED_BOX)}
        check(len(sizes) == 1, 'ticked and empty boxes print the same size')
    finally:
        for pdf in (blank_pdf, filled_pdf):
            if os.path.exists(pdf):
                os.remove(pdf)
    print('%d checks, %d failures' % (checks[0], len(failures)))


if __name__ == '__main__':
    main()
