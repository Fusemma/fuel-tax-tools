# Builds the Notification Certificate of Taxable Fuel Registrant template.
# Rebuild any time with:  python build_637_template.py
# Output lands beside this script. Wording mirrors the certificate as sent
# 2026-09-22. Every fill-in spot is a click-and-type field showing its own
# gray hint text until something is typed in it.
import datetime
import os

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Inches

from word_fields import field, finish

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'Notification Certificate - TEMPLATE.docx')

SIGLINE = '_' * 30

doc = Document()

# No author or program name in the file's properties
props = doc.core_properties
props.author = props.last_modified_by = props.comments = ''
props.created = props.modified = datetime.datetime.now(
    datetime.timezone.utc).replace(tzinfo=None, microsecond=0)

normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(12)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1

ph = doc.styles.add_style('Placeholder Text', WD_STYLE_TYPE.CHARACTER)
ph.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

# Title
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('NOTIFICATION CERTIFICATE OF TAXABLE FUEL REGISTRANT')
run.bold = True
run.font.size = Pt(14)
doc.add_paragraph()

# TO block
p = doc.add_paragraph('TO:\t')
field(p, 'Company name')
p = doc.add_paragraph('\t')
field(p, 'Street address')
p = doc.add_paragraph('\t')
field(p, 'City, State ZIP')
p = doc.add_paragraph('FEIN:\t')
field(p, 'Recipient FEIN')
doc.add_paragraph()
doc.add_paragraph()
doc.add_paragraph()

# Certification paragraph
p = doc.add_paragraph(
    'The undersigned taxable registrant (“Registrant”) hereby '
    'certifies under penalties of perjury that Registrant is registered by '
    'the Internal Revenue Service with registration number '
)
field(p, 'Registration number')
p.add_run(
    ' and that Registrant’s registration hasn’t been revoked or '
    'suspended by the Internal Revenue Service.'
)
doc.add_paragraph()

# Fraud paragraph
doc.add_paragraph(
    'Registrant understands that the fraudulent use of this certificate may '
    'subject Registrant and all parties making such fraudulent use of this '
    'certificate to a fine or imprisonment, or both, together with the costs '
    'of prosecution.'
)
for _ in range(5):
    doc.add_paragraph()

# Signature block: two columns, no borders. Each pair of rows is a fill-in
# line (or ink line) over its label.
rows = [
    (('field', 'Name of registrant'), ('field', 'Name of person signing')),
    (('label', 'Name of Registrant'), ('label', 'Name of Person Signing')),
    (('field', 'Registrant street address'), ('field', 'Title of person signing')),
    (('label', 'Address'), ('label', 'Title of Person Signing')),
    (('field', 'City / State / ZIP'), ('text', SIGLINE)),
    (('label', 'City/State/Zip'), ('label', 'Signature')),
    (('field', 'Registrant FEIN'), ('text', SIGLINE)),
    (('label', 'FEIN'), ('label', 'Date')),
    (('field', 'Telephone'), ('text', '')),
    (('label', 'Telephone'), ('text', '')),
]
table = doc.add_table(rows=len(rows), cols=2)
for col in table.columns:
    for cell in col.cells:
        cell.width = Inches(3.25)
for r, (left, right) in enumerate(rows):
    for c, (kind, value) in enumerate((left, right)):
        cell = table.cell(r, c)
        p = cell.paragraphs[0]
        if kind == 'field':
            field(p, value)
        elif value:
            p.add_run(value)
        if kind == 'label':
            p.paragraph_format.space_after = Pt(14)

finish(doc)
doc.save(OUT)

# Read it back as a sanity check
check = Document(OUT)
sdts = check.element.body.findall(
    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt')
print('saved:', OUT)
print('fields:', len(sdts))
